import Foundation
import Combine

struct ACPMessage: Identifiable {
    let id: String
    let role: String
    var content: String
}

@MainActor
final class ACPClient: ObservableObject {
    @Published private(set) var isConnected = false
    @Published private(set) var isConnecting = false
    @Published private(set) var isBusy = false
    @Published private(set) var messages: [ACPMessage] = []
    @Published private(set) var currentSessionId: String?
    @Published private(set) var availableModels: [String] = []
    @Published private(set) var selectedModel = "qwen3.5:9b"
    @Published private(set) var activity = ""
    @Published private(set) var errorMessage: String?

    private var process: Process?
    private var stdinPipe: Pipe?
    private var stdoutPipe: Pipe?
    private var stderrPipe: Pipe?
    private var lines = ACPLineBuffer()
    private var stderr = Data()
    private var requestId = 0
    private var pending: [Int: CheckedContinuation<JSONValue, Error>] = [:]
    private var assistantId: String?
    private var harnessDirectory: URL?
    private let defaults: UserDefaults
    private let storagePath: URL?
    private var sessionKey = ""
    private var intentionalDisconnect = false

    init(defaults: UserDefaults = .standard, storagePath: URL? = nil) {
        self.defaults = defaults
        self.storagePath = storagePath
    }

    func connect(harnessDirectory explicitDirectory: URL? = nil) async {
        guard process == nil else { return }
        isConnecting = true
        errorMessage = nil
        intentionalDisconnect = false
        messages = []
        currentSessionId = nil
        assistantId = nil
        lines = ACPLineBuffer()
        stderr = Data()
        let environment = ProcessInfo.processInfo.environment
        let sourceRoot = URL(fileURLWithPath: #filePath)
            .deletingLastPathComponent().deletingLastPathComponent()
            .deletingLastPathComponent().deletingLastPathComponent().deletingLastPathComponent()
        let directory = explicitDirectory
            ?? environment["MOFF_HARNESS_DIR"].map { URL(fileURLWithPath: $0) }
            ?? sourceRoot.appendingPathComponent("harness")
        harnessDirectory = directory
        sessionKey = "moff.session.\(storagePath?.path ?? environment["BOT_STORAGE_PATH"] ?? "default")"
        do {
            guard FileManager.default.fileExists(atPath: directory.appendingPathComponent("dist/cli.js").path) else {
                throw ACPError.message("Build the harness first with npm --prefix harness run build, or set MOFF_HARNESS_DIR.")
            }
            let support = try FileManager.default.url(for: .applicationSupportDirectory, in: .userDomainMask,
                appropriateFor: nil, create: true).appendingPathComponent("MoffCompanion")
            try FileManager.default.createDirectory(at: support, withIntermediateDirectories: true)
            let child = Process()
            child.executableURL = URL(fileURLWithPath: "/usr/bin/env")
            child.arguments = [environment["MOFF_NODE"] ?? "node", directory.appendingPathComponent("dist/cli.js").path, "acp", "--ollama"]
            child.currentDirectoryURL = directory
            var childEnvironment = environment
            childEnvironment["PATH"] = (environment["PATH"] ?? "/usr/bin:/bin") + ":/opt/homebrew/bin:/usr/local/bin"
            childEnvironment["BOT_STORAGE_PATH"] = storagePath?.path ?? environment["BOT_STORAGE_PATH"] ?? support.appendingPathComponent("pi-bot.sqlite").path
            child.environment = childEnvironment
            let input = Pipe()
            let output = Pipe()
            let errors = Pipe()
            child.standardInput = input
            child.standardOutput = output
            child.standardError = errors
            stdinPipe = input
            stdoutPipe = output
            stderrPipe = errors
            process = child
            output.fileHandleForReading.readabilityHandler = { [weak self] handle in
                let data = handle.availableData
                DispatchQueue.main.async { self?.receive(data) }
            }
            errors.fileHandleForReading.readabilityHandler = { [weak self] handle in
                let data = handle.availableData
                DispatchQueue.main.async {
                    self?.stderr.append(data)
                    if let count = self?.stderr.count, count > 8192 { self?.stderr.removeFirst(count - 8192) }
                }
            }
            child.terminationHandler = { [weak self] child in
                let status = child.terminationStatus
                DispatchQueue.main.async {
                    guard let self else { return }
                    self.stopped(status: status)
                }
            }
            try child.run()
            let initialized = try await request("initialize", [
                "protocolVersion": .number(1),
                "clientInfo": .object(["name": .string("MoffCompanion"), "version": .string("0.1.0")]),
                "clientCapabilities": .object([:]),
            ])
            guard initialized["protocolVersion"].number == 1 else {
                throw ACPError.message("The agent negotiated an unsupported ACP version.")
            }
            let previous = defaults.string(forKey: sessionKey)
            let result = try await request(previous == nil ? "session/new" : "session/load", sessionParams(previous))
            let session = previous ?? result["sessionId"].string
            guard let session else { throw ACPError.message("The agent did not return a session ID.") }
            currentSessionId = session
            defaults.set(session, forKey: sessionKey)
            updateConfig(result)
            isConnected = true
            isConnecting = false
        } catch {
            fail(error)
            disconnect()
        }
    }

    private func sessionParams(_ sessionId: String?) -> [String: JSONValue] {
        var params: [String: JSONValue] = [
            "cwd": .string(harnessDirectory?.path ?? "/"),
            "mcpServers": .array([]),
        ]
        if let sessionId { params["sessionId"] = .string(sessionId) }
        return params
    }

    func newConversation() async {
        guard !isConnecting && !isBusy else { return }
        if !isConnected {
            defaults.removeObject(forKey: sessionKey)
            await connect()
            return
        }
        do {
            let result = try await request("session/new", sessionParams(nil))
            guard let id = result["sessionId"].string else { throw ACPError.message("Missing session ID.") }
            messages = []
            assistantId = nil
            currentSessionId = id
            defaults.set(id, forKey: sessionKey)
            errorMessage = nil
            updateConfig(result)
        } catch { fail(error) }
    }

    func sendPrompt(_ text: String) async {
        guard isConnected, !isBusy, let session = currentSessionId else { return }
        let trimmed = text.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !trimmed.isEmpty else { return }
        isBusy = true
        errorMessage = nil
        activity = "Thinking..."
        messages.append(ACPMessage(id: UUID().uuidString, role: "user", content: trimmed))
        let id = UUID().uuidString
        assistantId = id
        messages.append(ACPMessage(id: id, role: "assistant", content: ""))
        defer { isBusy = false; activity = ""; assistantId = nil }
        do {
            let result = try await request("session/prompt", [
                "sessionId": .string(session),
                "prompt": .array([.object(["type": .string("text"), "text": .string(trimmed)])]),
            ])
            if result["stopReason"].string == "cancelled" { errorMessage = "Stopped. The partial answer was saved." }
            else if result["stopReason"].string != "end_turn" {
                throw ACPError.message("The agent ended unexpectedly: \(result["stopReason"].string ?? "no reason").")
            }
        } catch { fail(error) }
    }

    func cancel() {
        guard isBusy, let session = currentSessionId else { return }
        do {
            try send(.object(["jsonrpc": .string("2.0"), "method": .string("session/cancel"),
                              "params": .object(["sessionId": .string(session)])]))
            activity = "Stopping..."
        } catch { fail(error) }
    }

    func selectModel(_ model: String) async {
        guard isConnected, !isBusy, let session = currentSessionId else { return }
        errorMessage = nil
        do {
            let result = try await request("session/set_config_option", [
                "sessionId": .string(session), "configId": .string("model"), "value": .string(model),
            ])
            updateConfig(result)
        } catch { fail(error) }
    }

    private func updateConfig(_ result: JSONValue) {
        if let model = result["configOptions"].array?.first(where: { $0["id"].string == "model" }) {
            availableModels = model["options"].array?.compactMap { $0["value"].string } ?? []
            selectedModel = model["currentValue"].string ?? selectedModel
        }
    }

    private func request(_ method: String, _ params: [String: JSONValue]) async throws -> JSONValue {
        requestId += 1
        let id = requestId
        return try await withCheckedThrowingContinuation { continuation in
            pending[id] = continuation
            do {
                try send(.object(["jsonrpc": .string("2.0"), "id": .number(Double(id)),
                                  "method": .string(method), "params": .object(params)]))
            } catch {
                pending.removeValue(forKey: id)?.resume(throwing: error)
            }
        }
    }

    private func send(_ value: JSONValue) throws {
        guard let pipe = stdinPipe, process?.isRunning == true else { throw ACPError.message("The agent is not running. Reconnect to continue.") }
        var data = try JSONEncoder().encode(value)
        data.append(10)
        try pipe.fileHandleForWriting.write(contentsOf: data)
    }

    private func receive(_ data: Data) {
        guard !data.isEmpty else { return }
        do {
            for message in try lines.append(data) {
                if let number = message["id"].number, number >= 0, number < Double(Int.max) {
                    let continuation = pending.removeValue(forKey: Int(number))
                    if let error = message["error"].object {
                        continuation?.resume(throwing: ACPError.message(error["message"]?.string ?? "Unknown agent error"))
                    } else if message.object?["result"] != nil {
                        continuation?.resume(returning: message["result"])
                    } else {
                        continuation?.resume(throwing: ACPError.message("Malformed ACP response."))
                    }
                } else if message["method"].string == "session/update" {
                    let params = message["params"]
                    if let currentSessionId, params["sessionId"].string != currentSessionId { continue }
                    let update = params["update"]
                    switch update["sessionUpdate"].string {
                    case "agent_message_chunk", "user_message_chunk":
                        guard update["content"]["type"].string == "text",
                              let text = update["content"]["text"].string else { continue }
                        let role = update["sessionUpdate"].string == "user_message_chunk" ? "user" : "assistant"
                        let id = update["_meta"]["moff/entryId"].string ?? assistantId ?? UUID().uuidString
                        if let index = messages.firstIndex(where: { $0.id == id }) { messages[index].content += text }
                        else { messages.append(ACPMessage(id: id, role: role, content: text)) }
                    case "tool_call":
                        activity = update["title"].string ?? "Reading with Moff..."
                    case "tool_call_update":
                        activity = update["status"].string == "failed" ? "Tool failed; waiting for the agent..." : "Thinking..."
                    default: break
                    }
                }
            }
        } catch {
            fail(error)
            disconnect()
        }
    }

    private func fail(_ error: Error) { errorMessage = error.localizedDescription }

    func disconnect() {
        intentionalDisconnect = true
        isConnected = false
        isConnecting = false
        do { try stdinPipe?.fileHandleForWriting.close() }
        catch { fail(error) }
        rejectPending("Disconnected. Reconnect to resume the saved conversation.")
        if process?.isRunning != true { cleanup() }
    }

    private func rejectPending(_ message: String) {
        let requests = pending
        pending.removeAll()
        for continuation in requests.values { continuation.resume(throwing: ACPError.message(message)) }
    }

    private func stopped(status: Int32) {
        if !intentionalDisconnect {
            let detail = String(data: stderr, encoding: .utf8) ?? ""
            errorMessage = "Agent exited (\(status)). \(detail) Reconnect to resume."
        }
        rejectPending(errorMessage ?? "The agent stopped.")
        isConnected = false
        isConnecting = false
        isBusy = false
        cleanup()
    }

    private func cleanup() {
        stdoutPipe?.fileHandleForReading.readabilityHandler = nil
        stderrPipe?.fileHandleForReading.readabilityHandler = nil
        process = nil
        stdinPipe = nil
        stdoutPipe = nil
        stderrPipe = nil
    }
}
