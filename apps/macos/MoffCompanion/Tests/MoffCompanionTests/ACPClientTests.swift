import XCTest
@testable import MoffCompanion

final class ACPClientTests: XCTestCase {
    func testFramesHandleEveryByteBoundaryAndMultipleMessages() throws {
        let multibyte = String(decoding: [0xF0, 0x9F, 0xA6, 0x8B], as: UTF8.self)
        let frames = Data("{\"id\":1,\"result\":{\"text\":\"Hello \(multibyte)\"}}\n{\"id\":2,\"result\":{}}\n".utf8)
        for split in 0...frames.count {
            var buffer = ACPLineBuffer()
            let first = try buffer.append(Data(frames.prefix(split)))
            let second = try buffer.append(Data(frames.dropFirst(split)))
            let decoded = first + second
            XCTAssertEqual(decoded.count, 2)
            XCTAssertEqual(decoded[0]["id"].number, 1)
            XCTAssertEqual(decoded[0]["result"]["text"].string, "Hello \(multibyte)")
            XCTAssertEqual(decoded[1]["id"].number, 2)
        }
    }

    func testMalformedFramesFailExplicitly() {
        var buffer = ACPLineBuffer()
        XCTAssertThrowsError(try buffer.append(Data("not json\n".utf8)))
    }

    @MainActor
    func testMissingHarnessShowsActionableError() async {
        let client = ACPClient()
        await client.connect(harnessDirectory: URL(fileURLWithPath: "/nonexistent-moff-harness"))
        XCTAssertFalse(client.isConnected)
        XCTAssertFalse(client.isConnecting)
        XCTAssertTrue(client.errorMessage?.contains("Build the harness") == true)
    }

    @MainActor
    func testNativeClientCreatesAndReloadsARealACPSession() async throws {
        let directory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        let suite = "moff-test-\(UUID().uuidString)"
        let defaults = try XCTUnwrap(UserDefaults(suiteName: suite))
        let path = directory.appendingPathComponent("pi.sqlite")
        let first = ACPClient(defaults: defaults, storagePath: path)
        let second = ACPClient(defaults: defaults, storagePath: path)
        defer {
            first.disconnect()
            second.disconnect()
            defaults.removePersistentDomain(forName: suite)
            try? FileManager.default.removeItem(at: directory)
        }
        await first.connect()
        XCTAssertTrue(first.isConnected, first.errorMessage ?? "No connection")
        let session = try XCTUnwrap(first.currentSessionId)
        XCTAssertEqual(first.selectedModel, "qwen3.5:9b")
        XCTAssertTrue(first.availableModels.contains("qwen3.5:9b"))
        await first.selectModel("unavailable")
        XCTAssertTrue(first.errorMessage?.contains("Model not registered") == true)
        XCTAssertEqual(first.selectedModel, "qwen3.5:9b")
        await first.selectModel("llama3.2:latest")
        XCTAssertNil(first.errorMessage)
        XCTAssertEqual(first.selectedModel, "llama3.2:latest")
        first.disconnect()
        // Separate clients share the saved ID; loading also checks persistence in the real agent.
        await second.connect()
        XCTAssertTrue(second.isConnected, second.errorMessage ?? "No connection")
        XCTAssertEqual(second.currentSessionId, session)
        XCTAssertEqual(second.selectedModel, "llama3.2:latest")
        second.disconnect()
        try await Task.sleep(for: .milliseconds(300))
    }

    @MainActor
    func testLiveNativeClientStreamsQwenAnswer() async throws {
        guard ProcessInfo.processInfo.environment["MOFF_RUN_LIVE_TESTS"] == "1" else {
            throw XCTSkip("Set MOFF_RUN_LIVE_TESTS=1 to exercise the native client with local Qwen.")
        }
        let directory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        let suite = "moff-live-\(UUID().uuidString)"
        let defaults = try XCTUnwrap(UserDefaults(suiteName: suite))
        let client = ACPClient(defaults: defaults, storagePath: directory.appendingPathComponent("pi.sqlite"))
        defer {
            client.disconnect()
            defaults.removePersistentDomain(forName: suite)
            try? FileManager.default.removeItem(at: directory)
        }
        await client.connect()
        XCTAssertTrue(client.isConnected, client.errorMessage ?? "No connection")
        await client.sendPrompt("Reply with exactly NATIVE_ACP_OK and nothing else.")
        XCTAssertNil(client.errorMessage)
        XCTAssertFalse(client.isBusy)
        XCTAssertEqual(client.messages.last?.content.trimmingCharacters(in: .whitespacesAndNewlines), "NATIVE_ACP_OK")
        client.disconnect()
        try await Task.sleep(for: .milliseconds(300))
    }
}
