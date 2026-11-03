import SwiftUI

struct MoffCompanionView: View {
    @StateObject private var client = ACPClient()
    @State private var inputText: String = ""

    var body: some View {
        VStack(spacing: 0) {
            HStack {
                Label(client.isConnected ? "Moff · Read-only" : client.isConnecting ? "Connecting..." : "Disconnected",
                      systemImage: client.isConnected ? "checkmark.circle" : "network")
                    .font(.headline)

                Spacer()

                Button("New conversation", systemImage: "square.and.pencil") {
                    Task { await client.newConversation() }
                }
                .labelStyle(.iconOnly)
                .disabled(client.isConnecting || client.isBusy)
                Picker("Model", selection: Binding(
                    get: { client.selectedModel },
                    set: { model in Task { await client.selectModel(model) } }
                )) {
                    ForEach(client.availableModels, id: \.self) { model in
                        Text(model).tag(model)
                    }
                }
                .frame(width: 200)
                .disabled(!client.isConnected || client.isBusy)
            }
            .padding()
            .background(Color(NSColor.windowBackgroundColor))

            Divider()

            ScrollViewReader { proxy in
                ScrollView {
                    LazyVStack(alignment: .leading, spacing: 12) {
                        if client.messages.isEmpty {
                            ContentUnavailableView("Your Microsoft 365 assistant",
                                systemImage: "envelope",
                                description: Text("Ask Moff to read your latest emails. Conversations are saved locally with Pi Durable."))
                                .frame(maxWidth: .infinity)
                                .padding(.vertical, 40)
                        }
                        ForEach(client.messages) { msg in
                            HStack(alignment: .top) {
                                if msg.role == "user" {
                                    Spacer()
                                    Text(msg.content)
                                        .textSelection(.enabled)
                                        .padding(10)
                                        .background(Color.accentColor.opacity(0.2))
                                        .cornerRadius(8)
                                } else {
                                    VStack(alignment: .leading, spacing: 4) {
                                        Text("Moff")
                                            .font(.caption)
                                            .foregroundColor(.secondary)
                                        Text(msg.content.isEmpty ? (client.isBusy ? "Thinking..." : "No answer text.") : msg.content)
                                            .textSelection(.enabled)
                                            .padding(10)
                                            .background(Color(NSColor.controlBackgroundColor))
                                            .cornerRadius(8)
                                    }
                                    Spacer()
                                }
                            }
                            .id(msg.id)
                        }
                    }
                    .padding()
                }
                .onChange(of: client.messages.last?.content) {
                    if let last = client.messages.last {
                        proxy.scrollTo(last.id, anchor: .bottom)
                    }
                }
            }

            Divider()

            if let error = client.errorMessage {
                HStack(alignment: .top) {
                    Label(error, systemImage: "exclamationmark.triangle")
                        .textSelection(.enabled)
                    Spacer()
                    if !client.isConnected {
                        Button("Reconnect") { Task { await client.connect() } }
                            .disabled(client.isConnecting)
                    }
                }
                .font(.callout)
                .foregroundStyle(.red)
                .padding()
            }

            if client.isBusy {
                HStack {
                    ProgressView().controlSize(.small)
                    Text(client.activity).font(.callout).foregroundStyle(.secondary)
                    Spacer()
                    Button("Stop", action: client.cancel)
                }
                .padding(.horizontal)
                .padding(.top, 8)
            }

            HStack {
                TextField("Ask about your emails...", text: $inputText, axis: .vertical)
                    .lineLimit(1...5)
                    .textFieldStyle(RoundedBorderTextFieldStyle())
                    .disabled(!client.isConnected || client.isBusy)
                    .onSubmit {
                        sendMessage()
                    }

                Button("Send") {
                    sendMessage()
                }
                .keyboardShortcut(.return, modifiers: .command)
                .disabled(!client.isConnected || client.isBusy || inputText.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty)
            }
            .padding()
            .background(Color(NSColor.windowBackgroundColor))
        }
        .frame(minWidth: 600, minHeight: 450)
        .task { await client.connect() }
        .onDisappear { client.disconnect() }
    }

    private func sendMessage() {
        let text = inputText.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !text.isEmpty, client.isConnected, !client.isBusy else { return }
        inputText = ""
        Task { await client.sendPrompt(text) }
    }
}
