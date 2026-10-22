import SwiftUI

@main
struct MoffCompanionApp: App {
    var body: some Scene {
        Window("Moff Companion", id: "main") {
            MoffCompanionView()
        }
        .windowStyle(.titleBar)
        .defaultSize(width: 700, height: 500)
    }
}
