// swift-tools-version: 6.0
import PackageDescription

let package = Package(
    name: "MoffCompanion",
    platforms: [.macOS(.v14)],
    products: [
        .executable(name: "MoffCompanion", targets: ["MoffCompanion"])
    ],
    targets: [
        .executableTarget(
            name: "MoffCompanion",
            path: "Sources"
        ),
        .testTarget(name: "MoffCompanionTests", dependencies: ["MoffCompanion"])
    ]
)
