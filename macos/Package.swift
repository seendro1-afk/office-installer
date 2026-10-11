// swift-tools-version: 6.0

import PackageDescription

let package = Package(
    name: "OfficeInstaller",
    platforms: [
        .macOS(.v14)
    ],
    targets: [
        .executableTarget(
            name: "OfficeInstaller",
            path: "Sources/OfficeInstaller"
        )
    ]
)