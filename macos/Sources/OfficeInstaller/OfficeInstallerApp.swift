import SwiftUI

@main
struct OfficeInstallerApp: App {

    var body: some Scene {

        WindowGroup {

            TabView {

                ContentView()
                    .tabItem {
                        Label(
                            "Установка Office",
                            systemImage: "square.and.arrow.down"
                        )
                    }

                OfficeRepairView(
                    packageURL: existingOfficePackage
                )
                .tabItem {
                    Label(
                        "Восстановление Office",
                        systemImage: "wrench.and.screwdriver"
                    )
                }
            }
            .frame(
                minWidth: 620,
                minHeight: 760
            )
        }
    }

    private var existingOfficePackage: URL? {

        let package = FileManager.default
            .temporaryDirectory
            .appendingPathComponent(
                "OfficeInstallerDownloads"
            )
            .appendingPathComponent(
                "Microsoft_Office_Installer.pkg"
            )

        guard FileManager.default.fileExists(
            atPath: package.path
        ) else {
            return nil
        }

        return package
    }
}