import Foundation

struct OfficeRepairDiagnostic: Identifiable, Sendable {

    let id: String
    let name: String
    let installed: Bool
    let version: String?
    let packageRegistered: Bool
    let availableForInstallation: Bool
}

enum OfficeRepairManager {

    private struct OfficeApplication {

        let name: String
        let bundleName: String
        let componentIdentifier: String
        let packageIdentifier: String
    }

    private static let applications: [OfficeApplication] = [
        OfficeApplication(
            name: "Word",
            bundleName: "Microsoft Word.app",
            componentIdentifier: "com.microsoft.word",
            packageIdentifier: "com.microsoft.package.Microsoft_Word.app"
        ),
        OfficeApplication(
            name: "Excel",
            bundleName: "Microsoft Excel.app",
            componentIdentifier: "com.microsoft.excel",
            packageIdentifier: "com.microsoft.package.Microsoft_Excel.app"
        ),
        OfficeApplication(
            name: "PowerPoint",
            bundleName: "Microsoft PowerPoint.app",
            componentIdentifier: "com.microsoft.powerpoint",
            packageIdentifier: "com.microsoft.package.Microsoft_PowerPoint.app"
        ),
        OfficeApplication(
            name: "Outlook",
            bundleName: "Microsoft Outlook.app",
            componentIdentifier: "com.microsoft.outlook",
            packageIdentifier: "com.microsoft.package.Microsoft_Outlook.app"
        ),
        OfficeApplication(
            name: "OneNote",
            bundleName: "Microsoft OneNote.app",
            componentIdentifier: "com.microsoft.onenote.mac",
            packageIdentifier: "com.microsoft.package.Microsoft_OneNote.app"
        )
    ]

    static func diagnose(
        packageURL: URL
    ) throws -> [OfficeRepairDiagnostic] {

        let availableComponents =
            try OfficeInstallationManager.availableComponents(
                packageURL: packageURL
            )

        return applications.map { application in

            let appURL = URL(
                fileURLWithPath: "/Applications"
            ).appendingPathComponent(
                application.bundleName
            )

            let installed = FileManager.default.fileExists(
                atPath: appURL.path
            )

            let version = Bundle(
                url: appURL
            )?.infoDictionary?[
                "CFBundleShortVersionString"
            ] as? String

            let registered = isPackageRegistered(
                application.packageIdentifier
            )

            let available =
                availableComponents[
                    application.componentIdentifier
                ] ?? false

            return OfficeRepairDiagnostic(
                id: application.componentIdentifier,
                name: application.name,
                installed: installed,
                version: version,
                packageRegistered: registered,
                availableForInstallation: available
            )
        }
    }

    private static func isPackageRegistered(
        _ identifier: String
    ) -> Bool {

        let process = Process()

        process.executableURL = URL(
            fileURLWithPath: "/usr/sbin/pkgutil"
        )

        process.arguments = [
            "--pkg-info",
            identifier
        ]

        process.standardOutput = Pipe()
        process.standardError = Pipe()

        do {
            try process.run()
            process.waitUntilExit()

            return process.terminationStatus == 0

        } catch {
            return false
        }
    }
}