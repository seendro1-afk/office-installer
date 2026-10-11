import Foundation

struct OfficeApplication: Identifiable {
    let id: String
    let name: String
    let bundleID: String

    var isInstalled: Bool {
        let paths = [
            "/Applications/\(name).app",
            "\(NSHomeDirectory())/Applications/\(name).app"
        ]

        return paths.contains {
            FileManager.default.fileExists(atPath: $0)
        }
    }
}

enum OfficeCatalog {
    static let applications: [OfficeApplication] = [
        .init(
            id: "word",
            name: "Microsoft Word",
            bundleID: "com.microsoft.Word"
        ),
        .init(
            id: "excel",
            name: "Microsoft Excel",
            bundleID: "com.microsoft.Excel"
        ),
        .init(
            id: "powerpoint",
            name: "Microsoft PowerPoint",
            bundleID: "com.microsoft.Powerpoint"
        ),
        .init(
            id: "outlook",
            name: "Microsoft Outlook",
            bundleID: "com.microsoft.Outlook"
        ),
        .init(
            id: "onenote",
            name: "Microsoft OneNote",
            bundleID: "com.microsoft.onenote.mac"
        )
    ]
}