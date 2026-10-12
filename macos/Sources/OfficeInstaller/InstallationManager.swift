import Foundation

struct OfficeApplication: Identifiable {
    let id: String
    let name: String
    let bundleID: String

    private var applicationPaths: [String] {
        [
            "/Applications/\(name).app",
            "\(NSHomeDirectory())/Applications/\(name).app"
        ]
    }

    var isInstalled: Bool {
        applicationPaths.contains {
            FileManager.default.fileExists(atPath: $0)
        }
    }

    var installedVersion: String? {
        for path in applicationPaths {
            guard let bundle = Bundle(path: path),
                  bundle.bundleIdentifier == bundleID,
                  let version = bundle.object(
                    forInfoDictionaryKey: "CFBundleVersion"
                  ) as? String else {
                continue
            }

            return version
        }

        return nil
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


// MARK: - Сравнение версий Office

enum OfficeVersionComparator {

    static func isNewer(
        offeredVersion: String,
        installedVersion: String
    ) -> Bool {

        offeredVersion.compare(
            installedVersion,
            options: .numeric
        ) == .orderedDescending
    }
}

// MARK: - Проверка корпоративного пакета

enum LicenseVerificationError: LocalizedError {
    case fileNotFound
    case invalidFileType
    case verificationFailed(String)
    case untrustedPublisher

    var errorDescription: String? {
        switch self {
        case .fileNotFound:
            return "Файл лицензии не найден."

        case .invalidFileType:
            return "Необходимо выбрать установочный пакет .pkg."

        case .verificationFailed(let details):
            return "Не удалось подтвердить цифровую подпись: \(details)"

        case .untrustedPublisher:
            return "Пакет не подписан Microsoft Corporation."
        }
    }
}

enum LicenseVerifier {

    static func verify(_ url: URL) throws {

        guard url.pathExtension.lowercased() == "pkg" else {
            throw LicenseVerificationError.invalidFileType
        }

        guard FileManager.default.fileExists(atPath: url.path) else {
            throw LicenseVerificationError.fileNotFound
        }

        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/sbin/pkgutil")
        process.arguments = ["--check-signature", url.path]

        let outputPipe = Pipe()
        process.standardOutput = outputPipe
        process.standardError = outputPipe

        try process.run()

        let outputData = outputPipe.fileHandleForReading.readDataToEndOfFile()
        process.waitUntilExit()

        let output = String(
            data: outputData,
            encoding: .utf8
        ) ?? ""

        guard process.terminationStatus == 0 else {
            throw LicenseVerificationError.verificationFailed(output)
        }

        // Проверяем издателя по первому сертификату цепочки.
        // Важно: недостаточно просто искать слово Microsoft
        // в произвольном месте вывода.
        let expectedSigner =
            "1. Developer ID Installer: Microsoft Corporation (UBF8T346G9)"

        guard output.contains(expectedSigner) else {
            throw LicenseVerificationError.untrustedPublisher
        }

        // Требуем подтверждения доверенной подписи.
        guard output.contains(
            "Status: signed by a developer certificate issued by Apple for distribution"
        ) else {
            throw LicenseVerificationError.verificationFailed(
                "macOS не подтвердила доверенную подпись разработчика."
            )
        }

        // Требуем успешную проверку нотариального удостоверения.
        guard output.contains(
            "Notarization: trusted by the Apple notary service"
        ) else {
            throw LicenseVerificationError.verificationFailed(
                "Не подтверждена нотариальная проверка Apple."
            )
        }
    }
}