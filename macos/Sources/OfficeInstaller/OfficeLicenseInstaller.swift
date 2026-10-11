import Foundation
import Combine

enum OfficeLicenseInstallerError: LocalizedError {
    case fileNotFound
    case cancelled
    case installationFailed(String)

    var errorDescription: String? {
        switch self {
        case .fileNotFound:
            return "Файл корпоративной лицензии не найден."
        case .cancelled:
            return "Установка лицензии отменена пользователем."
        case .installationFailed(let message):
            return "Не удалось установить лицензию: \(message)"
        }
    }
}

private struct LicenseCommandResult: Sendable {
    let exitCode: Int32
    let output: String
}

@MainActor
final class OfficeLicenseInstaller: ObservableObject {

    @Published private(set) var isInstalling = false
    @Published private(set) var statusMessage = "Ожидание"

    func install(packageURL: URL) async throws {

        guard !isInstalling else {
            throw OfficeLicenseInstallerError.installationFailed(
                "Установка уже выполняется."
            )
        }

        isInstalling = true
        statusMessage = "Подготовка лицензии..."

        defer {
            isInstalling = false
        }

        let hasAccess = packageURL.startAccessingSecurityScopedResource()

        defer {
            if hasAccess {
                packageURL.stopAccessingSecurityScopedResource()
            }
        }

        guard FileManager.default.fileExists(atPath: packageURL.path) else {
            throw OfficeLicenseInstallerError.fileNotFound
        }

        // Проверяем оригинальный пакет Microsoft.
        statusMessage = "Проверка подписи Microsoft..."
        try LicenseVerifier.verify(packageURL)

        // Создаём отдельную временную папку.
        let temporaryDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(
                "OfficeLicense-\(UUID().uuidString)",
                isDirectory: true
            )

        try FileManager.default.createDirectory(
            at: temporaryDirectory,
            withIntermediateDirectories: true
        )

        defer {
            try? FileManager.default.removeItem(
                at: temporaryDirectory
            )
        }

        // Копируем пакет под простым именем.
        let temporaryPackage = temporaryDirectory
            .appendingPathComponent("OfficeLicense.pkg")

        statusMessage = "Подготовка установочного пакета..."

        try FileManager.default.copyItem(
            at: packageURL,
            to: temporaryPackage
        )

        // Повторно проверяем подпись копии.
        try LicenseVerifier.verify(temporaryPackage)

        // Правильно экранируем путь для shell.
        let quotedPath = shellQuote(temporaryPackage.path)

        let command = """
        /usr/sbin/installer -pkg \(quotedPath) -target /
        """

        // Правильно экранируем команду для AppleScript.
        let escapedCommand = command
            .replacingOccurrences(of: "\\", with: "\\\\")
            .replacingOccurrences(of: "\"", with: "\\\"")
            .replacingOccurrences(of: "\n", with: " ")

        let script = """
        do shell script "\(escapedCommand)" with administrator privileges
        """

        statusMessage = "Ожидание разрешения администратора..."

        let result = try await Task.detached {
            () throws -> LicenseCommandResult in

            let process = Process()

            process.executableURL = URL(
                fileURLWithPath: "/usr/bin/osascript"
            )

            process.arguments = [
                "-e",
                script
            ]

            let pipe = Pipe()

            process.standardOutput = pipe
            process.standardError = pipe

            try process.run()

            let data = pipe.fileHandleForReading.readDataToEndOfFile()

            process.waitUntilExit()

            return LicenseCommandResult(
                exitCode: process.terminationStatus,
                output: String(
                    data: data,
                    encoding: .utf8
                ) ?? ""
            )
        }.value

        print("VL Serializer exit code: \(result.exitCode)")
        print("VL Serializer output: \(result.output)")

        guard result.exitCode == 0 else {

            statusMessage = "Ошибка установки лицензии"

            if result.output.contains("User canceled") ||
               result.output.contains("(-128)") {
                throw OfficeLicenseInstallerError.cancelled
            }

            throw OfficeLicenseInstallerError.installationFailed(
                result.output
            )
        }

        statusMessage = "Установка лицензии завершена"
    }

    private func shellQuote(_ value: String) -> String {
        "'" + value.replacingOccurrences(
            of: "'",
            with: "'\\''"
        ) + "'"
    }
}
