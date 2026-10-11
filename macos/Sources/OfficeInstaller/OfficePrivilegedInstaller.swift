import Foundation
import Combine

enum OfficePrivilegedInstallerError: LocalizedError {
    case packageNotFound
    case configurationNotFound
    case authorizationCancelled
    case installationFailed(String)

    var errorDescription: String? {
        switch self {
        case .packageNotFound:
            return "Установочный пакет Office не найден."
        case .configurationNotFound:
            return "Файл конфигурации установки не найден."
        case .authorizationCancelled:
            return "Установка отменена пользователем."
        case .installationFailed(let details):
            return "Ошибка установки: \(details)"
        }
    }
}

private struct CommandResult: Sendable {
    let exitCode: Int32
    let output: String
}

@MainActor
final class OfficePrivilegedInstaller: ObservableObject {

    @Published private(set) var isInstalling = false
    @Published private(set) var statusMessage = "Ожидание установки"

    func install(
        packageURL: URL,
        choicesURL: URL
    ) async throws {

        guard FileManager.default.fileExists(atPath: packageURL.path) else {
            throw OfficePrivilegedInstallerError.packageNotFound
        }

        guard FileManager.default.fileExists(atPath: choicesURL.path) else {
            throw OfficePrivilegedInstallerError.configurationNotFound
        }

        guard !isInstalling else {
            throw OfficePrivilegedInstallerError.installationFailed(
                "Операция уже выполняется."
            )
        }

        isInstalling = true
        statusMessage = "Проверка прав администратора..."

        defer {
            isInstalling = false
        }

        // Повторно проверяем подпись пакета перед установкой.
        try LicenseVerifier.verify(packageURL)

        let command = [
            "/usr/sbin/installer",
            "-pkg", shellQuote(packageURL.path),
            "-target", "/",
            "-applyChoiceChangesXML", shellQuote(choicesURL.path)
        ].joined(separator: " ")

        let script = """
        do shell script \(appleScriptQuote(command)) with administrator privileges
        """

        let result = try await Task.detached {
            () throws -> CommandResult in

            let process = Process()
            process.executableURL = URL(
                fileURLWithPath: "/usr/bin/osascript"
            )
            process.arguments = ["-e", script]

            let pipe = Pipe()
            process.standardOutput = pipe
            process.standardError = pipe

            try process.run()

            let data = pipe.fileHandleForReading.readDataToEndOfFile()
            process.waitUntilExit()

            return CommandResult(
                exitCode: process.terminationStatus,
                output: String(
                    data: data,
                    encoding: .utf8
                ) ?? ""
            )
        }.value

        print("osascript exit code: \(result.exitCode)")
        print("osascript output: \(result.output)")

        guard result.exitCode == 0 else {
            statusMessage = "Ошибка авторизации"

            if result.output.contains("User canceled") ||
               result.output.contains("(-128)") {
                throw OfficePrivilegedInstallerError.authorizationCancelled
            }

            throw OfficePrivilegedInstallerError.installationFailed(
                result.output
            )
        }

        statusMessage = "Установка завершена"
        print("Установщик завершил работу успешно.")
        print(result.output)
    }

    private func shellQuote(_ value: String) -> String {
        "'" + value.replacingOccurrences(
            of: "'",
            with: "'\\''"
        ) + "'"
    }

    private func appleScriptQuote(_ value: String) -> String {
        let escaped = value
            .replacingOccurrences(of: "\\", with: "\\\\")
            .replacingOccurrences(of: "\"", with: "\\\"")
            .replacingOccurrences(of: "\n", with: " ")

        return "\"\(escaped)\""
    }

}
