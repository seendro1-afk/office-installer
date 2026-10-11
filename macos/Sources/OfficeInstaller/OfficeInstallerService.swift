import Foundation
import AppKit

enum OfficeInstallerServiceError: LocalizedError {
    case packageNotFound
    case choicesNotFound
    case installerLaunchFailed

    var errorDescription: String? {
        switch self {
        case .packageNotFound:
            return "Установочный пакет Office не найден."
        case .choicesNotFound:
            return "Не удалось подготовить настройки установки."
        case .installerLaunchFailed:
            return "Не удалось открыть системный установщик macOS."
        }
    }
}

@MainActor
final class OfficeInstallerService: ObservableObject {

    @Published private(set) var isPreparing = false
    @Published private(set) var statusMessage = "Ожидание установки"

    func prepareInstallation(
        packageURL: URL,
        selection: OfficeInstallationSelection
    ) async throws -> URL {

        guard FileManager.default.fileExists(atPath: packageURL.path) else {
            throw OfficeInstallerServiceError.packageNotFound
        }

        isPreparing = true
        statusMessage = "Проверка установочного пакета..."

        defer {
            isPreparing = false
        }

        let choicesURL = try await Task.detached {
            try LicenseVerifier.verify(packageURL)

            return try OfficeInstallationManager.prepareChoices(
                packageURL: packageURL,
                selection: selection
            )
        }.value

        statusMessage = "Конфигурация установки подготовлена"

        return choicesURL
    }

    func openSystemInstaller(packageURL: URL) throws {
        guard FileManager.default.fileExists(atPath: packageURL.path) else {
            throw OfficeInstallerServiceError.packageNotFound
        }

        let success = NSWorkspace.shared.open(packageURL)

        guard success else {
            throw OfficeInstallerServiceError.installerLaunchFailed
        }

        statusMessage = "Открыт системный установщик macOS"
    }
}
