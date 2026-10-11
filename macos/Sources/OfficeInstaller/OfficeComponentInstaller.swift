import Foundation

enum OfficeComponentInstallerError: LocalizedError {

    case packageNotFound
    case invalidPackage
    case noApplicationsSelected
    case invalidSelection
    case unavailableComponent(String)
    case configurationFailed(String)

    var errorDescription: String? {

        switch self {

        case .packageNotFound:
            return "Установочный пакет Office не найден."

        case .invalidPackage:
            return "Выбранный файл должен иметь расширение .pkg."

        case .noApplicationsSelected:
            return "Не выбран ни один компонент Office."

        case .invalidSelection:
            return "Выбранные компоненты не поддерживаются."

        case .unavailableComponent(let identifier):
            return "Компонент недоступен для установки: \(identifier)"

        case .configurationFailed(let message):
            return "Ошибка создания конфигурации: \(message)"
        }
    }
}

struct OfficeComponentInstaller {

    // MARK: - Основные приложения

    static let supportedApplications: Set<String> = [
        "com.microsoft.word",
        "com.microsoft.excel",
        "com.microsoft.powerpoint",
        "com.microsoft.outlook",
        "com.microsoft.onenote.mac"
    ]

    // MARK: - Дополнительные компоненты

    static let optionalComponents: Set<String> = [
        "com.microsoft.OneDrive",
        "com.microsoft.teams2",
        "com.microsoft.copilot.shim",
        "com.microsoft.defender.shim"
    ]

    // MARK: - Системные компоненты

    static let managedComponents: Set<String> = [
        "com.microsoft.office.frameworks",
        "com.microsoft.office.fonts",
        "com.microsoft.office.proofing",
        "com.microsoft.licensing"
    ]

    // MARK: - Все управляемые пользователем компоненты

    static var configurableComponents: Set<String> {
        supportedApplications.union(optionalComponents)
    }

    // MARK: - Создание конфигурации

    static func createConfiguration(
        packageURL: URL,
        selectedApplications: Set<String>,
        selectedOptionalComponents: Set<String> = []
    ) throws -> URL {

        guard FileManager.default.fileExists(
            atPath: packageURL.path
        ) else {
            throw OfficeComponentInstallerError.packageNotFound
        }

        guard packageURL.pathExtension.lowercased() == "pkg" else {
            throw OfficeComponentInstallerError.invalidPackage
        }

        guard selectedApplications.isSubset(
            of: supportedApplications
        ) else {
            throw OfficeComponentInstallerError.invalidSelection
        }

        guard selectedOptionalComponents.isSubset(
            of: optionalComponents
        ) else {
            throw OfficeComponentInstallerError.invalidSelection
        }

        let selectedComponents =
            selectedApplications.union(selectedOptionalComponents)

        guard !selectedComponents.isEmpty else {
            throw OfficeComponentInstallerError.noApplicationsSelected
        }

        // Получаем актуальную доступность компонентов
        // непосредственно от системного установщика macOS.

        let availableComponents =
            try OfficeInstallationManager.availableComponents(
                packageURL: packageURL
            )

        // Проверяем выбранные компоненты.

        for identifier in selectedComponents.sorted() {

            guard availableComponents[identifier] == true else {
                throw OfficeComponentInstallerError.unavailableComponent(
                    identifier
                )
            }
        }

        var changes: [[String: Any]] = []

        // Меняем состояние только доступных компонентов.
        // Недоступные компоненты не трогаем.

        for identifier in configurableComponents.sorted() {

            guard availableComponents[identifier] == true else {
                continue
            }

            changes.append([
                "choiceIdentifier": identifier,
                "choiceAttribute": "selected",
                "attributeSetting":
                    selectedComponents.contains(identifier) ? 1 : 0
            ])
        }

        // Frameworks, Fonts, Proofing, Licensing
        // остаются под управлением Microsoft Installer.

        let data: Data

        do {

            data = try PropertyListSerialization.data(
                fromPropertyList: changes,
                format: .xml,
                options: 0
            )

        } catch {

            throw OfficeComponentInstallerError.configurationFailed(
                error.localizedDescription
            )
        }

        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent(
                "OfficeInstallerConfigurations",
                isDirectory: true
            )

        do {

            try FileManager.default.createDirectory(
                at: directory,
                withIntermediateDirectories: true
            )

            let destination = directory.appendingPathComponent(
                UUID().uuidString + ".xml"
            )

            try data.write(
                to: destination,
                options: .atomic
            )

            return destination

        } catch {

            throw OfficeComponentInstallerError.configurationFailed(
                error.localizedDescription
            )
        }
    }
}