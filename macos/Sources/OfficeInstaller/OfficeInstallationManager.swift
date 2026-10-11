import Foundation

// MARK: - Выбор компонентов Office

struct OfficeInstallationSelection: Sendable {

    // Основные приложения

    var word = false
    var excel = false
    var powerpoint = false
    var outlook = false
    var onenote = false

    // Дополнительные компоненты

    var onedrive = false
    var teams = false
    var copilot = false
    var defender = false

    // MARK: Основные приложения

    var selectedIdentifiers: Set<String> {

        var result = Set<String>()

        if word {
            result.insert("com.microsoft.word")
        }

        if excel {
            result.insert("com.microsoft.excel")
        }

        if powerpoint {
            result.insert("com.microsoft.powerpoint")
        }

        if outlook {
            result.insert("com.microsoft.outlook")
        }

        if onenote {
            result.insert("com.microsoft.onenote.mac")
        }

        return result
    }

    // MARK: Дополнительные компоненты

    var selectedOptionalIdentifiers: Set<String> {

        var result = Set<String>()

        if onedrive {
            result.insert("com.microsoft.OneDrive")
        }

        if teams {
            result.insert("com.microsoft.teams2")
        }

        if copilot {
            result.insert("com.microsoft.copilot.shim")
        }

        if defender {
            result.insert("com.microsoft.defender.shim")
        }

        return result
    }

    // MARK: Все выбранные компоненты

    var allSelectedIdentifiers: Set<String> {
        selectedIdentifiers.union(selectedOptionalIdentifiers)
    }
}

// MARK: - Ошибки установки

enum OfficeInstallationError: LocalizedError {

    case packageNotFound
    case invalidPackage
    case noApplicationsSelected
    case missingComponent(String)
    case unavailableComponent(String)
    case rejectedComponent(String)
    case unexpectedComponent(String)
    case preparationFailed(String)

    var errorDescription: String? {

        switch self {

        case .packageNotFound:
            return "Установочный пакет Office не найден."

        case .invalidPackage:
            return "Не удалось прочитать установочный пакет Office."

        case .noApplicationsSelected:
            return "Не выбран ни один компонент Office."

        case .missingComponent(let identifier):
            return "В пакете отсутствует компонент: \(identifier)"

        case .unavailableComponent(let identifier):
            return "Компонент недоступен для установки: \(identifier)"

        case .rejectedComponent(let identifier):
            return "macOS отклонила выбранный компонент: \(identifier)"

        case .unexpectedComponent(let identifier):
            return "Обнаружен незапрошенный компонент: \(identifier)"

        case .preparationFailed(let details):
            return "Ошибка подготовки установки: \(details)"
        }
    }
}

// MARK: - Менеджер установки

struct OfficeInstallationManager {

    // MARK: Запуск системного установщика для проверки

    private static func runInstaller(
        arguments: [String]
    ) throws -> Data {

        let process = Process()

        process.executableURL = URL(
            fileURLWithPath: "/usr/sbin/installer"
        )

        process.arguments = arguments

        let outputPipe = Pipe()
        let errorPipe = Pipe()

        process.standardOutput = outputPipe
        process.standardError = errorPipe

        do {
            try process.run()
        } catch {
            throw OfficeInstallationError.preparationFailed(
                error.localizedDescription
            )
        }

        let output = outputPipe.fileHandleForReading
            .readDataToEndOfFile()

        let errors = errorPipe.fileHandleForReading
            .readDataToEndOfFile()

        process.waitUntilExit()

        guard process.terminationStatus == 0 else {

            let message = String(
                data: errors,
                encoding: .utf8
            ) ?? "Неизвестная ошибка системного установщика."

            throw OfficeInstallationError.preparationFailed(
                message
            )
        }

        return output
    }

    // MARK: Проверка файла PKG

    private static func validatePackage(
        _ packageURL: URL
    ) throws {

        guard FileManager.default.fileExists(
            atPath: packageURL.path
        ) else {
            throw OfficeInstallationError.packageNotFound
        }

        guard packageURL.pathExtension.lowercased() == "pkg" else {
            throw OfficeInstallationError.invalidPackage
        }
    }

    // MARK: Получение доступных компонентов

    static func availableComponents(
        packageURL: URL
    ) throws -> [String: Bool] {

        try validatePackage(packageURL)

        let output = try runInstaller(
            arguments: [
                "-showChoicesXML",
                "-pkg", packageURL.path,
                "-target", "/"
            ]
        )

        let plist: Any

        do {
            plist = try PropertyListSerialization.propertyList(
                from: output,
                options: [],
                format: nil
            )
        } catch {
            throw OfficeInstallationError.invalidPackage
        }

        guard let choices = plist as? [[String: Any]] else {
            throw OfficeInstallationError.invalidPackage
        }

        var components: [String: Bool] = [:]

        func scan(_ items: [[String: Any]]) {

            for item in items {

                if let identifier = item["choiceIdentifier"] as? String {

                    let enabled =
                        (item["choiceIsEnabled"] as? NSNumber)?
                        .boolValue ?? true

                    components[identifier] = enabled
                }

                if let children =
                    item["childItems"] as? [[String: Any]] {

                    scan(children)
                }
            }
        }

        scan(choices)

        return components
    }

    // MARK: Проверка выбранных компонентов

    static func validateSelection(
        packageURL: URL,
        selectedApplications: Set<String>
    ) throws {

        guard !selectedApplications.isEmpty else {
            throw OfficeInstallationError.noApplicationsSelected
        }

        let components = try availableComponents(
            packageURL: packageURL
        )

        for identifier in selectedApplications.sorted() {

            guard let enabled = components[identifier] else {

                throw OfficeInstallationError.missingComponent(
                    identifier
                )
            }

            guard enabled else {

                throw OfficeInstallationError.unavailableComponent(
                    identifier
                )
            }
        }
    }

    // MARK: Проверка итогового XML

    private static func verifyChoices(
        packageURL: URL,
        choicesURL: URL,
        selection: OfficeInstallationSelection
    ) throws {

        let output = try runInstaller(
            arguments: [
                "-showChoicesAfterApplyingChangesXML",
                choicesURL.path,
                "-pkg", packageURL.path,
                "-target", "/"
            ]
        )

        let plist: Any

        do {
            plist = try PropertyListSerialization.propertyList(
                from: output,
                options: [],
                format: nil
            )
        } catch {
            throw OfficeInstallationError.invalidPackage
        }

        guard let changes = plist as? [[String: Any]] else {
            throw OfficeInstallationError.invalidPackage
        }

        var enabled: [String: Bool] = [:]
        var selected: [String: Bool] = [:]

        for item in changes {

            guard
                let identifier =
                    item["choiceIdentifier"] as? String,
                let attribute =
                    item["choiceAttribute"] as? String
            else {
                continue
            }

            let value =
                (item["attributeSetting"] as? NSNumber)?
                .boolValue ?? false

            switch attribute {

            case "enabled":
                enabled[identifier] = value

            case "selected":
                selected[identifier] = value

            default:
                break
            }
        }

        // Все выбранные компоненты должны быть
        // доступны и действительно включены.

        for identifier in selection.allSelectedIdentifiers.sorted() {

            guard enabled[identifier] == true else {

                throw OfficeInstallationError.unavailableComponent(
                    identifier
                )
            }

            guard selected[identifier] == true else {

                throw OfficeInstallationError.rejectedComponent(
                    identifier
                )
            }
        }

        // Основные и дополнительные приложения,
        // которые пользователь НЕ выбирал,
        // не должны оказаться включёнными.

        let configurableComponents =
            OfficeComponentInstaller.supportedApplications
            .union(OfficeComponentInstaller.optionalComponents)

        for identifier in configurableComponents.sorted() {

            if !selection.allSelectedIdentifiers.contains(identifier),
               selected[identifier] == true {

                throw OfficeInstallationError.unexpectedComponent(
                    identifier
                )
            }
        }

        // Служебные компоненты Microsoft
        // (Frameworks, Fonts, Proofing, Licensing)
        // здесь намеренно не отключаются.
    }

    // MARK: Подготовка конфигурации установки

    static func prepareChoices(
        packageURL: URL,
        selection: OfficeInstallationSelection
    ) throws -> URL {

        let selected = selection.allSelectedIdentifiers

        try validateSelection(
            packageURL: packageURL,
            selectedApplications: selected
        )

        let choicesURL =
            try OfficeComponentInstaller.createConfiguration(
                packageURL: packageURL,
                selectedApplications: selection.selectedIdentifiers,
                selectedOptionalComponents:
                    selection.selectedOptionalIdentifiers
            )

        try verifyChoices(
            packageURL: packageURL,
            choicesURL: choicesURL,
            selection: selection
        )

        return choicesURL
    }
}