import SwiftUI
import AppKit
import UniformTypeIdentifiers

struct ContentView: View {

    // MARK: - Основные приложения

    @State private var word = false
    @State private var excel = false
    @State private var powerpoint = false
    @State private var outlook = false
    @State private var onenote = false

    // MARK: - Дополнительные компоненты

    @State private var onedrive = false
    @State private var teams = false
    @State private var copilot = false
    @State private var defender = false

    // MARK: - Доступность компонентов

    @State private var availableComponents: [String: Bool] = [:]
    @State private var componentVersions: [String: String] = [:]
    @State private var isCheckingComponents = false
    @State private var componentsChecked = false
    @State private var componentStatus = "Компоненты ещё не проверены."

    // MARK: - Лицензия

    @StateObject private var licenseInstaller = OfficeLicenseInstaller()

    @State private var licenseFile: URL?
    @State private var showLicensePicker = false
    @State private var isVerifyingLicense = false
    @State private var licenseVerified = false

    @State private var showLicenseInstallConfirmation = false
    @State private var licenseInstallResult = ""
    @State private var showLicenseInstallResult = false

    // MARK: - Загрузка

    @StateObject private var downloader = OfficeDownloader()

    @State private var downloadedPackage: URL?
    @State private var downloadedPackageVerified = false
    @State private var isVerifyingDownload = false

    // MARK: - Установка

    @State private var isInstalling = false
    @State private var installationStatus = ""

    // MARK: - Ошибки

    @State private var errorMessage = ""
    @State private var showError = false

    private let officeDownloadURL =
        "https://go.microsoft.com/fwlink/p/?linkid=2244846"

    // MARK: - Компоненты

    private let mainComponents: [(name: String, id: String)] = [
        ("Microsoft Word", "com.microsoft.word"),
        ("Microsoft Excel", "com.microsoft.excel"),
        ("Microsoft PowerPoint", "com.microsoft.powerpoint"),
        ("Microsoft Outlook", "com.microsoft.outlook"),
        ("Microsoft OneNote", "com.microsoft.onenote.mac")
    ]

    private let extraComponents: [(name: String, id: String)] = [
        ("Microsoft OneDrive", "com.microsoft.OneDrive"),
        ("Microsoft Teams", "com.microsoft.teams2"),
        ("Microsoft Copilot (компонент)", "com.microsoft.copilot.shim"),
        ("Microsoft Defender (компонент)", "com.microsoft.defender.shim")
    ]

    // MARK: - Выбор

    private var selection: OfficeInstallationSelection {
        OfficeInstallationSelection(
            word: word,
            excel: excel,
            powerpoint: powerpoint,
            outlook: outlook,
            onenote: onenote,
            onedrive: onedrive,
            teams: teams,
            copilot: copilot,
            defender: defender
        )
    }

    private var hasSelectedComponents: Bool {
        !selection.allSelectedIdentifiers.isEmpty
    }

    private var selectedNames: [String] {
        let selected = selection.allSelectedIdentifiers

        return (mainComponents + extraComponents)
            .filter { selected.contains($0.id) }
            .map(\.name)
    }

    private func isAvailable(_ identifier: String) -> Bool {
        componentsChecked && availableComponents[identifier] == true
    }

    // MARK: - Интерфейс

    var body: some View {

        ScrollView {

            VStack(alignment: .leading, spacing: 20) {

                Text("Office Installer LTSC 2024")
                    .font(.largeTitle)
                    .bold()

                Text("Установка Microsoft Office для macOS")
                    .foregroundStyle(.secondary)

                Divider()

                // MARK: Основные приложения

                Text("Основные приложения")
                    .font(.headline)

                Text("Выберите приложения, которые необходимо установить.")
                    .font(.caption)
                    .foregroundStyle(.secondary)

                VStack(alignment: .leading, spacing: 2) {

                    componentRow(
                        "Microsoft Word",
                        identifier: "com.microsoft.word",
                        selected: $word
                    )

                    componentRow(
                        "Microsoft Excel",
                        identifier: "com.microsoft.excel",
                        selected: $excel
                    )

                    componentRow(
                        "Microsoft PowerPoint",
                        identifier: "com.microsoft.powerpoint",
                        selected: $powerpoint
                    )

                    componentRow(
                        "Microsoft Outlook",
                        identifier: "com.microsoft.outlook",
                        selected: $outlook
                    )

                    componentRow(
                        "Microsoft OneNote",
                        identifier: "com.microsoft.onenote.mac",
                        selected: $onenote
                    )
                }

                Divider()

                // MARK: Дополнительные компоненты

                Text("Дополнительные компоненты")
                    .font(.headline)

                Text(
                    "Доступность зависит от установочного пакета Microsoft."
                )
                .font(.caption)
                .foregroundStyle(.secondary)

                VStack(alignment: .leading, spacing: 2) {

                    componentRow(
                        "Microsoft OneDrive",
                        identifier: "com.microsoft.OneDrive",
                        selected: $onedrive
                    )

                    componentRow(
                        "Microsoft Teams",
                        identifier: "com.microsoft.teams2",
                        selected: $teams
                    )

                    componentRow(
                        "Microsoft Copilot (компонент)",
                        identifier: "com.microsoft.copilot.shim",
                        selected: $copilot
                    )

                    componentRow(
                        "Microsoft Defender (компонент)",
                        identifier: "com.microsoft.defender.shim",
                        selected: $defender
                    )
                }

                HStack(spacing: 12) {

                    Button("Выбрать всё доступное") {
                        selectAllAvailable()
                    }
                    .disabled(!componentsChecked || isInstalling)

                    Button("Снять всё") {
                        deselectAll()
                    }
                    .disabled(isInstalling)
                }
                .buttonStyle(.bordered)

                Divider()

                // MARK: Проверка компонентов

                Text("Проверка компонентов")
                    .font(.headline)

                Button {
                    checkComponents()
                } label: {

                    Label(
                        "Проверить компоненты",
                        systemImage: "checkmark.shield"
                    )
                }
                .buttonStyle(.borderedProminent)
                .disabled(
                    !downloadedPackageVerified ||
                    isCheckingComponents ||
                    isInstalling
                )

                if isCheckingComponents {
                    ProgressView("Проверка компонентов Microsoft Installer...")
                }

                Text(componentStatus)
                    .font(.caption)
                    .foregroundStyle(.secondary)

                Divider()

                // MARK: Загрузка

                Text("Загрузка Office")
                    .font(.headline)

                Text(
                    "Установочный пакет загружается с официальных серверов Microsoft."
                )
                .font(.caption)
                .foregroundStyle(.secondary)

                Button {
                    startDownload()
                } label: {

                    Label(
                        downloader.isDownloading
                            ? "Загрузка..."
                            : "Скачать Office",
                        systemImage: "arrow.down.circle.fill"
                    )
                }
                .buttonStyle(.borderedProminent)
                .disabled(
                    downloader.isDownloading ||
                    isVerifyingDownload ||
                    isInstalling
                )

                if downloader.isDownloading {

                    VStack(alignment: .leading, spacing: 8) {

                        ProgressView(value: downloader.progress)

                        Text(downloader.statusMessage)
                            .font(.caption)
                            .foregroundStyle(.secondary)

                        Text("\(Int(downloader.progress * 100))%")
                            .font(.caption)
                    }
                }

                if isVerifyingDownload {
                    ProgressView("Проверка цифровой подписи...")
                }

                if let downloadedPackage {

                    Text(downloadedPackage.lastPathComponent)
                        .font(.caption)

                    if downloadedPackageVerified {

                        Label(
                            "Подпись Microsoft подтверждена",
                            systemImage: "checkmark.shield.fill"
                        )
                        .foregroundStyle(.green)
                        .font(.caption)

                    } else if !isVerifyingDownload {

                        Label(
                            "Подпись пакета не подтверждена",
                            systemImage: "xmark.shield.fill"
                        )
                        .foregroundStyle(.red)
                        .font(.caption)
                    }
                }

                Divider()

                // MARK: Установка

                Text("Установка Office")
                    .font(.headline)

                Button {
                    confirmInstallation()
                } label: {

                    Label(
                        isInstalling
                            ? "Установка..."
                            : "Установить выбранное",
                        systemImage: "square.and.arrow.down.fill"
                    )
                }
                .buttonStyle(.borderedProminent)
                .disabled(
                    !hasSelectedComponents ||
                    !componentsChecked ||
                    !downloadedPackageVerified ||
                    downloader.isDownloading ||
                    isVerifyingDownload ||
                    isCheckingComponents ||
                    isInstalling
                )

                if isInstalling {

                    HStack {

                        ProgressView()
                            .controlSize(.small)

                        Text(installationStatus)
                            .font(.caption)
                    }
                }

                Divider()

                // MARK: Лицензия

                Text("Корпоративная лицензия")
                    .font(.headline)

                Text(
                    "Для Office LTSC 2024 требуется действующая " +
                    "корпоративная лицензия Microsoft."
                )
                .font(.caption)
                .foregroundStyle(.secondary)

                Button("Выбрать файл лицензии (.pkg)") {
                    showLicensePicker = true
                }
                .buttonStyle(.bordered)

                if let licenseFile {

                    Text(licenseFile.lastPathComponent)
                        .font(.caption)

                    if isVerifyingLicense {

                        ProgressView("Проверка подписи...")

                    } else if licenseVerified {

                        Label(
                            "Подпись Microsoft подтверждена",
                            systemImage: "checkmark.shield.fill"
                        )
                        .foregroundStyle(.green)

                    } else {

                        Label(
                            "Подпись не подтверждена",
                            systemImage: "exclamationmark.shield.fill"
                        )
                        .foregroundStyle(.red)
                    }
                }

                Button("Применить корпоративную лицензию") {
                    showLicenseInstallConfirmation = true
                }
                .buttonStyle(.borderedProminent)
                .disabled(
                    !licenseVerified ||
                    licenseFile == nil ||
                    licenseInstaller.isInstalling
                )

                if licenseInstaller.isInstalling {

                    HStack {

                        ProgressView()
                            .controlSize(.small)

                        Text(licenseInstaller.statusMessage)
                            .font(.caption)
                    }
                }

                Text(
                    "Приложение не предоставляет лицензии Microsoft. " +
                    "Используйте только законно полученный VL Serializer."
                )
                .font(.caption)
                .foregroundStyle(.secondary)
            }
            .padding(30)
        }
        .frame(minWidth: 620, minHeight: 760)
        .onAppear(perform: restoreDownloadedOffice)

        // MARK: - Выбор лицензии

        .fileImporter(
            isPresented: $showLicensePicker,
            allowedContentTypes: [
                UTType(filenameExtension: "pkg")!
            ],
            allowsMultipleSelection: false
        ) { result in

            switch result {

            case .success(let urls):

                guard let url = urls.first else {
                    return
                }

                licenseFile = url
                licenseVerified = false

                verifyLicense(url)

            case .failure(let error):

                showErrorMessage(error.localizedDescription)
            }
        }

        // MARK: - Подтверждение лицензии

        .confirmationDialog(
            "Применить корпоративную лицензию?",
            isPresented: $showLicenseInstallConfirmation
        ) {

            Button("Продолжить") {

                guard let licenseFile else {
                    return
                }

                Task {

                    do {

                        try await licenseInstaller.install(
                            packageURL: licenseFile
                        )

                        licenseInstallResult =
                            "Установщик корпоративной лицензии " +
                            "завершил работу успешно."

                    } catch {

                        licenseInstallResult =
                            error.localizedDescription
                    }

                    showLicenseInstallResult = true
                }
            }

            Button("Отмена", role: .cancel) {}

        } message: {

            Text(
                "Будет запущен официальный Microsoft VL Serializer. " +
                "Для установки потребуются права администратора macOS."
            )
        }

        .alert(
            "Результат установки лицензии",
            isPresented: $showLicenseInstallResult
        ) {
            Button("OK", role: .cancel) {}
        } message: {
            Text(licenseInstallResult)
        }

        .alert(
            "Ошибка",
            isPresented: $showError
        ) {
            Button("OK", role: .cancel) {}
        } message: {
            Text(errorMessage)
        }
    }

    // MARK: - Строка компонента

    private func componentRow(
        _ name: String,
        identifier: String,
        selected: Binding<Bool>
    ) -> some View {

        let available = isAvailable(identifier)

        let application = OfficeCatalog.applications
            .first(where: { $0.name == name })

        let installed = application?.isInstalled ?? false
        let installedVersion = application?.installedVersion
        let offeredVersion = componentVersions[identifier]

        let status: String

        if !componentsChecked {
            status = "Не проверено"
        } else if !available {
            if installed,
               let installedVersion,
               let offeredVersion,
               !OfficeVersionComparator.isNewer(
                   offeredVersion: offeredVersion,
                   installedVersion: installedVersion
               ) {
                status = "Уже установлено (\(installedVersion))"
            } else {
                status = "Недоступно"
            }
        } else if installed {
            if let installedVersion,
               let offeredVersion,
               OfficeVersionComparator.isNewer(
                   offeredVersion: offeredVersion,
                   installedVersion: installedVersion
               ) {
                status = "Доступно обновление"
            } else {
                status = "Установлено"
            }
        } else {
            status = "Доступно для установки"
        }

        let iconNames: [String: String] = [
            "com.microsoft.word": "word",
            "com.microsoft.excel": "excel",
            "com.microsoft.powerpoint": "powerpoint",
            "com.microsoft.outlook": "outlook",
            "com.microsoft.onenote.mac": "onenote",
            "com.microsoft.OneDrive": "onedrive",
            "com.microsoft.teams2": "teams",
            "com.microsoft.copilot.shim": "copilot",
            "com.microsoft.defender.shim": "defender"
        ]

        return HStack(spacing: 14) {

            Group {
                if let iconName = iconNames[identifier],
                   let iconURL = (
                       Bundle.module.url(
                           forResource: iconName,
                           withExtension: "png"
                       )
                       ?? Bundle.module.url(
                           forResource: iconName,
                           withExtension: "png",
                           subdirectory: "Icons"
                       )
                   ),
                   let icon = NSImage(contentsOf: iconURL) {

                    Image(nsImage: icon)
                        .resizable()
                        .scaledToFit()

                } else {

                    Image(systemName: "square.grid.2x2.app.fill")
                        .resizable()
                        .scaledToFit()
                        .foregroundStyle(.secondary)
                }
            }
            .frame(width: 28, height: 28)

            Toggle(name, isOn: selected)
                .toggleStyle(.checkbox)
                .font(.system(size: 13))
                .disabled(!available || isInstalling)

            Spacer(minLength: 8)

            Circle()
                .fill(
                    !componentsChecked
                        ? Color.gray
                        : available
                            ? Color.blue
                            : installed
                                ? Color.green
                                : Color.gray
                )
                .frame(width: 6, height: 6)

            Text(status)
                .font(.system(size: 11))
                .foregroundStyle(.secondary)
                .lineLimit(1)
        }
        .frame(minHeight: 30)
        .padding(.horizontal, 8)
        .padding(.vertical, 2)
    }

    // MARK: - Выбрать всё доступное

    private func selectAllAvailable() {

        word = isAvailable("com.microsoft.word")
        excel = isAvailable("com.microsoft.excel")
        powerpoint = isAvailable("com.microsoft.powerpoint")
        outlook = isAvailable("com.microsoft.outlook")
        onenote = isAvailable("com.microsoft.onenote.mac")

        onedrive = isAvailable("com.microsoft.OneDrive")
        teams = isAvailable("com.microsoft.teams2")
        copilot = isAvailable("com.microsoft.copilot.shim")
        defender = isAvailable("com.microsoft.defender.shim")
    }

    // MARK: - Снять всё

    private func deselectAll() {

        word = false
        excel = false
        powerpoint = false
        outlook = false
        onenote = false

        onedrive = false
        teams = false
        copilot = false
        defender = false
    }

    // MARK: - Проверка компонентов

    private func checkComponents() {

        guard
            let package = downloadedPackage,
            downloadedPackageVerified
        else {
            return
        }

        isCheckingComponents = true
        componentsChecked = false
        componentStatus = "Проверка компонентов..."

        deselectAll()

        Task {

            let result: Result<([String: Bool], [String: String]), Error> =
                await Task.detached {

                    do {

                        let components =
                            try OfficeInstallationManager.availableComponents(
                                packageURL: package
                            )

                        let versions =
                            try OfficePackageMetadata.componentVersions(
                                packageURL: package
                            )

                        return .success((components, versions))

                    } catch {

                        return .failure(error)
                    }

                }.value

            isCheckingComponents = false

            switch result {

            case .success(let data):

                let (components, versions) = data

                availableComponents = components
                componentVersions = versions
                componentsChecked = true

                let selectableIdentifiers = [
                    "com.microsoft.word",
                    "com.microsoft.excel",
                    "com.microsoft.powerpoint",
                    "com.microsoft.outlook",
                    "com.microsoft.onenote.mac",
                    "com.microsoft.OneDrive",
                    "com.microsoft.teams2",
                    "com.microsoft.copilot.shim",
                    "com.microsoft.defender.shim"
                ]

                let count = selectableIdentifiers.filter {
                    components[$0] == true
                }.count

                componentStatus =
                    "Проверка завершена. Доступных компонентов: \(count)."

            case .failure(let error):

                availableComponents = [:]
                componentVersions = [:]
                componentsChecked = false

                componentStatus = "Ошибка проверки компонентов."

                showErrorMessage(error.localizedDescription)
            }
        }
    }

    // MARK: - Подтверждение установки

    private func confirmInstallation() {

        let alert = NSAlert()

        alert.messageText = "Установить выбранные компоненты Office?"

        alert.informativeText =
            "Выбраны компоненты:\n\n" +
            selectedNames.joined(separator: "\n") +
            "\n\nДля продолжения потребуются права администратора macOS."

        alert.alertStyle = .informational

        alert.addButton(withTitle: "Продолжить")
        alert.addButton(withTitle: "Отмена")

        if alert.runModal() == .alertFirstButtonReturn {
            prepareSelectedOfficeInstallation()
        }
    }

    // MARK: - Установка

    private func prepareSelectedOfficeInstallation() {

        guard
            let package = downloadedPackage,
            downloadedPackageVerified
        else {
            showErrorMessage("Установочный пакет Office не найден.")
            return
        }

        let currentSelection = selection
        let names = selectedNames

        isInstalling = true
        installationStatus = "Подготовка установки..."

        Task {

            defer {
                isInstalling = false
                installationStatus = ""
            }

            do {

                let choicesURL = try await Task.detached {

                    try LicenseVerifier.verify(package)

                    return try OfficeInstallationManager.prepareChoices(
                        packageURL: package,
                        selection: currentSelection
                    )

                }.value

                let confirmation = NSAlert()

                confirmation.messageText = "Начать установку Office?"

                confirmation.informativeText =
                    "Конфигурация успешно проверена.\n\n" +
                    "Выбраны компоненты:\n" +
                    names.joined(separator: ", ") +
                    "\n\nПотребуется пароль администратора macOS."

                confirmation.alertStyle = .warning

                confirmation.addButton(withTitle: "Установить")
                confirmation.addButton(withTitle: "Отмена")

                guard confirmation.runModal() == .alertFirstButtonReturn else {
                    return
                }

                installationStatus = "Установка Office..."

                let installer = OfficePrivilegedInstaller()

                try await installer.install(
                    packageURL: package,
                    choicesURL: choicesURL
                )

                showInformation(
                    title: "Установка завершена",
                    message: "Выбранные компоненты успешно установлены."
                )

                checkComponents()

            } catch {

                showErrorMessage(error.localizedDescription)
            }
        }
    }

    // MARK: - Восстановление скачанного Office

    private func restoreDownloadedOffice() {

        guard downloadedPackage == nil, !isVerifyingDownload else {
            return
        }

        let package = FileManager.default.temporaryDirectory
            .appendingPathComponent("OfficeInstallerDownloads")
            .appendingPathComponent("Microsoft_Office_Installer.pkg")

        guard FileManager.default.fileExists(
            atPath: package.path
        ) else {
            return
        }

        Task {

            isVerifyingDownload = true

            let result: Result<Void, Error> = await Task.detached {

                do {
                    try LicenseVerifier.verify(package)
                    return .success(())
                } catch {
                    return .failure(error)
                }

            }.value

            isVerifyingDownload = false

            switch result {

            case .success:

                downloadedPackage = package
                downloadedPackageVerified = true

                print("Office найден. Повторная загрузка не требуется.")

                checkComponents()

            case .failure(let error):

                downloadedPackage = nil
                downloadedPackageVerified = false

                showErrorMessage(error.localizedDescription)
            }
        }
    }

    // MARK: - Загрузка Office

    private func startDownload() {

        guard let url = URL(string: officeDownloadURL) else {
            showErrorMessage("Некорректная ссылка Microsoft.")
            return
        }

        downloadedPackage = nil
        downloadedPackageVerified = false
        availableComponents = [:]
        componentsChecked = false

        deselectAll()

        Task {

            do {

                let package = try await downloader.downloadPackage(
                    from: url,
                    fileName: "Microsoft_Office_Installer.pkg"
                )

                downloadedPackage = package
                isVerifyingDownload = true

                let result: Result<Void, Error> = await Task.detached {

                    do {
                        try LicenseVerifier.verify(package)
                        return .success(())
                    } catch {
                        return .failure(error)
                    }

                }.value

                isVerifyingDownload = false

                switch result {

                case .success:

                    downloadedPackageVerified = true
                    checkComponents()

                case .failure(let error):

                    downloadedPackageVerified = false
                    showErrorMessage(error.localizedDescription)
                }

            } catch {

                showErrorMessage(error.localizedDescription)
            }
        }
    }

    // MARK: - Проверка лицензии

    private func verifyLicense(_ url: URL) {

        isVerifyingLicense = true
        licenseVerified = false

        Task {

            let result: Result<Void, Error> = await Task.detached {

                let hasAccess = url.startAccessingSecurityScopedResource()

                defer {
                    if hasAccess {
                        url.stopAccessingSecurityScopedResource()
                    }
                }

                do {
                    try LicenseVerifier.verify(url)
                    return .success(())
                } catch {
                    return .failure(error)
                }

            }.value

            isVerifyingLicense = false

            switch result {

            case .success:
                licenseVerified = true

            case .failure(let error):
                licenseVerified = false
                showErrorMessage(error.localizedDescription)
            }
        }
    }

    // MARK: - Сообщения

    private func showInformation(
        title: String,
        message: String
    ) {

        let alert = NSAlert()

        alert.messageText = title
        alert.informativeText = message
        alert.alertStyle = .informational

        alert.addButton(withTitle: "OK")
        alert.runModal()
    }

    private func showErrorMessage(_ message: String) {
        errorMessage = message
        showError = true
    }
}