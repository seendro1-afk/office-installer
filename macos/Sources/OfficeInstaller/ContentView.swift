import SwiftUI
import UniformTypeIdentifiers

struct ContentView: View {

    // MARK: - Приложения

    @State private var word = true
    @State private var excel = true
    @State private var powerpoint = true
    @State private var outlook = true
    @State private var onenote = false

    // MARK: - Лицензия

    @State private var licenseFile: URL?
    @State private var showLicensePicker = false

    @State private var isVerifying = false
    @State private var licenseVerified = false

    @State private var errorMessage: String?
    @State private var showError = false

    var body: some View {

        VStack(alignment: .leading, spacing: 20) {

            Text("Office Installer LTSC 2024")
                .font(.largeTitle)
                .bold()

            Text("Выберите приложения Microsoft Office для установки")
                .foregroundStyle(.secondary)

            Divider()

            // MARK: - Список приложений

            VStack(alignment: .leading, spacing: 12) {

                appRow("Microsoft Word", selected: $word)
                appRow("Microsoft Excel", selected: $excel)
                appRow("Microsoft PowerPoint", selected: $powerpoint)
                appRow("Microsoft Outlook", selected: $outlook)
                appRow("Microsoft OneNote", selected: $onenote)
            }

            Divider()

            // MARK: - Корпоративная лицензия

            Text("Корпоративная лицензия")
                .font(.headline)

            Text("Выберите официальный VL Serializer для Office LTSC 2024.")
                .font(.caption)
                .foregroundStyle(.secondary)

            Button("Выбрать файл лицензии (.pkg)") {
                showLicensePicker = true
            }
            .buttonStyle(.bordered)

            if let licenseFile {

                VStack(alignment: .leading, spacing: 10) {

                    Text(licenseFile.lastPathComponent)
                        .font(.caption)
                        .lineLimit(1)
                        .truncationMode(.middle)

                    if isVerifying {

                        HStack {

                            ProgressView()
                                .controlSize(.small)

                            Text("Проверка цифровой подписи...")
                                .font(.caption)
                                .foregroundStyle(.secondary)
                        }

                    } else if licenseVerified {

                        Label(
                            "Подпись Microsoft подтверждена",
                            systemImage: "checkmark.shield.fill"
                        )
                        .font(.caption)
                        .foregroundStyle(.green)

                    } else {

                        Label(
                            "Подпись не подтверждена",
                            systemImage: "exclamationmark.shield.fill"
                        )
                        .font(.caption)
                        .foregroundStyle(.red)
                    }
                }
            }

            Divider()

            // MARK: - Установка

            Button("Установить выбранное") {

                // Реальную установку добавим позже.
                print("Нажата кнопка установки")

                print("Подпись пакета подтверждена: \(licenseVerified)")
            }
            .buttonStyle(.borderedProminent)
            .disabled(
                isVerifying ||
                !licenseVerified ||
                (
                    !word &&
                    !excel &&
                    !powerpoint &&
                    !outlook &&
                    !onenote
                )
            )

            Text(
                "Для использования Office LTSC 2024 необходима " +
                "действующая корпоративная лицензия Microsoft."
            )
            .font(.caption)
            .foregroundStyle(.secondary)
            .fixedSize(horizontal: false, vertical: true)
        }
        .padding(30)
        .frame(width: 560)

        // MARK: - Выбор файла

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

                errorMessage = error.localizedDescription
                showError = true
            }
        }

        // MARK: - Ошибки

        .alert(
            "Ошибка проверки",
            isPresented: $showError
        ) {
            Button("OK", role: .cancel) {}
        } message: {
            Text(errorMessage ?? "Неизвестная ошибка")
        }
    }

    // MARK: - Проверка подписи

    private func verifyLicense(_ url: URL) {

        isVerifying = true
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

            isVerifying = false

            switch result {

            case .success:
                licenseVerified = true

            case .failure(let error):
                licenseVerified = false
                errorMessage = error.localizedDescription
                showError = true
            }
        }
    }

    // MARK: - Строка приложения

    private func appRow(
        _ name: String,
        selected: Binding<Bool>
    ) -> some View {

        HStack {

            Toggle(name, isOn: selected)
                .toggleStyle(.checkbox)

            Spacer()

            let installed = OfficeCatalog.applications
                .first(where: { $0.name == name })?
                .isInstalled ?? false

            Text(
                installed
                ? "Установлено"
                : "Не установлено"
            )
            .font(.caption)
            .foregroundStyle(
                installed
                ? Color.green
                : Color.secondary
            )
        }
    }
}