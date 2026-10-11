import SwiftUI

struct OfficeRepairView: View {

    let packageURL: URL?

    @State private var diagnostics: [OfficeRepairDiagnostic] = []
    @State private var isChecking = false
    @State private var errorMessage: String?
    @State private var hasChecked = false

    var body: some View {

        ScrollView {

            VStack(alignment: .leading, spacing: 20) {

                Text("Восстановление Office")
                    .font(.largeTitle)
                    .bold()

                Text(
                    "Проверка установленных приложений и возможности их повторной установки."
                )
                .foregroundStyle(.secondary)

                Divider()

                Button {
                    checkOffice()
                } label: {
                    Label(
                        isChecking ? "Проверка..." : "Проверить Office",
                        systemImage: "stethoscope"
                    )
                }
                .buttonStyle(.borderedProminent)
                .disabled(isChecking || packageURL == nil)

                if packageURL == nil {
                    Text(
                        "Сначала загрузите установочный пакет Microsoft во вкладке «Установка Office»."
                    )
                    .font(.caption)
                    .foregroundStyle(.secondary)
                }

                if isChecking {
                    ProgressView("Проверяем приложения Office...")
                }

                if let errorMessage {
                    Text(errorMessage)
                        .foregroundStyle(.red)
                }

                if hasChecked && diagnostics.isEmpty && errorMessage == nil {
                    Text("Данные диагностики отсутствуют.")
                        .foregroundStyle(.secondary)
                }

                ForEach(diagnostics) { diagnostic in

                    VStack(alignment: .leading, spacing: 8) {

                        HStack {
                            Text("Microsoft \(diagnostic.name)")
                                .font(.headline)

                            Spacer()

                            Label(
                                diagnostic.installed
                                    ? "Установлено"
                                    : "Не установлено",
                                systemImage: diagnostic.installed
                                    ? "checkmark.circle.fill"
                                    : "xmark.circle"
                            )
                            .foregroundStyle(
                                diagnostic.installed
                                    ? Color.green
                                    : Color.secondary
                            )
                        }

                        Text(
                            "Версия: \(diagnostic.version ?? "Не определена")"
                        )

                        Text(
                            "Пакет зарегистрирован: " +
                            (diagnostic.packageRegistered ? "Да" : "Нет")
                        )

                        Text(
                            "Компонент доступен в установщике: " +
                            (diagnostic.availableForInstallation ? "Да" : "Нет")
                        )
                    }
                    .padding(14)
                    .background(
                        RoundedRectangle(cornerRadius: 12)
                            .fill(Color.secondary.opacity(0.08))
                    )
                }

                if hasChecked && !diagnostics.isEmpty {
                    Text(
                        "Диагностика не изменяет приложения, документы, настройки или лицензии Office."
                    )
                    .font(.caption)
                    .foregroundStyle(.secondary)
                }
            }
            .padding(30)
        }
    }

    private func checkOffice() {

        guard let packageURL else {
            return
        }

        isChecking = true
        errorMessage = nil
        diagnostics = []
        hasChecked = false

        Task {
            do {
                let result = try await Task.detached {
                    try OfficeRepairManager.diagnose(
                        packageURL: packageURL
                    )
                }.value

                diagnostics = result
                hasChecked = true

            } catch {
                errorMessage = error.localizedDescription
            }

            isChecking = false
        }
    }
}
