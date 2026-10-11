import SwiftUI

struct ContentView: View {
    @State private var word = true
    @State private var excel = true
    @State private var powerpoint = true
    @State private var outlook = true
    @State private var onenote = false

    var body: some View {
        VStack(alignment: .leading, spacing: 20) {
            Text("Office Installer 2024")
                .font(.largeTitle)
                .bold()

            Text("Выберите приложения Microsoft Office для установки")
                .foregroundStyle(.secondary)

            Divider()

            VStack(alignment: .leading, spacing: 12) {
                appRow("Microsoft Word", selected: $word)
                appRow("Microsoft Excel", selected: $excel)
                appRow("Microsoft PowerPoint", selected: $powerpoint)
                appRow("Microsoft Outlook", selected: $outlook)
                appRow("Microsoft OneNote", selected: $onenote)
            }

            Divider()

            Button("Установить выбранное") {
                print("Установка пока не реализована")
            }
            .buttonStyle(.borderedProminent)
            .disabled(
                !word && !excel && !powerpoint &&
                !outlook && !onenote
            )

            Text("Для использования Office необходима действующая лицензия Microsoft.")
                .font(.caption)
                .foregroundStyle(.secondary)
        }
        .padding(30)
        .frame(width: 560)
    }

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

            Text(installed ? "Установлено" : "Не установлено")
                .font(.caption)
                .foregroundStyle(installed ? .green : .secondary)
        }
    }
}