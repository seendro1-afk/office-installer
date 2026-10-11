import Foundation
import Combine

enum OfficeDownloadError: LocalizedError {
    case invalidURL
    case untrustedHost
    case invalidResponse
    case invalidStatusCode(Int)
    case invalidFileName
    case emptyFile
    case alreadyDownloading
    case invalidPackage

    var errorDescription: String? {
        switch self {
        case .invalidURL:
            return "Некорректный адрес загрузки."
        case .untrustedHost:
            return "Адрес загрузки не принадлежит разрешённому серверу Microsoft."
        case .invalidResponse:
            return "Сервер вернул некорректный ответ."
        case .invalidStatusCode(let code):
            return "Ошибка сервера: HTTP \(code)."
        case .invalidFileName:
            return "Некорректное имя установочного пакета."
        case .emptyFile:
            return "Скачанный файл пуст."
        case .alreadyDownloading:
            return "Загрузка уже выполняется."
        case .invalidPackage:
            return "Загруженный файл не распознан как установочный пакет."
        }
    }
}

// MARK: - Потокобезопасный делегат

private final class OfficeDownloadDelegate: NSObject,
    URLSessionDownloadDelegate, @unchecked Sendable {

    private let lock = NSLock()

    private let destination: URL
    private let progressHandler: @Sendable (Double) -> Void
    private let completionHandler: @Sendable (Result<URL, Error>) -> Void

    private var completed = false

    private static let allowedHosts: Set<String> = [
        "go.microsoft.com",
        "officecdn.microsoft.com",
        "officecdnmac.microsoft.com",
        "res.public.onecdn.static.microsoft"
    ]

    init(
        destination: URL,
        progressHandler: @escaping @Sendable (Double) -> Void,
        completionHandler: @escaping @Sendable (Result<URL, Error>) -> Void
    ) {
        self.destination = destination
        self.progressHandler = progressHandler
        self.completionHandler = completionHandler
    }

    static func isTrustedURL(_ url: URL) -> Bool {
        guard url.scheme?.lowercased() == "https",
              let host = url.host?.lowercased(),
              url.user == nil,
              url.password == nil
        else {
            return false
        }

        return allowedHosts.contains(host)
    }

    private func finish(_ result: Result<URL, Error>) {
        lock.lock()

        guard !completed else {
            lock.unlock()
            return
        }

        completed = true
        lock.unlock()

        completionHandler(result)
    }

    // MARK: - Прогресс

    func urlSession(
        _ session: URLSession,
        downloadTask: URLSessionDownloadTask,
        didWriteData bytesWritten: Int64,
        totalBytesWritten: Int64,
        totalBytesExpectedToWrite: Int64
    ) {
        guard totalBytesExpectedToWrite > 0 else {
            return
        }

        let value = min(
            Double(totalBytesWritten) /
            Double(totalBytesExpectedToWrite),
            1
        )

        progressHandler(value)
    }

    // MARK: - Проверка перенаправлений

    func urlSession(
        _ session: URLSession,
        task: URLSessionTask,
        willPerformHTTPRedirection response: HTTPURLResponse,
        newRequest request: URLRequest,
        completionHandler: @escaping (URLRequest?) -> Void
    ) {
        guard let url = request.url,
              Self.isTrustedURL(url)
        else {
            completionHandler(nil)
            return
        }

        completionHandler(request)
    }

    // MARK: - Сохранение файла

    func urlSession(
        _ session: URLSession,
        downloadTask: URLSessionDownloadTask,
        didFinishDownloadingTo location: URL
    ) {
        guard let response = downloadTask.response as? HTTPURLResponse else {
            finish(.failure(OfficeDownloadError.invalidResponse))
            return
        }

        guard (200...299).contains(response.statusCode) else {
            finish(
                .failure(
                    OfficeDownloadError.invalidStatusCode(
                        response.statusCode
                    )
                )
            )
            return
        }

        guard let finalURL = response.url,
              Self.isTrustedURL(finalURL)
        else {
            finish(.failure(OfficeDownloadError.untrustedHost))
            return
        }

        do {
            let fileManager = FileManager.default

            if fileManager.fileExists(atPath: destination.path) {
                try fileManager.removeItem(at: destination)
            }

            try fileManager.moveItem(
                at: location,
                to: destination
            )

            let attributes = try fileManager.attributesOfItem(
                atPath: destination.path
            )

            let size = attributes[.size] as? Int64 ?? 0

            guard size > 0 else {
                try? fileManager.removeItem(at: destination)
                throw OfficeDownloadError.emptyFile
            }

            finish(.success(destination))

        } catch {
            finish(.failure(error))
        }
    }

    // MARK: - Сетевые ошибки

    func urlSession(
        _ session: URLSession,
        task: URLSessionTask,
        didCompleteWithError error: Error?
    ) {
        if let error {
            finish(.failure(error))
        }
    }
}

// MARK: - Основной загрузчик

@MainActor
final class OfficeDownloader: ObservableObject {

    @Published private(set) var progress: Double = 0
    @Published private(set) var isDownloading = false
    @Published private(set) var statusMessage = "Ожидание загрузки"

    private var session: URLSession?
    private var delegate: OfficeDownloadDelegate?

    func downloadPackage(
        from url: URL,
        fileName: String
    ) async throws -> URL {

        guard !isDownloading else {
            throw OfficeDownloadError.alreadyDownloading
        }

        guard OfficeDownloadDelegate.isTrustedURL(url) else {
            throw OfficeDownloadError.untrustedHost
        }

        guard fileName.lowercased().hasSuffix(".pkg"),
              URL(fileURLWithPath: fileName).lastPathComponent == fileName,
              fileName != ".pkg"
        else {
            throw OfficeDownloadError.invalidFileName
        }

        let directory = FileManager.default.temporaryDirectory
            .appendingPathComponent(
                "OfficeInstallerDownloads",
                isDirectory: true
            )

        try FileManager.default.createDirectory(
            at: directory,
            withIntermediateDirectories: true
        )

        let destination = directory.appendingPathComponent(fileName)

        progress = 0
        isDownloading = true
        statusMessage = "Подключение к Microsoft..."

        let delegate = OfficeDownloadDelegate(
            destination: destination,
            progressHandler: { [weak self] value in
                Task { @MainActor in
                    self?.progress = value
                    self?.statusMessage = "Загрузка Office..."
                }
            },
            completionHandler: { _ in }
        )

        // Делегат и продолжение создаются вместе ниже.
        self.delegate = delegate

        let configuration = URLSessionConfiguration.ephemeral
        configuration.timeoutIntervalForRequest = 60

        defer {
            isDownloading = false
            session?.invalidateAndCancel()
            session = nil
            self.delegate = nil
        }

        return try await withCheckedThrowingContinuation { continuation in

            let activeDelegate = OfficeDownloadDelegate(
                destination: destination,
                progressHandler: { [weak self] value in
                    Task { @MainActor in
                        self?.progress = value
                        self?.statusMessage = "Загрузка Office..."
                    }
                },
                completionHandler: { result in
                    continuation.resume(with: result)
                }
            )

            self.delegate = activeDelegate

            let activeSession = URLSession(
                configuration: configuration,
                delegate: activeDelegate,
                delegateQueue: nil
            )

            self.session = activeSession
            activeSession.downloadTask(with: url).resume()
        }
    }
}
