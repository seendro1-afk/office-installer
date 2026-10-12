import Foundation

enum OfficePackageMetadataError: LocalizedError {
    case extractionFailed(String)
    case distributionNotFound
    case invalidXML

    var errorDescription: String? {
        switch self {
        case .extractionFailed(let details):
            return "Не удалось прочитать метаданные Office: \(details)"
        case .distributionNotFound:
            return "В установочном пакете отсутствует Distribution."
        case .invalidXML:
            return "Не удалось обработать Distribution установщика Office."
        }
    }
}

enum OfficePackageMetadata {

    static func componentVersions(packageURL: URL) throws -> [String: String] {

        let temporaryDirectory = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)

        try FileManager.default.createDirectory(
            at: temporaryDirectory,
            withIntermediateDirectories: true
        )

        defer {
            try? FileManager.default.removeItem(at: temporaryDirectory)
        }

        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/xar")
        process.arguments = [
            "-xf",
            packageURL.path,
            "Distribution"
        ]
        process.currentDirectoryURL = temporaryDirectory

        let outputPipe = Pipe()
        process.standardOutput = outputPipe
        process.standardError = outputPipe

        try process.run()

        let outputData = outputPipe.fileHandleForReading.readDataToEndOfFile()
        process.waitUntilExit()

        guard process.terminationStatus == 0 else {
            let details = String(data: outputData, encoding: .utf8) ?? ""
            throw OfficePackageMetadataError.extractionFailed(details)
        }

        let distributionURL = temporaryDirectory
            .appendingPathComponent("Distribution")

        guard FileManager.default.fileExists(atPath: distributionURL.path) else {
            throw OfficePackageMetadataError.distributionNotFound
        }

        let data = try Data(contentsOf: distributionURL)

        let parserDelegate = DistributionVersionParser()
        let parser = XMLParser(data: data)
        parser.delegate = parserDelegate

        guard parser.parse() else {
            throw OfficePackageMetadataError.invalidXML
        }

        return parserDelegate.versions
    }
}

private final class DistributionVersionParser: NSObject, XMLParserDelegate {

    private(set) var versions: [String: String] = [:]

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        guard elementName == "pkg-ref",
              let identifier = attributeDict["id"],
              let version = attributeDict["version"],
              !version.isEmpty else {
            return
        }

        versions[identifier] = version
    }
}
