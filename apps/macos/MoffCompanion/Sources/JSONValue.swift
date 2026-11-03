import Foundation

enum JSONValue: Codable, Sendable {
    case object([String: JSONValue])
    case array([JSONValue])
    case string(String)
    case number(Double)
    case bool(Bool)
    case null

    init(from decoder: Decoder) throws {
        let container = try decoder.singleValueContainer()
        if container.decodeNil() { self = .null }
        else if let value = try? container.decode(Bool.self) { self = .bool(value) }
        else if let value = try? container.decode(Double.self) { self = .number(value) }
        else if let value = try? container.decode(String.self) { self = .string(value) }
        else if let value = try? container.decode([JSONValue].self) { self = .array(value) }
        else { self = .object(try container.decode([String: JSONValue].self)) }
    }

    func encode(to encoder: Encoder) throws {
        var container = encoder.singleValueContainer()
        switch self {
        case .object(let value): try container.encode(value)
        case .array(let value): try container.encode(value)
        case .string(let value): try container.encode(value)
        case .number(let value): try container.encode(value)
        case .bool(let value): try container.encode(value)
        case .null: try container.encodeNil()
        }
    }

    subscript(_ key: String) -> JSONValue { object?[key] ?? .null }
    var object: [String: JSONValue]? { if case .object(let value) = self { value } else { nil } }
    var array: [JSONValue]? { if case .array(let value) = self { value } else { nil } }
    var string: String? { if case .string(let value) = self { value } else { nil } }
    var number: Double? { if case .number(let value) = self { value } else { nil } }
}

enum ACPError: LocalizedError {
    case message(String)
    var errorDescription: String? {
        switch self { case .message(let text): text }
    }
}

struct ACPLineBuffer {
    private var buffer = Data()

    mutating func append(_ data: Data) throws -> [JSONValue] {
        buffer.append(data)
        var messages: [JSONValue] = []
        while let newline = buffer.firstIndex(of: 10) {
            let line = buffer[..<newline]
            if line.count > 16 * 1024 * 1024 {
                throw ACPError.message("The agent sent an oversized protocol frame.")
            }
            if !line.isEmpty { messages.append(try JSONDecoder().decode(JSONValue.self, from: Data(line))) }
            buffer.removeSubrange(...newline)
        }
        if buffer.count > 16 * 1024 * 1024 {
            throw ACPError.message("The agent sent an oversized protocol frame. Reconnect to start again.")
        }
        return messages
    }
}
