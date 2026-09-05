import Cocoa
import Vision
import ImageIO

struct Window: Encodable { let number: UInt32; let title: String; let x: CGFloat; let y: CGFloat; let width: CGFloat; let height: CGFloat }
struct Observation: Encodable { let text: String; let x: CGFloat; let y: CGFloat; let w: CGFloat; let h: CGFloat }

if CommandLine.arguments.dropFirst().first == "windows" {
  let raw = CGWindowListCopyWindowInfo([.optionOnScreenOnly, .excludeDesktopElements], kCGNullWindowID) as? [[String: Any]] ?? []
  let windows = raw.compactMap { item -> Window? in
    guard let owner = item[kCGWindowOwnerName as String] as? String, owner == "Microsoft Outlook", let number = item[kCGWindowNumber as String] as? UInt32, let bounds = item[kCGWindowBounds as String] as? [String: CGFloat] else { return nil }
    return Window(number: number, title: item[kCGWindowName as String] as? String ?? "", x: bounds["X"] ?? 0, y: bounds["Y"] ?? 0, width: bounds["Width"] ?? 0, height: bounds["Height"] ?? 0)
  }
  print(String(data: try! JSONEncoder().encode(windows), encoding: .utf8)!)
} else if CommandLine.arguments.count == 3, CommandLine.arguments[1] == "ocr", let image = CGImageSourceCreateWithURL(URL(fileURLWithPath: CommandLine.arguments[2]) as CFURL, nil).flatMap({ CGImageSourceCreateImageAtIndex($0, 0, nil) }) {
  let request = VNRecognizeTextRequest()
  request.recognitionLevel = .accurate
  request.usesLanguageCorrection = true
  request.regionOfInterest = CGRect(x: 0.18, y: 0.10, width: 0.50, height: 0.82)
  let handler = VNImageRequestHandler(cgImage: image)
  try! handler.perform([request])
  let output = (request.results ?? []).compactMap { result -> Observation? in
    guard let candidate = result.topCandidates(1).first else { return nil }
    let b = result.boundingBox
    return Observation(text: candidate.string, x: b.origin.x, y: b.origin.y, w: b.width, h: b.height)
  }
  print(String(data: try! JSONEncoder().encode(output), encoding: .utf8)!)
} else { exit(64) }
