import AppKit
import Vision
import Foundation

struct Result: Codable {
  let text: String
  let x: Double
  let y: Double
  let width: Double
  let height: Double
  let confidence: Float
}

guard CommandLine.arguments.count >= 2 else {
  fputs("usage: window_ocr <image-path>\n", stderr)
  exit(2)
}
let url = URL(fileURLWithPath: CommandLine.arguments[1])
guard let image = NSImage(contentsOf: url) else {
  fputs("unable to read image\n", stderr)
  exit(3)
}
var rect = NSRect(origin: .zero, size: image.size)
guard let cgImage = image.cgImage(forProposedRect: &rect, context: nil, hints: nil) else {
  fputs("unable to create cg image\n", stderr)
  exit(4)
}
let request = VNRecognizeTextRequest()
request.recognitionLevel = .accurate
request.recognitionLanguages = ["zh-Hans", "en-US"]
request.usesLanguageCorrection = false
try VNImageRequestHandler(cgImage: cgImage).perform([request])
let results: [Result] = (request.results ?? []).compactMap { observation in
  guard let candidate = observation.topCandidates(1).first else { return nil }
  let box = observation.boundingBox
  return Result(
    text: candidate.string,
    x: box.origin.x,
    y: box.origin.y,
    width: box.size.width,
    height: box.size.height,
    confidence: candidate.confidence
  )
}
let encoder = JSONEncoder()
encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
FileHandle.standardOutput.write(try encoder.encode(results))
FileHandle.standardOutput.write(Data("\n".utf8))
