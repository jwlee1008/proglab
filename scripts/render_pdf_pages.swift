import AppKit
import PDFKit

let args = CommandLine.arguments
guard args.count == 3 else {
    fputs("usage: swift render_pdf_pages.swift input.pdf output_dir\n", stderr)
    exit(2)
}

let input = URL(fileURLWithPath: args[1])
let outputDir = URL(fileURLWithPath: args[2], isDirectory: true)
try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)

guard let pdf = PDFDocument(url: input) else {
    fputs("failed to open pdf\n", stderr)
    exit(1)
}

for index in 0..<pdf.pageCount {
    guard let page = pdf.page(at: index) else { continue }
    let pageRect = page.bounds(for: .mediaBox)
    let scale: CGFloat = 2.0
    let size = NSSize(width: pageRect.width * scale, height: pageRect.height * scale)
    let image = NSImage(size: size)

    image.lockFocus()
    NSColor.white.set()
    NSBezierPath(rect: NSRect(origin: .zero, size: size)).fill()
    guard let context = NSGraphicsContext.current?.cgContext else {
        image.unlockFocus()
        continue
    }
    context.saveGState()
    context.scaleBy(x: scale, y: scale)
    page.draw(with: .mediaBox, to: context)
    context.restoreGState()
    image.unlockFocus()

    guard let tiff = image.tiffRepresentation,
          let bitmap = NSBitmapImageRep(data: tiff),
          let png = bitmap.representation(using: .png, properties: [:]) else {
        fputs("failed to encode page \(index + 1)\n", stderr)
        continue
    }

    let out = outputDir.appendingPathComponent(String(format: "page-%d.png", index + 1))
    try png.write(to: out)
    print(out.path)
}
