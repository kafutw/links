import Foundation
import Vision
import CoreImage
import AppKit

let args = CommandLine.arguments
guard args.count >= 3 else { print("usage: cutout in out.png"); exit(1) }
let url = URL(fileURLWithPath: args[1])
guard let ci = CIImage(contentsOf: url) else { print("cannot read"); exit(1) }
let req = VNGenerateForegroundInstanceMaskRequest()
let handler = VNImageRequestHandler(ciImage: ci, options: [:])
do {
  try handler.perform([req])
  guard let obs = req.results?.first else { print("no foreground"); exit(2) }
  let buf = try obs.generateMaskedImage(ofInstances: obs.allInstances, from: handler, croppedToInstancesExtent: true)
  let out = CIImage(cvPixelBuffer: buf)
  let ctx = CIContext()
  try ctx.writePNGRepresentation(of: out, to: URL(fileURLWithPath: args[2]), format: .RGBA8, colorSpace: CGColorSpaceCreateDeviceRGB())
  print("ok", out.extent.size, "instances", obs.allInstances.count)
} catch { print("err", error); exit(3) }
