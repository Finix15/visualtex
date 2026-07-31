import CoreGraphics
import Foundation

func usage() -> Never {
    FileHandle.standardError.write(
        Data("Usage: swift macos_physical_double_click.swift <screen-x> <screen-y> [--appkit-y]\n".utf8)
    )
    exit(2)
}

guard (CommandLine.arguments.count == 3 || CommandLine.arguments.count == 4),
      let x = Double(CommandLine.arguments[1]),
      let sourceY = Double(CommandLine.arguments[2]),
      x.isFinite,
      sourceY.isFinite else {
    usage()
}
let appKitY = CommandLine.arguments.count == 4
    ? CommandLine.arguments[3] == "--appkit-y"
    : false
if CommandLine.arguments.count == 4 && !appKitY {
    usage()
}
let mainDisplayBounds = CGDisplayBounds(CGMainDisplayID())
let y = appKitY
    ? mainDisplayBounds.maxY - sourceY
    : sourceY
let point = CGPoint(x: x, y: y)
let source = CGEventSource(stateID: .hidSystemState)

func post(_ type: CGEventType, clickState: Int64) {
    guard let event = CGEvent(
        mouseEventSource: source,
        mouseType: type,
        mouseCursorPosition: point,
        mouseButton: .left
    ) else {
        FileHandle.standardError.write(Data("Unable to create Quartz mouse event\n".utf8))
        exit(1)
    }
    event.setIntegerValueField(.mouseEventClickState, value: clickState)
    event.post(tap: .cghidEventTap)
}

post(.mouseMoved, clickState: 0)
usleep(80_000)
post(.leftMouseDown, clickState: 1)
post(.leftMouseUp, clickState: 1)
usleep(90_000)
post(.leftMouseDown, clickState: 2)
post(.leftMouseUp, clickState: 2)
FileHandle.standardOutput.write(
    Data("CLICK|\(x)|\(y)|appKitY=\(appKitY)\n".utf8)
)
