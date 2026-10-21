import AppKit
import UniformTypeIdentifiers
import Vision

// CGWindowListCreateImage 在 SDK 中已标记 obsoleted，但运行时仍可用，
// 用动态声明绕过编译期限制（用于采样悬浮窗背后的屏幕亮度）。
@_silgen_name("CGWindowListCreateImage")
private func legacyCGWindowListCreateImage(_ bounds: CGRect, _ listOption: UInt32, _ windowID: UInt32, _ imageOption: UInt32) -> CGImage?
private let kWindowListOptionOnScreenBelowWindow: UInt32 = 1 << 0
private let kWindowListOptionOnScreenOnly: UInt32 = 1 << 0   // kCGNullWindowID 时生效
private let kWindowListOptionIncludingWindow: UInt32 = 1 << 3   // 指定窗口 ID 时：只截该窗口（否则会截成它下面的窗口）
private let kWindowImageBestResolution: UInt32 = 1 << 0

// MARK: - 鼠标取题：OCR 与文本匹配

private func debugLog(_ msg: String) {
    let line = "[\(Int(Date().timeIntervalSince1970))] \(msg)\n"
    let path = "/tmp/quizbar_pick.log"
    if let fh = FileHandle(forWritingAtPath: path) {
        fh.seekToEndOfFile(); fh.write(line.data(using: .utf8)!); fh.closeFile()
    } else {
        try? line.write(toFile: path, atomically: true, encoding: .utf8)
    }
}

/// 截取指定全局区域（用户鼠标选区）并 OCR 识别文字
/// 返回 captured：截图本身是否成功
private func ocrRegion(global: CGRect, windowID: CGWindowID) -> (captured: Bool, text: String?) {
    var image = legacyCGWindowListCreateImage(global, kWindowListOptionIncludingWindow, windowID, kWindowImageBestResolution)
    if image == nil {
        debugLog("单窗口截图为空，回退全屏采集")
        image = legacyCGWindowListCreateImage(global, kWindowListOptionOnScreenOnly, 0, kWindowImageBestResolution)
    }
    guard let img = image else {
        debugLog("截图失败（nil）")
        return (false, nil)
    }
    if let png = NSBitmapImageRep(cgImage: img).representation(using: .png, properties: [:]) {
        try? png.write(to: URL(fileURLWithPath: "/tmp/quizbar_sel.png"))
    }
    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.recognitionLanguages = ["zh-Hans", "en-US"]
    request.usesLanguageCorrection = true       // 中文纠错，减少 OCR 错字
    request.minimumTextHeight = 0.025           // 过滤小干扰元素
    if VNRecognizeTextRequest.supportedRevisions.contains(VNRecognizeTextRequestRevision3) {
        request.revision = VNRecognizeTextRequestRevision3   // 新版中文识别模型
    }
    let handler = VNImageRequestHandler(cgImage: img, options: [:])
    do {
        try handler.perform([request])
    } catch {
        debugLog("Vision 错误: \(error)")
        return (true, nil)
    }
    let lines = (request.results ?? []).compactMap { $0.topCandidates(1).first?.string }
    debugLog("截图 \(img.width)x\(img.height) 识别行数 \(lines.count)")
    return (true, lines.isEmpty ? nil : lines.joined(separator: " "))
}

/// 仅保留中文、字母、数字，供模糊匹配
private func normalizedText(_ s: String) -> String {
    var out = String.UnicodeScalarView()
    for c in s.unicodeScalars where (c.value >= 0x4E00 && c.value <= 0x9FFF)
        || (c.value >= 0x30 && c.value <= 0x39)
        || (c.value >= 0x41 && c.value <= 0x5A)
        || (c.value >= 0x61 && c.value <= 0x7A) {
        out.append(c)
    }
    return String(out).lowercased()
}

/// 编辑距离相似度（0~1）
private func textSimilarity(_ a: String, _ b: String) -> Double {
    let x = Array(a), y = Array(b)
    guard !x.isEmpty, !y.isEmpty else { return 0 }
    var prev = Array(0...y.count)
    var cur = [Int](repeating: 0, count: y.count + 1)
    for i in 1...x.count {
        cur[0] = i
        for j in 1...y.count {
            let cost = x[i - 1] == y[j - 1] ? 0 : 1
            cur[j] = Swift.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
        }
        swap(&prev, &cur)
    }
    return 1 - Double(prev[y.count]) / Double(Swift.max(x.count, y.count))
}

// MARK: - XLSX 解析

final class XlsxParser: NSObject, XMLParserDelegate {
    static func parse(path: String) throws -> [(sheet: String, cells: [String])] {
        let tmp = URL(fileURLWithPath: NSTemporaryDirectory()).appendingPathComponent("ExcelFloat-\(UUID().uuidString)")
        try FileManager.default.createDirectory(at: tmp, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: tmp) }

        let ditto = Process()
        ditto.executableURL = URL(fileURLWithPath: "/usr/bin/ditto")
        ditto.arguments = ["-x", "-k", path, tmp.path]
        try ditto.run()
        ditto.waitUntilExit()
        guard ditto.terminationStatus == 0 else { throw NSError(domain: "Xlsx", code: 1, userInfo: [NSLocalizedDescriptionKey: "无法解压该文件，请确认是 .xlsx 格式"]) }

        let sharedStrings = try Self.parseWith(StringsParser.self, at: tmp.appendingPathComponent("xl/sharedStrings.xml"))?.strings ?? []
        let workbook = try Self.parseWith(WorkbookParser.self, at: tmp.appendingPathComponent("xl/workbook.xml"))
        let rels = try Self.parseWith(RelsParser.self, at: tmp.appendingPathComponent("xl/_rels/workbook.xml.rels"))?.rels ?? [:]
        let styles = try Self.parseWith(StylesParser.self, at: tmp.appendingPathComponent("xl/styles.xml"))

        guard let wb = workbook, !wb.sheets.isEmpty else { throw NSError(domain: "Xlsx", code: 2, userInfo: [NSLocalizedDescriptionKey: "文件中没有工作表"]) }

        let dateFormats = Self.dateFormatIds(styles)

        // 依次读取所有工作表，合并轮播
        var rows: [(sheet: String, cells: [String])] = []
        for s in wb.sheets {
            guard let target = rels[s.rid] else { continue }
            var sheetPath = "xl/" + target
            if target.hasPrefix("/") { sheetPath = String(target.dropFirst()) }
            let sheetURL = tmp.appendingPathComponent(sheetPath)
            guard FileManager.default.fileExists(atPath: sheetURL.path),
                  let sp = try Self.parseWith(SheetParser.self, at: sheetURL) else { continue }
            for row in sp.buildRows() {
                var cells: [String] = []
                for cell in row {
                    var text = cell.text
                    if cell.isSharedIndex {
                        let idx = Int(cell.text) ?? -1
                        text = idx >= 0 && idx < sharedStrings.count ? sharedStrings[idx] : ""
                    } else if Self.styleIsDate(styles, cell.style, dateFormats), let serial = Double(cell.text) {
                        text = Self.formatExcelDate(serial)
                    } else if text.hasSuffix(".0") {
                        text = String(text.dropLast(2))
                    }
                    cells.append(text.trimmingCharacters(in: CharacterSet.whitespacesAndNewlines))
                }
                if cells.contains(where: { !$0.isEmpty }) { rows.append((s.name, cells)) }
            }
        }
        guard !rows.isEmpty else { throw NSError(domain: "Xlsx", code: 5, userInfo: [NSLocalizedDescriptionKey: "所有工作表都是空的"]) }
        return rows
    }

    private static func styleIsDate(_ styles: StylesParser?, _ style: Int, _ dateFmtIds: Set<Int>) -> Bool {
        guard let styles = styles else { return false }
        return dateFmtIds.contains(styles.numFmtId(forStyle: style))
    }

    private static func parseWith<T: NSObject & XMLParserDelegate>(_ type: T.Type, at url: URL) throws -> T? {
        guard FileManager.default.fileExists(atPath: url.path) else { return nil }
        let p = XMLParser(contentsOf: url)
        let delegate = T()
        p?.delegate = delegate
        guard p?.parse() == true else { throw NSError(domain: "Xlsx", code: 10, userInfo: [NSLocalizedDescriptionKey: "XML 解析失败: \(url.lastPathComponent)"]) }
        return delegate
    }

    private static func dateFormatIds(_ styles: StylesParser?) -> Set<Int> {
        let builtin: Set<Int> = [14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]
        var ids = builtin
        guard let styles = styles else { return ids }
        for (fmtId, code) in styles.customFormats {
            let lower = code.lowercased()
            if lower.contains("y") || lower.contains("d") { ids.insert(fmtId) }
        }
        return ids
    }

    private static func formatExcelDate(_ serial: Double) -> String {
        let date = Date(timeIntervalSince1970: (serial - 25569) * 86400)
        let hasTime = abs(serial.rounded() - serial) > 0.000001
        let fmt = DateFormatter()
        fmt.timeZone = TimeZone(secondsFromGMT: 0)   // Excel 序列号本身不含时区，按 UTC 解析避免时间偏移
        fmt.dateFormat = hasTime ? "yyyy/M/d HH:mm" : "yyyy/M/d"
        return fmt.string(from: date)
    }

    // MARK: 各 XML 的代理

    final class StringsParser: NSObject, XMLParserDelegate {
        var strings: [String] = []
        private var current = ""
        private var inSi = false
        func parser(_ p: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName: String?, attributes: [String: String]) {
            if name == "si" { inSi = true; current = "" }
        }
        func parser(_ p: XMLParser, foundCharacters string: String) { if inSi { current += string } }
        func parser(_ p: XMLParser, didEndElement name: String, namespaceURI: String?, qualifiedName: String?) {
            if name == "si" { strings.append(current); inSi = false }
        }
    }

    final class WorkbookParser: NSObject, XMLParserDelegate {
        var sheets: [(name: String, rid: String)] = []
        func parser(_ p: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName: String?, attributes: [String: String]) {
            if name == "sheet", let n = attributes["name"], let rid = attributes["r:id"] {
                sheets.append((n, rid))
            }
        }
    }

    final class RelsParser: NSObject, XMLParserDelegate {
        var rels: [String: String] = [:]
        func parser(_ p: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName: String?, attributes: [String: String]) {
            if name == "Relationship", let id = attributes["Id"], let target = attributes["Target"] {
                rels[id] = target
            }
        }
    }

    final class StylesParser: NSObject, XMLParserDelegate {
        var customFormats: [Int: String] = [:]
        private var inNumFmts = false
        private var xfNumFmtIds: [Int] = []
        private var inCellXfs = false
        func parser(_ p: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName: String?, attributes: [String: String]) {
            switch name {
            case "numFmts": inNumFmts = true
            case "numFmt":
                if inNumFmts, let id = attributes["numFmtId"].flatMap({ Int($0) }), let code = attributes["formatCode"] {
                    customFormats[id] = code
                }
            case "cellXfs": inCellXfs = true
            case "xf":
                if inCellXfs { xfNumFmtIds.append(Int(attributes["numFmtId"] ?? "") ?? 0) }
            default: break
            }
        }
        func parser(_ p: XMLParser, didEndElement name: String, namespaceURI: String?, qualifiedName: String?) {
            if name == "numFmts" { inNumFmts = false }
            if name == "cellXfs" { inCellXfs = false }
        }
        func numFmtId(forStyle style: Int) -> Int { style < xfNumFmtIds.count ? xfNumFmtIds[style] : 0 }
    }

    struct Cell { var text = ""; var style = 0; var isSharedIndex = false }

    final class SheetParser: NSObject, XMLParserDelegate {
        private var rows: [(num: Int, cells: [Int: Cell])] = []
        private var curRowNum = 0
        private var curRow: [Int: Cell] = [:]
        private var curCol = 0
        private var curCell = Cell()
        private var inV = false
        private var inInlineT = false

        func parser(_ p: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName: String?, attributes: [String: String]) {
            switch name {
            case "row":
                curRowNum = Int(attributes["r"] ?? "") ?? (rows.map(\.num).max().map { $0 + 1 } ?? 1)
                curRow = [:]
            case "c":
                curCell = Cell()
                if let ref = attributes["r"] {
                    curCol = Self.columnIndex(from: ref)
                } else {
                    curCol += 1
                }
                if attributes["t"] == "s" { curCell.isSharedIndex = true }
                curCell.style = Int(attributes["s"] ?? "") ?? 0
            case "v": inV = true
            case "t": if !inV { inInlineT = true }
            default: break
            }
        }

        func parser(_ p: XMLParser, foundCharacters string: String) {
            if inV || inInlineT { curCell.text += string }
        }

        func parser(_ p: XMLParser, didEndElement name: String, namespaceURI: String?, qualifiedName: String?) {
            switch name {
            case "v": inV = false
            case "t": inInlineT = false
            case "c": curRow[curCol] = curCell
            case "row": rows.append((curRowNum, curRow))
            default: break
            }
        }

        func buildRows() -> [[Cell]] {
            let sorted = rows.sorted { $0.num < $1.num }
            let maxCol = sorted.map { $0.cells.keys.max() ?? 0 }.max() ?? 0
            return sorted.map { row in
                (0...maxCol).map { row.cells[$0] ?? Cell() }
            }
        }

        static func columnIndex(from ref: String) -> Int {
            var n = 0
            for ch in ref {
                guard ch.isLetter else { break }
                n = n * 26 + (Int(ch.asciiValue! - 64))
            }
            return n - 1
        }
    }
}

// MARK: - 悬浮面板

final class FloatPanel: NSPanel {
    override var canBecomeKey: Bool { true }
}

final class CardView: NSView {
    var onFileDropped: ((URL) -> Void)?

    override init(frame: NSRect) {
        super.init(frame: frame)
        registerForDraggedTypes([.fileURL])
    }
    required init?(coder: NSCoder) { fatalError() }

    override func draggingEntered(_ sender: NSDraggingInfo) -> NSDragOperation {
        let urls = sender.draggingPasteboard.readObjects(forClasses: [NSURL.self]) as? [URL] ?? []
        return urls.contains { $0.pathExtension.lowercased() == "xlsx" } ? .copy : []
    }
    override func performDragOperation(_ sender: NSDraggingInfo) -> Bool {
        let urls = sender.draggingPasteboard.readObjects(forClasses: [NSURL.self]) as? [URL] ?? []
        if let url = urls.first(where: { $0.pathExtension.lowercased() == "xlsx" }) {
            onFileDropped?(url)
            return true
        }
        return false
    }
}

@MainActor
final class AppDelegate: NSObject, NSApplicationDelegate, NSWindowDelegate {
    private var panel: FloatPanel!
    private var effectView: NSVisualEffectView!
    private var label: NSTextField!
    private var footer: NSTextField!
    private var playMenuItem: NSMenuItem!
    private var intervalItems: [NSMenuItem] = []
    private var fontSizeItems: [NSMenuItem] = []
    private var pickBtn: NSButton!
    private var capBtn: NSButton!

    private var rows: [(sheet: String, cells: [String])] = []
    private var index = 0
    private var timer: Timer?
    private var contrastTimer: Timer?
    private var filePath: String?
    private var fileMTime: Date?
    private var paused = false
    private var lightBackground = false
    private var captureMenuItem: NSMenuItem!
    private var blockCapture: Bool {
        get { UserDefaults.standard.object(forKey: "blockCapture") as? Bool ?? true }
        set { UserDefaults.standard.set(newValue, forKey: "blockCapture") }
    }
    private var fgColor = NSColor.white
    private var accentColor = NSColor(red: 0.54, green: 0.71, blue: 0.95, alpha: 1)
    private let cardWidth: CGFloat = 440
    private var isFixingFrame = false
    private var anchorRight: CGFloat = 0   // 右上角锚点（全局坐标）
    private var anchorTop: CGFloat = 0
    private var contentTextHeight: CGFloat = 20
    private var contentFooterHeight: CGFloat = 12
    private var pickMenuItem: NSMenuItem!
    private var pickTimer: Timer?
    private var isPicking = false
    private var selStartPoint: NSPoint?   // 鼠标按下起点（检测划选）
    private var lastPickText = ""
    private var captureFailStreak = 0
    private var noPermission = false   // 通过窗口标题可见性探测屏幕录制权限

    private var pickMode: Bool {
        get { UserDefaults.standard.object(forKey: "pickMode") as? Bool ?? false }
        set { UserDefaults.standard.set(newValue, forKey: "pickMode") }
    }

    private var interval: TimeInterval {
        get { UserDefaults.standard.double(forKey: "interval") == 0 ? 3 : UserDefaults.standard.double(forKey: "interval") }
        set { UserDefaults.standard.set(newValue, forKey: "interval") }
    }
    private var fontSize: CGFloat {
        get { CGFloat(UserDefaults.standard.object(forKey: "fontSize") as? Double ?? 24) }
        set { UserDefaults.standard.set(Double(newValue), forKey: "fontSize") }
    }

    func applicationDidFinishLaunching(_ notification: Notification) {
        // 单实例保护：已有运行中的实例时直接退出当前进程
        let others = NSRunningApplication.runningApplications(withBundleIdentifier: "com.onejiall.quizbar")
            .filter { $0 != NSRunningApplication.current }
        if others.first != nil {
            NSApp.terminate(nil)
            return
        }
        buildUI()
        applyTheme()
        startTimer()
        startContrastTimer()
        if pickMode { startPickTimer() }
        NotificationCenter.default.addObserver(self, selector: #selector(screenChanged),
                                               name: NSApplication.didChangeScreenParametersNotification,
                                               object: nil)
        NotificationCenter.default.addObserver(self, selector: #selector(windowResized),
                                               name: NSWindow.didResizeNotification,
                                               object: nil)
        if let saved = UserDefaults.standard.string(forKey: "lastFile"), FileManager.default.fileExists(atPath: saved) {
            loadFile(path: saved)
        } else {
            label.stringValue = "把 .xlsx 表格拖到这张卡片上\n或右键 → 打开 Excel 文件…"
            footer.stringValue = "点击暂停/继续 · 可拖动位置 · 右键更多选项"
        }
    }

    // MARK: 选中取题

    private func startPickTimer() {
        pickTimer?.invalidate()
        pickTimer = Timer.scheduledTimer(withTimeInterval: 0.1, repeats: true) { [weak self] _ in
            DispatchQueue.main.async { self?.pickPoll() }
        }
    }

    private func stopPickTimer() {
        pickTimer?.invalidate()
        pickTimer = nil
    }

    /// 轮询鼠标：按下记起点，松开后若形成划选（位移足够）则识别选区
    private func pickPoll() {
        guard pickMode, !isPicking, !rows.isEmpty else { return }
        let loc = NSEvent.mouseLocation
        let leftDown = NSEvent.pressedMouseButtons & 1 != 0
        if leftDown {
            if selStartPoint == nil { selStartPoint = loc }
            return
        }
        guard let start = selStartPoint else { return }
        selStartPoint = nil
        // 点击（几乎无位移）不算划选
        guard abs(loc.x - start.x) > 12 || abs(loc.y - start.y) > 6 else { return }

        let mid = NSPoint(x: (start.x + loc.x) / 2, y: (start.y + loc.y) / 2)
        guard let (winID, _) = pickWindow(at: mid) else {
            debugLog("选区中心无窗口")
            return
        }
        // 权限真实判据：无屏幕录制权限时读不到其他窗口的标题
        if noPermission {
            debugLog("无屏幕录制权限（窗口标题不可见），请求授权")
            CGRequestScreenCaptureAccess()
            render()
            return
        }
        let primaryMaxY = NSScreen.screens.first?.frame.maxY ?? 0
        // 鼠标按下/松开的点在文字行中部，垂直方向外扩整行高度，
        // 保证首行和末行的字形完整（否则上下截断导致识别不全）
        let padX: CGFloat = 4
        let padY: CGFloat = 14
        let minX = min(start.x, loc.x) - padX
        let maxX = max(start.x, loc.x) + padX
        let minY = min(start.y, loc.y) - padY
        let maxY = max(start.y, loc.y) + padY
        let global = CGRect(x: minX, y: primaryMaxY - maxY, width: maxX - minX, height: maxY - minY)

        isPicking = true
        DispatchQueue.global(qos: .userInitiated).async { [weak self] in
            let result = ocrRegion(global: global, windowID: winID)
            DispatchQueue.main.async {
                guard let self = self else { return }
                self.isPicking = false
                guard self.pickMode else { return }
                if result.captured {
                    self.captureFailStreak = 0
                } else {
                    self.captureFailStreak += 1
                }
                if self.captureFailStreak == 3 { self.render() }
                guard result.captured, let text = result.text, !text.isEmpty else {
                    debugLog("选区无文字")
                    return
                }
                debugLog("选区OCR: \(text.prefix(60))")
                self.matchAndShow(text)
            }
        }
    }

    /// 找到鼠标所在的最上层普通窗口（排除自身卡片），返回窗口 ID 与宽度
    /// 找到鼠标所在的最上层普通窗口（排除自身卡片），返回窗口 ID 与宽度
    /// 同时用"能否读到窗口标题"探测屏幕录制权限（无权限时 kCGWindowName 全为空）
    private func pickWindow(at cocoaPoint: NSPoint) -> (CGWindowID, CGFloat)? {
        guard let list = CGWindowListCopyWindowInfo([.optionOnScreenOnly], kCGNullWindowID) as? [[String: Any]] else { return nil }
        let primaryMaxY = NSScreen.screens.first?.frame.maxY ?? 0
        let pt = CGPoint(x: cocoaPoint.x, y: primaryMaxY - cocoaPoint.y)
        var sawAnyTitle = false
        for w in list {
            if (w["kCGWindowName"] as? String)?.isEmpty == false { sawAnyTitle = true }
            guard (w["kCGWindowLayer"] as? Int) == 0 else { continue }   // 普通窗口层
            guard let ownerPID = w["kCGWindowOwnerPID"] as? Int, ownerPID != ProcessInfo.processInfo.processIdentifier else { continue }
            guard let dict = w["kCGWindowBounds"] as? [String: Any],
                  let b = CGRect(dictionaryRepresentation: dict as CFDictionary) else { continue }
            if b.contains(pt) {
                noPermission = !sawAnyTitle
                return (CGWindowID((w["kCGWindowNumber"] as? Int) ?? 0), b.width)
            }
        }
        noPermission = !sawAnyTitle
        return nil
    }

    /// 用识别到的文字与所有题干做模糊匹配，命中则切换到该题
    private func matchAndShow(_ picked: String) {
        let target = normalizedText(picked)
        guard target.count >= 4 else { return }
        let snapshot = rows
        DispatchQueue.global(qos: .userInitiated).async { [weak self] in
            var best = -1
            var bestScore = 0.0
            for (i, row) in snapshot.enumerated() {
                // 题干取第一个非空单元格（部分表格首列为空）
                let stem = normalizedText(row.cells.first(where: { !$0.isEmpty }) ?? "")
                guard !stem.isEmpty else { continue }
                let score = (stem.contains(target) || target.contains(stem)) ? 1.0 : textSimilarity(target, stem)
                if score > bestScore { bestScore = score; best = i }
            }
            DispatchQueue.main.async {
                guard let self = self, self.pickMode else { return }
                if best < 0 {
                    debugLog("匹配失败 rows=\(snapshot.count)")
                    return
                }
                debugLog("匹配 best=\(best) score=\(String(format: "%.2f", bestScore)) applied=\(bestScore >= 0.72 && self.index != best)")
                guard bestScore >= 0.72, self.index != best else { return }
                self.index = best
                self.render()
            }
        }
    }

    @objc private func togglePickMode() {
        pickMode.toggle()
        if pickMode {
            startPickTimer()
            captureFailStreak = 0
        } else {
            stopPickTimer()
        }
        if !rows.isEmpty { render() }
    }

    @objc private func screenChanged() {
        // 显示器变化时保持用户锚点，仅按可视区钳制位置
        applyAnchor(height: panel?.frame.height ?? 100, animate: true)
    }

    @objc private func resetAnchor() {
        resetAnchorToScreenCorner()
        UserDefaults.standard.set(Double(anchorRight), forKey: "anchorRight")
        UserDefaults.standard.set(Double(anchorTop), forKey: "anchorTop")
        applyAnchor(height: panel?.frame.height ?? 100, animate: true)
    }

    @objc private func windowResized() {
        guard !isFixingFrame, let panel = panel, panel.isVisible else { return }
        // 宽度被意外拉伸时异步纠正（不能在布局通知里同步 setFrame，会递归爆栈）
        if abs(panel.frame.width - cardWidth) > 0.5 {
            isFixingFrame = true
            DispatchQueue.main.async { [weak self] in
                guard let self = self else { return }
                self.applyAnchor(height: self.panel?.frame.height ?? 100)
                self.isFixingFrame = false
            }
        }
    }

    /// 双屏时始终定位到 Mac 内建屏（合盖时回退主屏）
    private func preferredScreen() -> NSScreen? {
        let builtin = NSScreen.screens.first { scr in
            guard let num = scr.deviceDescription[NSDeviceDescriptionKey("NSScreenNumber")] as? NSNumber else { return false }
            return CGDisplayIsBuiltin(num.uint32Value) != 0
        }
        return builtin ?? NSScreen.screens.first ?? NSScreen.main
    }

    private func preferredVisibleFrame() -> NSRect? { preferredScreen()?.visibleFrame }

    /// 按最终窗口高度手动摆放子视图（顶部图标条、题干、底部信息栏、拖拽热区铺满）
    private func layoutContent(windowHeight h: CGFloat) {
        guard let effect = effectView else { return }
        let textH = max(contentTextHeight, 20)
        let footerH = max(contentFooterHeight, 12)
        let topReserve: CGFloat = 38   // 顶部图标条区域
        label.frame = NSRect(x: 20, y: h - topReserve - textH, width: cardWidth - 40, height: textH)
        footer.frame = NSRect(x: 20, y: 12, width: cardWidth - 40, height: footerH)
        pickBtn.frame = NSRect(x: cardWidth - 14 - 26, y: h - 36, width: 26, height: 26)
        capBtn.frame = NSRect(x: cardWidth - 14 - 26 - 4 - 26, y: h - 36, width: 26, height: 26)
        for sub in effect.subviews where sub is CardView { sub.frame = effect.bounds }
        updateModeIcons()
    }

    private func makeIconButton(symbol: String, tooltip: String, action: Selector) -> NSButton {
        let cfg = NSImage.SymbolConfiguration(pointSize: 12, weight: .medium)
        let btn = NSButton()
        btn.image = NSImage(systemSymbolName: symbol, accessibilityDescription: nil)?
            .withSymbolConfiguration(cfg)
        btn.isBordered = false
        btn.imagePosition = .imageOnly
        btn.wantsLayer = true
        btn.layer?.cornerRadius = 13
        btn.layer?.masksToBounds = true
        btn.imageScaling = .scaleProportionallyDown
        btn.toolTip = tooltip
        btn.target = self
        btn.action = action
        return btn
    }

    /// 图标状态：激活时亮色圆底 + 白色图标，未激活时半透明灰底
    private func updateModeIcons() {
        guard let pickBtn = pickBtn, let capBtn = capBtn else { return }
        let idleBg = lightBackground
            ? NSColor.black.withAlphaComponent(0.10)
            : NSColor.white.withAlphaComponent(0.14)

        pickBtn.layer?.backgroundColor = (pickMode ? NSColor.systemBlue : idleBg).cgColor
        pickBtn.contentTintColor = pickMode ? .white : .secondaryLabelColor
        let pickSymbol = pickMode ? "text.viewfinder" : "viewfinder"
        let cfg = NSImage.SymbolConfiguration(pointSize: 12, weight: .medium)
        pickBtn.image = NSImage(systemSymbolName: pickSymbol, accessibilityDescription: "选中取题")?
            .withSymbolConfiguration(cfg)

        capBtn.layer?.backgroundColor = (blockCapture
            ? NSColor.systemGreen.withAlphaComponent(0.85)
            : idleBg).cgColor
        capBtn.contentTintColor = blockCapture ? .white : .secondaryLabelColor
        capBtn.image = NSImage(systemSymbolName: blockCapture ? "eye.slash" : "eye", accessibilityDescription: "防拍")?
            .withSymbolConfiguration(cfg)
    }

    /// 重置锚点：内建屏右上角
    private func resetAnchorToScreenCorner() {
        guard let vf = preferredVisibleFrame() else { return }
        anchorRight = vf.maxX - 12
        anchorTop = vf.maxY - 12
    }

    /// 以锚点右上角为基准摆放窗口，向下伸展；锚点或卡片过高时自动钳制在屏幕内
    private func applyAnchor(height: CGFloat, animate: Bool = false) {
        guard let vf = preferredVisibleFrame(), let panel = panel else { return }
        let h = min(height, vf.height - 24)
        let right = min(max(anchorRight, vf.minX + cardWidth + 4), vf.maxX - 4)
        let top = min(max(anchorTop, vf.minY + h + 4), vf.maxY - 4)
        var f = panel.frame
        f.size.width = cardWidth
        f.size.height = h
        f.origin.x = right - cardWidth
        f.origin.y = top - h
        panel.setFrame(f, display: true, animate: animate)
        layoutContent(windowHeight: panel.frame.height)
    }

    func application(_ application: NSApplication, open urls: [URL]) {
        if let url = urls.first(where: { $0.pathExtension.lowercased() == "xlsx" }) {
            loadFile(path: url.path)
        }
    }

    // MARK: UI

    private func softFont(_ size: CGFloat, _ weight: NSFont.Weight = .medium) -> NSFont {
        let base = NSFont.systemFont(ofSize: size, weight: weight)
        return base.fontDescriptor.withDesign(.rounded).flatMap { NSFont(descriptor: $0, size: size) } ?? base
    }

    private func buildUI() {
        let width = cardWidth
        let effect = NSVisualEffectView(frame: NSRect(x: 0, y: 0, width: width, height: 120))
        effect.material = .hudWindow
        effect.blendingMode = .behindWindow
        effect.state = .active
        effect.wantsLayer = true
        effect.layer?.cornerRadius = 20
        effect.layer?.masksToBounds = true

        label = NSTextField(wrappingLabelWithString: "")
        label.alignment = .center
        label.font = softFont(fontSize)
        label.textColor = .white
        label.isEditable = false
        label.isSelectable = false
        label.drawsBackground = false

        footer = NSTextField(labelWithString: "")
        footer.alignment = .center
        footer.font = softFont(11, .regular)
        footer.textColor = NSColor.white.withAlphaComponent(0.45)

        effect.addSubview(label)
        effect.addSubview(footer)

        let card = CardView(frame: effect.bounds)
        card.autoresizingMask = [.width, .height]
        card.onFileDropped = { [weak self] url in self?.loadFile(path: url.path) }
        effect.addSubview(card)

        // 右上角模式图标：点击切换 取题 / 防拍
        pickBtn = makeIconButton(symbol: "viewfinder", tooltip: "选中取题 开/关", action: #selector(togglePickMode))
        capBtn = makeIconButton(symbol: "eye.slash", tooltip: "防拍 开/关", action: #selector(toggleCaptureBlock))
        effect.addSubview(pickBtn)
        effect.addSubview(capBtn)

        // 手动布局（不使用 Auto Layout）：Auto Layout 的"窗口适配内容"机制
        // 会在 setFrame 后改写窗口尺寸，导致卡片被文本撑宽、右边出屏。
        // 窗口尺寸只由 applyAnchor 决定，子视图 frame 在 layoutContent 中手动计算。

        panel = FloatPanel(contentRect: effect.frame, styleMask: [.borderless, .nonactivatingPanel], backing: .buffered, defer: false)
        effectView = effect
        panel.level = .screenSaver
        panel.isOpaque = false
        panel.backgroundColor = .clear
        panel.hasShadow = true
        panel.isMovableByWindowBackground = true
        panel.sharingType = blockCapture ? .none : .readOnly   // 屏蔽录屏 / 投屏 / 屏幕共享采集（可开关）
        panel.collectionBehavior = [.canJoinAllSpaces, .fullScreenAuxiliary]
        panel.contentView = effect
        panel.delegate = self
        panel.contentView?.menu = buildMenu()

        let click = NSClickGestureRecognizer(target: self, action: #selector(togglePause))
        effect.addGestureRecognizer(click)

        // 初始化：从未拖动过 → 内建屏右上角；拖动过 → 恢复上次锚点位置
        panel.orderFrontRegardless()
        if UserDefaults.standard.object(forKey: "anchorRight") != nil,
           UserDefaults.standard.object(forKey: "anchorTop") != nil {
            anchorRight = CGFloat(UserDefaults.standard.double(forKey: "anchorRight"))
            anchorTop = CGFloat(UserDefaults.standard.double(forKey: "anchorTop"))
        } else {
            resetAnchorToScreenCorner()
        }
        applyAnchor(height: panel.frame.height)
    }

    private func buildMenu() -> NSMenu {
        let menu = NSMenu()
        playMenuItem = NSMenuItem(title: "暂停", action: #selector(togglePause), keyEquivalent: " ")
        playMenuItem.target = self
        menu.addItem(playMenuItem)

        let prev = NSMenuItem(title: "上一条", action: #selector(prevRow), keyEquivalent: "")
        prev.target = self
        menu.addItem(prev)
        let next = NSMenuItem(title: "下一条", action: #selector(nextRow), keyEquivalent: "")
        next.target = self
        menu.addItem(next)
        menu.addItem(.separator())

        let open = NSMenuItem(title: "打开 Excel 文件…", action: #selector(openFile), keyEquivalent: "o")
        open.target = self
        menu.addItem(open)
        let resetPos = NSMenuItem(title: "复位到屏幕右上角", action: #selector(resetAnchor), keyEquivalent: "")
        resetPos.target = self
        menu.addItem(resetPos)
        menu.addItem(.separator())

        let intervalMenu = NSMenu()
        for sec in [1.0, 2.0, 3.0, 5.0, 10.0, 30.0] {
            let item = NSMenuItem(title: "\(Int(sec)) 秒", action: #selector(changeInterval(_:)), keyEquivalent: "")
            item.target = self
            item.tag = Int(sec)
            intervalMenu.addItem(item)
            intervalItems.append(item)
        }
        let intervalParent = NSMenuItem(title: "播放间隔", action: nil, keyEquivalent: "")
        intervalParent.submenu = intervalMenu
        menu.addItem(intervalParent)

        let fontMenu = NSMenu()
        for size in [16.0, 20.0, 24.0, 30.0, 38.0] {
            let item = NSMenuItem(title: "\(Int(size)) 号", action: #selector(changeFontSize(_:)), keyEquivalent: "")
            item.target = self
            item.tag = Int(size)
            fontMenu.addItem(item)
            fontSizeItems.append(item)
        }
        let fontParent = NSMenuItem(title: "字号", action: nil, keyEquivalent: "")
        fontParent.submenu = fontMenu
        menu.addItem(fontParent)
        menu.addItem(.separator())

        captureMenuItem = NSMenuItem(title: "屏蔽录屏/截图", action: #selector(toggleCaptureBlock), keyEquivalent: "")
        captureMenuItem.target = self
        menu.addItem(captureMenuItem)

        pickMenuItem = NSMenuItem(title: "选中取题", action: #selector(togglePickMode), keyEquivalent: "")
        pickMenuItem.target = self
        menu.addItem(pickMenuItem)
        menu.addItem(.separator())

        let quit = NSMenuItem(title: "退出", action: #selector(quitApp), keyEquivalent: "q")
        quit.target = self
        menu.addItem(quit)

        menu.delegate = self
        return menu
    }

    // MARK: 位置锚点

    func windowDidMove(_ notification: Notification) {
        syncAnchorFromFrame()
    }

    /// 窗口移动后（含用户拖动），以当前位置的右上角为新锚点，并钳制在可视区内
    private func syncAnchorFromFrame() {
        guard let panel = panel else { return }
        var f = panel.frame
        if let vf = preferredVisibleFrame() {
            f.origin.x = min(max(f.origin.x, vf.minX + 4), vf.maxX - f.width - 4)
            f.origin.y = min(max(f.origin.y, vf.minY + 4), vf.maxY - f.height - 4)
            if f != panel.frame { panel.setFrame(f, display: true, animate: false) }
        }
        if abs(panel.frame.width - cardWidth) > 0.5 {
            var fixed = panel.frame
            fixed.size.width = cardWidth
            panel.setFrame(fixed, display: true, animate: true)
        }
        anchorRight = panel.frame.maxX
        anchorTop = panel.frame.maxY
        UserDefaults.standard.set(Double(anchorRight), forKey: "anchorRight")
        UserDefaults.standard.set(Double(anchorTop), forKey: "anchorTop")
    }

    // MARK: 数据加载

    private func loadFile(path: String) {
        DispatchQueue.global(qos: .userInitiated).async { [weak self] in
            do {
                let rows = try XlsxParser.parse(path: path)
                let mtime = (try? FileManager.default.attributesOfItem(atPath: path))?[.modificationDate] as? Date
                DispatchQueue.main.async {
                    guard let self = self else { return }
                    self.rows = rows
                    self.index = 0
                    self.paused = false
                    self.filePath = path
                    self.fileMTime = mtime
                    UserDefaults.standard.set(path, forKey: "lastFile")
                    self.render()
                }
            } catch {
                DispatchQueue.main.async {
                    self?.label.stringValue = "读取失败：\((error as NSError).userInfo[NSLocalizedDescriptionKey] as? String ?? "未知错误")"
                }
            }
        }
    }

    private func render() {
        guard !rows.isEmpty else {
            label.stringValue = "表格为空"
            footer.stringValue = ""
            return
        }
        index = (index + rows.count) % rows.count
        let row = rows[index]
        let parsed = buildRowAttr(row.cells)
        label.attributedStringValue = parsed.attr
        label.font = softFont(fontSize)
        let name = filePath.map { ( $0 as NSString).lastPathComponent } ?? ""
        var parts = ["\(index + 1)/\(rows.count)"]
        if !parsed.meta.isEmpty, !pickMode { parts.append(parsed.meta) }
        parts.append(row.sheet)
        if pickMode {
            parts.append(noPermission ? "无录屏权限" : (captureFailStreak >= 3 ? "取题不可用" : "选中取题"))
        } else {
            parts.append(name)
        }
        if paused { parts.append("已暂停") }
        parts.append(blockCapture ? "防拍" : "可拍")
        footer.stringValue = parts.joined(separator: "  ·  ")
        playMenuItem.title = paused ? "继续" : "暂停"
        resizeToFit()
    }

    // MARK: 行排版：识别题干 / 难度 / 分类 / 选项 / 答案（含判断题）

    private func judgeAnswer(_ raw: String) -> Bool? {
        switch raw.trimmingCharacters(in: .whitespaces).lowercased() {
        case "对", "正确", "是", "√", "t", "true", "y", "yes": return true
        case "错", "错误", "否", "×", "x", "f", "false", "n", "no": return false
        default: return nil
        }
    }

    private func buildRowAttr(_ cells: [String]) -> (attr: NSAttributedString, meta: String) {
        let nonEmpty = cells.filter { !$0.isEmpty }

        // 找答案列：单个 A~F 字母（单选）或 对/错/√/× 等（判断）
        var ansIdx = -1
        var judgeMode = false
        for (i, c) in nonEmpty.enumerated() where i > 0 {
            let t = c.trimmingCharacters(in: .whitespaces)
            let up = t.uppercased()
            if judgeAnswer(t) != nil {
                // T/F/Y/N 与选项字母有歧义：判断题答案后面通常不再有选项
                if ["T", "F", "Y", "N"].contains(up) {
                    if nonEmpty.count - i - 1 < 2 { ansIdx = i; judgeMode = true; break }
                } else {
                    ansIdx = i; judgeMode = true; break
                }
            } else if up.range(of: "^[A-F]{1,6}$", options: .regularExpression) != nil {
                ansIdx = i; break   // 单选 A~F 或多选如 ABD
            }
        }

        let basePara = NSMutableParagraphStyle()
        basePara.lineSpacing = 3
        basePara.paragraphSpacing = 5

        if ansIdx > 0 {
            let metaParts = Array(nonEmpty[1..<ansIdx])
            let optPara = NSMutableParagraphStyle()
            optPara.lineSpacing = 3
            optPara.paragraphSpacing = 3
            optPara.headIndent = 16

            if judgeMode {
                let yes = judgeAnswer(nonEmpty[ansIdx])!
                let s = NSMutableAttributedString()
                s.append(NSAttributedString(string: nonEmpty[0], attributes: [
                    .font: softFont(fontSize), .foregroundColor: fgColor, .paragraphStyle: basePara,
                ]))
                for (text, correct) in [("正确", yes), ("错误", !yes)] {
                    s.append(NSAttributedString(string: "\n\(correct ? "✓ " : "")\(text)", attributes: [
                        .font: softFont(max(fontSize - 2, 12), .regular),
                        .foregroundColor: correct ? accentColor : fgColor,
                        .paragraphStyle: optPara,
                    ]))
                }
                let ap = NSMutableParagraphStyle()
                ap.paragraphSpacing = 0
                ap.paragraphSpacingBefore = 7
                s.append(NSAttributedString(string: "\n答案：\(yes ? "正确" : "错误")", attributes: [
                    .font: softFont(max(fontSize - 5, 12), .semibold),
                    .foregroundColor: accentColor,
                    .paragraphStyle: ap,
                ]))
                return (s, metaParts.joined(separator: " · "))
            }

            let options = Array(nonEmpty[(ansIdx + 1)...])
            let answer = nonEmpty[ansIdx].trimmingCharacters(in: .whitespaces).uppercased()
            let answerDisplay = answer.map(String.init).joined(separator: " ")
            let letters = Array("ABCDEF")

            let s = NSMutableAttributedString()
            s.append(NSAttributedString(string: nonEmpty[0], attributes: [
                .font: softFont(fontSize), .foregroundColor: fgColor, .paragraphStyle: basePara,
            ]))
            for (j, opt) in options.enumerated() {
                let letter = j < letters.count ? letters[j] : "?"
                let p = NSMutableParagraphStyle()
                p.lineSpacing = 3
                p.paragraphSpacing = 3
                p.headIndent = 16
                let correct = answer.contains(letter)
                var text = "\(letter). \(opt)"
                if opt.range(of: "^[A-Fa-f][.、．)]", options: .regularExpression) != nil { text = opt }
                if correct { text = "✓ " + text }
                s.append(NSAttributedString(string: "\n" + text, attributes: [
                    .font: softFont(max(fontSize - 2, 12), .regular),
                    .foregroundColor: correct ? accentColor : fgColor,
                    .paragraphStyle: p,
                ]))
            }
            let ap = NSMutableParagraphStyle()
            ap.paragraphSpacing = 0
            ap.paragraphSpacingBefore = 7
            s.append(NSAttributedString(string: "\n答案：\(answerDisplay)", attributes: [
                .font: softFont(max(fontSize - 5, 12), .semibold),
                .foregroundColor: accentColor,
                .paragraphStyle: ap,
            ]))
            return (s, metaParts.joined(separator: " · "))
        }

        // 未识别出选项/判断：末列为长文本时按简答题排版，否则回退单行
        let last = nonEmpty.last ?? ""
        if nonEmpty.count >= 2, last.count >= 6, last != nonEmpty[0] {
            let metaParts = Array(nonEmpty[1..<(nonEmpty.count - 1)])
            let ap = NSMutableParagraphStyle()
            ap.lineSpacing = 4
            ap.paragraphSpacingBefore = 7
            let answer = last.count > 300 ? String(last.prefix(300)) + "…" : last
            let s = NSMutableAttributedString()
            s.append(NSAttributedString(string: nonEmpty[0], attributes: [
                .font: softFont(fontSize), .foregroundColor: fgColor, .paragraphStyle: basePara,
            ]))
            s.append(NSAttributedString(string: "\n答案：\(answer)", attributes: [
                .font: softFont(max(fontSize - 4, 12), .regular),
                .foregroundColor: accentColor,
                .paragraphStyle: ap,
            ]))
            return (s, metaParts.joined(separator: " · "))
        }

        // 普通数据行：保持单行样式
        return (NSAttributedString(string: nonEmpty.joined(separator: "　|　"), attributes: [
            .font: softFont(fontSize), .foregroundColor: fgColor, .paragraphStyle: basePara,
        ]), "")
    }

    private func resizeToFit() {
        let textHeight = label.attributedStringValue.boundingRect(
            with: NSSize(width: cardWidth - 40, height: .greatestFiniteMagnitude),
            options: [.usesLineFragmentOrigin]
        ).height
        contentTextHeight = ceil(textHeight) + 4
        contentFooterHeight = ceil(footer.intrinsicContentSize.height)
        let height = max(76, contentTextHeight + 6 + contentFooterHeight + 28 + 38)
        applyAnchor(height: height)   // 切题时以锚点右上角为准，向下伸缩
    }

    private func startTimer() {
        timer?.invalidate()
        timer = Timer.scheduledTimer(withTimeInterval: interval, repeats: true) { [weak self] _ in
            DispatchQueue.main.async { self?.tick() }
        }
    }

    // MARK: 背景明暗自适应

    private func startContrastTimer() {
        contrastTimer?.invalidate()
        contrastTimer = Timer.scheduledTimer(withTimeInterval: 2.0, repeats: true) { [weak self] _ in
            DispatchQueue.main.async { self?.updateContrast() }
        }
    }

    private func updateContrast() {
        guard let view = panel?.contentView, view.window != nil else { return }
        let screenMaxY = NSScreen.screens.first?.frame.maxY ?? 0
        let f = panel.frame
        let cgRect = CGRect(x: f.minX, y: screenMaxY - f.maxY, width: f.width, height: f.height)
        guard let image = legacyCGWindowListCreateImage(cgRect, kWindowListOptionOnScreenBelowWindow, UInt32(panel.windowNumber), kWindowImageBestResolution) else { return }
        guard let ctx = CGContext(data: nil, width: 1, height: 1, bitsPerComponent: 8, bytesPerRow: 4,
                                  space: CGColorSpaceCreateDeviceRGB(),
                                  bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue) else { return }
        ctx.interpolationQuality = .medium
        ctx.draw(image, in: CGRect(x: 0, y: 0, width: 1, height: 1))
        guard let data = ctx.data else { return }
        let px = data.bindMemory(to: UInt8.self, capacity: 4)
        let lum = (0.299 * Double(px[0]) + 0.587 * Double(px[1]) + 0.114 * Double(px[2])) / 255.0
        // 滞回阈值防抖：0.7 亮→暗色卡片，0.55 暗→亮色卡片
        let wantLight = lightBackground ? lum > 0.55 : lum > 0.7
        if wantLight != lightBackground {
            lightBackground = wantLight
            applyTheme()
        }
    }

    private func applyTheme() {
        if lightBackground {
            effectView.material = .popover
            effectView.appearance = NSAppearance(named: .vibrantLight)
            fgColor = NSColor(red: 0.26, green: 0.25, blue: 0.24, alpha: 1)
            accentColor = NSColor(red: 0.23, green: 0.44, blue: 0.83, alpha: 1)
            label.textColor = fgColor
            footer.textColor = NSColor.black.withAlphaComponent(0.42)
            updateModeIcons()
        } else {
            effectView.material = .hudWindow
            effectView.appearance = NSAppearance(named: .vibrantDark)
            fgColor = NSColor(red: 0.95, green: 0.94, blue: 0.92, alpha: 1)
            accentColor = NSColor(red: 0.56, green: 0.73, blue: 0.95, alpha: 1)
            label.textColor = fgColor
            footer.textColor = NSColor.white.withAlphaComponent(0.5)
            updateModeIcons()
        }
    }

    @objc private func tick() {
        if let path = filePath,
           let mtime = (try? FileManager.default.attributesOfItem(atPath: path))?[.modificationDate] as? Date,
           mtime != fileMTime {
            loadFile(path: path)   // 文件被修改后自动重新加载
            return
        }
        guard !pickMode, !paused, !rows.isEmpty else { return }   // 取题模式下不自动轮播
        index += 1
        render()
    }

    // MARK: 动作

    @objc private func togglePause() {
        guard !rows.isEmpty else { return }
        paused.toggle()
        render()
    }
    @objc private func prevRow() { index -= 1; render() }
    @objc private func nextRow() { index += 1; render() }

    @objc private func openFile() {
        NSApp.activate(ignoringOtherApps: true)
        let open = NSOpenPanel()
        open.allowedContentTypes = {
            if #available(macOS 11.0, *) { return [UTType(filenameExtension: "xlsx") ?? .data] }
            return [.data]
        }()
        open.canChooseDirectories = false
        open.begin { [weak self] resp in
            if resp == .OK, let url = open.url { self?.loadFile(path: url.path) }
        }
    }

    @objc private func changeInterval(_ sender: NSMenuItem) {
        interval = TimeInterval(sender.tag)
        startTimer()
    }
    @objc private func changeFontSize(_ sender: NSMenuItem) {
        fontSize = CGFloat(sender.tag)
        render()
    }
    @objc private func toggleCaptureBlock() {
        blockCapture.toggle()
        panel.sharingType = blockCapture ? .none : .readOnly
        if !rows.isEmpty { render() }
    }
    @objc private func quitApp() { NSApp.terminate(nil) }
}

extension AppDelegate: NSMenuDelegate {
    func menuNeedsUpdate(_ menu: NSMenu) {
        for item in intervalItems { item.state = item.tag == Int(interval) ? .on : .off }
        for item in fontSizeItems { item.state = item.tag == Int(fontSize) ? .on : .off }
        captureMenuItem.state = blockCapture ? .on : .off
        pickMenuItem.state = pickMode ? .on : .off
    }
}

// MARK: - 入口

let cliArgs = CommandLine.arguments
if cliArgs.count >= 3, cliArgs[1] == "--parse-test" {
    do {
        let rows = try XlsxParser.parse(path: cliArgs[2])
        for row in rows { print("[\(row.sheet)] " + row.cells.joined(separator: " | ")) }
        exit(0)
    } catch {
        print("ERROR: \((error as NSError).userInfo[NSLocalizedDescriptionKey] as? String ?? "\(error)")")
        exit(1)
    }
}

MainActor.assumeIsolated {
    let app = NSApplication.shared
    let delegate = AppDelegate()
    app.delegate = delegate
    app.setActivationPolicy(.accessory)
    app.run()
}
