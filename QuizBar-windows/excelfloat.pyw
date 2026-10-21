# -*- coding: utf-8 -*-
"""ExcelFloat Windows 版：Excel 悬浮轮播卡片
- 置顶悬浮（高于一切应用），半透明卡片，逐行循环播放
- 左键点击 = 暂停/继续；按住拖动 = 移动位置；右键 = 菜单
- 屏蔽录屏 / 投屏 / 屏幕共享（Windows 10 2004+）
- 选中取题：任意应用里划选题目文字，松开后 OCR 识别并匹配题库
- 单选/多选/判断/简答题智能排版，背景明暗自适应
依赖：pip install openpyxl pillow winsdk
"""
import os
import re
import io
import sys
import json
import ctypes
import datetime
import threading
import tkinter as tk
from tkinter import filedialog

try:
    import openpyxl
except ImportError:
    openpyxl = None

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
WDA_EXCLUDEFROMCAPTURE = 0x11
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x80

user32 = ctypes.windll.user32 if sys.platform == "win32" else None
gdi32 = ctypes.windll.gdi32 if sys.platform == "win32" else None
kernel32 = ctypes.windll.kernel32 if sys.platform == "win32" else None


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", ctypes.c_uint32), ("biWidth", ctypes.c_int32),
                ("biHeight", ctypes.c_int32), ("biPlanes", ctypes.c_uint16),
                ("biBitCount", ctypes.c_uint16), ("biCompression", ctypes.c_uint32),
                ("biSizeImage", ctypes.c_uint32), ("biXPelsPerMeter", ctypes.c_int32),
                ("biYPelsPerMeter", ctypes.c_int32), ("biClrUsed", ctypes.c_uint32),
                ("biClrImportant", ctypes.c_uint32)]


def load_config():
    cfg = {"file": "", "x": -1, "y": -1, "interval": 3, "font": 18,
           "block_capture": True, "pick_mode": False}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg.update(json.load(f))
    except Exception:
        pass
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False)
    except Exception:
        pass


def fmt_cell(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        if v.hour or v.minute or v.second:
            return v.strftime("%Y/%m/%d %H:%M")
        return v.strftime("%Y/%m/%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def normalized_text(s):
    out = []
    for c in s:
        o = ord(c)
        if (0x4E00 <= o <= 0x9FFF) or (0x30 <= o <= 0x39) or (0x41 <= o <= 0x5A) or (0x61 <= o <= 0x7A):
            out.append(c.lower())
    return "".join(out)


def text_similarity(a, b):
    x, y = list(a), list(b)
    if not x or not y:
        return 0.0
    prev = list(range(len(y) + 1))
    cur = [0] * (len(y) + 1)
    for i in range(1, len(x) + 1):
        cur[0] = i
        for j in range(1, len(y) + 1):
            cost = 0 if x[i - 1] == y[j - 1] else 1
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
        prev, cur = cur, prev
    return 1 - prev[len(y)] / max(len(x), len(y))


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.rows = []              # [(sheet, cells)]
        self.index = 0
        self.paused = False
        self.path = ""
        self.mtime = None
        self.hwnd = 0
        self.capture_blocked = False
        self.light_bg = False
        self.accent = "#8fb4f2"

        self.interval_var = tk.DoubleVar(value=self.cfg["interval"])
        self.font_var = tk.IntVar(value=self.cfg["font"])
        self.cap_var = tk.BooleanVar(value=bool(self.cfg.get("block_capture", True)))
        self.pick_var = tk.BooleanVar(value=bool(self.cfg.get("pick_mode", False)))

        self.pick_timer = None
        self.is_picking = False
        self.sel_start = None
        self.last_pick_text = ""
        self.capture_fail_streak = 0
        self.ocr_engine = None
        self.ocr_checked = False

        self.width = 460
        self.bg = "#24262c"
        self.fg = "#f0eee9"

        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.88)
        root.configure(bg=self.bg)

        self.label = tk.Label(root, text="", font=("Microsoft YaHei UI", self.cfg["font"]),
                              fg=self.fg, bg=self.bg, wraplength=self.width - 44,
                              justify="left", anchor="w")
        self.label.pack(padx=22, pady=(14, 2), fill="x")

        self.opts = tk.Label(root, text="", font=("Microsoft YaHei UI", max(self.cfg["font"] - 2, 10)),
                             fg=self.fg, bg=self.bg, wraplength=self.width - 60,
                             justify="left", anchor="w")
        self.opts.pack(padx=38, pady=0, fill="x")

        self.ans = tk.Label(root, text="", font=("Microsoft YaHei UI", max(self.cfg["font"] - 4, 10), "bold"),
                            fg="#8fb4f2", bg=self.bg, justify="left", anchor="w",
                            wraplength=self.width - 44)
        self.ans.pack(padx=22, pady=(6, 2), fill="x")

        self.footer = tk.Label(root, text="", font=("Microsoft YaHei UI", 9),
                               fg="#9d9a94", bg=self.bg)
        self.footer.pack(padx=22, pady=(0, 10), fill="x")

        # 右上角模式图标（点击切换）
        self.pick_icon = tk.Label(root, text="🔍", font=("Segoe UI Emoji", 12),
                                  bg=self.bg, fg=self.fg, cursor="hand2")
        self.pick_icon.place(relx=1.0, x=-30, y=6, anchor="ne")
        self.cap_icon = tk.Label(root, text="🙈", font=("Segoe UI Emoji", 12),
                                 bg=self.bg, fg=self.fg, cursor="hand2")
        self.cap_icon.place(relx=1.0, x=-62, y=6, anchor="ne")
        self.pick_icon.bind("<Button-1>", lambda e: self.toggle_pick())
        self.cap_icon.bind("<Button-1>", lambda e: self.toggle_capture())

        sw = root.winfo_screenwidth()
        x = self.cfg["x"] if self.cfg["x"] >= 0 else sw - self.width - 40
        y = self.cfg["y"] if self.cfg["y"] >= 0 else 80
        root.geometry("%dx120+%d+%d" % (self.width, int(x), int(y)))

        self.build_menu()
        root.bind("<Button-1>", self.on_press)
        root.bind("<B1-Motion>", self.on_drag)
        root.bind("<ButtonRelease-1>", self.on_release)
        root.bind("<Button-3>", self.show_menu)

        root.after(200, self.on_ready)

    # ---------- 初始化 ----------

    def on_ready(self):
        self.apply_toolwindow()
        if self.cap_var.get():
            self.apply_capture_block(True)
        sw = self.root.winfo_screenwidth()
        nx = min(max(self.root.winfo_x(), 0), sw - self.width - 4)
        ny = min(max(self.root.winfo_y(), 0), self.root.winfo_screenheight() - self.root.winfo_height() - 4)
        self.root.geometry("+%d+%d" % (nx, ny))

        arg_file = sys.argv[1] if len(sys.argv) > 1 else ""
        start_file = arg_file or self.cfg.get("file", "")
        if start_file and os.path.exists(start_file):
            self.load_file(start_file)
        else:
            self.label.config(text="把 .xlsx 拖到本程序图标上打开\n或右键 → 打开 Excel 文件…")
            self.footer.config(text="点击暂停/继续 · 拖动位置 · 右键更多选项")
        self.schedule()
        self.update_mode_icons()
        if self.pick_var.get():
            self.start_pick_timer()

    def apply_toolwindow(self):
        # 从 Alt+Tab / 任务栏隐藏
        try:
            self.hwnd = user32.GetParent(self.root.winfo_id()) or self.root.winfo_id()
            style = user32.GetWindowLongW(self.hwnd, GWL_EXSTYLE)
            user32.SetWindowLongW(self.hwnd, GWL_EXSTYLE, style | WS_EX_TOOLWINDOW)
        except Exception:
            pass

    def apply_capture_block(self, enable):
        if user32 is None:
            return
        try:
            if enable:
                self.capture_blocked = bool(user32.SetWindowDisplayAffinity(
                    self.hwnd, WDA_EXCLUDEFROMCAPTURE))
            else:
                user32.SetWindowDisplayAffinity(self.hwnd, 0x0)
                self.capture_blocked = False
        except Exception:
            self.capture_blocked = False
        self.cfg["block_capture"] = enable
        save_config(self.cfg)

    # ---------- 背景明暗自适应 ----------

    def contrast_tick(self):
        try:
            lum = self.sample_luminance()
            if lum is not None:
                want = True if lum > 0.7 else (False if lum < 0.55 else None)
                if want is not None and want != self.light_bg:
                    self.light_bg = want
                    self.apply_theme()
        finally:
            self.root.after(2000, self.contrast_tick)

    def sample_luminance(self):
        """采样卡片四周一圈屏幕内容的平均亮度（0~1）"""
        if user32 is None:
            return None
        try:
            s = 32
            x, y = self.root.winfo_x(), self.root.winfo_y()
            w, h = self.root.winfo_width(), self.root.winfo_height()
            m = 12
            pts = [(x + w // 2 - s // 2, y - m - s),          # 上
                   (x + w // 2 - s // 2, y + h + m),          # 下
                   (x - m - s, y + h // 2 - s // 2),          # 左
                   (x + w + m, y + h // 2 - s // 2)]          # 右
            bmi = BITMAPINFOHEADER()
            bmi.biSize, bmi.biWidth, bmi.biHeight = 40, s, -s
            bmi.biPlanes, bmi.biBitCount, bmi.biCompression = 1, 32, 0
            hdc = user32.GetDC(0)
            mem = gdi32.CreateCompatibleDC(hdc)
            bmp = gdi32.CreateCompatibleBitmap(hdc, s, s)
            old = gdi32.SelectObject(mem, bmp)
            total, count = 0.0, 0
            try:
                for (px, py) in pts:
                    if px < 0 or py < 0:
                        continue
                    gdi32.BitBlt(mem, 0, 0, s, s, hdc, px, py, 0x00CC0020)
                    gdi32.SelectObject(mem, old)
                    buf = ctypes.create_string_buffer(s * s * 4)
                    if gdi32.GetDIBits(mem, bmp, 0, s, buf, ctypes.byref(bmi), 0) != s:
                        continue
                    gdi32.SelectObject(mem, bmp)
                    for i in range(0, s * s * 4, 4):
                        total += 0.114 * buf[i] + 0.587 * buf[i + 1] + 0.299 * buf[i + 2]
                        count += 1
            finally:
                gdi32.SelectObject(mem, old)
                gdi32.DeleteObject(bmp)
                gdi32.DeleteDC(mem)
                user32.ReleaseDC(0, hdc)
            return total / count / 255.0 if count else None
        except Exception:
            return None

    def apply_theme(self):
        if self.light_bg:
            bg, fg, sub = "#f3f2f0", "#45423e", "#8f8b85"
            accent = "#3b6fd4"
        else:
            bg, fg, sub = "#24262c", "#f0eee9", "#9d9a94"
            accent = "#8fb4f2"
        self.accent = accent
        self.root.configure(bg=bg)
        self.label.configure(bg=bg, fg=fg)
        self.opts.configure(bg=bg, fg=fg)
        self.ans.configure(bg=bg, fg=accent)
        self.footer.configure(bg=bg, fg=sub)
        for icon in (self.pick_icon, self.cap_icon):
            icon.configure(bg=bg)
        self.update_mode_icons()

    def update_mode_icons(self):
        pick_on = bool(self.pick_var.get())
        cap_on = bool(self.cap_var.get())
        self.pick_icon.config(text="🎯" if pick_on else "🔍")
        self.cap_icon.config(text="🙈" if cap_on else "👁")
        self.pick_icon.config(fg="#8fb4f2" if pick_on else self.fg)
        self.cap_icon.config(fg="#7ec97e" if cap_on else self.fg)

    # ---------- 数据 ----------

    def load_file(self, path):
        if openpyxl is None:
            self.label.config(text="缺少依赖，请在命令行执行：\npip install openpyxl")
            return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            rows = []
            for ws in wb.worksheets:  # 多 sheet 依次合并轮播
                for row in ws.iter_rows(values_only=True):
                    cells = [fmt_cell(v) for v in row]
                    while cells and not cells[-1]:
                        cells.pop()
                    if cells:
                        rows.append((ws.title, cells))
            wb.close()
            self.rows = rows
            self.index = 0
            self.paused = False
            self.path = path
            self.mtime = os.path.getmtime(path)
            self.cfg["file"] = path
            save_config(self.cfg)
            self.render()
        except Exception as e:
            self.label.config(text="读取失败：%s" % e)

    def parse_row(self, cells):
        """识别单选/多选/判断题，未识别返回 None"""
        ne = [c for c in cells if c]
        ans_idx = -1
        judge = None
        for i in range(1, len(ne)):
            t = ne[i].strip()
            up = t.upper()
            j = self._judge(t)
            if j is not None:
                # T/F/Y/N 与选项字母有歧义：判断题答案后面通常不再有选项
                if up in ("T", "F", "Y", "N"):
                    if len(ne) - i - 1 < 2:
                        ans_idx, judge = i, j
                        break
                else:
                    ans_idx, judge = i, j
                    break
            if re.fullmatch(r"[A-F]{1,6}", up):  # 单选 A~F 或多选如 ABD
                ans_idx = i
                break
        if ans_idx <= 0:
            return None
        meta = ne[1:ans_idx]
        if judge is not None:
            return {"stem": ne[0], "opts": None, "judge": judge,
                    "ans": "正确" if judge else "错误", "meta": " · ".join(meta)}
        letter = ne[ans_idx].strip().upper()
        return {"stem": ne[0], "opts": ne[ans_idx + 1:], "judge": None,
                "ans": letter, "meta": " · ".join(meta)}

    @staticmethod
    def _judge(raw):
        t = raw.strip().lower()
        if t in ("对", "正确", "是", "√", "t", "true", "y", "yes"):
            return True
        if t in ("错", "错误", "否", "×", "x", "f", "false", "n", "no"):
            return False
        return None

    def parse_essay(self, cells):
        """简答题：无选项/判断答案，末列为长文本答案"""
        ne = [c for c in cells if c]
        if len(ne) < 2:
            return None
        last = ne[-1]
        if len(last) < 6 or last == ne[0]:
            return None
        return {"stem": ne[0], "opts": None, "judge": None,
                "ans": last[:300] + ("…" if len(last) > 300 else ""),
                "meta": " · ".join(ne[1:-1])}

    def render(self):
        if not self.rows:
            self.footer.config(text="")
            return
        self.index %= len(self.rows)
        sheet, cells = self.rows[self.index]
        parsed = self.parse_row(cells) or self.parse_essay(cells)
        self.label.config(font=("Microsoft YaHei UI", self.cfg["font"]))
        self.opts.config(font=("Microsoft YaHei UI", max(self.cfg["font"] - 2, 10)))
        self.ans.config(font=("Microsoft YaHei UI", max(self.cfg["font"] - 4, 10), "bold"))
        meta = ""
        if parsed:
            self.label.config(text=parsed["stem"])
            if parsed["opts"] is not None:  # 单选/多选题
                letters = "ABCDEF"
                lines = []
                for j, opt in enumerate(parsed["opts"]):
                    letter = letters[j] if j < len(letters) else "?"
                    text = opt if re.match(r"^[A-Fa-f][.、．)]", opt) else "%s. %s" % (letter, opt)
                    if letter in parsed["ans"]:
                        text = "✓ " + text
                    lines.append(text)
                self.opts.config(text="\n".join(lines))
            elif parsed["judge"] is not None:  # 判断题
                yes = parsed["judge"]
                l1 = ("✓ " if yes else "") + "正确"
                l2 = ("✓ " if not yes else "") + "错误"
                self.opts.config(text=l1 + "\n" + l2)
            else:  # 简答题
                self.opts.config(text="")
            ans_text = " ".join(parsed["ans"]) if parsed["opts"] is not None else parsed["ans"]
            self.ans.config(text="答案：%s" % ans_text)
            meta = parsed["meta"]
        else:
            self.label.config(text="  |  ".join(c for c in cells if c))
            self.opts.config(text="")
            self.ans.config(text="")
        name = os.path.basename(self.path)
        parts = ["%d/%d" % (self.index + 1, len(self.rows))]
        if meta:
            parts.append(meta)
        parts.append(sheet)
        if self.pick_var.get():
            parts.append("取题不可用" if self.capture_fail_streak >= 3 else "选中取题")
        else:
            parts.append(name)
        if self.paused:
            parts.append("已暂停")
        parts.append("防拍" if self.cap_var.get() else "可拍")
        self.footer.config(text="  ·  ".join(parts))
        self.menu.entryconfig(self.play_index, label="继续" if self.paused else "暂停")
        self.update_mode_icons()
        self.resize()

    def resize(self):
        self.root.update_idletasks()
        h = (self.label.winfo_reqheight() + self.opts.winfo_reqheight()
             + self.ans.winfo_reqheight() + self.footer.winfo_reqheight() + 26)
        self.root.geometry("%dx%d+%d+%d" % (self.width, h,
                                            self.root.winfo_x(), self.root.winfo_y()))

    # ---------- 轮播 ----------

    def schedule(self):
        self.root.after(int(self.cfg["interval"] * 1000), self.tick)

    def tick(self):
        try:
            if self.path and os.path.exists(self.path):
                m = os.path.getmtime(self.path)
                if m != self.mtime:
                    self.load_file(self.path)  # 文件被修改后自动重新加载
            if not self.pick_var.get() and not self.paused and self.rows:  # 取题模式下不自动轮播
                self.index += 1
                self.render()
        finally:
            self.root.attributes("-topmost", True)  # 防止置顶被抢占
            self.schedule()

    # ---------- 选中取题（OCR） ----------

    def start_pick_timer(self):
        self.root.after(100, self.pick_poll)

    def pick_poll(self):
        if not self.pick_var.get():
            return   # 取题关闭时停止轮询链，开启时由 toggle 重新启动
        try:
            if user32 is None or self.is_picking or not self.rows:
                return
            # 左键状态：0x01 VK_LBUTTON，高位表示按下
            left_down = user32.GetAsyncKeyState(0x01) & 0x8000
            px, py = self.root.winfo_pointerx(), self.root.winfo_pointery()
            if left_down:
                if self.sel_start is None:
                    self.sel_start = (px, py)
                return
            start = self.sel_start
            self.sel_start = None
            if not start:
                return
            if abs(px - start[0]) <= 12 and abs(py - start[1]) <= 6:
                return  # 点击不算划选
            x1, x2 = sorted((start[0], px))
            y1, y2 = sorted((start[1], py))
            bbox = (max(x1 - 4, 0), max(y1 - 14, 0), x2 + 4, y2 + 14)
            if bbox[2] - bbox[0] < 20 or bbox[3] - bbox[1] < 8:
                return
            self.is_picking = True
            threading.Thread(target=self._ocr_work, args=(bbox,), daemon=True).start()
        except Exception:
            pass
        finally:
            self.start_pick_timer()

    def _ocr_work(self, bbox):
        text = None
        try:
            from PIL import ImageGrab
            img = ImageGrab.grab(bbox=bbox)
            text = run_windows_ocr(img)
        except Exception as e:
            print("OCR error:", e)
            text = None
        self.root.after(0, lambda: self._ocr_done(text))

    def _ocr_done(self, text):
        self.is_picking = False
        if not self.pick_var.get():
            return
        if text is None:
            self.capture_fail_streak += 1
        else:
            self.capture_fail_streak = 0
        if self.capture_fail_streak == 3:
            self.render()
        if not text or text == self.last_pick_text:
            return
        self.last_pick_text = text
        self.match_and_show(text)

    def get_ocr_engine(self):
        if self.ocr_checked:
            return self.ocr_engine
        self.ocr_checked = True
        try:
            from winsdk.windows.media.ocr import OcrEngine
            from winsdk.windows.globalization import Language
            self.ocr_engine = (OcrEngine.try_create_from_language(Language("zh-Hans"))
                               or OcrEngine.try_create_from_user_profile_languages())
        except ImportError:
            self.ocr_engine = None
            self.label.config(text="选中取题需要依赖，请执行：\npip install winsdk pillow")
        except Exception:
            self.ocr_engine = None
        return self.ocr_engine

    def match_and_show(self, picked):
        target = normalized_text(picked)
        if len(target) < 4:
            return
        snapshot = self.rows
        best, best_score = -1, 0.0
        for i, (sheet, cells) in enumerate(snapshot):
            stem = normalized_text(next((c for c in cells if c), ""))
            if not stem:
                continue
            score = 1.0 if (stem in target or target in stem) else text_similarity(target, stem)
            if score > best_score:
                best_score, best = score, i
        if best >= 0 and best_score >= 0.72 and self.pick_var.get():
            if self.index != best:
                self.index = best
                self.render()

    def toggle_pick(self):
        self.pick_var.set(not self.pick_var.get())
        self.cfg["pick_mode"] = self.pick_var.get()
        save_config(self.cfg)
        self.capture_fail_streak = 0
        if self.pick_var.get():
            self.start_pick_timer()
            if self.get_ocr_engine() is None:
                self.footer.config(text="缺少 OCR 依赖：pip install winsdk pillow")
        if self.rows:
            self.render()

    def toggle_capture(self):
        self.cap_var.set(not self.cap_var.get())
        self.apply_capture_block(bool(self.cap_var.get()))
        if self.rows:
            self.render()

    # ---------- 交互 ----------

    def on_press(self, e):
        self.press_x = e.x_root
        self.press_y = e.y_root
        self.moved = False

    def on_drag(self, e):
        self.moved = True
        dx = e.x_root - self.press_x
        dy = e.y_root - self.press_y
        nx = self.root.winfo_x() + dx
        ny = self.root.winfo_y() + dy
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        h = self.root.winfo_height()
        nx = min(max(nx, 0), sw - self.width - 4)
        ny = min(max(ny, 0), sh - h - 4)
        self.root.geometry("+%d+%d" % (nx, ny))
        self.press_x = e.x_root
        self.press_y = e.y_root

    def on_release(self, e):
        if self.moved:
            self.cfg["x"] = self.root.winfo_x()
            self.cfg["y"] = self.root.winfo_y()
            save_config(self.cfg)
        elif self.rows:
            self.paused = not self.paused
            self.render()

    def show_menu(self, e):
        self.menu.tk_popup(e.x_root, e.y_root)

    def build_menu(self):
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label="暂停", command=self.toggle_pause)
        self.play_index = 0
        m.add_command(label="上一条", command=lambda: self.step(-1))
        m.add_command(label="下一条", command=lambda: self.step(1))
        m.add_separator()
        m.add_command(label="打开 Excel 文件…", command=self.open_file)
        m.add_separator()

        interval_menu = tk.Menu(m, tearoff=0)
        for s in (1, 2, 3, 5, 10, 30):
            interval_menu.add_radiobutton(label="%d 秒" % s, value=s,
                                          variable=self.interval_var,
                                          command=lambda s=s: self.set_interval(s))
        m.add_cascade(label="播放间隔", menu=interval_menu)

        font_menu = tk.Menu(m, tearoff=0)
        for s in (14, 18, 22, 26, 32):
            font_menu.add_radiobutton(label="%d 号" % s, value=s,
                                      variable=self.font_var,
                                      command=lambda s=s: self.set_font(s))
        m.add_cascade(label="字号", menu=font_menu)

        m.add_checkbutton(label="屏蔽录屏/投屏", variable=self.cap_var,
                          command=self.toggle_capture)
        m.add_checkbutton(label="选中取题", variable=self.pick_var,
                          command=self.toggle_pick)
        m.add_separator()
        m.add_command(label="退出", command=self.quit)
        self.menu = m

    def toggle_pause(self):
        if not self.rows:
            return
        self.paused = not self.paused
        self.render()

    def step(self, d):
        if not self.rows:
            return
        self.index += d
        self.render()

    def open_file(self):
        p = filedialog.askopenfilename(title="选择 Excel 文件",
                                       filetypes=[("Excel 工作表", "*.xlsx")])
        if p:
            self.load_file(p)

    def set_interval(self, s):
        self.cfg["interval"] = s
        save_config(self.cfg)

    def set_font(self, s):
        self.cfg["font"] = s
        save_config(self.cfg)
        self.render()

    def quit(self):
        self.cfg["x"] = self.root.winfo_x()
        self.cfg["y"] = self.root.winfo_y()
        save_config(self.cfg)
        self.root.destroy()


# ---------- 选中取题 OCR（Windows 自带引擎，无需权限） ----------

def run_windows_ocr(img):
    """用 Windows.Media.Ocr 识别 PIL Image，返回文本（后台线程中调用）"""
    import asyncio
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.globalization import Language
    from winsdk.windows.storage.streams import InMemoryRandomAccessStream, DataWriter
    from winsdk.windows.graphics.imaging import BitmapDecoder

    engine = (OcrEngine.try_create_from_language(Language("zh-Hans"))
              or OcrEngine.try_create_from_user_profile_languages())
    if engine is None:
        raise RuntimeError("系统缺少中文 OCR 语言包")

    async def _run():
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream)
        writer.write_bytes(buf.getvalue())
        await writer.store_async()
        writer.detach_stream()
        decoder = BitmapDecoder.create(stream)
        bitmap = await decoder.get_software_bitmap_async()
        result = await engine.recognize_async(bitmap)
        return " ".join(line.text for line in result.lines)

    return asyncio.run(_run())


def main():
    if sys.platform == "win32" and kernel32 is not None:
        # 单实例保护：已有运行实例时直接退出
        kernel32.CreateMutexW(None, False, "QuizBar_SingleInstance_Mutex")
        if kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            sys.exit(0)
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
