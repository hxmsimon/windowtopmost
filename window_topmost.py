"""
window_topmost.py  —  窗口置顶工具 v10

v10: 稳定版 - 纯 Tkinter，零 Win32 消息循环依赖
  - 移除全局鼠标钩子（ACCESS_VIOLATION）
  - 移除 pystray/PIL（无额外依赖）
  - 移除系统托盘图标（避免 Win32 消息回调导致闪退/残留）
  - 标准窗口 + 关闭按钮退出，简洁可靠
  - 双击列表项切换置顶
"""

import tkinter as tk
from tkinter import ttk
import ctypes
from ctypes import wintypes, WINFUNCTYPE, c_int, Structure, POINTER, byref, sizeof
import threading
import time
import os

# ═══════════════════════════════════════════════════════════════════════════════
#  Win32 API（仅用于窗口枚举和置顶，不涉及消息循环）
# ═══════════════════════════════════════════════════════════════════════════════
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, c_int, c_int, c_int, c_int, wintypes.UINT]
user32.SetWindowPos.restype = wintypes.BOOL
user32.IsWindow.argtypes = [wintypes.HWND]
user32.IsWindow.restype = wintypes.BOOL
user32.GetWindowLongW.argtypes = [wintypes.HWND, c_int]
user32.GetWindowLongW.restype = ctypes.c_long
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = c_int
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, c_int]
user32.GetWindowTextW.restype = c_int
user32.IsWindowVisible.argtypes = [wintypes.HWND]
user32.IsWindowVisible.restype = wintypes.BOOL
user32.ShowWindow.argtypes = [wintypes.HWND, c_int]
user32.ShowWindow.restype = wintypes.BOOL
user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.SetForegroundWindow.restype = wintypes.BOOL
user32.GetParent.argtypes = [wintypes.HWND]
user32.GetParent.restype = wintypes.HWND
EnumWindowsProcType = WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
user32.EnumWindows.argtypes = [EnumWindowsProcType, wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL

# ═══════════════════════════════════════════════════════════════════════════════
#  常量
# ═══════════════════════════════════════════════════════════════════════════════
HWND_TOPMOST   = wintypes.HWND(-1)
HWND_NOTOPMOST = wintypes.HWND(-2)
SWP_NOSIZE     = 0x0001
SWP_NOMOVE     = 0x0002
SWP_NOACTIVATE = 0x0010
GWL_EXSTYLE    = -20
WS_EX_TOPMOST  = 0x0008
SW_RESTORE     = 9

# ═══════════════════════════════════════════════════════════════════════════════
#  辅助函数
# ═══════════════════════════════════════════════════════════════════════════════
def set_topmost(hwnd, enable):
    flag = HWND_TOPMOST if enable else HWND_NOTOPMOST
    return bool(user32.SetWindowPos(hwnd, flag, 0, 0, 0, 0,
                                     SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE))

def is_topmost(hwnd):
    return bool(user32.GetWindowLongW(hwnd, GWL_EXSTYLE) & WS_EX_TOPMOST)

def get_window_title(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    if n == 0: return ""
    buf = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, buf, n + 1)
    return buf.value.strip()

def get_top_level_hwnd(hwnd):
    while True:
        parent = user32.GetParent(hwnd)
        if not parent or parent == 0: return hwnd
        hwnd = parent

def get_windows():
    result = []
    def _cb(hwnd, _):
        if not user32.IsWindowVisible(hwnd): return True
        title = get_window_title(hwnd)
        if title: result.append((int(hwnd), title, is_topmost(hwnd)))
        return True
    user32.EnumWindows(EnumWindowsProcType(_cb), 0)
    result.sort(key=lambda x: (not x[2], x[1].lower()))
    return result

# ═══════════════════════════════════════════════════════════════════════════════
#  App
# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📌 窗口置顶工具")
        self.geometry("680x540")
        self.resizable(True, True)
        self.configure(bg="#1e1e2e")

        # 居中
        self.update_idletasks()
        w, h = 680, 540
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.wm_attributes("-topmost", True)
        self.update()
        try:
            self.wm_attributes("-alpha", 0.98)
        except: pass

        self._topmost_set = set()
        self._quit_flag = False

        self._build_ui()
        self._refresh_list()
        self._start_auto_refresh()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI ───────────────────────────────────────────────────
    def _build_ui(self):
        BG, FG, DIM = "#1e1e2e", "#cdd6f4", "#6c7086"

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", background="#2a2a3e", foreground=FG,
                        rowheight=26, fieldbackground="#2a2a3e", font=("微软雅黑", 10))
        style.configure("Treeview.Heading", background="#313244", foreground="#cba6f7",
                        font=("微软雅黑", 10, "bold"))
        style.map("Treeview", background=[("selected", "#45475a")])

        # 提示
        tk.Label(self, text="双击列表项可切换置顶状态", bg=BG, fg=DIM,
                 font=("微软雅黑", 9)).pack(fill=tk.X, padx=10, pady=(6, 2))

        # 搜索
        sf = tk.Frame(self, bg=BG)
        sf.pack(fill=tk.X, padx=10, pady=(2, 4))
        tk.Label(sf, text="🔍", bg=BG, fg=FG, font=("微软雅黑", 10)).pack(side=tk.LEFT)
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._filter_list())
        tk.Entry(sf, textvariable=self._search_var, bg="#313244", fg=FG,
                 insertbackground=FG, relief=tk.FLAT, font=("微软雅黑", 10)
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0), ipady=3)

        # 列表
        lf = tk.Frame(self, bg=BG)
        lf.pack(fill=tk.BOTH, expand=True, padx=10, pady=3)
        self._tree = ttk.Treeview(lf, columns=("pin", "title"), show="headings", selectmode="browse")
        self._tree.heading("pin", text="置顶"); self._tree.heading("title", text="窗口标题")
        self._tree.column("pin", width=60, anchor=tk.CENTER, stretch=False)
        self._tree.column("title", width=580, anchor=tk.W)
        self._tree.tag_configure("pinned", foreground="#a6e3a1", font=("微软雅黑", 10, "bold"))
        vsb = ttk.Scrollbar(lf, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y); self._tree.pack(fill=tk.BOTH, expand=True)
        self._tree.bind("<Double-1>", lambda e: self._toggle_pin())
        self._tree.bind("<Return>", lambda e: self._toggle_pin())

        # 按钮
        bf = tk.Frame(self, bg=BG); bf.pack(fill=tk.X, padx=10, pady=(3, 4))
        for txt, cmd, clr in [("📌 置顶/取消", self._toggle_pin, "#a6e3a1"),
                                ("🪟 显示该窗口", self._bring_front, "#89dceb"),
                                ("🔄 刷新", self._refresh_list, "#cdd6f4"),
                                ("🔓 全部取消置顶", self._unpin_all, "#fab387")]:
            tk.Button(bf, text=txt, command=cmd, bg="#313244", fg=clr,
                      activebackground="#45475a", activeforeground=clr, relief=tk.FLAT,
                      font=("微软雅黑", 10), padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=2)
        self._auto_var = tk.BooleanVar(value=True)
        tk.Checkbutton(bf, text="自动刷新", variable=self._auto_var, bg=BG, fg=FG,
                       selectcolor="#313244", activebackground=BG, activeforeground=FG,
                       font=("微软雅黑", 9)).pack(side=tk.RIGHT, padx=6)

        # 状态栏
        self._status = tk.StringVar(value="就绪 — 双击列表项切换置顶")
        tk.Label(self, textvariable=self._status, bg="#181825", fg=DIM,
                 font=("微软雅黑", 9), anchor=tk.W).pack(fill=tk.X, padx=10, pady=(0, 2))

    # ── 列表 ─────────────────────────────────────────────────
    def _refresh_list(self):
        self._filter_list()

    def _filter_list(self):
        kw = self._search_var.get().lower()
        sel = self._tree.selection()
        sel_iid = sel[0] if sel else None
        self._tree.delete(*self._tree.get_children())
        for hwnd_int, title, _ in get_windows():
            if kw and kw not in title.lower(): continue
            pin = is_topmost(wintypes.HWND(hwnd_int))
            self._tree.insert("", tk.END, iid=str(hwnd_int),
                              values=("📌" if pin else "", title),
                              tags=("pinned",) if pin else ())
        if sel_iid and self._tree.exists(sel_iid):
            self._tree.selection_set(sel_iid); self._tree.see(sel_iid)

    def _selected_hwnd(self):
        sel = self._tree.selection()
        if not sel: return None
        try: return int(sel[0])
        except: return None

    # ── 置顶 ─────────────────────────────────────────────────
    def _toggle_pin(self):
        hwnd_int = self._selected_hwnd()
        if hwnd_int is None:
            self._status.set("⚠  请先选择一个窗口"); return
        hwnd = wintypes.HWND(hwnd_int)
        cur = is_topmost(hwnd)
        ok = set_topmost(hwnd, not cur)
        iid = str(hwnd_int)
        title = self._tree.set(iid, "title") if self._tree.exists(iid) else str(hwnd_int)
        if ok:
            if not cur:
                self._topmost_set.add(hwnd_int)
                self._status.set(f"📌 已置顶：{title}")
            else:
                self._topmost_set.discard(hwnd_int)
                self._status.set(f"取消置顶：{title}")
        else:
            self._status.set(f"⚠ 置顶失败：{title}")
        self._filter_list()

    def _bring_front(self):
        hwnd_int = self._selected_hwnd()
        if hwnd_int is None:
            self._status.set("⚠  请先选择一个窗口"); return
        hwnd = wintypes.HWND(hwnd_int)
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)

    def _unpin_all(self):
        for v in list(self._topmost_set):
            hw = wintypes.HWND(v)
            if user32.IsWindow(hw): set_topmost(hw, False)
        self._topmost_set.clear()
        self._status.set("🔓 已取消所有置顶"); self._filter_list()

    # ── 自动刷新 ────────────────────────────────────────────
    def _start_auto_refresh(self):
        def _loop():
            while not self._quit_flag:
                time.sleep(3)
                if self._auto_var.get():
                    try: self.after(0, self._refresh_list)
                    except: break
        threading.Thread(target=_loop, daemon=True).start()

    # ── 关闭 ─────────────────────────────────────────────────
    def _on_close(self):
        self._quit_flag = True
        for v in list(self._topmost_set):
            hw = wintypes.HWND(v)
            if user32.IsWindow(hw): set_topmost(hw, False)
        self.destroy()

# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    App().mainloop()
