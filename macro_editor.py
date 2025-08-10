# macro_editor.py
import time
import tkinter as tk
from tkinter import ttk

import customtkinter as ctk
import pyautogui as pag

from utils import KEY_LIST, flush_modifiers, busy_wait, esc_pressed


class MacroEditor(ctk.CTkFrame):
    """
    マクロエディタ（完全版）
    - ブロック追加 / 配線（ポートドラッグ or 接続モード） / 実行
    - ダブルクリックでインライン名称編集
    - 複数選択（Shift+クリック / マーキー）・複製・削除
    - 右バナー（インスペクタ）で詳細編集＆削除
    - 入力系フォーカス中はショートカット無効化（Delete誤爆防止）
    - グリッドスナップ（10px）
    - アクション：左クリック / 右クリック / ダブルクリック / キー入力 / マウス移動（絶対/相対, X/Y, 時間）
    - ショートカット：
        A: アクション切替, P: 押し方切替, +/=: 回数+1, -: 回数-1,
        [: 間隔0.5x, ]: 間隔2x, .: 間隔=0.01, D: 複製, Delete: 削除, G: グリッドスナップ切替
    """

    # ======================== 初期化 ========================
    def __init__(self, master, stop_flag_ref, **kwargs):
        super().__init__(master, **kwargs)
        self.stop_flag_ref = stop_flag_ref  # callable: 実行停止フラグを返す

        # 状態
        self.blocks = {}            # bid -> meta
        self.connections = []       # (from_bid, to_bid, line_id)
        self.block_counter = 0
        self.current_block_id = None
        self.multi_selected = set()
        self.grid_snap = True

        # 一時オブジェクト
        self.connect_mode = False
        self.marquee_rect = None
        self.drag_select_origin = None
        self.wire_preview = None
        self.wire_from = None   # (bid, side 'L'|'R')

        self.recent_keys = []

        # ツールバー
        toolbar = ctk.CTkFrame(self)
        toolbar.pack(fill="x", padx=8, pady=(8, 6))
        ctk.CTkButton(toolbar, text="ブロック追加",
                      command=lambda: self.add_block(select_after_add=True)).pack(side="left", padx=4)
        ctk.CTkButton(toolbar, text="接続モード",
                      command=self.toggle_connect).pack(side="left", padx=4)
        ctk.CTkButton(toolbar, text="マクロ実行",
                      command=self.run_macro).pack(side="left", padx=4)

        # 本体
        body = ctk.CTkFrame(self)
        body.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        # キャンバス
        c_area = ctk.CTkFrame(body)
        c_area.pack(side="left", fill="both", expand=True)
        self.canvas = ctk.CTkCanvas(c_area, bg="#2A2A2A", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        # インスペクタ（右バナー）
        self.inspector = ctk.CTkFrame(body, width=300)
        self.inspector.pack(side="right", fill="y", padx=(8, 0))
        self._build_inspector(self.inspector)

        # バインド：ブロック
        self.canvas.tag_bind("block", "<Button-1>", self._on_block_click)
        self.canvas.tag_bind("block", "<B1-Motion>", self._on_block_drag)
        self.canvas.tag_bind("block", "<ButtonRelease-1>", self._on_block_release)
        self.canvas.tag_bind("block", "<Double-Button-1>", self._on_block_rename)

        # バインド：キャンバス
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        self.canvas.bind("<Button-1>", self._on_canvas_mousedown)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_mouseup)

        # バインド：ショートカット（入力欄フォーカス中は無効化）
        self.canvas.bind_all("<Key>", self._on_key_shortcuts)
        # 座標ラベル（右寄せ）
        self.cursor_label = ctk.CTkLabel(toolbar, text="(x, y) = (---, ---)", text_color="#A0A0A0")
        self.cursor_label.pack(side="right", padx=6)

        # 座標アップデート開始
        self._tick_cursor()
    def _tick_cursor(self):
        """マウスカーソル座標を 50ms 間隔で表示更新"""
        try:
            x, y = pag.position()
            # 必要ならここでマルチモニタやスケーリング補正を入れる
            self.cursor_label.configure(text=f"(x, y) = ({x}, {y})")
        except Exception:
            # 一瞬取得できないことがある環境でも落ちないように
            pass
        # 50ms 後に再実行（負荷が気になるなら 100ms でもOK）
        self.after(50, self._tick_cursor)



    # ======================== インスペクタ（右バナー） ========================
    def _build_inspector(self, parent):
        ctk.CTkLabel(parent, text="インスペクタ", font=("TkDefaultFont", 14, "bold")).pack(pady=(10, 6))
        self.lbl_sel = ctk.CTkLabel(parent, text="選択中: なし")
        self.lbl_sel.pack(pady=(0, 8))

        # 共通変数
        self.var_action = ctk.StringVar(value="左クリック")
        self.var_press = ctk.StringVar(value="短押し")
        self.var_seconds = ctk.StringVar(value="1.0")
        self.var_count = ctk.StringVar(value="1")
        self.var_interval = ctk.StringVar(value="0.5")
        self.var_key = ctk.StringVar(value="enter")

        # マウス移動
        self.var_move_mode = ctk.StringVar(value="絶対座標")  # or 相対座標
        self.var_move_x = ctk.StringVar(value="0")
        self.var_move_y = ctk.StringVar(value="0")
        self.var_move_time = ctk.StringVar(value="0")  # 0 で瞬間移動

        # アクション
        ctk.CTkLabel(parent, text="アクション").pack(anchor="w", padx=12)
        ctk.CTkOptionMenu(
            parent,
            values=["左クリック", "右クリック", "ダブルクリック", "キー入力", "マウス移動"],
            variable=self.var_action,
            command=lambda *_: (self._apply_inspector(), self._switch_inspector_fields())
        ).pack(fill="x", padx=12, pady=(0, 8))

        # 押し方
        self.row_press = ctk.CTkFrame(parent)
        self.row_press.pack(fill="x", padx=12, pady=(2, 4))
        ctk.CTkLabel(self.row_press, text="押し方").pack(side="left")
        ctk.CTkOptionMenu(
            self.row_press,
            values=["短押し", "長押し"],
            variable=self.var_press,
            command=lambda *_: (self._apply_inspector(), self._switch_inspector_fields())
        ).pack(side="right")

        # 長押し
        self.row_seconds = ctk.CTkFrame(parent)
        self.row_seconds.pack(fill="x", padx=12, pady=(2, 4))
        ctk.CTkLabel(self.row_seconds, text="長押し(秒)").pack(side="left")
        e1 = ctk.CTkEntry(self.row_seconds, textvariable=self.var_seconds, width=80)
        e1.pack(side="right")
        e1.bind("<FocusOut>", lambda *_: self._apply_inspector())

        # 短押し：回数/間隔
        self.row_count = ctk.CTkFrame(parent)
        self.row_count.pack(fill="x", padx=12, pady=(2, 4))
        ctk.CTkLabel(self.row_count, text="回数").pack(side="left")
        e2 = ctk.CTkEntry(self.row_count, textvariable=self.var_count, width=80)
        e2.pack(side="right")
        e2.bind("<FocusOut>", lambda *_: self._apply_inspector())

        self.row_interval = ctk.CTkFrame(parent)
        self.row_interval.pack(fill="x", padx=12, pady=(2, 10))
        ctk.CTkLabel(self.row_interval, text="間隔(秒)").pack(side="left")
        e3 = ctk.CTkEntry(self.row_interval, textvariable=self.var_interval, width=80)
        e3.pack(side="right")
        e3.bind("<FocusOut>", lambda *_: self._apply_inspector())

        # キー入力
        self.row_key = ctk.CTkFrame(parent)
        ctk.CTkLabel(self.row_key, text="キー").pack(anchor="w", padx=0, pady=(0, 2))
        try:
            from customtkinter import CTkComboBox
            self.key_widget = CTkComboBox(self.row_key, values=KEY_LIST, variable=self.var_key)
            self.key_widget.pack(fill="x")
            self.key_widget.bind("<KeyRelease>", lambda *_: self._apply_inspector())
            self.key_widget.bind("<<ComboboxSelected>>", lambda *_: self._apply_inspector())
        except Exception:
            self.key_widget = ctk.CTkEntry(self.row_key, textvariable=self.var_key)
            self.key_widget.pack(fill="x")
            self.key_widget.bind("<KeyRelease>", lambda *_: self._apply_inspector())
        self.row_key_recent = ctk.CTkFrame(self.row_key)
        self.row_key_recent.pack(fill="x", pady=(4, 0))

        # マウス移動
        self.row_move_mode = ctk.CTkFrame(parent)
        self.row_move_mode.pack(fill="x", padx=12, pady=(2, 4))
        ctk.CTkLabel(self.row_move_mode, text="移動方法").pack(side="left")
        ctk.CTkOptionMenu(
            self.row_move_mode,
            values=["絶対座標", "相対座標"],
            variable=self.var_move_mode,
            command=lambda *_: self._apply_inspector()
        ).pack(side="right")

        self.row_move_xy = ctk.CTkFrame(parent)
        self.row_move_xy.pack(fill="x", padx=12, pady=(2, 4))
        ctk.CTkLabel(self.row_move_xy, text="X").grid(row=0, column=0, padx=4, sticky="w")
        e4 = ctk.CTkEntry(self.row_move_xy, textvariable=self.var_move_x, width=80)
        e4.grid(row=0, column=1)
        e4.bind("<FocusOut>", lambda *_: self._apply_inspector())
        ctk.CTkLabel(self.row_move_xy, text="Y").grid(row=0, column=2, padx=(12, 4), sticky="w")
        e5 = ctk.CTkEntry(self.row_move_xy, textvariable=self.var_move_y, width=80)
        e5.grid(row=0, column=3)
        e5.bind("<FocusOut>", lambda *_: self._apply_inspector())

        self.row_move_time = ctk.CTkFrame(parent)
        self.row_move_time.pack(fill="x", padx=12, pady=(2, 10))
        ctk.CTkLabel(self.row_move_time, text="移動時間(秒)").pack(side="left")
        e6 = ctk.CTkEntry(self.row_move_time, textvariable=self.var_move_time, width=80)
        e6.pack(side="right")
        e6.bind("<FocusOut>", lambda *_: self._apply_inspector())

        # 削除ボタン（右バナー）
        self.frame_delete = ctk.CTkFrame(parent)
        self.frame_delete.pack(fill="x", padx=12, pady=(14, 12))
        self.btn_delete = ctk.CTkButton(
            self.frame_delete,
            text="ブロックを削除",
            fg_color="#7A1F1F",
            hover_color="#992525",
            state="disabled",
            command=self._delete_from_banner,
        )
        self.btn_delete.pack(fill="x")

        # 初期の表示切替と最近キーUI
        self._switch_inspector_fields()
        self._update_recent_keys_ui()
        self._update_delete_button()

    def _switch_inspector_fields(self):
        act = self.var_action.get()
        prs = self.var_press.get()

        # 一旦全部隠す
        for f in (self.row_press, self.row_seconds, self.row_count, self.row_interval,
                  self.row_key, self.row_move_mode, self.row_move_xy, self.row_move_time):
            f.pack_forget()

        if act in ("左クリック", "右クリック", "ダブルクリック", "キー入力"):
            self.row_press.pack(fill="x", padx=12, pady=(2, 4))
            if prs == "長押し":
                self.row_seconds.pack(fill="x", padx=12, pady=(2, 4))
            else:
                self.row_count.pack(fill="x", padx=12, pady=(2, 4))
                self.row_interval.pack(fill="x", padx=12, pady=(2, 10))
            if act == "キー入力":
                self.row_key.pack(fill="x", padx=12, pady=(2, 10))
        else:  # マウス移動
            self.row_move_mode.pack(fill="x", padx=12, pady=(2, 4))
            self.row_move_xy.pack(fill="x", padx=12, pady=(2, 4))
            self.row_move_time.pack(fill="x", padx=12, pady=(2, 10))

    def _update_recent_keys_ui(self):
        for w in self.row_key_recent.winfo_children():
            w.destroy()
        if not self.recent_keys:
            return
        ctk.CTkLabel(self.row_key_recent, text="最近:").pack(side="left")
        for k in self.recent_keys[:5]:
            ctk.CTkButton(self.row_key_recent, text=k, width=44,
                          command=lambda kk=k: (self.var_key.set(kk), self._apply_inspector())
                          ).pack(side="left", padx=2)

    def _apply_inspector(self):
        bid = self.current_block_id
        if not bid or bid not in self.blocks:
            return
        cfg = self.blocks[bid]['config']
        act = self.var_action.get()
        cfg['action'] = act

        if act in ("左クリック", "右クリック", "ダブルクリック", "キー入力"):
            cfg['press_type'] = self.var_press.get()
            try:
                cfg['seconds'] = float(self.var_seconds.get())
            except Exception:
                pass
            try:
                cfg['repeat_count'] = max(1, int(self.var_count.get()))
            except Exception:
                pass
            try:
                cfg['repeat_interval'] = max(0.0, float(self.var_interval.get()))
            except Exception:
                pass

            if act == "キー入力":
                k = (self.var_key.get() or 'enter').strip()
                cfg['key'] = k
                if k:
                    if k in self.recent_keys:
                        self.recent_keys.remove(k)
                    self.recent_keys.insert(0, k)
                    self._update_recent_keys_ui()

        else:  # マウス移動
            cfg['move_mode'] = self.var_move_mode.get()
            try:
                cfg['move_x'] = int(self.var_move_x.get())
            except Exception:
                cfg['move_x'] = 0
            try:
                cfg['move_y'] = int(self.var_move_y.get())
            except Exception:
                cfg['move_y'] = 0
            try:
                mv = float(self.var_move_time.get())
                cfg['move_time'] = max(0.0, mv)
            except Exception:
                cfg['move_time'] = 0.0

        self._refresh_block_label(bid)

    def _refresh_block_label(self, bid):
        cfg = self.blocks[bid]['config']
        act = cfg.get('action', "左クリック")
        if act == "キー入力":
            text = f"Key: {cfg.get('key', 'enter')}"
        elif act == "マウス移動":
            mode = cfg.get('move_mode', '絶対座標')
            x, y = cfg.get('move_x', 0), cfg.get('move_y', 0)
            if mode == "絶対座標":
                text = f"Move: ({x},{y})"
            else:
                text = f"Move: d{ x:+},{ y:+}"
        else:
            text = act
        self.canvas.itemconfig(self.blocks[bid]['text_id'], text=text)

    def _update_delete_button(self):
        count = (1 if self.current_block_id else 0) + len(self.multi_selected)
        if count == 0:
            self.btn_delete.configure(text="ブロックを削除", state="disabled")
            return
        label = "このブロックを削除" if count == 1 else f"{count}個のブロックを削除"
        self.btn_delete.configure(text=label, state="normal")

    def _delete_from_banner(self):
        targets = set(self.multi_selected)
        if self.current_block_id:
            targets.add(self.current_block_id)
        if not targets:
            return
        self._delete_blocks(targets)
        self._update_delete_button()

    # ======================== ブロックUI ========================
    def add_block(self, select_after_add=False):
        x, y, w, h = 60, 60, 180, 54
        bid = f"block_{self.block_counter}"
        self.block_counter += 1

        r = self.canvas.create_rectangle(
            x, y, x + w, y + h, fill="#3A3A3A", outline="#5A5A5A", width=2, tags=("block", bid)
        )
        t = self.canvas.create_text(x + w / 2, y + h / 2, text="左クリック", fill="white", tags=("block", bid))

        # 左右ポート
        lp = self.canvas.create_oval(x - 6, y + h / 2 - 6, x + 6, y + h / 2 + 6,
                                     fill="#7AA2F7", outline="", tags=("port", bid, "portL"))
        rp = self.canvas.create_oval(x + w - 6, y + h / 2 - 6, x + w + 6, y + h / 2 + 6,
                                     fill="#7AA2F7", outline="", tags=("port", bid, "portR"))
        self.canvas.tag_bind(lp, "<Button-1>", lambda e, b=bid: self._start_wire(b, "L"))
        self.canvas.tag_bind(rp, "<Button-1>", lambda e, b=bid: self._start_wire(b, "R"))
        self.canvas.tag_bind("port", "<B1-Motion>", self._drag_wire)
        self.canvas.tag_bind("port", "<ButtonRelease-1>", self._finish_wire)

        self.blocks[bid] = {
            'rect_id': r, 'text_id': t, 'x': x, 'y': y, 'w': w, 'h': h, 'dragging': False,
            'ports': {'L': lp, 'R': rp},
            'config': {
                'action': "左クリック",
                'press_type': "短押し",
                'seconds': 1.0,
                'repeat_count': 1,
                'repeat_interval': 0.5,
                'key': 'enter',
                'move_mode': '絶対座標',
                'move_x': 0, 'move_y': 0, 'move_time': 0.0,
            }
        }
        if select_after_add:
            self._select_block(bid)

    def toggle_connect(self):
        self.connect_mode = not self.connect_mode
        self.canvas.configure(cursor="tcross" if self.connect_mode else "")
        print(f"接続モード：{'ON' if self.connect_mode else 'OFF'}")

    def _on_block_click(self, event):
        items = self.canvas.find_withtag("current")
        if not items:
            return
        tags = self.canvas.gettags(items[0])
        bid = next((t for t in tags if t.startswith("block_")), None)
        if not bid or bid not in self.blocks:
            return

        if self.connect_mode:
            if self.current_block_id and self.current_block_id != bid:
                self._draw_connection(self.current_block_id, bid)
            self._select_block(bid)
            return

        bx, by = self.blocks[bid]['x'], self.blocks[bid]['y']
        self.blocks[bid]['drag_offset_x'] = event.x - bx
        self.blocks[bid]['drag_offset_y'] = event.y - by
        self.blocks[bid]['dragging'] = True

        # Shift なら複数選択に追加
        if self._shift(event) and self.current_block_id != bid:
            self.multi_selected.add(bid)
            self._highlight(bid, True)
            self._update_delete_button()
        else:
            self._select_block(bid)

    def _on_block_drag(self, event):
        items = self.canvas.find_withtag("current")
        if not items:
            return
        tags = self.canvas.gettags(items[0])
        bid = next((t for t in tags if t.startswith("block_")), None)
        if not bid or not self.blocks[bid].get('dragging'):
            return

        ox, oy = self.blocks[bid]['drag_offset_x'], self.blocks[bid]['drag_offset_y']
        nx, ny = event.x - ox, event.y - oy
        w, h = self.blocks[bid]['w'], self.blocks[bid]['h']

        self.canvas.coords(self.blocks[bid]['rect_id'], nx, ny, nx + w, ny + h)
        self.canvas.coords(self.blocks[bid]['text_id'], nx + w / 2, ny + h / 2)
        self.blocks[bid]['x'], self.blocks[bid]['y'] = nx, ny

        lp = self.blocks[bid]['ports']['L']
        rp = self.blocks[bid]['ports']['R']
        self.canvas.coords(lp, nx - 6, ny + h / 2 - 6, nx + 6, ny + h / 2 + 6)
        self.canvas.coords(rp, nx + w - 6, ny + h / 2 - 6, nx + w + 6, ny + h / 2 + 6)

        # 線更新
        for (f, t, lid) in list(self.connections):
            if f in self.blocks and t in self.blocks:
                fx = self.blocks[f]['x'] + self.blocks[f]['w'] / 2
                fy = self.blocks[f]['y'] + self.blocks[f]['h'] / 2
                tx = self.blocks[t]['x'] + self.blocks[t]['w'] / 2
                ty = self.blocks[t]['y'] + self.blocks[t]['h'] / 2
                self.canvas.coords(lid, fx, fy, tx, ty)

    def _on_block_release(self, event):
        items = self.canvas.find_withtag("current")
        if not items:
            return
        tags = self.canvas.gettags(items[0])
        bid = next((t for t in tags if t.startswith("block_")), None)
        if not bid or bid not in self.blocks:
            return
        self.blocks[bid]['dragging'] = False

        if self.grid_snap:
            x, y = self.blocks[bid]['x'], self.blocks[bid]['y']
            nx, ny = round(x / 10) * 10, round(y / 10) * 10
            dx, dy = nx - x, ny - y
            if dx or dy:
                self._nudge_block(bid, dx, dy)

    def _nudge_block(self, bid, dx, dy):
        meta = self.blocks[bid]
        x, y, w, h = meta['x'] + dx, meta['y'] + dy, meta['w'], meta['h']
        self.canvas.coords(meta['rect_id'], x, y, x + w, y + h)
        self.canvas.coords(meta['text_id'], x + w / 2, y + h / 2)
        lp = meta['ports']['L']
        rp = meta['ports']['R']
        self.canvas.coords(lp, x - 6, y + h / 2 - 6, x + 6, y + h / 2 + 6)
        self.canvas.coords(rp, x + w - 6, y + h / 2 - 6, x + w + 6, y + h / 2 + 6)
        meta['x'], meta['y'] = x, y

        # 線更新
        for (f, t, lid) in list(self.connections):
            if f in self.blocks and t in self.blocks:
                fx = self.blocks[f]['x'] + self.blocks[f]['w'] / 2
                fy = self.blocks[f]['y'] + self.blocks[f]['h'] / 2
                tx = self.blocks[t]['x'] + self.blocks[t]['w'] / 2
                ty = self.blocks[t]['y'] + self.blocks[t]['h'] / 2
                self.canvas.coords(lid, fx, fy, tx, ty)

    # ======================== キャンバス空白 ========================
    def _on_canvas_mousedown(self, event):
        # 既存マーキー掃除
        if self.marquee_rect and self._item_exists(self.marquee_rect):
            try:
                self.canvas.delete(self.marquee_rect)
            except Exception:
                pass
        self.marquee_rect = None
        self.drag_select_origin = None

        # 空白ならマーキー開始
        item = self.canvas.find_withtag("current")
        if not item:
            self.drag_select_origin = (event.x, event.y)
            self.marquee_rect = self.canvas.create_rectangle(
                event.x, event.y, event.x, event.y, outline="#58A6FF", dash=(2, 2)
            )

    def _on_canvas_drag(self, event):
        if not self.marquee_rect or not self._item_exists(self.marquee_rect) or not self.drag_select_origin:
            return
        x0, y0 = self.drag_select_origin
        self.canvas.coords(self.marquee_rect, x0, y0, event.x, event.y)

    def _on_canvas_mouseup(self, event):
        if not self.marquee_rect or not self._item_exists(self.marquee_rect):
            self.marquee_rect = None
            self.drag_select_origin = None
            return
        coords = self.canvas.coords(self.marquee_rect) or []
        try:
            x0, y0, x1, y1 = coords
        except Exception:
            try:
                self.canvas.delete(self.marquee_rect)
            except Exception:
                pass
            self.marquee_rect = None
            self.drag_select_origin = None
            return

        try:
            self.canvas.delete(self.marquee_rect)
        except Exception:
            pass
        self.marquee_rect = None
        self.drag_select_origin = None

        # 範囲内ブロックを選択
        self.multi_selected.clear()
        xmin, xmax = min(x0, x1), max(x0, x1)
        ymin, ymax = min(y0, y1), max(y0, y1)
        for bid, meta in self.blocks.items():
            x, y, w, h = meta['x'], meta['y'], meta['w'], meta['h']
            if (x >= xmin) and (y >= ymin) and (x + w <= xmax) and (y + h <= ymax):
                self.multi_selected.add(bid)
                self._highlight(bid, True)
        self._update_delete_button()

    # ======================== リネーム ========================
    def _on_block_rename(self, event):
        items = self.canvas.find_withtag("current")
        if not items:
            return
        tags = self.canvas.gettags(items[0])
        bid = next((t for t in tags if t.startswith("block_")), None)
        if not bid or bid not in self.blocks:
            return

        try:
            self.canvas.delete("inline_edit")
        except Exception:
            pass

        meta = self.blocks[bid]
        text_id = meta.get('text_id')

        if not text_id or not self._item_exists(text_id):
            x, y, w, h = meta['x'], meta['y'], meta['w'], meta['h']
            text_id = self.canvas.create_text(x + w / 2, y + h / 2, text="Block", fill="white", tags=("block", bid))
            meta['text_id'] = text_id

        coords = self.canvas.coords(text_id) if self._item_exists(text_id) else []
        if not coords or len(coords) < 2:
            x, y, w, h = meta['x'], meta['y'], meta['w'], meta['h']
            tx, ty = x + w / 2, y + h / 2
        else:
            tx, ty = coords[0], coords[1]

        current_text = self.canvas.itemcget(text_id, "text") or "Block"
        entry = ctk.CTkEntry(self.canvas, width=150)
        entry.insert(0, current_text)
        self.canvas.create_window(tx, ty, window=entry, tags=("inline_edit",))
        entry.focus()

        def _commit(_=None):
            new = entry.get().strip() or "Block"
            self.canvas.itemconfig(text_id, text=new)
            try:
                self.canvas.delete("inline_edit")
            except Exception:
                pass

        entry.bind("<Return>", _commit)
        entry.bind("<Escape>", lambda e: self.canvas.delete("inline_edit"))
        entry.bind("<FocusOut>", _commit)

    # ======================== 配線（ワイヤ） ========================
    def _start_wire(self, bid, side):
        if self.wire_preview and self._item_exists(self.wire_preview):
            try:
                self.canvas.delete(self.wire_preview)
            except Exception:
                pass
        self.wire_preview = None
        self.wire_from = (bid, side)
        x, y = self._port_center(bid, side)
        self.wire_preview = self.canvas.create_line(x, y, x, y, fill="white", dash=(3, 2))

    def _drag_wire(self, event):
        if not self.wire_preview or not self._item_exists(self.wire_preview):
            self.wire_preview = None
            return
        coords = self.canvas.coords(self.wire_preview) or []
        if len(coords) < 2:
            return
        x0, y0 = coords[:2]
        self.canvas.coords(self.wire_preview, x0, y0, event.x, event.y)

    def _finish_wire(self, event):
        if not self.wire_preview or not self._item_exists(self.wire_preview):
            self.wire_preview = None
            self.wire_from = None
            return
        try:
            self.canvas.delete(self.wire_preview)
        except Exception:
            pass
        self.wire_preview = None
        if not self.wire_from:
            return
        bid_from, _ = self.wire_from
        self.wire_from = None

        item = self.canvas.find_withtag("current")
        if not item:
            # 配線をやめた場合も接続モードを解除しておくと誤操作が減る
            self.connect_mode = False
            self.canvas.configure(cursor="")
            return

        tags = self.canvas.gettags(item)
        bid_to = next((t for t in tags if t.startswith("block_")), None)
        if bid_to and bid_to != bid_from:
            self._draw_connection(bid_from, bid_to)

        # ★ここでモード解除
        self.connect_mode = False
        self.canvas.configure(cursor="")
        print("接続完了 → 接続モード OFF")


    def _port_center(self, bid, side):
        x, y, w, h = self.blocks[bid]['x'], self.blocks[bid]['y'], self.blocks[bid]['w'], self.blocks[bid]['h']
        return (x, y + h / 2) if side == "L" else (x + w, y + h / 2)

    def _draw_connection(self, b1, b2):
        if b1 not in self.blocks or b2 not in self.blocks:
            return
        x1 = self.blocks[b1]['x'] + self.blocks[b1]['w'] / 2
        y1 = self.blocks[b1]['y'] + self.blocks[b1]['h'] / 2
        x2 = self.blocks[b2]['x'] + self.blocks[b2]['w'] / 2
        y2 = self.blocks[b2]['y'] + self.blocks[b2]['h'] / 2
        lid = self.canvas.create_line(x1, y1, x2, y2, arrow="last", fill="white", width=2)
        self.connections.append((b1, b2, lid))

    # ======================== 選択系 ========================
    def _select_block(self, bid):
        if self.current_block_id and self.current_block_id in self.blocks:
            self._highlight(self.current_block_id, False)
        for b in list(self.multi_selected):
            self._highlight(b, False)
        self.multi_selected.clear()

        self.current_block_id = bid
        self._highlight(bid, True)
        self.lbl_sel.configure(text=f"選択中: {bid}")
        self._load_to_inspector(bid)
        self._update_delete_button()

    def _highlight(self, bid, on):
        self.canvas.itemconfig(
            self.blocks[bid]['rect_id'],
            outline="#58A6FF" if on else "#5A5A5A",
            width=3 if on else 2
        )

    def _load_to_inspector(self, bid):
        cfg = self.blocks[bid]['config']
        act = cfg.get('action', "左クリック")
        self.var_action.set(act)

        if act in ("左クリック", "右クリック", "ダブルクリック", "キー入力"):
            self.var_press.set(cfg.get('press_type', "短押し"))
            self.var_seconds.set(str(cfg.get('seconds', 1.0)))
            self.var_count.set(str(cfg.get('repeat_count', 1)))
            self.var_interval.set(str(cfg.get('repeat_interval', 0.5)))
            self.var_key.set(cfg.get('key', 'enter'))
        else:
            self.var_move_mode.set(cfg.get('move_mode', '絶対座標'))
            self.var_move_x.set(str(cfg.get('move_x', 0)))
            self.var_move_y.set(str(cfg.get('move_y', 0)))
            self.var_move_time.set(str(cfg.get('move_time', 0.0)))

        self._switch_inspector_fields()

    def _shift(self, event):
        try:
            return (event.state & 0x0001) != 0
        except Exception:
            return False

    # ======================== ショートカット ========================
    def _is_typing_widget(self, w) -> bool:
        """フォーカス中が入力系（Entry/Combobox/Text等）なら True"""
        if w is None:
            return False
        typing_types = (
            ctk.CTkEntry,
            getattr(ctk, "CTkComboBox", tuple()),
            tk.Entry, tk.Text, ttk.Entry, ttk.Combobox,
        )
        try:
            return isinstance(w, typing_types)
        except Exception:
            return False

    def _on_key_shortcuts(self, event):
        # 入力欄にフォーカスがある間はショートカット無効化
        if self._is_typing_widget(self.focus_get()):
            return
        if not self.current_block_id:
            return

        bids = {self.current_block_id} | set(self.multi_selected)

        def apply(fn):
            for b in bids:
                if b not in self.blocks:
                    continue
                fn(self.blocks[b]['config'])
                self._refresh_block_label(b)

        key = (event.keysym or "").lower()

        if key == 'a':
            order = ["左クリック", "右クリック", "ダブルクリック", "キー入力", "マウス移動"]
            def _next(cfg):
                i = order.index(cfg.get('action', "左クリック"))
                cfg['action'] = order[(i + 1) % len(order)]
            apply(_next)
            self._load_to_inspector(self.current_block_id)

        elif key == 'p':
            def _toggle(cfg):
                cfg['press_type'] = "長押し" if cfg.get('press_type', "短押し") == "短押し" else "短押し"
            apply(_toggle)
            self._load_to_inspector(self.current_block_id)

        elif key in ('plus', 'equal'):
            def _inc(cfg): cfg['repeat_count'] = max(1, cfg.get('repeat_count', 1) + 1)
            apply(_inc)

        elif key == 'minus':
            def _dec(cfg): cfg['repeat_count'] = max(1, cfg.get('repeat_count', 1) - 1)
            apply(_dec)

        elif key == 'bracketleft':
            def _half(cfg): cfg['repeat_interval'] = max(0.0, cfg.get('repeat_interval', 0.5) * 0.5)
            apply(_half)

        elif key == 'bracketright':
            def _double(cfg): cfg['repeat_interval'] = cfg.get('repeat_interval', 0.5) * 2
            apply(_double)

        elif key == 'period':
            def _fast(cfg): cfg['repeat_interval'] = 0.01
            apply(_fast)

        elif key == 'd':
            self._duplicate_blocks(bids)

        elif key in ('delete', 'backspace'):
            self._delete_blocks(bids)

        elif key == 'g':
            self.grid_snap = not self.grid_snap
            print(f"グリッドスナップ: {'ON' if self.grid_snap else 'OFF'}")

    def _duplicate_blocks(self, bids):
        dx, dy = 20, 20
        mapping = {}
        for b in bids:
            if b not in self.blocks:
                continue
            meta = self.blocks[b]
            x, y, w, h = meta['x'] + dx, meta['y'] + dy, meta['w'], meta['h']
            nb = f"block_{self.block_counter}"
            self.block_counter += 1
            r = self.canvas.create_rectangle(x, y, x + w, y + h, fill="#3A3A3A", outline="#5A5A5A", width=2,
                                             tags=("block", nb))
            t = self.canvas.create_text(x + w / 2, y + h / 2,
                                        text=self.canvas.itemcget(meta['text_id'], "text"),
                                        fill="white", tags=("block", nb))
            lp = self.canvas.create_oval(x - 6, y + h / 2 - 6, x + 6, y + h / 2 + 6,
                                         fill="#7AA2F7", outline="", tags=("port", nb, "portL"))
            rp = self.canvas.create_oval(x + w - 6, y + h / 2 - 6, x + w + 6, y + h / 2 + 6,
                                         fill="#7AA2F7", outline="", tags=("port", nb, "portR"))
            self.canvas.tag_bind(lp, "<Button-1>", lambda e, b=nb: self._start_wire(b, "L"))
            self.canvas.tag_bind(rp, "<Button-1>", lambda e, b=nb: self._start_wire(b, "R"))
            self.canvas.tag_bind("port", "<B1-Motion>", self._drag_wire)
            self.canvas.tag_bind("port", "<ButtonRelease-1>", self._finish_wire)

            self.blocks[nb] = {
                'rect_id': r, 'text_id': t, 'x': x, 'y': y, 'w': w, 'h': h, 'dragging': False,
                'ports': {'L': lp, 'R': rp},
                'config': dict(meta['config'])
            }
            mapping[b] = nb

        # 選択内の接続を複製
        for (f, t, _) in list(self.connections):
            if f in bids and t in bids:
                nf, nt = mapping.get(f), mapping.get(t)
                if nf and nt:
                    self._draw_connection(nf, nt)

    def _delete_blocks(self, bids):
        bids = set(bids)

        # プレビューやマーキーを掃除
        if self.wire_preview and self._item_exists(self.wire_preview):
            try:
                self.canvas.delete(self.wire_preview)
            except Exception:
                pass
        self.wire_preview = None
        self.wire_from = None
        if self.marquee_rect and self._item_exists(self.marquee_rect):
            try:
                self.canvas.delete(self.marquee_rect)
            except Exception:
                pass
        self.marquee_rect = None
        self.drag_select_origin = None

        # 線削除
        for (f, t, lid) in list(self.connections):
            if f in bids or t in bids:
                try:
                    self.canvas.delete(lid)
                except Exception:
                    pass
                self.connections.remove((f, t, lid))

        # ブロック削除
        for b in bids:
            meta = self.blocks.get(b)
            if not meta:
                continue
            for iid in (meta['rect_id'], meta['text_id'], meta['ports']['L'], meta['ports']['R']):
                try:
                    self.canvas.delete(iid)
                except Exception:
                    pass
            self.blocks.pop(b, None)
            self.multi_selected.discard(b)

        if self.current_block_id in bids:
            self.current_block_id = None
            self.lbl_sel.configure(text="選択中: なし")

        self._update_delete_button()

    # ======================== 実行 ========================
    def run_macro(self):
        stop = self.stop_flag_ref
        incoming = {t for (_, t, _) in self.connections}
        start_nodes = [bid for bid in self.blocks if bid not in incoming]
        if not start_nodes:
            print("スタートブロックがありません。")
            return

        visited = set()

        def dfs(bid):
            if bid in visited:
                return
            visited.add(bid)
            self._exec_block(bid, stop)
            for (f, t, _) in self.connections:
                if f == bid:
                    dfs(t)

        for s in start_nodes:
            dfs(s)

    def _exec_block(self, bid, stop):
        cfg = self.blocks[bid]['config']
        act = cfg.get('action', "左クリック")

        # 実行直前：修飾キー離れ待ち（mac の  対策）
        flush_modifiers()
        if stop():
            return

        if act in ("左クリック", "右クリック", "ダブルクリック"):
            press = cfg.get('press_type', "短押し")
            secs = cfg.get('seconds', 1.0)
            count = cfg.get('repeat_count', 1)
            itv = cfg.get('repeat_interval', 0.5)

            if press == "短押し":
                x, y = pag.position()
                for _ in range(max(1, count)):
                    if stop():
                        return
                    if act == "左クリック":
                        pag.click(x, y)
                    elif act == "右クリック":
                        pag.rightClick(x, y)
                    else:
                        pag.doubleClick(x, y)
                    if busy_wait(itv, stop):
                        return
            else:
                x, y = pag.position()
                pag.moveTo(x, y)
                pag.mouseDown()
                t0 = time.time()
                while time.time() - t0 < secs:
                    if stop() or esc_pressed():
                        pag.mouseUp()
                        return
                    time.sleep(0.01)
                pag.mouseUp()

        elif act == "キー入力":
            press = cfg.get('press_type', "短押し")
            secs = cfg.get('seconds', 1.0)
            count = cfg.get('repeat_count', 1)
            itv = cfg.get('repeat_interval', 0.5)
            key = cfg.get('key', 'enter')

            if press == "短押し":
                for _ in range(max(1, count)):
                    if stop():
                        return
                    pag.press(key)
                    if busy_wait(itv, stop):
                        return
            else:
                pag.keyDown(key)
                t0 = time.time()
                while time.time() - t0 < secs:
                    if stop() or esc_pressed():
                        pag.keyUp(key)
                        return
                    time.sleep(0.01)
                pag.keyUp(key)

        elif act == "マウス移動":
            mode = cfg.get('move_mode', '絶対座標')
            x = int(cfg.get('move_x', 0))
            y = int(cfg.get('move_y', 0))
            dur = float(cfg.get('move_time', 0.0))
            try:
                if mode == "絶対座標":
                    pag.moveTo(x, y, duration=max(0.0, dur))
                else:
                    pag.moveRel(x, y, duration=max(0.0, dur))
            except Exception as e:
                print("マウス移動エラー:", e)
                return

        print(f"{bid} 実行完了")

    # ======================== 安全ユーティリティ ========================
    def _item_exists(self, item_id) -> bool:
        if not item_id:
            return False
        try:
            return self.canvas.type(item_id) != ""
        except Exception:
            return False
