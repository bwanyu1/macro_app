# macro_editor.py
import customtkinter as ctk
import pyautogui as pag
import time

from utils import KEY_LIST, flush_modifiers, busy_wait, esc_pressed

class MacroEditor(ctk.CTkFrame):
    """
    マクロタブ：キャンバス＋右インスペクタ常設（ポップアップ廃止）
    - ブロックのインライン編集（ダブルクリック）
    - 右/左ポートからドラッグで配線（ワンクリック接続）
    - 複数選択（Shift+クリック／ドラッグマーキー）
    - ショートカット: A(アクション切替), P(押し方切替), +/- (回数±1),
                      [ / ] (間隔 0.5x / 2x), .(間隔=0.01), D(複製), Delete(削除), G(グリッドスナップ切替)
    - グリッドスナップ（10px）
    """
    def __init__(self, master, stop_flag_ref, **kwargs):
        super().__init__(master, **kwargs)
        self.stop_flag_ref = stop_flag_ref

        # 状態
        self.blocks = {}             # id -> meta
        self.connections = []        # (from_id, to_id, line_id)
        self.connect_mode = False    # （配線はポートでもできるので任意）
        self.block_counter = 0
        self.current_block_id = None
        self.recent_keys = []

        # 追加：操作性向上用
        self.multi_selected = set()   # 複数選択
        self.grid_snap = True         # G で切替
        self.marquee_rect = None
        self.drag_select_origin = None
        self.wire_preview = None
        self.wire_from = None  # (block_id, side)

        # ツールバー
        toolbar = ctk.CTkFrame(self); toolbar.pack(fill="x", padx=8, pady=(8,6))
        ctk.CTkButton(toolbar, text="ブロック追加", command=lambda: self.add_block(select_after_add=True)).pack(side="left", padx=4)
        ctk.CTkButton(toolbar, text="接続モード", command=self.toggle_connect).pack(side="left", padx=4)
        ctk.CTkButton(toolbar, text="マクロ実行", command=self.run_macro).pack(side="left", padx=4)

        # 本体
        body = ctk.CTkFrame(self); body.pack(fill="both", expand=True, padx=8, pady=(0,8))

        # キャンバス
        c_area = ctk.CTkFrame(body); c_area.pack(side="left", fill="both", expand=True)
        self.canvas = ctk.CTkCanvas(c_area, bg="#2A2A2A", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        # インスペクタ
        self.inspector = ctk.CTkFrame(body, width=280)
        self.inspector.pack(side="right", fill="y", padx=(8,0))
        self._build_inspector(self.inspector)

        # バインド（ブロック）
        self.canvas.tag_bind("block", "<Button-1>", self._on_block_click)
        self.canvas.tag_bind("block", "<B1-Motion>", self._on_block_drag)
        self.canvas.tag_bind("block", "<ButtonRelease-1>", self._on_block_release)
        self.canvas.tag_bind("block", "<Double-Button-1>", self._on_block_rename)

        # バインド（キャンバス・全体）
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        self.canvas.bind("<Button-1>", self._on_canvas_mousedown)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_mouseup)

        # ショートカット
        self.canvas.bind_all("<Key>", self._on_key_shortcuts)
    def _item_exists(self, item_id) -> bool:
        """Canvas に item_id がまだ存在するかを安全に判定"""
        if not item_id:
            return False
        try:
            return self.canvas.type(item_id) != ""
        except Exception:
            return False
    # ========= インスペクタ =========
    def _build_inspector(self, parent):
        ctk.CTkLabel(parent, text="インスペクタ", font=("TkDefaultFont", 14, "bold")).pack(pady=(10,6))
        self.lbl_sel = ctk.CTkLabel(parent, text="選択中: なし"); self.lbl_sel.pack(pady=(0,8))

        self.var_action   = ctk.StringVar(value="左クリック")
        self.var_press    = ctk.StringVar(value="短押し")
        self.var_seconds  = ctk.StringVar(value="1.0")
        self.var_count    = ctk.StringVar(value="1")
        self.var_interval = ctk.StringVar(value="0.5")
        self.var_key      = ctk.StringVar(value="enter")

        ctk.CTkLabel(parent, text="アクション").pack(anchor="w", padx=12)
        ctk.CTkOptionMenu(parent, values=["左クリック","右クリック","ダブルクリック","キー入力"],
                          variable=self.var_action, command=lambda *_: self._apply_inspector()
                          ).pack(fill="x", padx=12, pady=(0,8))

        ctk.CTkLabel(parent, text="押し方").pack(anchor="w", padx=12)
        ctk.CTkOptionMenu(parent, values=["短押し","長押し"],
                          variable=self.var_press, command=lambda *_: self._apply_inspector()
                          ).pack(fill="x", padx=12, pady=(0,8))

        row1 = ctk.CTkFrame(parent); row1.pack(fill="x", padx=12, pady=(2,4))
        ctk.CTkLabel(row1, text="長押し(秒)").pack(side="left")
        e1 = ctk.CTkEntry(row1, textvariable=self.var_seconds, width=80); e1.pack(side="right")
        e1.bind("<FocusOut>", lambda *_: self._apply_inspector())

        row2 = ctk.CTkFrame(parent); row2.pack(fill="x", padx=12, pady=(2,4))
        ctk.CTkLabel(row2, text="回数").pack(side="left")
        e2 = ctk.CTkEntry(row2, textvariable=self.var_count, width=80); e2.pack(side="right")
        e2.bind("<FocusOut>", lambda *_: self._apply_inspector())

        row3 = ctk.CTkFrame(parent); row3.pack(fill="x", padx=12, pady=(2,10))
        ctk.CTkLabel(row3, text="間隔(秒)").pack(side="left")
        e3 = ctk.CTkEntry(row3, textvariable=self.var_interval, width=80); e3.pack(side="right")
        e3.bind("<FocusOut>", lambda *_: self._apply_inspector())

        ctk.CTkLabel(parent, text="キー").pack(anchor="w", padx=12)
        try:
            from customtkinter import CTkComboBox
            self.key_widget = CTkComboBox(parent, values=KEY_LIST, variable=self.var_key)
            self.key_widget.pack(fill="x", padx=12, pady=(0,4))
            self.key_widget.bind("<KeyRelease>", lambda *_: self._apply_inspector())
            self.key_widget.bind("<<ComboboxSelected>>", lambda *_: self._apply_inspector())
        except Exception:
            self.key_widget = ctk.CTkEntry(parent, textvariable=self.var_key)
            self.key_widget.pack(fill="x", padx=12, pady=(0,4))
            self.key_widget.bind("<KeyRelease>", lambda *_: self._apply_inspector())

        self.frame_recent = ctk.CTkFrame(parent); self.frame_recent.pack(fill="x", padx=12, pady=(4,8))
        ctk.CTkLabel(self.frame_recent, text="最近:").pack(side="left")
        self._update_recent_keys_ui()

    def _update_recent_keys_ui(self):
        for w in self.frame_recent.winfo_children()[1:]:
            w.destroy()
        for k in self.recent_keys[:5]:
            ctk.CTkButton(self.frame_recent, text=k, width=40,
                          command=lambda kk=k: (self.var_key.set(kk), self._apply_inspector())
                          ).pack(side="left", padx=2)

    def _apply_inspector(self):
        bid = self.current_block_id
        if not bid or bid not in self.blocks:
            return
        cfg = self.blocks[bid]['config']
        cfg['action'] = self.var_action.get()
        cfg['press_type'] = self.var_press.get()
        try: cfg['seconds'] = float(self.var_seconds.get())
        except: pass
        try: cfg['repeat_count'] = int(self.var_count.get())
        except: pass
        try: cfg['repeat_interval'] = float(self.var_interval.get())
        except: pass
        k = (self.var_key.get() or 'enter').strip()
        cfg['key'] = k

        if k and k in KEY_LIST:
            if k in self.recent_keys:
                self.recent_keys.remove(k)
            self.recent_keys.insert(0, k)
            self._update_recent_keys_ui()

        # 表示名更新
        short = cfg['action'] if cfg['action'] != "キー入力" else f"Key: {cfg['key']}"
        self.canvas.itemconfig(self.blocks[bid]['text_id'], text=short)

    # ========= キャンバス・ブロック =========
    def add_block(self, select_after_add=False):
        x, y, w, h = 60, 60, 160, 50
        bid = f"block_{self.block_counter}"; self.block_counter += 1

        r = self.canvas.create_rectangle(x, y, x+w, y+h,
                                         fill="#3A3A3A", outline="#5A5A5A", width=2,
                                         tags=("block", bid))
        t = self.canvas.create_text(x+w/2, y+h/2, text="左クリック", fill="white",
                                    tags=("block", bid))

        # 左右ポート（小円）
        lp = self.canvas.create_oval(x-6, y+h/2-6, x+6, y+h/2+6,
                                     fill="#7AA2F7", outline="", tags=("port", bid, "portL"))
        rp = self.canvas.create_oval(x+w-6, y+h/2-6, x+w+6, y+h/2+6,
                                     fill="#7AA2F7", outline="", tags=("port", bid, "portR"))
        self.canvas.tag_bind(lp, "<Button-1>", lambda e, b=bid: self._start_wire(b, "L"))
        self.canvas.tag_bind(rp, "<Button-1>", lambda e, b=bid: self._start_wire(b, "R"))
        self.canvas.tag_bind("port", "<B1-Motion>", self._drag_wire)
        self.canvas.tag_bind("port", "<ButtonRelease-1>", self._finish_wire)

        self.blocks[bid] = {
            'rect_id': r, 'text_id': t, 'x': x, 'y': y, 'w': w, 'h': h, 'dragging': False,
            'ports': {'L': lp, 'R': rp},
            'config': {'action': "左クリック", 'press_type': "短押し", 'seconds': 1.0,
                       'repeat_count': 1, 'repeat_interval': 0.5, 'key': 'enter'}
        }
        if select_after_add:
            self._select_block(bid)

    def toggle_connect(self):
        self.connect_mode = not self.connect_mode
        print(f"接続モード：{'ON' if self.connect_mode else 'OFF'}")

    def _on_block_click(self, event):
        items = self.canvas.find_withtag("current")
        if not items: return
        tags = self.canvas.gettags(items[0])
        bid = next((t for t in tags if t.startswith("block_")), None)
        if not bid or bid not in self.blocks: return

        if self.connect_mode:
            if self.current_block_id and self.current_block_id != bid:
                self._draw_connection(self.current_block_id, bid)
            self._select_block(bid)
        else:
            bx, by = self.blocks[bid]['x'], self.blocks[bid]['y']
            self.blocks[bid]['drag_offset_x'] = event.x - bx
            self.blocks[bid]['drag_offset_y'] = event.y - by
            self.blocks[bid]['dragging'] = True
            # Shift+クリックで複数選択
            if self._shift(event) and self.current_block_id != bid:
                self.multi_selected.add(bid)
                self._highlight(bid, True)
            else:
                self._select_block(bid)

    def _on_block_drag(self, event):
        items = self.canvas.find_withtag("current")
        if not items: return
        tags = self.canvas.gettags(items[0])
        bid = next((t for t in tags if t.startswith("block_")), None)
        if not bid or bid not in self.blocks: return
        if not self.blocks[bid].get('dragging'): return

        # 単体移動
        ox, oy = self.blocks[bid]['drag_offset_x'], self.blocks[bid]['drag_offset_y']
        nx, ny = event.x - ox, event.y - oy
        w, h = self.blocks[bid]['w'], self.blocks[bid]['h']

        self.canvas.coords(self.blocks[bid]['rect_id'], nx, ny, nx+w, ny+h)
        self.canvas.coords(self.blocks[bid]['text_id'], nx + w/2, ny + h/2)
        self.blocks[bid]['x'], self.blocks[bid]['y'] = nx, ny

        # ポートも追従
        lp = self.blocks[bid]['ports']['L']; rp = self.blocks[bid]['ports']['R']
        self.canvas.coords(lp, nx-6, ny+h/2-6, nx+6, ny+h/2+6)
        self.canvas.coords(rp, nx+w-6, ny+h/2-6, nx+w+6, ny+h/2+6)

        # 接続線更新
        for (f, t, lid) in list(self.connections):
            if f in self.blocks and t in self.blocks:
                fx = self.blocks[f]['x'] + self.blocks[f]['w']/2
                fy = self.blocks[f]['y'] + self.blocks[f]['h']/2
                tx = self.blocks[t]['x'] + self.blocks[t]['w']/2
                ty = self.blocks[t]['y'] + self.blocks[t]['h']/2
                self.canvas.coords(lid, fx, fy, tx, ty)

    def _on_block_release(self, event):
        items = self.canvas.find_withtag("current")
        if not items: return
        tags = self.canvas.gettags(items[0])
        bid = next((t for t in tags if t.startswith("block_")), None)
        if not bid or bid not in self.blocks: return
        self.blocks[bid]['dragging'] = False

        # グリッドスナップ
        if self.grid_snap:
            x, y = self.blocks[bid]['x'], self.blocks[bid]['y']
            nx, ny = round(x/10)*10, round(y/10)*10
            dx, dy = nx-x, ny-y
            if dx or dy:
                self._nudge_block(bid, dx, dy)

    def _nudge_block(self, bid, dx, dy):
        meta = self.blocks[bid]
        x, y, w, h = meta['x']+dx, meta['y']+dy, meta['w'], meta['h']
        self.canvas.coords(meta['rect_id'], x, y, x+w, y+h)
        self.canvas.coords(meta['text_id'], x+w/2, y+h/2)
        lp = meta['ports']['L']; rp = meta['ports']['R']
        self.canvas.coords(lp, x-6, y+h/2-6, x+6, y+h/2+6)
        self.canvas.coords(rp, x+w-6, y+h/2-6, x+w+6, y+h/2+6)
        meta['x'], meta['y'] = x, y
        # 線を更新
        for (f, t, lid) in list(self.connections):
            if f in self.blocks and t in self.blocks:
                fx = self.blocks[f]['x'] + self.blocks[f]['w']/2
                fy = self.blocks[f]['y'] + self.blocks[f]['h']/2
                tx = self.blocks[t]['x'] + self.blocks[t]['w']/2
                ty = self.blocks[t]['y'] + self.blocks[t]['h']/2
                self.canvas.coords(lid, fx, fy, tx, ty)

    # ========= キャンバス（空白側） =========
    def _on_canvas_mousedown(self, event):
        # 既存のマーキーが残っていたら掃除
        if self.marquee_rect and self._item_exists(self.marquee_rect):
            try:
                self.canvas.delete(self.marquee_rect)
            except Exception:
                pass
        self.marquee_rect = None
        self.drag_select_origin = None

        # 背景クリック時のみマーキー開始
        item = self.canvas.find_withtag("current")
        if not item:
            self.drag_select_origin = (event.x, event.y)
            self.marquee_rect = self.canvas.create_rectangle(
                event.x, event.y, event.x, event.y,
                outline="#58A6FF", dash=(2, 2)
            )
            return
        # 既存のブロッククリック時の挙動は今まで通り（何もしない）
        # Shift+クリックの複数選択は _on_block_click 側で処理済み

    def _on_canvas_drag(self, event):
        if not self.marquee_rect or not self._item_exists(self.marquee_rect) or not self.drag_select_origin:
            return
        x0, y0 = self.drag_select_origin
        self.canvas.coords(self.marquee_rect, x0, y0, event.x, event.y)

    def _on_canvas_mouseup(self, event):
        if not self.marquee_rect or not self._item_exists(self.marquee_rect):
            # すでに消えている（または未作成）なら何もしないでクリーンアップ
            self.marquee_rect = None
            self.drag_select_origin = None
            return

        coords = self.canvas.coords(self.marquee_rect) or []
        try:
            x0, y0, x1, y1 = coords
        except Exception:
            # 期待する4要素でなければスキップ
            try:
                self.canvas.delete(self.marquee_rect)
            except Exception:
                pass
            self.marquee_rect = None
            self.drag_select_origin = None
            return

        # マーキー消去
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

    # ========= インラインリネーム =========
    def _on_block_rename(self, event):
        items = self.canvas.find_withtag("current")
        if not items: return
        tags = self.canvas.gettags(items[0])
        bid = next((t for t in tags if t.startswith("block_")), None)
        if not bid: return
        tx, ty = self.canvas.coords(self.blocks[bid]['text_id'])
        entry = ctk.CTkEntry(self.canvas, width=140)
        entry.insert(0, self.canvas.itemcget(self.blocks[bid]['text_id'], "text"))
        self.canvas.create_window(tx, ty, window=entry, tags=("inline_edit",))
        entry.focus()

        def _commit(_=None):
            new = entry.get().strip() or "Block"
            self.canvas.itemconfig(self.blocks[bid]['text_id'], text=new)
            self.canvas.delete("inline_edit")
        entry.bind("<Return>", _commit)
        entry.bind("<Escape>", lambda e: self.canvas.delete("inline_edit"))
        entry.bind("<FocusOut>", _commit)

    # ========= 配線（ポート→ドラッグ→接続） =========
    def _start_wire(self, bid, side):
        # 既存プレビューが残っていれば掃除
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
        # 途中で release されて削除済みの場合がある
        if not self.wire_preview or not self._item_exists(self.wire_preview):
            self.wire_preview = None
            return
        coords = self.canvas.coords(self.wire_preview) or []
        if len(coords) < 2:
            return
        x0, y0 = coords[:2]
        self.canvas.coords(self.wire_preview, x0, y0, event.x, event.y)

    def _finish_wire(self, event):
        # 既に消えている可能性があるのでガード
        if not self.wire_preview or not self._item_exists(self.wire_preview):
            self.wire_preview = None
            self.wire_from = None
            return

        # プレビュー線を消す
        try:
            self.canvas.delete(self.wire_preview)
        except Exception:
            pass
        self.wire_preview = None

        # 開始元が無ければ終了
        if not self.wire_from:
            return
        bid_from, _ = self.wire_from
        self.wire_from = None

        # ドロップ先がブロックなら接続
        item = self.canvas.find_withtag("current")
        if not item:
            return
        tags = self.canvas.gettags(item)
        bid_to = next((t for t in tags if t.startswith("block_")), None)
        if bid_to and bid_to != bid_from:
            self._draw_connection(bid_from, bid_to)


    def _port_center(self, bid, side):
        x, y, w, h = self.blocks[bid]['x'], self.blocks[bid]['y'], self.blocks[bid]['w'], self.blocks[bid]['h']
        if side == "L":
            return (x, y+h/2)
        return (x+w, y+h/2)

    def _draw_connection(self, b1, b2):
        if b1 not in self.blocks or b2 not in self.blocks: return
        x1 = self.blocks[b1]['x'] + self.blocks[b1]['w']/2
        y1 = self.blocks[b1]['y'] + self.blocks[b1]['h']/2
        x2 = self.blocks[b2]['x'] + self.blocks[b2]['w']/2
        y2 = self.blocks[b2]['y'] + self.blocks[b2]['h']/2
        lid = self.canvas.create_line(x1, y1, x2, y2, arrow="last", fill="white", width=2)
        self.connections.append((b1, b2, lid))

    # ========= 選択＆ハイライト =========
    def _select_block(self, bid):
        # 既存のハイライト解除
        if self.current_block_id and self.current_block_id in self.blocks:
            self._highlight(self.current_block_id, False)
        for b in list(self.multi_selected):
            self._highlight(b, False)
        self.multi_selected.clear()

        self.current_block_id = bid
        self._highlight(bid, True)
        self.lbl_sel.configure(text=f"選択中: {bid}")
        self._load_to_inspector(bid)

    def _highlight(self, bid, on):
        self.canvas.itemconfig(self.blocks[bid]['rect_id'],
                               outline="#58A6FF" if on else "#5A5A5A",
                               width=3 if on else 2)

    def _load_to_inspector(self, bid):
        cfg = self.blocks[bid]['config']
        self.var_action.set(cfg.get('action', "左クリック"))
        self.var_press.set(cfg.get('press_type', "短押し"))
        self.var_seconds.set(str(cfg.get('seconds', 1.0)))
        self.var_count.set(str(cfg.get('repeat_count', 1)))
        self.var_interval.set(str(cfg.get('repeat_interval', 0.5)))
        self.var_key.set(cfg.get('key', 'enter'))

    def _shift(self, event):  # Tk の修飾キー状態を簡易判定
        try:
            return (event.state & 0x0001) != 0
        except:
            return False

    # ========= ショートカット =========
    def _on_key_shortcuts(self, event):
        if not self.current_block_id:
            return
        bids = {self.current_block_id} | set(self.multi_selected)

        def apply(fn):
            for b in bids:
                if b not in self.blocks: continue
                fn(self.blocks[b]['config'])
                cfg = self.blocks[b]['config']
                short = cfg['action'] if cfg['action'] != "キー入力" else f"Key: {cfg['key']}"
                self.canvas.itemconfig(self.blocks[b]['text_id'], text=short)

        key = (event.keysym or "").lower()
        if key == 'a':  # アクション切替
            order = ["左クリック","右クリック","ダブルクリック","キー入力"]
            def _next(cfg):
                i = order.index(cfg.get('action',"左クリック"))
                cfg['action'] = order[(i+1)%len(order)]
            apply(_next)

        elif key == 'p':  # 押し方切替
            def _toggle(cfg):
                cfg['press_type'] = "長押し" if cfg.get('press_type',"短押し")=="短押し" else "短押し"
            apply(_toggle)

        elif key in ('plus','equal'):  # +
            def _inc(cfg): cfg['repeat_count'] = max(1, cfg.get('repeat_count',1)+1)
            apply(_inc)

        elif key == 'minus':  # -
            def _dec(cfg): cfg['repeat_count'] = max(1, cfg.get('repeat_count',1)-1)
            apply(_dec)

        elif key == 'bracketleft':   # [
            def _half(cfg): cfg['repeat_interval'] = max(0.0, cfg.get('repeat_interval',0.5)*0.5)
            apply(_half)

        elif key == 'bracketright':  # ]
            def _double(cfg): cfg['repeat_interval'] = cfg.get('repeat_interval',0.5)*2
            apply(_double)

        elif key == 'period':  # .
            def _fast(cfg): cfg['repeat_interval'] = 0.01
            apply(_fast)

        elif key == 'd':  # 複製
            self._duplicate_blocks(bids)

        elif key in ('delete','backspace'):  # 削除
            self._delete_blocks(bids)

        elif key == 'g':  # グリッドスナップ切替
            self.grid_snap = not self.grid_snap
            print(f"グリッドスナップ: {'ON' if self.grid_snap else 'OFF'}")

    def _duplicate_blocks(self, bids):
        dx, dy = 20, 20
        mapping = {}
        for b in bids:
            if b not in self.blocks: continue
            meta = self.blocks[b]
            x, y, w, h = meta['x']+dx, meta['y']+dy, meta['w'], meta['h']
            nb = f"block_{self.block_counter}"; self.block_counter += 1
            r = self.canvas.create_rectangle(x,y,x+w,y+h, fill="#3A3A3A", outline="#5A5A5A", width=2, tags=("block", nb))
            t = self.canvas.create_text(x+w/2,y+h/2, text=self.canvas.itemcget(meta['text_id'], "text"), fill="white", tags=("block", nb))
            lp = self.canvas.create_oval(x-6, y+h/2-6, x+6, y+h/2+6, fill="#7AA2F7", outline="", tags=("port", nb, "portL"))
            rp = self.canvas.create_oval(x+w-6, y+h/2-6, x+w+6, y+h/2+6, fill="#7AA2F7", outline="", tags=("port", nb, "portR"))
            self.canvas.tag_bind(lp, "<Button-1>", lambda e, b=nb: self._start_wire(b, "L"))
            self.canvas.tag_bind(rp, "<Button-1>", lambda e, b=nb: self._start_wire(b, "R"))
            self.canvas.tag_bind("port", "<B1-Motion>", self._drag_wire)
            self.canvas.tag_bind("port", "<ButtonRelease-1>", self._finish_wire)

            self.blocks[nb] = {
                'rect_id': r, 'text_id': t, 'x': x, 'y': y, 'w': w, 'h': h, 'dragging': False,
                'ports': {'L': lp, 'R': rp},
                'config': dict(self.blocks[b]['config'])
            }
            mapping[b] = nb
        # 選択内接続のみ複製
        for (f, t, _) in list(self.connections):
            if f in bids and t in bids:
                nf, nt = mapping.get(f), mapping.get(t)
                if nf and nt:
                    self._draw_connection(nf, nt)

    def _delete_blocks(self, bids):
        bids = set(bids)

        # 進行中のプレビューやマーキーがあれば消去
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

    # ========= 実行 =========
    def run_macro(self):
        stop = self.stop_flag_ref
        incoming = {t for (_, t, _) in self.connections}
        start_nodes = [bid for bid in self.blocks if bid not in incoming]
        if not start_nodes:
            print("スタートブロックがありません。"); return

        visited = set()
        def dfs(bid):
            if bid in visited: return
            visited.add(bid)
            self._exec_block(bid, stop)
            for (f, t, _) in self.connections:
                if f == bid:
                    dfs(t)
        for s in start_nodes:
            dfs(s)

    def _exec_block(self, bid, stop):
        cfg = self.blocks[bid]['config']
        action = cfg.get('action', "左クリック")
        press  = cfg.get('press_type', "短押し")
        secs   = cfg.get('seconds', 1.0)
        count  = cfg.get('repeat_count', 1)
        itv    = cfg.get('repeat_interval', 0.5)
        key    = cfg.get('key', 'enter')

        if stop(): return
        flush_modifiers()  # 実行直前の修飾キー離れ待ち（mac の  対策）

        if action in ("左クリック","右クリック","ダブルクリック"):
            if press == "短押し":
                x, y = pag.position()
                for _ in range(max(1, count)):
                    if stop(): return
                    if action == "左クリック": pag.click(x,y)
                    elif action == "右クリック": pag.rightClick(x,y)
                    else: pag.doubleClick(x,y)
                    if busy_wait(itv, stop): return
            else:
                x, y = pag.position()
                pag.moveTo(x,y); pag.mouseDown()
                t0 = time.time()
                while time.time() - t0 < secs:
                    if stop() or esc_pressed():
                        pag.mouseUp(); return
                    time.sleep(0.01)
                pag.mouseUp()
        else:
            if press == "短押し":
                for _ in range(max(1, count)):
                    if stop(): return
                    pag.press(key)
                    if busy_wait(itv, stop): return
            else:
                pag.keyDown(key)
                t0 = time.time()
                while time.time() - t0 < secs:
                    if stop() or esc_pressed():
                        pag.keyUp(key); return
                    time.sleep(0.01)
                pag.keyUp(key)

        print(f"{bid} 実行完了")
