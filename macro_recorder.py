import json
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pyautogui

# Make PyAutoGUI as fast as possible (removes built-in per-call pauses)
pyautogui.PAUSE = 0
pyautogui.MINIMUM_DURATION = 0
pyautogui.MINIMUM_SLEEP = 0

try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except Exception:
    KEYBOARD_AVAILABLE = False


class MacroApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Simple Macro Recorder")

        # State
        self.recording = False
        self.events = []  # {"key": str, "delay": float(seconds)}
        self._last_time = None
        self._play_thread = None
        self._stop_playback = threading.Event()

        # Inline editor state
        self._edit_entry = None
        self._edit_iid = None

        # Options
        self.ignore_keys = {"f9", "f10", "esc"}
        self.use_hotkeys = tk.BooleanVar(value=True)
        self.playback_speed = tk.DoubleVar(value=1.0)

        # Toggle key: can be "f8" or "scan:<code>"
        self.play_toggle_key = tk.StringVar(value="f8")

        # Repeat options
        self.repeat_enabled = tk.BooleanVar(value=False)
        self.repeat_delay_ms = tk.IntVar(value=250)

        # Capture-mode flag for setting toggle key
        self._capturing_toggle_key = False

        # NEW: robust toggle detection (manual, via hook)
        self.toggle_scan_code = None
        self.toggle_key_name = None
        self._toggle_pressed_guard = False
        self._resolve_toggle_key()

        # UI
        self._build_ui()

        # Hook keyboard if possible
        if KEYBOARD_AVAILABLE:
            keyboard.hook(self._on_key_event)
            self._setup_hotkeys()  # only F9/F10/ESC; toggle handled manually
        else:
            self.use_hotkeys.set(False)
            self._set_status("keyboard module not available. Recording hotkeys disabled.")

        # Clean shutdown
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------- Helpers: ms formatting ----------------

    @staticmethod
    def sec_to_ms_int(seconds: float) -> int:
        return int(round(max(0.0, float(seconds)) * 1000.0))

    @staticmethod
    def ms_int_to_sec(ms: int) -> float:
        return max(0, int(ms)) / 1000.0

    def _resolve_toggle_key(self):
        """
        Parse play_toggle_key into either a scan code (best for dead keys) or a name.
        Examples:
          "scan:41" -> scan code 41
          "f8"      -> name "f8"
        """
        val = str(self.play_toggle_key.get()).strip().lower()
        self.toggle_scan_code = None
        self.toggle_key_name = None

        if val.startswith("scan:"):
            try:
                self.toggle_scan_code = int(val.split(":", 1)[1])
            except Exception:
                self.toggle_scan_code = None
        elif val:
            self.toggle_key_name = val

    # ---------------- UI ----------------

    def _build_ui(self):
        frm = ttk.Frame(self.root, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Buttons row
        btns = ttk.Frame(frm)
        btns.grid(row=0, column=0, sticky="ew")

        self.btn_record = ttk.Button(btns, text="Start Recording", command=self.toggle_recording)
        self.btn_record.grid(row=0, column=0, padx=(0, 8))

        self.btn_play = ttk.Button(btns, text="Play", command=self.play_macro)
        self.btn_play.grid(row=0, column=1, padx=(0, 8))

        self.btn_stop = ttk.Button(btns, text="Stop", command=self.stop_playback)
        self.btn_stop.grid(row=0, column=2, padx=(0, 8))

        self.btn_clear = ttk.Button(btns, text="Clear", command=self.clear_macro)
        self.btn_clear.grid(row=0, column=3, padx=(0, 8))

        self.btn_save = ttk.Button(btns, text="Save", command=self.save_macro)
        self.btn_save.grid(row=0, column=4, padx=(0, 8))

        self.btn_load = ttk.Button(btns, text="Load", command=self.load_macro)
        self.btn_load.grid(row=0, column=5, padx=(0, 8))

        opts = ttk.Frame(frm)
        opts.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        opts.columnconfigure(6, weight=1)

        self.chk_hotkeys = ttk.Checkbutton(
            opts,
            text="Enable hotkeys (F9 record, F10 play, ESC stop playback; toggle handled by manual hook)",
            variable=self.use_hotkeys,
            command=self._hotkeys_toggled
        )
        self.chk_hotkeys.grid(row=0, column=0, sticky="w", columnspan=7)

        ttk.Label(opts, text="Playback speed:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.speed = ttk.Scale(opts, from_=0.25, to=3.0, variable=self.playback_speed, orient="horizontal")
        self.speed.grid(row=1, column=1, sticky="ew", pady=(8, 0), padx=(8, 8), columnspan=2)

        self.speed_val = ttk.Label(opts, text="1.00x")
        self.speed_val.grid(row=1, column=3, sticky="w", pady=(8, 0))
        self.speed.bind("<Motion>", lambda _e: self.speed_val.config(text=f"{self.playback_speed.get():.2f}x"))
        self.speed.bind("<ButtonRelease-1>", lambda _e: self.speed_val.config(text=f"{self.playback_speed.get():.2f}x"))

        # Toggle hotkey controls
        ttk.Label(opts, text="Play toggle hotkey:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.toggle_entry = ttk.Entry(opts, textvariable=self.play_toggle_key, width=14)
        self.toggle_entry.grid(row=2, column=1, padx=(8, 8), sticky="w", pady=(8, 0))

        self.btn_apply_hotkey = ttk.Button(opts, text="Apply", command=self.apply_toggle_hotkey)
        self.btn_apply_hotkey.grid(row=2, column=2, padx=(0, 8), sticky="w", pady=(8, 0))

        self.btn_capture_hotkey = ttk.Button(opts, text="Set (press key)", command=self.capture_toggle_hotkey)
        self.btn_capture_hotkey.grid(row=2, column=3, padx=(0, 8), sticky="w", pady=(8, 0))

        # Repeat controls
        self.chk_repeat = ttk.Checkbutton(opts, text="Repeat playback", variable=self.repeat_enabled)
        self.chk_repeat.grid(row=2, column=4, sticky="w", pady=(8, 0))

        ttk.Label(opts, text="Repeat delay (ms):").grid(row=2, column=5, padx=(16, 0), sticky="w", pady=(8, 0))
        self.repeat_delay_entry = ttk.Entry(opts, textvariable=self.repeat_delay_ms, width=8)
        self.repeat_delay_entry.grid(row=2, column=6, padx=(8, 0), sticky="w", pady=(8, 0))

        # Treeview
        list_frame = ttk.LabelFrame(frm, text="Recorded steps (Delay is ms; double-click Delay to edit)")
        list_frame.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        frm.rowconfigure(3, weight=1)

        columns = ("step", "key", "delay_ms")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("step", text="#")
        self.tree.heading("key", text="Key")
        self.tree.heading("delay_ms", text="Delay (ms)")
        self.tree.column("step", width=60, anchor="e", stretch=False)
        self.tree.column("key", width=180, anchor="w", stretch=True)
        self.tree.column("delay_ms", width=120, anchor="e", stretch=False)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", pady=8)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Button-1>", self._on_tree_single_click)

        self.status = ttk.Label(frm, text="Ready.", anchor="w")
        self.status.grid(row=4, column=0, sticky="ew", pady=(10, 0))

        hint = (
            "Toggle key works reliably even for dead keys (like ^) because it uses scan codes.\n"
            "Use 'Set (press key)' to capture the key. Press once to start, again to stop.\n"
            "Delays are edited in milliseconds."
        )
        ttk.Label(frm, text=hint, foreground="#555", anchor="w", justify="left").grid(
            row=5, column=0, sticky="ew", pady=(8, 0)
        )

        if not KEYBOARD_AVAILABLE:
            self.chk_hotkeys.state(["disabled"])
            self.btn_apply_hotkey.state(["disabled"])
            self.btn_capture_hotkey.state(["disabled"])
            self._set_status("keyboard module not installed/usable. Install 'keyboard' to record keystrokes globally.")

    def _set_status(self, msg: str):
        self.status.config(text=msg)

    # ---------------- Recording ----------------

    def toggle_recording(self):
        if not KEYBOARD_AVAILABLE:
            messagebox.showerror("Not available", "Global keystroke recording requires the 'keyboard' module.")
            return
        if self.recording:
            self.stop_recording()
        else:
            self.start_recording()

    def start_recording(self):
        if self._play_thread and self._play_thread.is_alive():
            messagebox.showwarning("Busy", "Stop playback before recording.")
            return

        self._end_inline_edit(commit=True)
        self.events.clear()
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        self.recording = True
        self._last_time = time.perf_counter()
        self.btn_record.config(text="Stop Recording")
        self._set_status("Recording ON — press keys now (F9 to stop if hotkeys enabled).")

    def stop_recording(self):
        self.recording = False
        self._last_time = None
        self.btn_record.config(text="Start Recording")
        self._set_status(f"Recording OFF — captured {len(self.events)} steps.")

    # ---------------- Key hook ----------------

    def _on_key_event(self, e):
        # 1) Capture toggle hotkey (scan code) when in capture mode
        if self._capturing_toggle_key and e.event_type == "down":
            sc = getattr(e, "scan_code", None)
            if sc is not None:
                self.root.after(0, lambda: self._finish_capture_toggle_key(f"scan:{sc}"))
            else:
                name = (e.name or "").lower()
                if name:
                    self.root.after(0, lambda: self._finish_capture_toggle_key(name))
            return

        # 2) Manual toggle detection (robust; works even after other keys pressed)
        if self.use_hotkeys.get():
            is_toggle = False

            if self.toggle_scan_code is not None and getattr(e, "scan_code", None) == self.toggle_scan_code:
                is_toggle = True
            elif self.toggle_scan_code is None and self.toggle_key_name and (e.name or "").lower() == self.toggle_key_name:
                is_toggle = True

            if is_toggle:
                # Fire only on key DOWN, with guard to prevent repeats while held
                if e.event_type == "down" and not self._toggle_pressed_guard:
                    self._toggle_pressed_guard = True
                    self.root.after(0, self.toggle_playback)
                elif e.event_type == "up":
                    self._toggle_pressed_guard = False
                return

        # 3) Normal macro recording
        if not self.recording or e.event_type != "down":
            return

        key = (e.name or "").lower()
        if not key or key in self.ignore_keys:
            return

        t = time.perf_counter()
        delay_sec = t - (self._last_time if self._last_time is not None else t)
        self._last_time = t

        self.events.append({"key": key, "delay": float(delay_sec)})

        idx = len(self.events)
        delay_ms = self.sec_to_ms_int(delay_sec)
        self.root.after(0, lambda: self.tree.insert("", "end", values=(f"{idx:03d}", key, f"{delay_ms}")))

    # ---------------- Inline editing (Delay column, ms) ----------------

    def _on_tree_single_click(self, event):
        if self._edit_entry is not None:
            region = self.tree.identify("region", event.x, event.y)
            if region != "cell":
                self._end_inline_edit(commit=True)
            else:
                col = self.tree.identify_column(event.x)
                if col != "#3":
                    self._end_inline_edit(commit=True)

    def _on_tree_double_click(self, event):
        if self.recording:
            return
        row_iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_iid or col != "#3":
            return
        self._begin_edit_delay_cell(row_iid)

    def _begin_edit_delay_cell(self, iid: str):
        self._end_inline_edit(commit=True)
        bbox = self.tree.bbox(iid, column="delay_ms")
        if not bbox:
            return
        x, y, w, h = bbox
        current_ms_text = self.tree.set(iid, "delay_ms")

        self._edit_iid = iid
        self._edit_entry = ttk.Entry(self.tree)
        self._edit_entry.place(x=x, y=y, width=w, height=h)
        self._edit_entry.insert(0, current_ms_text)
        self._edit_entry.select_range(0, tk.END)
        self._edit_entry.focus()

        self._edit_entry.bind("<Return>", lambda _e: self._end_inline_edit(commit=True))
        self._edit_entry.bind("<Escape>", lambda _e: self._end_inline_edit(commit=False))
        self._edit_entry.bind("<FocusOut>", lambda _e: self._end_inline_edit(commit=True))

    def _end_inline_edit(self, commit: bool):
        if self._edit_entry is None or self._edit_iid is None:
            return

        iid = self._edit_iid
        entry = self._edit_entry
        new_text = entry.get().strip()

        self._edit_entry = None
        self._edit_iid = None
        entry.destroy()

        if not commit:
            return

        try:
            if new_text == "":
                raise ValueError
            new_ms = int(new_text)
            if new_ms < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid delay", "Enter a non-negative integer (milliseconds), e.g. 55 or 555.")
            return

        self.tree.set(iid, "delay_ms", str(new_ms))

        children = list(self.tree.get_children(""))
        try:
            idx = children.index(iid)
        except ValueError:
            return

        if 0 <= idx < len(self.events):
            self.events[idx]["delay"] = self.ms_int_to_sec(new_ms)

    # ---------------- Playback ----------------

    def toggle_playback(self):
        if self._play_thread and self._play_thread.is_alive():
            self.stop_playback()
        else:
            self.play_macro()

    def play_macro(self):
        if self.recording:
            messagebox.showwarning("Recording", "Stop recording before playback.")
            return
        if not self.events:
            messagebox.showinfo("Empty", "No macro recorded.")
            return
        if self._play_thread and self._play_thread.is_alive():
            return

        self._end_inline_edit(commit=True)

        if self.repeat_enabled.get():
            try:
                rd = int(self.repeat_delay_ms.get())
                if rd < 0:
                    raise ValueError
            except Exception:
                messagebox.showerror("Invalid repeat delay", "Repeat delay must be a non-negative integer (ms).")
                return

        self._stop_playback.clear()
        self._play_thread = threading.Thread(target=self._play_worker, daemon=True)
        self._play_thread.start()
        self._set_status("Playing... (toggle key stops)")

    def _play_worker(self):
        speed = max(0.01, float(self.playback_speed.get()))
        repeat = bool(self.repeat_enabled.get())
        repeat_delay_s = self.ms_int_to_sec(int(self.repeat_delay_ms.get() or 0))

        while True:
            for step in self.events:
                if self._stop_playback.is_set():
                    break
                delay = step["delay"] / speed
                time.sleep(max(0.0, delay))
                try:
                    pyautogui.press(step["key"])
                except Exception:
                    pass

            if self._stop_playback.is_set():
                break
            if not repeat:
                break

            if repeat_delay_s > 0:
                end_t = time.time() + repeat_delay_s
                while time.time() < end_t:
                    if self._stop_playback.is_set():
                        break
                    time.sleep(0.02)
                if self._stop_playback.is_set():
                    break

        self.root.after(0, lambda: self._set_status(
            "Playback finished." if not self._stop_playback.is_set() else "Playback stopped."
        ))

    def stop_playback(self):
        self._stop_playback.set()

    # ---------------- Hotkey capture for dead keys ----------------

    def capture_toggle_hotkey(self):
        if not KEYBOARD_AVAILABLE:
            messagebox.showerror("Not available", "Hotkeys require the 'keyboard' module.")
            return
        self._capturing_toggle_key = True
        self._set_status("Press the key you want to use as Play Toggle hotkey...")

    def _finish_capture_toggle_key(self, key_id: str):
        self._capturing_toggle_key = False
        self.play_toggle_key.set(str(key_id).strip().lower())
        self._resolve_toggle_key()
        self._set_status(f"Play toggle hotkey set to: {self.play_toggle_key.get()}")

    # ---------------- Save/Load/Clear ----------------

    def clear_macro(self):
        if self.recording:
            messagebox.showwarning("Recording", "Stop recording first.")
            return
        self._end_inline_edit(commit=True)
        self.events.clear()
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self._set_status("Cleared.")

    def save_macro(self):
        if not self.events:
            messagebox.showinfo("Empty", "Nothing to save.")
            return
        self._end_inline_edit(commit=True)

        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="Save macro"
        )
        if not path:
            return

        data = {
            "version": 2,
            "events": self.events,
            "repeat_enabled": bool(self.repeat_enabled.get()),
            "repeat_delay_ms": int(self.repeat_delay_ms.get()),
            "play_toggle_key": str(self.play_toggle_key.get()).strip().lower() or "f8",
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        self._set_status(f"Saved to {path}")

    def load_macro(self):
        if self.recording:
            messagebox.showwarning("Recording", "Stop recording first.")
            return
        self._end_inline_edit(commit=True)

        path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            title="Load macro"
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            events = data.get("events", [])
            if not isinstance(events, list):
                raise ValueError("Invalid file format")

            cleaned = []
            for ev in events:
                if not isinstance(ev, dict):
                    continue
                k = str(ev.get("key", "")).lower()
                d = float(ev.get("delay", 0.0))
                if k:
                    cleaned.append({"key": k, "delay": max(0.0, d)})
            self.events = cleaned

            self.repeat_enabled.set(bool(data.get("repeat_enabled", False)))
            try:
                self.repeat_delay_ms.set(max(0, int(data.get("repeat_delay_ms", 250))))
            except Exception:
                pass

            tkey = str(data.get("play_toggle_key", "f8")).strip().lower()
            if tkey:
                self.play_toggle_key.set(tkey)
            self._resolve_toggle_key()

            for iid in self.tree.get_children():
                self.tree.delete(iid)

            for i, ev in enumerate(self.events, start=1):
                ms = self.sec_to_ms_int(ev["delay"])
                self.tree.insert("", "end", values=(f"{i:03d}", ev["key"], str(ms)))

            if KEYBOARD_AVAILABLE:
                self._setup_hotkeys()

            self._set_status(f"Loaded {len(self.events)} steps from {path}")
        except Exception as ex:
            messagebox.showerror("Load failed", str(ex))

    # ---------------- Hotkeys (F9/F10/ESC only) ----------------

    def apply_toggle_hotkey(self):
        if not KEYBOARD_AVAILABLE:
            messagebox.showerror("Not available", "Hotkeys require the 'keyboard' module.")
            return
        key = str(self.play_toggle_key.get()).strip().lower()
        if not key:
            messagebox.showerror("Invalid", "Toggle hotkey cannot be empty.")
            return

        self.play_toggle_key.set(key)
        self._resolve_toggle_key()
        self._set_status(f"Applied play toggle hotkey: {key}")

    def _setup_hotkeys(self):
        # Only set F9/F10/ESC here. Toggle key is handled in _on_key_event manually.
        try:
            keyboard.clear_all_hotkeys()
        except Exception:
            pass

        if not self.use_hotkeys.get():
            return

        keyboard.add_hotkey("f9", lambda: self.root.after(0, self.toggle_recording))
        keyboard.add_hotkey("f10", lambda: self.root.after(0, self.play_macro))
        keyboard.add_hotkey("esc", lambda: self.root.after(0, self.stop_playback))

    def _hotkeys_toggled(self):
        if not KEYBOARD_AVAILABLE:
            return
        self._setup_hotkeys()
        self._set_status("Hotkeys enabled." if self.use_hotkeys.get() else "Hotkeys disabled.")

    def on_close(self):
        try:
            self.stop_playback()
        except Exception:
            pass
        try:
            if KEYBOARD_AVAILABLE:
                keyboard.unhook_all()
                keyboard.clear_all_hotkeys()
        except Exception:
            pass
        self.root.destroy()

def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    MacroApp(root)
    root.geometry("880x540")
    root.mainloop()

if __name__ == "__main__":
    main()
