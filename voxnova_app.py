import os
import sys
import json
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import scrolledtext

import pyttsx3
import ttkbootstrap as tb
from ttkbootstrap.style import Style as TBStyle

# App constants
APP_NAME = "VoxNova Studio"
APP_TITLE = f"{APP_NAME} — Text to Speech"
CONFIG_BASENAME = ".voxnova_config.json"

# Fix ttkbootstrap stale style when re-running in the same process (Jupyter/Spyder)
def _reset_ttkbootstrap_style_if_stale():
    try:
        inst = getattr(TBStyle, "instance", None)
        if inst is not None:
            alive = True
            try:
                alive = bool(inst.master and inst.master.winfo_exists())
            except Exception:
                alive = False
            if not alive:
                TBStyle.instance = None
    except Exception:
        pass


class VoxNovaApp:
    CONFIG_PATH = os.path.join(os.path.expanduser("~"), CONFIG_BASENAME)

    # -------------------- Init --------------------
    def __init__(self):
        _reset_ttkbootstrap_style_if_stale()

        cfg = self._load_config_static()
        theme = "flatly" if not cfg.get("dark_mode", False) else "darkly"

        try:
            self.root = tb.Window(themename=theme)
        except Exception:
            self.root = tk.Tk()
            try:
                tb.Style().theme_use(theme)
            except Exception:
                pass

        self.root.title(APP_TITLE)

        # Minimum sizes
        self.MIN_W, self.MIN_H = 880, 560

        # Geometry restore (width/height/x/y + clamp + zoom state)
        self._normal_w = cfg.get("w")
        self._normal_h = cfg.get("h")
        self._normal_x = cfg.get("x")
        self._normal_y = cfg.get("y")

        # Back-compat: parse old "normal_geometry" string
        if (self._normal_w is None or self._normal_h is None) and cfg.get("normal_geometry"):
            try:
                size_part = cfg["normal_geometry"].split("+")[0]
                w_str, h_str = size_part.split("x")
                self._normal_w = int(w_str)
                self._normal_h = int(h_str)
            except Exception:
                pass

        w = self._normal_w if isinstance(self._normal_w, int) else 1040
        h = self._normal_h if isinstance(self._normal_h, int) else 690
        x = self._normal_x if isinstance(self._normal_x, int) else None
        y = self._normal_y if isinstance(self._normal_y, int) else None

        self._apply_initial_geometry(w, h, x, y)
        try:
            if cfg.get("zoomed", False):
                self._set_zoomed(True)
        except Exception:
            pass
        self.root.minsize(self.MIN_W, self.MIN_H)
        self.root.bind("<Configure>", self._on_configure)

        # State
        self.is_busy = False
        self.engine = None
        self.engine_lock = threading.Lock()

        self.voices = []
        self.voice_display_map = {}   # display -> voice.id
        self.selected_voice_id = cfg.get("voice_id", None)
        self._desired_voice_id = self.selected_voice_id
        self.dark_mode = cfg.get("dark_mode", False)

        # Control variables
        self.rate_var = tk.IntVar(value=cfg.get("rate", 160))     # approx WPM
        self.volume_var = tk.IntVar(value=cfg.get("volume", 100)) # 0–100
        self.filter_var = tk.StringVar(value="")
        self.gender_var = tk.StringVar(value="Any")               # Any/Male/Female
        self.font_size = tk.IntVar(value=cfg.get("font_size", 13))

        # Debounce for stats update
        self._stats_after_id = None

        # File defaults
        self.default_ext, self.file_filters = self._audio_defaults()

        # Build UI
        self._build_ui()
        self._bind_shortcuts()

        # Apply theme text colors and font
        self._apply_text_look()
        self.root.after_idle(self._refresh_button_styles)
        self.root.after(150, self._refresh_button_styles)

        # Load system voices asynchronously
        threading.Thread(target=self._load_voices, daemon=True).start()

        # Events
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

        # Initial stats
        self._update_stats()

    # -------------------- Config --------------------
    @classmethod
    def _load_config_static(cls):
        try:
            if os.path.exists(cls.CONFIG_PATH):
                with open(cls.CONFIG_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_config(self):
        # Save numeric width/height/x/y and zoomed state
        try:
            cur_state = self.root.state()
            if cur_state == "normal" and not self._is_zoomed():
                self._normal_w = max(self.MIN_W, int(self.root.winfo_width()))
                self._normal_h = max(self.MIN_H, int(self.root.winfo_height()))
                self._normal_x = int(self.root.winfo_x())
                self._normal_y = int(self.root.winfo_y())
        except Exception:
            pass

        data = {
            "rate": int(self.rate_var.get()),
            "volume": int(self.volume_var.get()),
            "voice_id": self.selected_voice_id,
            "dark_mode": self.dark_mode,
            "font_size": int(self.font_size.get()),
            "w": int(self._normal_w) if self._normal_w else 1040,
            "h": int(self._normal_h) if self._normal_h else 690,
            "x": int(self._normal_x) if self._normal_x is not None else None,
            "y": int(self._normal_y) if self._normal_y is not None else None,
            "zoomed": self._is_zoomed(),
            "normal_geometry": f"{int(self._normal_w) if self._normal_w else 1040}x{int(self._normal_h) if self._normal_h else 690}",
        }
        try:
            with open(self.CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def _on_configure(self, _=None):
        try:
            if not self._is_zoomed() and self.root.state() == "normal":
                w = max(self.MIN_W, int(self.root.winfo_width()))
                h = max(self.MIN_H, int(self.root.winfo_height()))
                x = int(self.root.winfo_x())
                y = int(self.root.winfo_y())
                self._normal_w, self._normal_h, self._normal_x, self._normal_y = w, h, x, y
        except Exception:
            pass

    # Zoom helpers
    def _is_zoomed(self) -> bool:
        try:
            if self.root.state() == "zoomed":
                return True
        except Exception:
            pass
        try:
            val = self.root.attributes("-zoomed")
            if isinstance(val, bool):
                return val
        except Exception:
            pass
        return False

    def _set_zoomed(self, value: bool):
        try:
            self.root.state("zoomed" if value else "normal")
            return
        except Exception:
            pass
        try:
            self.root.attributes("-zoomed", bool(value))
        except Exception:
            pass

    def _apply_initial_geometry(self, w: int, h: int, x: int | None, y: int | None):
        try:
            self.root.update_idletasks()
            sw = int(self.root.winfo_screenwidth())
            sh = int(self.root.winfo_screenheight())

            w = max(self.MIN_W, min(int(w), max(self.MIN_W, sw - 80)))
            h = max(self.MIN_H, min(int(h), max(self.MIN_H, sh - 120)))

            def valid_pos(px, py):
                return px is not None and py is not None and (0 <= px <= sw - w) and (0 <= py <= sh - h)

            if not valid_pos(x, y):
                x = max(0, (sw - w) // 2)
                y = max(0, (sh - h) // 2)

            self.root.geometry(f"{w}x{h}+{x}+{y}")
            self._normal_w, self._normal_h, self._normal_x, self._normal_y = w, h, x, y
        except Exception:
            self.root.geometry("1040x690")

    # -------------------- Platform Helpers --------------------
    def _audio_defaults(self):
        if sys.platform == "darwin":
            return ".aiff", [("AIFF Audio File", "*.aiff"), ("WAV Audio File", "*.wav")]
        else:
            return ".wav", [("WAV Audio File", "*.wav")]

    # -------------------- UI --------------------
    def _build_ui(self):
        header = tb.Frame(self.root, bootstyle="primary", padding=(16, 12))
        header.pack(side=tk.TOP, fill=tk.X)
        tb.Label(header, text=APP_TITLE,
                 font=("Segoe UI", 20, "bold"),
                 bootstyle="inverse-primary").pack(anchor=tk.W)
        tb.Label(header, text="Offline TTS powered by pyttsx3 • Choose a voice, then Speak or Save",
                 bootstyle="inverse-primary").pack(anchor=tk.W, pady=(2, 0))

        toolbar = ttk.Frame(self.root, padding=(16, 10, 16, 8))
        toolbar.pack(side=tk.TOP, fill=tk.X)

        self.btn_speak = tb.Button(toolbar, text="🔊 Speak  (Ctrl+Enter)",
                                   bootstyle="primary", command=self.start_speaking, width=20)
        self.btn_stop  = tb.Button(toolbar, text="⏹ Stop  (Esc)",
                                   bootstyle="danger", command=self.stop_now, state="disabled", width=16)
        self.btn_save  = tb.Button(toolbar, text="💾 Save Audio  (Ctrl+S)",
                                   bootstyle="success", command=self.start_saving, width=20)
        self.btn_open  = tb.Button(toolbar, text="📂 Open File  (Ctrl+O)",
                                   bootstyle="info", command=self.open_text_file, width=18)
        self.btn_clear = tb.Button(toolbar, text="🧹 Clear  (Ctrl+L)",
                                   bootstyle="warning", command=self.clear_text, width=14)
        self.btn_theme = tb.Button(toolbar, text=("☀️ Light Mode" if self.dark_mode else "🌙 Dark Mode"),
                                   bootstyle="secondary", command=self.toggle_theme, width=16)

        self.btn_speak.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_stop.pack(side=tk.LEFT, padx=8)
        self.btn_save.pack(side=tk.LEFT, padx=8)
        self.btn_open.pack(side=tk.LEFT, padx=8)
        self.btn_clear.pack(side=tk.LEFT, padx=8)
        self.btn_theme.pack(side=tk.RIGHT)

        paned = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        paned.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=12, pady=8)

        left = ttk.Frame(paned)
        paned.add(left, weight=0)
        right = ttk.Frame(paned)
        paned.add(right, weight=1)

        controls = tb.Labelframe(left, text="Controls", padding=(12, 10, 12, 12), bootstyle="secondary")
        controls.pack(side=tk.TOP, fill=tk.Y, padx=(4, 8), pady=(0, 8))

        ttk.Label(controls, text="Search Voice").grid(row=0, column=0, sticky=tk.W, pady=(0, 2))
        filter_entry = ttk.Entry(controls, textvariable=self.filter_var)
        filter_entry.grid(row=1, column=0, sticky="we")
        self.filter_var.trace_add("write", lambda *_: self._apply_voice_filter())
        tooltip(filter_entry, "Filter voices by name/id")

        row2 = ttk.Frame(controls)
        row2.grid(row=2, column=0, sticky="we", pady=(8, 2))
        ttk.Label(row2, text="Gender:").pack(side=tk.LEFT)
        self.gender_combo = ttk.Combobox(row2, state="readonly", width=12, values=["Any", "Male", "Female"])
        self.gender_combo.set(self.gender_var.get())
        self.gender_combo.pack(side=tk.LEFT, padx=(6, 0))
        self.gender_combo.bind("<<ComboboxSelected>>", lambda _=None: (self.gender_var.set(self.gender_combo.get()), self._apply_voice_filter()))

        ttk.Label(controls, text="Voice").grid(row=3, column=0, sticky=tk.W, pady=(8, 2))
        self.voice_combo = ttk.Combobox(controls, state="readonly", width=48, values=["Loading voices..."])
        self.voice_combo.grid(row=4, column=0, sticky="we")
        self.voice_combo.bind("<<ComboboxSelected>>", self._on_voice_change)

        voice_actions = ttk.Frame(controls)
        voice_actions.grid(row=5, column=0, sticky="we", pady=(8, 8))
        tb.Button(voice_actions, text="🔄 Rescan", bootstyle="secondary",
                  command=lambda: threading.Thread(target=self._load_voices, daemon=True).start()).pack(side=tk.LEFT)
        tb.Button(voice_actions, text="🎙 Preview", bootstyle="secondary", command=self.preview_voice).pack(side=tk.LEFT, padx=(8, 0))
        tb.Button(voice_actions, text="➕ Get Voices", bootstyle="link", command=self._open_voice_settings).pack(side=tk.LEFT, padx=(8, 0))
        # MS Mark button removed

        ttk.Label(controls, text="Rate (WPM)").grid(row=6, column=0, sticky=tk.W)
        rate_row = ttk.Frame(controls)
        rate_row.grid(row=7, column=0, sticky="we")
        self.rate_scale = ttk.Scale(rate_row, from_=80, to=220, orient=tk.HORIZONTAL, command=lambda _=None: self._on_rate_change())
        self.rate_scale.set(self.rate_var.get())
        self.rate_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.rate_value_label = ttk.Label(rate_row, text=f"{self.rate_var.get()}")
        self.rate_value_label.pack(side=tk.LEFT, padx=(6, 0))

        ttk.Label(controls, text="Volume (%)").grid(row=8, column=0, sticky=tk.W, pady=(8, 0))
        vol_row = ttk.Frame(controls)
        vol_row.grid(row=9, column=0, sticky="we")
        self.volume_scale = ttk.Scale(vol_row, from_=0, to=100, orient=tk.HORIZONTAL, command=lambda _=None: self._on_volume_change())
        self.volume_scale.set(self.volume_var.get())
        self.volume_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.volume_value_label = ttk.Label(vol_row, text=f"{self.volume_var.get()}%")
        self.volume_value_label.pack(side=tk.LEFT, padx=(6, 0))

        ttk.Label(controls, text="Text Size").grid(row=10, column=0, sticky=tk.W, pady=(8, 0))
        size_row = ttk.Frame(controls)
        size_row.grid(row=11, column=0, sticky="we")
        tb.Button(size_row, text="A-", bootstyle="secondary", width=4, command=lambda: self._bump_font(-1)).pack(side=tk.LEFT)
        ttk.Label(size_row, textvariable=self.font_size, width=3, anchor="center").pack(side=tk.LEFT, padx=4)
        tb.Button(size_row, text="A+", bootstyle="secondary", width=4, command=lambda: self._bump_font(1)).pack(side=tk.LEFT)

        stats = tb.Labelframe(left, text="Stats", padding=(12, 10, 12, 10), bootstyle="secondary")
        stats.pack(side=tk.TOP, fill=tk.X, padx=(4, 8))
        self.stats_label = ttk.Label(stats, text="Words: 0   Chars: 0   ~00:00")
        self.stats_label.pack(anchor=tk.W)

        controls.columnconfigure(0, weight=1)

        text_frame = ttk.Frame(right)
        text_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=(8, 4))

        self.text_box = scrolledtext.ScrolledText(
            text_frame, wrap="word", font=("Segoe UI", self.font_size.get()),
            undo=True, padx=10, pady=10, bd=0, relief="flat", height=10
        )
        self.text_box.pack(fill=tk.BOTH, expand=True)
        self.text_box.insert("1.0", "Type or paste your text here...")
        self.text_box.bind("<<Modified>>", self._on_text_change)
        self._attach_text_context_menu()

        status_bar = ttk.Frame(self.root, padding=(16, 6, 16, 6))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label = tb.Label(status_bar, text="📝 Ready", bootstyle="secondary")
        self.status_label.pack(side=tk.LEFT)
        try:
            self.progress = tb.Progressbar(status_bar, mode="indeterminate", length=180, bootstyle="info-striped")
        except tk.TclError:
            self.progress = tb.Progressbar(status_bar, mode="indeterminate", length=180, bootstyle="info")
        except Exception:
            self.progress = ttk.Progressbar(status_bar, mode="indeterminate", length=180)
        self.progress.pack(side=tk.RIGHT)

    def _attach_text_context_menu(self):
        menu = tk.Menu(self.text_box, tearoff=0)
        menu.add_command(label="Cut", command=lambda: self.text_box.event_generate("<<Cut>>"))
        menu.add_command(label="Copy", command=lambda: self.text_box.event_generate("<<Copy>>"))
        menu.add_command(label="Paste", command=lambda: self.text_box.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="Select All", command=lambda: self.text_box.event_generate("<<SelectAll>>"))

        def popup(event):
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
        self.text_box.bind("<Button-3>", popup)

    def _bind_shortcuts(self):
        self.root.bind("<Control-Return>", lambda e: self.start_speaking())
        self.root.bind("<Control-s>", lambda e: self.start_saving())
        self.root.bind("<Control-S>", lambda e: self.start_saving())
        self.root.bind("<Control-o>", lambda e: self.open_text_file())
        self.root.bind("<Control-O>", lambda e: self.open_text_file())
        self.root.bind("<Control-l>", lambda e: self.clear_text())
        self.root.bind("<Control-L>", lambda e: self.clear_text())
        self.root.bind("<Escape>", lambda e: self.stop_now())
        self.root.bind("<Control-q>", lambda e: self.on_exit())
        self.root.bind("<Control-Q>", lambda e: self.on_exit())
        self.root.bind("<Control-a>", lambda e: (self.text_box.event_generate("<<SelectAll>>"), "break"))

    # -------------------- Voice Loading --------------------
    def _load_voices(self):
        try:
            engine = pyttsx3.init('sapi5') if sys.platform.startswith("win") else pyttsx3.init()
            voices = engine.getProperty("voices") or []
            current_id = engine.getProperty("voice")
            try:
                engine.stop()
            except Exception:
                pass

            self.voices = voices
            display_to_id = {}
            display_list = []

            for v in voices:
                gender = self._detect_gender(v).capitalize()  # Male/Female/Unknown
                lang = ""
                try:
                    if hasattr(v, "languages") and v.languages:
                        lang_val = v.languages[0]
                        if isinstance(lang_val, bytes):
                            lang_val = lang_val.decode("utf-8", errors="ignore")
                        lang = f" ({lang_val})"
                except Exception:
                    lang = ""
                label = f"{getattr(v, 'name', 'Voice')}{lang} — {gender} — {self._short_voice_id(getattr(v, 'id', ''))}"
                display_list.append(label)
                display_to_id[label] = getattr(v, "id", "")

            def apply():
                self.voice_display_map = display_to_id
                values = display_list if display_list else ["(No voices found)"]
                self.voice_combo["values"] = values

                selected_display = None
                # prefer saved voice id
                if self._desired_voice_id:
                    for disp, vid in display_to_id.items():
                        if vid == self._desired_voice_id:
                            selected_display = disp
                            break

                # else use current engine voice
                if not selected_display:
                    for disp, vid in display_to_id.items():
                        if vid == current_id:
                            selected_display = disp
                            break

                # else first
                if not selected_display and values and values[0] != "(No voices found)":
                    selected_display = values[0]

                if selected_display:
                    self.voice_combo.set(selected_display)
                    self.selected_voice_id = display_to_id[selected_display]
                else:
                    self.voice_combo.set("(No voices found)")
                    self.selected_voice_id = None

                self._apply_voice_filter()
                self._set_status("📝 Ready", "secondary")

            self.root.after(0, apply)
        except Exception as e:
            self.root.after(0, lambda: self._set_status(f"❌ Voice load error: {e}", "danger"))

    @staticmethod
    def _short_voice_id(vid: str) -> str:
        if not vid:
            return "default"
        sep = ":" if ":" in vid else "."
        return vid.split(sep)[-1]

    def _detect_gender(self, v) -> str:
        try:
            g = getattr(v, "gender", None)
            if isinstance(g, str) and g.strip():
                g = g.lower()
                if "female" in g or g == "f":
                    return "female"
                if "male" in g or g == "m":
                    return "male"
        except Exception:
            pass
        try:
            text = f"{getattr(v,'name','')} {getattr(v,'id','')}".lower()
            female_tokens = ["female", "zira", "hazel", "samantha", "eva", "catherine"]
            male_tokens = ["male", "david", "mark", "alex", "daniel", "fred"]
            if "+f" in text or any(t in text for t in female_tokens):
                return "female"
            if "+m" in text or any(t in text for t in male_tokens):
                return "male"
        except Exception:
            pass
        return "unknown"

    def _on_voice_change(self, _=None):
        display = self.voice_combo.get()
        self.selected_voice_id = self.voice_display_map.get(display, None)
        self._save_config()

    def _apply_voice_filter(self):
        if not self.voice_display_map:
            return
        q = self.filter_var.get().strip().lower()
        gender = self.gender_var.get().lower()  # any/male/female
        all_vals = list(self.voice_display_map.keys())

        def gender_ok(label: str) -> bool:
            if gender == "any":
                return True
            lab = label.lower()
            if gender == "male":
                return " — male — " in lab
            if gender == "female":
                return " — female — " in lab
            return True

        if q:
            filtered = [d for d in all_vals if q in d.lower() and gender_ok(d)]
        else:
            filtered = [d for d in all_vals if gender_ok(d)]

        current = self.voice_combo.get()
        self.voice_combo["values"] = filtered if filtered else ["(No match)"]
        if current in filtered:
            pass
        elif filtered:
            self.voice_combo.set(filtered[0])
            self.selected_voice_id = self.voice_display_map.get(filtered[0])
        else:
            self.voice_combo.set("(No match)")
            self.selected_voice_id = None

    # -------------------- Actions --------------------
    def start_speaking(self):
        if self.is_busy:
            return
        text = self._get_text().strip()
        if not text:
            self._set_status("⚠️ Please type some text", "warning")
            return

        self._sync_controls()
        self._set_busy(True)
        self._set_status("🔊 Speaking...", "info")
        threading.Thread(target=self._run_tts, args=(text, "speak", None), daemon=True).start()

    def start_saving(self):
        if self.is_busy:
            return
        text = self._get_text().strip()
        if not text:
            self._set_status("⚠️ Please type text to save", "warning")
            return

        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        initial_name = f"{APP_NAME.replace(' ', '')}_{timestamp}{self.default_ext}"
        file_path = filedialog.asksaveasfilename(
            title="Save Audio",
            defaultextension=self.default_ext,
            filetypes=self.file_filters,
            initialfile=initial_name,
        )
        if not file_path:
            return

        self._sync_controls()
        self._set_busy(True)
        threading.Thread(target=self._run_tts, args=(text, "save", file_path), daemon=True).start()
        self._set_status("💾 Saving audio...", "info")

    def stop_now(self):
        with self.engine_lock:
            eng = self.engine
            self.engine = None
        try:
            if eng:
                eng.stop()
        except Exception:
            pass
        self._set_status("⏹️ Stopped", "warning")
        self._set_busy(False)

    def open_text_file(self):
        if self.is_busy:
            self._set_status("⚠️ Cannot open file while busy", "warning")
            return
        file_path = filedialog.askopenfilename(
            title="Open Text File",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
            self.text_box.delete("1.0", "end")
            self.text_box.insert("1.0", content)
            self._set_status(f"📂 Loaded {os.path.basename(file_path)}", "success")
            self._update_stats()
        except Exception as e:
            self._set_status(f"❌ Error opening file: {e}", "danger")

    def clear_text(self):
        if self.is_busy:
            self.stop_now()
        self.text_box.delete("1.0", "end")
        self._set_status("🧹 Cleared. Ready.", "secondary")
        self._update_stats()

    def on_exit(self):
        if self.is_busy:
            try:
                self.stop_now()
            except Exception:
                pass
        self._save_config()
        try:
            TBStyle.instance = None
        except Exception:
            pass
        self.root.destroy()

    def preview_voice(self):
        if self.is_busy:
            return
        if not self.selected_voice_id:
            self._set_status("⚠️ Select a voice to preview", "warning")
            return
        sample = "This is a preview of the selected voice. The quick brown fox jumps over the lazy dog."
        self._sync_controls()
        self._set_busy(True)
        self._set_status("🎙 Previewing voice...", "info")
        threading.Thread(target=self._run_tts, args=(sample, "speak", None), daemon=True).start()

    # -------------------- TTS Worker --------------------
    def _run_tts(self, text: str, mode: str, file_path):
        local_engine = None
        try:
            local_engine = pyttsx3.init('sapi5') if sys.platform.startswith("win") else pyttsx3.init()

            try:
                local_engine.setProperty("rate", int(self.rate_var.get()))
            except Exception:
                pass
            try:
                local_engine.setProperty("volume", max(0.0, min(1.0, float(self.volume_var.get()) / 100.0)))
            except Exception:
                pass
            if self.selected_voice_id:
                try:
                    local_engine.setProperty("voice", self.selected_voice_id)
                except Exception:
                    pass

            with self.engine_lock:
                self.engine = local_engine

            if mode == "speak":
                local_engine.say(text)
            else:
                local_engine.save_to_file(text, file_path)

            local_engine.runAndWait()

            if mode == "speak":
                self.root.after(0, lambda: self._set_status("✅ Done", "success"))
            else:
                filename = os.path.basename(file_path) if file_path else "file"
                self.root.after(0, lambda: self._set_status(f"✅ Saved as {filename}", "success"))
        except Exception as e:
            self.root.after(0, lambda: self._set_status(f"❌ Error: {e}", "danger"))
        finally:
            with self.engine_lock:
                if self.engine is local_engine:
                    self.engine = None
            self.root.after(0, lambda: self._set_busy(False))

    # -------------------- Helpers --------------------
    def _get_text(self) -> str:
        return self.text_box.get("1.0", "end")

    def _on_text_change(self, _=None):
        try:
            self.text_box.edit_modified(False)
        except Exception:
            pass
        if self._stats_after_id:
            try:
                self.root.after_cancel(self._stats_after_id)
            except Exception:
                pass
        self._stats_after_id = self.root.after(180, self._update_stats)

    def _on_rate_change(self):
        try:
            self.rate_var.set(int(float(self.rate_scale.get())))
        except Exception:
            pass
        self.rate_value_label.config(text=str(self.rate_var.get()))
        self._update_stats()
        self._save_config()

    def _on_volume_change(self):
        try:
            self.volume_var.set(int(float(self.volume_scale.get())))
        except Exception:
            pass
        self.volume_value_label.config(text=f"{self.volume_var.get()}%")
        self._save_config()

    def _update_stats(self):
        text = self._get_text()
        words = len(text.split())
        chars = len(text)
        try:
            current_rate = int(self.rate_var.get())
        except Exception:
            current_rate = 160
        current_rate = max(60, min(300, current_rate))
        seconds = int((words / max(1, current_rate)) * 60)
        mm = seconds // 60
        ss = seconds % 60
        self.stats_label.config(text=f"Words: {words}   Chars: {chars}   ~{mm:02d}:{ss:02d}")

    def _set_status(self, message: str, style: str = "secondary"):
        try:
            self.status_label.configure(text=message, bootstyle=style)
        except Exception:
            self.status_label.config(text=message)

    def _set_busy(self, value: bool):
        self.is_busy = value
        self.btn_speak.config(state="disabled" if value else "normal")
        self.btn_save.config(state="disabled" if value else "normal")
        self.btn_open.config(state="disabled" if value else "normal")
        self.btn_clear.config(state="disabled" if value else "normal")
        self.voice_combo.config(state="disabled" if value else "readonly")
        try:
            self.rate_scale.config(state="disabled" if value else "normal")
            self.volume_scale.config(state="disabled" if value else "normal")
        except Exception:
            pass
        self.btn_theme.config(state="disabled" if value else "normal")
        self.btn_stop.config(state="normal" if value else "disabled")

        if value:
            try:
                self.progress.start(12)
            except Exception:
                pass
        else:
            try:
                self.progress.stop()
            except Exception:
                pass

    def _bump_font(self, delta: int):
        new_size = max(8, min(28, self.font_size.get() + delta))
        self.font_size.set(new_size)
        self.text_box.configure(font=("Segoe UI", new_size))
        self._apply_text_look()
        self._save_config()

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        new_theme = "darkly" if self.dark_mode else "flatly"
        try:
            tb.Style().theme_use(new_theme)
        except Exception:
            pass
        self.btn_theme.config(text=("☀️ Light Mode" if self.dark_mode else "🌙 Dark Mode"))
        self._apply_text_look()
        self._refresh_button_styles()
        self._save_config()

    def _apply_text_look(self):
        if self.dark_mode:
            self.text_box.config(
                bg="#111827", fg="#E5E7EB", insertbackground="#E5E7EB",
                selectbackground="#374151", selectforeground="#E5E7EB",
            )
        else:
            self.text_box.config(
                bg="#FFFFFF", fg="#111827", insertbackground="#111827",
                selectbackground="#D1D5DB", selectforeground="#111827",
            )

    def _refresh_button_styles(self):
        try:
            self.btn_speak.configure(bootstyle="primary")
            self.btn_stop.configure(bootstyle="danger")
            self.btn_save.configure(bootstyle="success")
            self.btn_open.configure(bootstyle="info")
            self.btn_clear.configure(bootstyle="warning")
            try:
                self.progress.configure(bootstyle="info-striped")
            except tk.TclError:
                self.progress.configure(bootstyle="info")
        except Exception:
            pass

    def _sync_controls(self):
        try:
            self.rate_var.set(int(float(self.rate_scale.get())))
        except Exception:
            pass
        try:
            self.volume_var.set(int(float(self.volume_scale.get())))
        except Exception:
            pass

    def _open_voice_settings(self):
        try:
            if sys.platform.startswith("win"):
                os.startfile("ms-settings:speech")
            elif sys.platform == "darwin":
                os.system('open "x-apple.systempreferences:com.apple.preference.accessibility?Speech"')
            else:
                messagebox.showinfo("Get Voices", "Linux: install additional voices (e.g. espeak-ng, mbrola).")
        except Exception as e:
            messagebox.showerror("Get Voices", f"Could not open voice settings.\n{e}")

    def _show_about(self):
        messagebox.showinfo(
            f"About {APP_NAME}",
            f"{APP_NAME} — Text to Speech\nColorful UI with ttkbootstrap\nWorks offline via pyttsx3",
        )

    # -------------------- Run --------------------
    def run(self):
        self.root.mainloop()


def tooltip(widget, text: str):
    tip = tk.Toplevel(widget)
    tip.withdraw()
    tip.overrideredirect(True)
    try:
        tip.attributes("-topmost", True)
    except Exception:
        pass

    label = tk.Label(
        tip, text=text, padx=8, pady=4,
        bg="#111111", fg="#FFFFFF", bd=0,
        font=("Segoe UI", 9)
    )
    label.pack()

    def enter(_):
        try:
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + widget.winfo_height() + 8
            tip.geometry(f"+{x}+{y}")
            tip.deiconify()
        except Exception:
            pass

    def leave(_):
        tip.withdraw()

    widget.bind("<Enter>", enter)
    widget.bind("<Leave>", leave)


if __name__ == "__main__":
    if sys.platform.startswith("win"):
        try:
            import multiprocessing as _mp
            _mp.freeze_support()
        except Exception:
            pass

    _reset_ttkbootstrap_style_if_stale()
    app = VoxNovaApp()
    app.run()


