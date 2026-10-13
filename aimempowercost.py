"""
AIMemPowerCost
==============
Real-time GPU + CPU power monitor for Windows 11.

Shows live GPU utilization, VRAM usage, power draw (W), temperature, and —
for the CPU — utilization, temperature (when readable) and a calibrated
power estimate. Integrates power over time (Wh) and converts energy into
money using a configurable electricity tariff (default: Moscow, commercial,
2026).

CPU power has no clean non-admin Windows API, so it is an *estimate*: a
user-calibratable idle→max range scaled by the live CPU load. CPU
temperature is read best-effort from the ACPI thermal zone (needs admin);
when unavailable the UI shows an honest "н/д" state.

Energy is persisted to %APPDATA%\\AIMemPowerCost\\state.json, so day/month
totals survive app restarts (gaps while the app was closed are not tracked).
"""

import base64
import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
from collections import deque
from datetime import datetime, timedelta
from tkinter import messagebox

# ---------------------------------------------------------------- NVML ----

try:
    import pynvml
    pynvml.nvmlInit()
    HAS_NVML = True
    NVML_ERR = ""
except Exception as e:  # driver / library missing
    HAS_NVML = False
    NVML_ERR = str(e)

# ---------------------------------------------------------------- psutil --
# CPU load (psutil.cpu_percent). Falls back to WMI Win32_Processor when
# psutil is absent so the CPU card still shows something honest.
try:
    import psutil
    HAS_PSUTIL = True
except Exception:
    HAS_PSUTIL = False

# -------------------------------------------------------------- WMI (CPU) --
# Best-effort CPU model name + ACPI thermal-zone temperature. The thermal
# class lives in root/wmi and typically needs admin; we degrade gracefully.
_wmi = None
try:
    import win32com.client  # pywin32
    _wmi = win32com.client
except Exception:
    _wmi = None


def _cpu_name():
    """CPU model string, best effort (empty string if unreadable)."""
    try:
        c = _wmi.GetObject("winmgmts:\\\\.\\root\\cimv2")
        for p in c.InstancesOf("Win32_Processor"):
            n = (p.Name or "").strip()
            if n:
                return n
    except Exception:
        pass
    return ""


def _cpu_temp_c():
    """
    CPU/package temperature in °C from the ACPI thermal zone, or None.
    Requires admin on most systems; returns None rather than guessing.
    """
    if not _wmi:
        return None
    try:
        c = _wmi.GetObject("winmgmts:\\\\.\\root\\wmi")
        best = None
        for t in c.InstancesOf("MSAcpi_ThermalZoneTemperature"):
            v = getattr(t, "CurrentTemperature", None)
            if isinstance(v, (int, float)):
                cel = v / 10.0 - 273.15
                if best is None or cel > best:
                    best = cel
        return best
    except Exception:
        return None


# ---------------------------------------------------------------- Config ----

APP_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "AIMemPowerCost")
STATE_FILE = os.path.join(APP_DIR, "state.json")
APP_VERSION = "0.2"

# Default tariff, rubles per kWh.
# Moscow, non-residential (commercial), 2026: the official single-rate
# transmission tariff (low voltage, 1st half of 2026) is 1.28366 RUB/kWh
# (Moscow Dept. of Economic Policy order DPR-TR-436/25 of 29.12.2025);
# full delivered price = generation (market) + transmission + supplier
# markup + VAT and typically lands around 7-8 RUB/kWh.
# The app lets you override this in the UI; the value is saved in state.json.
DEFAULT_TARIFF = 7.90

# Default CPU power estimate range (W). Calibratable in the UI; saved to
# state.json. An i9-13900K idles ~40-70 W and can draw 250+ W under load.
DEFAULT_CPU_IDLE_W = 45.0
DEFAULT_CPU_MAX_W = 280.0

SAMPLE_SECONDS = 2.0        # sampling period
CHART_SECONDS = 300         # chart window: last 5 minutes
MINUTE_BUCKET_TTL_HOURS = 48  # keep minute buckets for 2 days (1h window + buffer)

# ----------------------------------------------------------------- Icon ----
# Embedded 48x48 PNG (base64) so the header renders even in --onefile builds.
# Injected at build time from icon.png; see README "Сборка".
_ICON_B64 = "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"


# ------------------------------------------------------------ Energy store --

class EnergyStore:
    """Persists energy (Wh) in minute buckets (rolling ~48 h) and day totals."""

    def __init__(self, path: str):
        self.path = path
        self.minutes: dict[str, float] = {}   # "YYYY-MM-DD HH:MM" -> Wh
        self.days: dict[str, float] = {}      # "YYYY-MM-DD" -> Wh
        self.tariff: float = DEFAULT_TARIFF   # RUB per kWh
        self.cpu_idle_w: float = DEFAULT_CPU_IDLE_W
        self.cpu_max_w: float = DEFAULT_CPU_MAX_W
        self._pending_wh = 0.0
        self._pending_min: str | None = None
        self.load()

    # -- persistence --------------------------------------------------

    def load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                d = json.load(f)
            self.minutes = {k: float(v) for k, v in d.get("minutes", {}).items()}
            self.days = {k: float(v) for k, v in d.get("days", {}).items()}
            self.tariff = float(d.get("tariff", DEFAULT_TARIFF))
            self.cpu_idle_w = float(d.get("cpu_idle_w", DEFAULT_CPU_IDLE_W))
            self.cpu_max_w = float(d.get("cpu_max_w", DEFAULT_CPU_MAX_W))
        except (FileNotFoundError, ValueError, TypeError):
            pass
        self.prune()

    def save(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "minutes": self.minutes,
                    "days": self.days,
                    "tariff": self.tariff,
                    "cpu_idle_w": self.cpu_idle_w,
                    "cpu_max_w": self.cpu_max_w,
                },
                f,
            )
        os.replace(tmp, self.path)

    def prune(self):
        cutoff = (datetime.now() - timedelta(hours=MINUTE_BUCKET_TTL_HOURS)).strftime(
            "%Y-%m-%d %H:%M"
        )
        self.minutes = {k: v for k, v in self.minutes.items() if k >= cutoff}

    # -- accumulation ---------------------------------------------------

    def add_wh(self, wh: float, now: datetime | None = None):
        """Accumulate Wh; flushes the previous minute bucket on rollover."""
        now = now or datetime.now()
        self._pending_wh += wh
        key = now.strftime("%Y-%m-%d %H:%M")
        if self._pending_min is None:
            self._pending_min = key
        elif key != self._pending_min:
            self.flush()

    def flush(self):
        """Commit the in-progress minute to minute+day totals and save."""
        if self._pending_min and self._pending_wh > 0:
            self.minutes[self._pending_min] = (
                self.minutes.get(self._pending_min, 0.0) + self._pending_wh
            )
            day = self._pending_min[:10]
            self.days[day] = self.days.get(day, 0.0) + self._pending_wh
        self._pending_min = None
        self._pending_wh = 0.0
        self.prune()
        self.save()

    # -- queries (include the in-progress minute so numbers update live) --

    def wh_last_hour(self, now: datetime | None = None) -> float:
        now = now or datetime.now()
        cutoff = (now - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M")
        total = sum(v for k, v in self.minutes.items() if k >= cutoff)
        return total + self._pending_wh

    def wh_today(self, now: datetime | None = None) -> float:
        now = now or datetime.now()
        return self.days.get(now.strftime("%Y-%m-%d"), 0.0) + self._pending_wh

    def wh_month(self, now: datetime | None = None) -> float:
        now = now or datetime.now()
        prefix = now.strftime("%Y-%m")
        return sum(v for k, v in self.days.items() if k.startswith(prefix)) + self._pending_wh

    def cost(self, wh: float) -> float:
        return wh / 1000.0 * self.tariff

    def reset(self):
        self.minutes = {}
        self.days = {}
        self._pending_wh = 0.0
        self._pending_min = None
        self.save()


# -------------------------------------------------------------- Collector --

class GpuCollector(threading.Thread):
    """Background thread: polls NVML every SAMPLE_SECONDS, integrates power."""

    def __init__(self, store: EnergyStore, gpu_index: int,
                 stop_event: threading.Event, result_queue):
        super().__init__(daemon=True, name=f"gpu-collector-{gpu_index}")
        self.store = store
        self.gpu_index = gpu_index
        self.stop_event = stop_event
        self.result_queue = result_queue

    def run(self):
        handle = pynvml.nvmlDeviceGetHandleByIndex(self.gpu_index)
        last = time.monotonic()
        while not self.stop_event.is_set():
            try:
                util = pynvml.nvmlDeviceGetUtilizationRates(handle)
                mem = pynvml.nvmlDeviceGetMemoryInfo(handle)
                power_w = pynvml.nvmlDeviceGetPowerUsage(handle) / 1000.0
                temp_c = pynvml.nvmlDeviceGetTemperature(handle, pynvml.NVML_TEMPERATURE_GPU)
                pmax_w = pynvml.nvmlDeviceGetPowerManagementLimit(handle) / 1000.0
            except Exception as e:
                self.result_queue.put({"error": f"gpu: {e}"})
                time.sleep(SAMPLE_SECONDS)
                continue

            now = time.monotonic()
            dt = now - last
            last = now
            if dt > 0:
                self.store.add_wh(power_w * dt / 3600.0)

            self.result_queue.put({
                "kind": "gpu",
                "ts": datetime.now().strftime("%H:%M:%S"),
                "util": util.gpu,
                "mem_used_gb": mem.used / 1024 ** 3,
                "mem_total_gb": mem.total / 1024 ** 3,
                "power_w": power_w,
                "pmax_w": pmax_w,
                "temp_c": temp_c,
            })
            self.stop_event.wait(SAMPLE_SECONDS)


class CpuCollector(threading.Thread):
    """
    Background thread: polls CPU load every SAMPLE_SECONDS, estimates power
    (calibratable idle→max scaled by load) and integrates it. CPU temperature
    is read best-effort (cached; the WMI query is slow and may need admin).
    """

    def __init__(self, store: EnergyStore, stop_event: threading.Event, result_queue,
                 temp_cache: dict):
        super().__init__(daemon=True, name="cpu-collector")
        self.store = store
        self.stop_event = stop_event
        self.result_queue = result_queue
        self.temp_cache = temp_cache  # {"t": monotonic, "c": float|None}

    @staticmethod
    def _cpu_percent():
        if HAS_PSUTIL:
            return psutil.cpu_percent()
        try:
            c = _wmi.GetObject("winmgmts:\\\\.\\root\\cimv2")
            return float(c.InstancesOf("Win32_Processor")[0].LoadPercentage)
        except Exception:
            return 0.0

    @staticmethod
    def _refresh_temp(cache: dict):
        cache["t"] = time.monotonic()
        cache["c"] = _cpu_temp_c()

    def run(self):
        if not HAS_PSUTIL and _wmi:
            psutil.cpu_percent(interval=None)  # prime the baseline (no-op if absent)
        last = time.monotonic()
        last_temp = 0.0
        while not self.stop_event.is_set():
            try:
                # CPU temperature is slow / admin-gated: refresh at most every 10 s.
                if time.monotonic() - last_temp >= 10.0:
                    self._refresh_temp(self.temp_cache)
                    last_temp = time.monotonic()
                temp_c = self.temp_cache.get("c")
            except Exception:
                temp_c = None

            try:
                load_pct = self._cpu_percent()
            except Exception as e:
                self.result_queue.put({"kind": "cpu", "error": f"cpu: {e}"})
                self.stop_event.wait(SAMPLE_SECONDS)
                continue

            now = time.monotonic()
            dt = now - last
            last = now
            idle = max(0.0, min(self.store.cpu_idle_w, self.store.cpu_max_w))
            peak = max(idle, self.store.cpu_max_w)
            frac = max(0.0, min(1.0, load_pct / 100.0))
            power_w = idle + (peak - idle) * frac
            if dt > 0:
                self.store.add_wh(power_w * dt / 3600.0)

            self.result_queue.put({
                "kind": "cpu",
                "ts": datetime.now().strftime("%H:%M:%S"),
                "load_pct": load_pct,
                "power_w": power_w,
                "temp_c": temp_c,
            })
            self.stop_event.wait(SAMPLE_SECONDS)


# ------------------------------------------------------------------- App ---

BG = "#15181e"
CARD = "#1e232c"
CARD_HI = "#262d39"
FG = "#e9ecf2"
MUTED = "#8b93a7"
ACCENT = "#4cc2ff"
GREEN = "#3ddc84"
AMBER = "#ffb454"
RED = "#ff6b6b"


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(f"AIMemPowerCost {APP_VERSION}")
        root.configure(bg=BG)
        self._icon_photo = self._load_icon(root)
        if self._icon_photo and sys.platform == "win32":
            ico = os.path.join(
                getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))),
                "icon.ico",
            )
            if not os.path.exists(ico):
                ico = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
            if os.path.exists(ico):
                try:
                    root.iconbitmap(ico)
                except Exception:
                    pass

        self.store = EnergyStore(STATE_FILE)
        self.gpu_count = 0
        self.gpu_names: list[str] = []
        if HAS_NVML:
            self.gpu_count = pynvml.nvmlDeviceGetCount()
            for i in range(self.gpu_count):
                h = pynvml.nvmlDeviceGetHandleByIndex(i)
                self.gpu_names.append(pynvml.nvmlDeviceGetName(h))
        self.cpu_name = _cpu_name()

        self.samples: deque[tuple[float, float]] = deque(maxlen=CHART_SECONDS // int(SAMPLE_SECONDS))
        self.cpu_samples: deque[tuple[float, float]] = deque(maxlen=CHART_SECONDS // int(SAMPLE_SECONDS))
        self.result_queue: "queue.Queue" = queue.Queue()
        self.stop_event = threading.Event()
        self.cpu_stop_event = threading.Event()
        self.collector: GpuCollector | None = None
        self.cpu_collector: CpuCollector | None = None
        self.temp_cache: dict = {"t": 0.0, "c": None}
        self._gpu_var = tk.StringVar(value=self.gpu_names[0] if self.gpu_names else "")

        self._build_ui()
        self._fit_window()
        # CPU collector always runs (psutil/WMI are independent of NVML).
        self._start_cpu_collector()
        if HAS_NVML and self.gpu_count:
            self._start_collector(0)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(300, self._tick)

    # -- icon -----------------------------------------------------------

    @staticmethod
    def _load_icon(root: tk.Tk):
        try:
            data = base64.b64decode(_ICON_B64)
            if len(data) < 64:
                return None
            from tkinter import PhotoImage
            im = PhotoImage(data=data)
            return im
        except Exception:
            return None

    # -- UI ---------------------------------------------------------------

    def _fit_window(self):
        """Size the window to its content (DPI-awareness is set in main())."""
        self.root.update_idletasks()
        w = self.root.winfo_reqwidth() + 16
        h = self.root.winfo_reqheight() + 32
        self.root.geometry(f"{w}x{h}")
        self.root.minsize(max(620, int(w * 0.72)), max(420, int(h * 0.72)))

    def _build_ui(self):
        r = self.root
        pad = dict(padx=14, pady=6)

        # header
        head = tk.Frame(r, bg=BG)
        head.pack(fill="x", **pad)
        if self._icon_photo is not None:
            tk.Label(head, image=self._icon_photo, bg=BG).pack(side="left", padx=(0, 8))
        tk.Label(head, text="AIMemPowerCost", font=("Segoe UI", 16, "bold"),
                 bg=BG, fg=ACCENT).pack(side="left")
        tk.Label(head, text=f"v{APP_VERSION}", font=("Segoe UI", 9),
                 bg=BG, fg=MUTED).pack(side="left", padx=(6, 0), pady=(4, 0))
        if HAS_NVML and self.gpu_count:
            menu = tk.Menu(head, tearoff=0, font=("Segoe UI", 10))
            for name in self.gpu_names:
                menu.add_command(label=name)
            self.gpu_menu = menu
            om = tk.OptionMenu(head, self._gpu_var, *self.gpu_names)
            om["menu"] = menu
            om.config(font=("Segoe UI", 10), bg=CARD_HI, fg=FG,
                      activebackground=CARD_HI, activeforeground=FG,
                      relief="flat", highlightthickness=0, bd=0, padx=8, pady=3)
            om.pack(side="right", ipady=2)
            self._gpu_var.trace_add("write", self._on_gpu_change)
        else:
            tk.Label(head, text="NVML недоступен", font=("Segoe UI", 10),
                     bg=BG, fg=RED).pack(side="right")

        # GPU live card (grid: 4 stat rows x 2 columns)
        live = tk.Frame(r, bg=CARD, highlightbackground=CARD, highlightthickness=1)
        live.pack(fill="x", **pad)
        for c in (0, 1):
            live.columnconfigure(c, weight=1, uniform="stat")
        self.l_util = self._stat_row(live, "Нагрузка GPU", 0, 0)
        self.l_mem = self._stat_row(live, "Видеопамять", 0, 1)
        self.l_power = self._stat_row(live, "Мощность", 1, 0)
        self.l_temp = self._stat_row(live, "Температура", 1, 1)

        # CPU live card (same style as the GPU card)
        cpu = tk.Frame(r, bg=CARD, highlightbackground=CARD, highlightthickness=1)
        cpu.pack(fill="x", **pad)
        for c in (0, 1):
            cpu.columnconfigure(c, weight=1, uniform="stat")
        self.c_load = self._stat_row(cpu, "НАГРУЗКА CPU", 0, 0)
        self.c_temp = self._stat_row(cpu, "ТЕМПЕРАТУРА CPU", 0, 1)
        self.c_power = self._stat_row(cpu, "Мощность CPU", 1, 0)
        self.c_name = self._stat_row(cpu, "Процессор", 1, 1, with_bar=False)
        self.c_name["value"].config(text=self.cpu_name or "—")

        # GPU chart
        chart_frame = tk.Frame(r, bg=CARD, highlightbackground=CARD, highlightthickness=1)
        chart_frame.pack(fill="x", **pad)
        tk.Label(chart_frame, text="Мощность GPU, Вт (последние 5 минут)",
                 font=("Segoe UI", 9), bg=CARD, fg=MUTED).pack(anchor="w", padx=12, pady=(8, 0))
        self.chart = tk.Canvas(chart_frame, height=96, bg="#181c23", highlightthickness=0)
        self.chart.pack(fill="x", padx=10, pady=(2, 10))

        # CPU chart (same style)
        cpu_chart = tk.Frame(r, bg=CARD, highlightbackground=CARD, highlightthickness=1)
        cpu_chart.pack(fill="x", **pad)
        tk.Label(cpu_chart, text="Мощность CPU (оценка), Вт (последние 5 минут)",
                 font=("Segoe UI", 9), bg=CARD, fg=MUTED).pack(anchor="w", padx=12, pady=(8, 0))
        self.cpu_chart = tk.Canvas(cpu_chart, height=96, bg="#181c23", highlightthickness=0)
        self.cpu_chart.pack(fill="x", padx=10, pady=(2, 10))

        # money cards
        money = tk.Frame(r, bg=BG)
        money.pack(fill="x", **pad)
        for c in range(3):
            money.columnconfigure(c, weight=1, uniform="money")
        self.card_day = self._money_card(money, "Сегодня", 0)
        self.card_hour = self._money_card(money, "За час", 1)
        self.card_month = self._money_card(money, "За месяц", 2)

        # settings
        st = tk.Frame(r, bg=BG)
        st.pack(fill="x", **pad)
        tk.Label(st, text="Тариф, ₽/кВт·ч:", font=("Segoe UI", 10),
                 bg=BG, fg=MUTED).pack(side="left")
        self.tariff_var = tk.StringVar(value=f"{self.store.tariff:.2f}")
        ent = tk.Entry(st, textvariable=self.tariff_var, width=8, font=("Segoe UI", 10),
                       bg=CARD_HI, fg=FG, insertbackground=FG, relief="flat", justify="center")
        ent.pack(side="left", padx=8, ipady=3)
        tk.Button(st, text="Сохранить", command=self._save_tariff,
                  font=("Segoe UI", 9), bg=CARD_HI, fg=FG, relief="flat",
                  activebackground=ACCENT, activeforeground="#10131a"
                  ).pack(side="left")
        tk.Button(st, text="Сбросить счётчики", command=self._reset_counters,
                  font=("Segoe UI", 9), bg=CARD_HI, fg=AMBER, relief="flat",
                  activebackground=CARD_HI, activeforeground=AMBER
                  ).pack(side="left", padx=8)
        self.status = tk.Label(st, text="запуск...", font=("Segoe UI", 9),
                               bg=BG, fg=MUTED)
        self.status.pack(side="right")

        # CPU power calibration (estimate range)
        cal = tk.Frame(r, bg=BG)
        cal.pack(fill="x", **pad)
        tk.Label(cal, text="Оценка мощности CPU, Вт — idle:", font=("Segoe UI", 10),
                 bg=BG, fg=MUTED).pack(side="left")
        self.cpu_idle_var = tk.StringVar(value=f"{self.store.cpu_idle_w:.0f}")
        e1 = tk.Entry(cal, textvariable=self.cpu_idle_var, width=6, font=("Segoe UI", 10),
                      bg=CARD_HI, fg=FG, insertbackground=FG, relief="flat", justify="center")
        e1.pack(side="left", padx=8, ipady=3)
        tk.Label(cal, text="макс:", font=("Segoe UI", 10),
                 bg=BG, fg=MUTED).pack(side="left")
        self.cpu_max_var = tk.StringVar(value=f"{self.store.cpu_max_w:.0f}")
        e2 = tk.Entry(cal, textvariable=self.cpu_max_var, width=6, font=("Segoe UI", 10),
                      bg=CARD_HI, fg=FG, insertbackground=FG, relief="flat", justify="center")
        e2.pack(side="left", padx=8, ipady=3)
        tk.Button(cal, text="Применить", command=self._save_cpu_power,
                  font=("Segoe UI", 9), bg=CARD_HI, fg=FG, relief="flat",
                  activebackground=ACCENT, activeforeground="#10131a"
                  ).pack(side="left")
        tk.Label(cal, text="(оценка: idle→макс по нагрузке; температура может требовать администратора)",
                 font=("Segoe UI", 8), bg=BG, fg=MUTED).pack(side="left", padx=(10, 0))

        if not HAS_NVML:
            tk.Label(r, text=f"Не удалось инициализировать NVML: {NVML_ERR}\n"
                             "Установите драйвер NVIDIA и перезапустите.",
                     font=("Segoe UI", 10), bg=BG, fg=RED, justify="left"
                     ).pack(pady=10)
        elif self.gpu_count == 0:
            tk.Label(r, text="GPU NVIDIA не обнаружен.", font=("Segoe UI", 10),
                     bg=BG, fg=RED).pack(pady=10)

    def _stat_row(self, parent, title, row, col, with_bar=True) -> dict:
        r = tk.Frame(parent, bg=CARD)
        r.grid(row=row, column=col, sticky="nsew", padx=10, pady=4)
        tk.Label(r, text=title, font=("Segoe UI", 10), bg=CARD, fg=MUTED
                 ).grid(row=0, column=0, sticky="w")
        val = tk.Label(r, text="—", font=("Segoe UI", 12, "bold"),
                       bg=CARD, fg=FG)
        val.grid(row=1, column=0, sticky="w", pady=(0, 4))
        bar = None
        if with_bar:
            bar = tk.Canvas(r, width=200, height=10, bg="#2a3140", highlightthickness=0)
            bar.grid(row=0, column=1, rowspan=2, sticky="e", padx=10, pady=(2, 4))
            bar.create_rectangle(0, 0, 0, 10, fill=ACCENT, outline="")
        return {"value": val, "bar": bar, "pct": 0.0}

    def _money_card(self, parent, title, col) -> dict:
        f = tk.Frame(parent, bg=CARD_HI, highlightbackground=CARD, highlightthickness=1)
        f.grid(row=0, column=col, sticky="nsew", padx=4, pady=4)
        tk.Label(f, text=title, font=("Segoe UI", 10), bg=CARD_HI, fg=MUTED
                 ).pack(anchor="w", padx=14, pady=(10, 0))
        cost = tk.Label(f, text="0.00 ₽", font=("Segoe UI", 18, "bold"),
                        bg=CARD_HI, fg=GREEN)
        cost.pack(anchor="w", padx=14)
        wh = tk.Label(f, text="0.000 кВт·ч", font=("Segoe UI", 9),
                      bg=CARD_HI, fg=MUTED)
        wh.pack(anchor="w", padx=14, pady=(0, 10))
        return {"cost": cost, "wh": wh}

    # -- logic ------------------------------------------------------------

    def _start_collector(self, gpu_index: int):
        if self.collector:
            self.stop_event.set()
            self.collector.join(timeout=3)
        self.stop_event = threading.Event()
        self.collector = GpuCollector(self.store, gpu_index, self.stop_event, self.result_queue)
        self.collector.start()

    def _start_cpu_collector(self):
        if self.cpu_collector:
            self.cpu_stop_event.set()
            self.cpu_collector.join(timeout=3)
        self.cpu_stop_event = threading.Event()
        self.cpu_collector = CpuCollector(self.store, self.cpu_stop_event,
                                          self.result_queue, self.temp_cache)
        self.cpu_collector.start()

    def _gpu_index(self) -> int | None:
        """Safe GPU index from the OptionMenu value (None if not in list)."""
        if not self.gpu_count:
            return None
        try:
            return self.gpu_names.index(self._gpu_var.get())
        except ValueError:
            return None

    def _on_gpu_change(self, *_):
        if HAS_NVML and self.gpu_count:
            idx = self._gpu_index()
            if idx is not None:
                self._start_collector(idx)

    def _tick(self):
        try:
            while True:
                try:
                    item = self.result_queue.get_nowait()
                except Exception:
                    break
                if "error" in item:
                    self.status.config(text=f"ошибка: {item['error'][:60]}", fg=RED)
                    continue
                self._render(item)
        except Exception as e:
            self.status.config(text=f"ошибка: {e}", fg=RED)
        self.root.after(250, self._tick)

    def _render(self, d):
        if d.get("kind") == "cpu":
            self._render_cpu(d)
        else:
            self._render_gpu(d)

        now = datetime.now()
        for card, wh in ((self.card_day, self.store.wh_today(now)),
                         (self.card_hour, self.store.wh_last_hour(now)),
                         (self.card_month, self.store.wh_month(now))):
            card["cost"].config(text=f"{self.store.cost(wh):.2f} ₽")
            card["wh"].config(text=f"{wh / 1000.0:.3f} кВт·ч")

        gi = self._gpu_index()
        self.status.config(text=f"обновлено {d['ts']}  ·  GPU {gi if gi is not None else '—'}",
                           fg=MUTED)

    def _render_gpu(self, d):
        self.l_util["value"].config(text=f"{d['util']} %")
        self._set_bar(self.l_util, d["util"] / 100.0)
        self.l_mem["value"].config(text=f"{d['mem_used_gb']:.1f} / {d['mem_total_gb']:.1f} ГБ")
        self._set_bar(self.l_mem, d["mem_used_gb"] / d["mem_total_gb"] if d["mem_total_gb"] else 0)
        self.l_power["value"].config(text=f"{d['power_w']:.0f} Вт")
        self._set_bar(self.l_power, d["power_w"] / d["pmax_w"] if d["pmax_w"] else 0,
                      color=AMBER if d["power_w"] > 0.8 * d["pmax_w"] else ACCENT)
        self.l_temp["value"].config(text=f"{d['temp_c']} °C")
        self._set_bar(self.l_temp, d["temp_c"] / 90.0,
                      color=RED if d["temp_c"] >= 85 else GREEN)

        self.samples.append((time.monotonic(), d["power_w"]))
        self._draw_chart(self.chart, self.samples, "пик: ")

    def _render_cpu(self, d):
        self.c_load["value"].config(text=f"{d['load_pct']:.0f} %")
        self._set_bar(self.c_load, d["load_pct"] / 100.0)

        temp = d.get("temp_c")
        if temp is None:
            self.c_temp["value"].config(text="н/д · нужен админ", fg=MUTED)
            self._set_bar(self.c_temp, 0.0, color="#2a3140")
        else:
            self.c_temp["value"].config(text=f"{temp:.0f} °C", fg=FG)
            self._set_bar(self.c_temp, temp / 90.0,
                          color=RED if temp >= 85 else GREEN)

        self.c_power["value"].config(text=f"≈ {d['power_w']:.0f} Вт")
        frac = max(0.0, min(1.0, d["power_w"] / max(1.0, self.store.cpu_max_w)))
        self._set_bar(self.c_power, frac,
                      color=AMBER if frac > 0.8 else ACCENT)

        self.cpu_samples.append((time.monotonic(), d["power_w"]))
        self._draw_chart(self.cpu_chart, self.cpu_samples, "пик: ")

    def _set_bar(self, row, frac, color=ACCENT):
        if row["bar"] is None:
            return
        frac = max(0.0, min(1.0, frac))
        row["bar"].delete("all")
        row["bar"].create_rectangle(0, 0, int(200 * frac), 10, fill=color, outline="")
        row["pct"] = frac

    def _draw_chart(self, c, samples, peak_prefix):
        c.delete("all")
        w = c.winfo_width() or 700
        if len(samples) < 2:
            return
        pmax = max(p for _, p in samples)
        pmax = max(pmax * 1.15, 20.0)
        t0, t1 = samples[0][0], samples[-1][0]
        span = max(t1 - t0, 1.0)

        def xy(t, p):
            x = (t - t0) / span * (w - 4) + 2
            y = 94 - (p / pmax) * 88
            return x, y

        pts = [xy(t, p) for t, p in samples]
        c.create_polygon([(pts[0][0], 94)] + pts + [(pts[-1][0], 94)],
                         fill="#27405a", outline="")
        c.create_line(*pts, fill=ACCENT, width=2)
        c.create_text(6, 6, anchor="w", text=f"{peak_prefix}{pmax / 1.15:.0f} Вт",
                      fill=MUTED, font=("Segoe UI", 8))

    def _save_tariff(self):
        try:
            v = float(self.tariff_var.get().replace(",", "."))
            if v <= 0:
                raise ValueError
            self.store.tariff = v
            self.store.save()
            self.status.config(text=f"тариф сохранён: {v:.2f} ₽/кВт·ч", fg=GREEN)
        except ValueError:
            messagebox.showerror("AIMemPowerCost", "Введите корректный тариф, например 7.90")

    def _save_cpu_power(self):
        try:
            idle = float(self.cpu_idle_var.get().replace(",", "."))
            mx = float(self.cpu_max_var.get().replace(",", "."))
            if idle < 0 or mx <= 0 or mx < idle:
                raise ValueError
            self.store.cpu_idle_w = idle
            self.store.cpu_max_w = mx
            self.store.save()
            self.status.config(text=f"оценка мощности CPU: {idle:.0f}–{mx:.0f} Вт", fg=GREEN)
        except ValueError:
            messagebox.showerror(
                "AIMemPowerCost",
                "Введите корректную оценку мощности CPU: 0 ≤ idle ≤ макс, макс > 0.")

    def _reset_counters(self):
        if messagebox.askyesno("AIMemPowerCost", "Сбросить накопленные счётчики (день/месяц)?"):
            self.store.reset()
            self.samples.clear()
            self.cpu_samples.clear()
            self.status.config(text="счётчики сброшены", fg=AMBER)

    def on_close(self):
        try:
            if self.collector:
                self.stop_event.set()
                self.collector.join(timeout=3)
            if self.cpu_collector:
                self.cpu_stop_event.set()
                self.cpu_collector.join(timeout=3)
            self.store.flush()
        finally:
            self.root.destroy()


# ------------------------------------------------------------------ main ----

def main():
    if sys.platform != "win32" and "--allow-any-os" not in sys.argv:
        print("AIMemPowerCost: целевая платформа — Windows (NVML).")
        sys.exit(1)

    if "--selftest" in sys.argv:
        # Headless integration test: sample the real GPU + CPU for ~6 s.
        store = EnergyStore(os.path.join(APP_DIR, "selftest-state.json"))
        store.reset()
        stop = threading.Event()
        cpu_stop = threading.Event()
        q = queue.Queue()
        cols = []
        if HAS_NVML and pynvml.nvmlDeviceGetCount():
            cols.append(GpuCollector(store, 0, stop, q))
        cols.append(CpuCollector(store, cpu_stop, q, {"t": 0.0, "c": None}))
        for c in cols:
            c.start()
        last = time.monotonic()
        samples = 0
        while time.monotonic() - last < 6:
            try:
                item = q.get(timeout=1)
                if "error" not in item:
                    samples += 1
            except Exception:
                pass
        for ev in (stop, cpu_stop):
            ev.set()
        for c in cols:
            c.join(timeout=3)
        store.flush()
        wh = store.wh_today()
        ok = 0 < wh < 8.0  # GPU ~100-450 W + CPU ~45-280 W * 6 s / 3600
        print(f"selftest: {samples} samples, accumulated {wh:.4f} Wh (expected 0.1-8 Wh)")
        try:
            os.remove(os.path.join(APP_DIR, "selftest-state.json"))
        except OSError:
            pass
        sys.exit(0 if ok else 1)

    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
