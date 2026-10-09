"""
AIMemPowerCost
==============
Real-time GPU power monitor for Windows 11.

Shows live GPU utilization, VRAM usage, power draw (W) and temperature,
integrates power over time (Wh) and converts energy into money using a
configurable electricity tariff (default: Moscow, commercial, 2026).

Energy is persisted to %APPDATA%\\AIMemPowerCost\\state.json, so day/month
totals survive app restarts (gaps while the app was closed are not tracked).
"""

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

# ---------------------------------------------------------------- Config ----

APP_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "AIMemPowerCost")
STATE_FILE = os.path.join(APP_DIR, "state.json")

# Default tariff, rubles per kWh.
# Moscow, non-residential (commercial), 2026: the official single-rate
# transmission tariff (low voltage, 1st half of 2026) is 1.28366 RUB/kWh
# (Moscow Dept. of Economic Policy order DPR-TR-436/25 of 29.12.2025);
# full delivered price = generation (market) + transmission + supplier
# markup + VAT and typically lands around 7-8 RUB/kWh.
# The app lets you override this in the UI; the value is saved in state.json.
DEFAULT_TARIFF = 7.90

SAMPLE_SECONDS = 2.0        # sampling period
CHART_SECONDS = 300         # chart window: last 5 minutes
MINUTE_BUCKET_TTL_HOURS = 48  # keep minute buckets for 2 days (1h window + buffer)


# ------------------------------------------------------------ Energy store --

class EnergyStore:
    """Persists energy (Wh) in minute buckets (rolling ~48 h) and day totals."""

    def __init__(self, path: str):
        self.path = path
        self.minutes: dict[str, float] = {}   # "YYYY-MM-DD HH:MM" -> Wh
        self.days: dict[str, float] = {}      # "YYYY-MM-DD" -> Wh
        self.tariff: float = DEFAULT_TARIFF   # RUB per kWh
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
        except (FileNotFoundError, ValueError, TypeError):
            pass
        self.prune()

    def save(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(
                {"minutes": self.minutes, "days": self.days, "tariff": self.tariff},
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
                self.result_queue.put({"error": str(e)})
                time.sleep(SAMPLE_SECONDS)
                continue

            now = time.monotonic()
            dt = now - last
            last = now
            if dt > 0:
                self.store.add_wh(power_w * dt / 3600.0)

            self.result_queue.put({
                "ts": datetime.now().strftime("%H:%M:%S"),
                "util": util.gpu,
                "mem_used_gb": mem.used / 1024 ** 3,
                "mem_total_gb": mem.total / 1024 ** 3,
                "power_w": power_w,
                "pmax_w": pmax_w,
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
        root.title("AIMemPowerCost")
        root.configure(bg=BG)
        root.geometry("760x560")
        root.minsize(700, 520)
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
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

        self.samples: deque[tuple[float, float]] = deque(maxlen=CHART_SECONDS // int(SAMPLE_SECONDS))
        self.result_queue: "queue.Queue" = queue.Queue()
        self.stop_event = threading.Event()
        self.collector: GpuCollector | None = None
        self._gpu_var = tk.StringVar()

        self._build_ui()
        if HAS_NVML and self.gpu_count:
            self._start_collector(0)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(300, self._tick)

    # -- UI ---------------------------------------------------------------

    def _build_ui(self):
        r = self.root
        pad = dict(padx=14, pady=6)

        # header
        head = tk.Frame(r, bg=BG)
        head.pack(fill="x", **pad)
        tk.Label(head, text="AIMemPowerCost", font=("Segoe UI", 16, "bold"),
                 bg=BG, fg=ACCENT).pack(side="left")
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

        # live card
        live = tk.Frame(r, bg=CARD, highlightbackground=CARD, highlightthickness=1)
        live.pack(fill="x", **pad)

        self.l_util = self._stat_row(live, "Нагрузка GPU")
        self.l_mem = self._stat_row(live, "Видеопамять")
        self.l_power = self._stat_row(live, "Мощность")
        self.l_temp = self._stat_row(live, "Температура")

        # chart
        chart_frame = tk.Frame(r, bg=CARD, highlightbackground=CARD, highlightthickness=1)
        chart_frame.pack(fill="x", **pad)
        tk.Label(chart_frame, text="Мощность, Вт (последние 5 минут)",
                 font=("Segoe UI", 9), bg=CARD, fg=MUTED).pack(anchor="w", padx=12, pady=(8, 0))
        self.chart = tk.Canvas(chart_frame, height=96, bg="#181c23", highlightthickness=0)
        self.chart.pack(fill="x", padx=10, pady=(2, 10))

        # money cards
        money = tk.Frame(r, bg=BG)
        money.pack(fill="x", **pad)
        self.card_day = self._money_card(money, "Сегодня")
        self.card_hour = self._money_card(money, "За час")
        self.card_month = self._money_card(money, "За месяц")

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

        if not HAS_NVML:
            tk.Label(r, text=f"Не удалось инициализировать NVML: {NVML_ERR}\n"
                             "Установите драйвер NVIDIA и перезапустите.",
                     font=("Segoe UI", 10), bg=BG, fg=RED, justify="left"
                     ).pack(pady=10)
        elif self.gpu_count == 0:
            tk.Label(r, text="GPU NVIDIA не обнаружен.", font=("Segoe UI", 10),
                     bg=BG, fg=RED).pack(pady=10)

    def _stat_row(self, parent, title) -> dict:
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x", padx=12, pady=4)
        tk.Label(row, text=title, font=("Segoe UI", 10), bg=CARD, fg=MUTED
                 ).grid(row=0, column=0, sticky="w")
        val = tk.Label(row, text="—", font=("Segoe UI", 12, "bold"),
                       bg=CARD, fg=FG)
        val.grid(row=1, column=0, sticky="w", pady=(0, 4))
        bar = tk.Canvas(row, width=420, height=10, bg="#2a3140", highlightthickness=0)
        bar.grid(row=0, column=1, rowspan=2, sticky="e", padx=12, pady=(2, 4))
        bar.create_rectangle(0, 0, 0, 10, fill=ACCENT, outline="")
        return {"value": val, "bar": bar, "pct": 0.0}

    def _money_card(self, parent, title) -> dict:
        f = tk.Frame(parent, bg=CARD_HI, highlightbackground=CARD, highlightthickness=1)
        f.pack(side="left", expand=True, fill="both", padx=(0 if parent.winfo_children() else 0, 8))
        if f is parent.winfo_children()[-1] and len(parent.winfo_children()) == 1:
            f.pack(padx=0)
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

    def _on_gpu_change(self, *_):
        if HAS_NVML and self.gpu_count:
            self._start_collector(self.gpu_names.index(self._gpu_var.get()))

    def _tick(self):
        try:
            while True:
                try:
                    item = self.result_queue.get_nowait()
                except Exception:
                    break
                if "error" in item:
                    self.status.config(text=f"ошибка NVML: {item['error'][:60]}", fg=RED)
                    continue
                self._render(item)
        except Exception as e:
            self.status.config(text=f"ошибка: {e}", fg=RED)
        self.root.after(250, self._tick)

    def _render(self, d):
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
        self._draw_chart()

        now = datetime.now()
        for card, wh in ((self.card_day, self.store.wh_today(now)),
                         (self.card_hour, self.store.wh_last_hour(now)),
                         (self.card_month, self.store.wh_month(now))):
            card["cost"].config(text=f"{self.store.cost(wh):.2f} ₽")
            card["wh"].config(text=f"{wh / 1000.0:.3f} кВт·ч")

        self.status.config(text=f"обновлено {d['ts']}  ·  GPU {self.gpu_names.index(self._gpu_var.get()) if self.gpu_count else '—'}",
                           fg=MUTED)

    def _set_bar(self, row, frac, color=ACCENT):
        frac = max(0.0, min(1.0, frac))
        row["bar"].delete("all")
        row["bar"].create_rectangle(0, 0, int(420 * frac), 10, fill=color, outline="")
        row["pct"] = frac

    def _draw_chart(self):
        c = self.chart
        c.delete("all")
        w = c.winfo_width() or 700
        if len(self.samples) < 2:
            return
        pmax = max(p for _, p in self.samples)
        pmax = max(pmax * 1.15, 50.0)
        t0, t1 = self.samples[0][0], self.samples[-1][0]
        span = max(t1 - t0, 1.0)

        def xy(t, p):
            x = (t - t0) / span * (w - 4) + 2
            y = 94 - (p / pmax) * 88
            return x, y

        pts = [xy(t, p) for t, p in self.samples]
        c.create_polygon([(pts[0][0], 94)] + pts + [(pts[-1][0], 94)],
                         fill="#27405a", outline="")
        c.create_line(*pts, fill=ACCENT, width=2)
        c.create_text(6, 6, anchor="w", text=f"пик: {pmax / 1.15:.0f} Вт",
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

    def _reset_counters(self):
        if messagebox.askyesno("AIMemPowerCost", "Сбросить накопленные счётчики (день/месяц)?"):
            self.store.reset()
            self.samples.clear()
            self.status.config(text="счётчики сброшены", fg=AMBER)

    def on_close(self):
        try:
            if self.collector:
                self.stop_event.set()
                self.collector.join(timeout=3)
            self.store.flush()
        finally:
            self.root.destroy()


# ------------------------------------------------------------------ main ----

def main():
    if sys.platform != "win32" and "--allow-any-os" not in sys.argv:
        print("AIMemPowerCost: целевая платформа — Windows (NVML).")
        sys.exit(1)

    if "--selftest" in sys.argv:
        # Headless integration test: sample the real GPU for ~6 s.
        store = EnergyStore(os.path.join(APP_DIR, "selftest-state.json"))
        store.reset()
        stop = threading.Event()
        q = queue.Queue()
        col = GpuCollector(store, 0, stop, q)
        col.start()
        last = time.monotonic()
        samples = 0
        while time.monotonic() - last < 6:
            try:
                item = q.get(timeout=1)
                samples += 1
            except Exception:
                pass
        stop.set()
        col.join(timeout=3)
        store.flush()
        wh = store.wh_today()
        expected_w = 160  # ballpark idle load of a 4090
        ok = 0 < wh < 2.0  # ~100-450 W * 6 s / 3600 = 0.17-0.75 Wh
        print(f"selftest: {samples} samples, accumulated {wh:.4f} Wh (expected 0.05-2 Wh)")
        try:
            os.remove(os.path.join(APP_DIR, "selftest-state.json"))
        except OSError:
            pass
        sys.exit(0 if ok else 1)

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
