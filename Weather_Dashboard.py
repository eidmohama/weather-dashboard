#!/usr/bin/env python
# coding: utf-8

# In[1]:


"""
weather_dashboard.py
====================
Professional Weather Dashboard - M602 Computer Programming
Gisma University of Applied Sciences | Winter 2026
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import requests
import json
import csv
import os
from datetime import datetime
from collections import defaultdict

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

#  CONFIGURATION

API_KEY     = "755250ee033eec013ab2b25660d75944"
FAV_FILE    = "favorites.json"
EXPORT_FILE = "weather_report.csv"

# Colour palette
C = {
    "bg":      "#0D1117",
    "panel":   "#161B22",
    "border":  "#21262D",
    "accent":  "#58A6FF",
    "green":   "#3FB950",
    "warn":    "#D29922",
    "danger":  "#F85149",
    "text":    "#E6EDF3",
    "muted":   "#8B949E",
    "hot":     "#FF7B54",
    "cold":    "#74C0FC",
}

FONT_BIG   = ("Segoe UI", 32, "bold")
FONT_H2    = ("Segoe UI", 16, "bold")
FONT_H3    = ("Segoe UI", 11, "bold")
FONT_BODY  = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI",  9)
FONT_MONO  = ("Consolas",  9)

# Map condition keywords to short text labels
CONDITION_MAP = {
    "thunderstorm": "[Storm]",
    "drizzle":      "[Drizzle]",
    "rain":         "[Rain]",
    "snow":         "[Snow]",
    "mist":         "[Mist]",
    "fog":          "[Fog]",
    "haze":         "[Haze]",
    "smoke":        "[Smoke]",
    "dust":         "[Dust]",
    "tornado":      "[Tornado]",
    "clear":        "[Clear]",
    "cloud":        "[Cloudy]",
}

def weather_label(description: str) -> str:
    """Return a short condition label based on the description string."""
    desc = description.lower()
    for key, label in CONDITION_MAP.items():
        if key in desc:
            return label
    return "[Weather]"


#  API CLASS
class WeatherAPIError(Exception):
    """Raised for any OpenWeatherMap API or network error."""
    pass


class WeatherAPI:
    """
    Thin wrapper around the OpenWeatherMap REST API.
    Provides current weather and 5-day forecast.
    """
    BASE = "https://api.openweathermap.org/data/2.5"

    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("API key must not be empty.")
        self._key = api_key

    def _get(self, endpoint: str, params: dict) -> dict:
        """Internal GET helper – raises WeatherAPIError on any failure."""
        params["appid"] = self._key
        try:
            r = requests.get(f"{self.BASE}/{endpoint}", params=params, timeout=10)
            data = r.json()
            if r.status_code != 200:
                raise WeatherAPIError(data.get("message", "Unknown API error."))
            return data
        except requests.exceptions.ConnectionError:
            raise WeatherAPIError("No internet connection.")
        except requests.exceptions.Timeout:
            raise WeatherAPIError("Request timed out. Try again.")
        except requests.exceptions.RequestException as exc:
            raise WeatherAPIError(str(exc))

    def get_current(self, city: str, unit: str = "metric") -> dict:
        """Fetch current weather for a city."""
        return self._get("weather", {"q": city, "units": unit})

    def get_forecast(self, city: str, unit: str = "metric") -> dict:
        """Fetch 5-day / 3-hour forecast (40 slots) for a city."""
        return self._get("forecast", {"q": city, "units": unit, "cnt": 40})


#  FAVOURITES MANAGER

class FavouritesManager:
    """
    Manages a persistent list of favourite cities stored in a JSON file.
    """

    def __init__(self, path: str = FAV_FILE):
        self._path = path
        self._cities: list = []
        self._load()

    def _load(self) -> None:
        if not os.path.exists(self._path):
            return
        try:
            with open(self._path, "r", encoding="utf-8") as fh:
                raw = json.load(fh)
                self._cities = raw if isinstance(raw, list) else []
        except (json.JSONDecodeError, OSError):
            self._cities = []

    def _save(self) -> None:
        try:
            with open(self._path, "w", encoding="utf-8") as fh:
                json.dump(self._cities, fh, indent=2)
        except OSError as exc:
            raise IOError(f"Could not save favourites: {exc}")

    def add(self, city: str) -> bool:
        city = city.strip().title()
        if not city or city in self._cities:
            return False
        self._cities.append(city)
        self._save()
        return True

    def remove(self, city: str) -> bool:
        if city not in self._cities:
            return False
        self._cities.remove(city)
        self._save()
        return True

    @property
    def cities(self) -> list:
        return list(self._cities)


#  CHART FUNCTIONS  (each returns a matplotlib Figure)

def _make_fig(w: float = 7.0, h: float = 3.8):
    """Create a dark-themed figure/axes pair."""
    fig, ax = plt.subplots(figsize=(w, h))
    fig.patch.set_facecolor(C["panel"])
    ax.set_facecolor(C["bg"])
    ax.tick_params(colors=C["muted"], labelsize=8)
    ax.xaxis.label.set_color(C["muted"])
    ax.yaxis.label.set_color(C["muted"])
    ax.title.set_color(C["text"])
    for spine in ax.spines.values():
        spine.set_edgecolor(C["border"])
    ax.grid(axis="y", color=C["border"], linewidth=0.6, linestyle="--", alpha=0.8)
    return fig, ax


def make_temperature_chart(slots: list) -> plt.Figure:
    """Line chart: temperature and feels-like for next 24 hours."""
    data  = slots[:8]
    times = [s["dt_txt"].split()[1][:5] for s in data]
    temps = [s["main"]["temp"]       for s in data]
    feels = [s["main"]["feels_like"] for s in data]
    xs    = list(range(len(times)))

    fig, ax = _make_fig()
    ax.plot(xs, temps, color=C["accent"], lw=2.5, marker="o", ms=6,
            label="Temperature", zorder=3)
    ax.plot(xs, feels, color=C["warn"],   lw=1.8, marker="s", ms=4,
            linestyle="--", label="Feels Like", zorder=3)
    ax.fill_between(xs, temps, alpha=0.10, color=C["accent"])

    # Value labels on each point
    for x, t in zip(xs, temps):
        ax.annotate(f"{t:.0f}", (x, t), textcoords="offset points",
                    xytext=(0, 7), ha="center", color=C["accent"], fontsize=7)

    ax.set_xticks(xs)
    ax.set_xticklabels(times)
    ax.set_ylabel("Temperature")
    ax.set_title("Temperature - Next 24 Hours", pad=10, fontsize=11, fontweight="bold")
    ax.legend(facecolor=C["panel"], edgecolor=C["border"],
              labelcolor=C["text"], fontsize=8, loc="upper right")
    fig.tight_layout(pad=1.5)
    return fig


def make_humidity_chart(slots: list) -> plt.Figure:
    """Bar chart: humidity percentage for next 24 hours."""
    data     = slots[:8]
    times    = [s["dt_txt"].split()[1][:5] for s in data]
    humidity = [s["main"]["humidity"] for s in data]

    fig, ax = _make_fig()
    bar_colors = [C["green"] if h < 60 else C["warn"] if h < 80 else C["danger"]
                  for h in humidity]
    bars = ax.bar(times, humidity, color=bar_colors, alpha=0.85,
                  width=0.55, zorder=3)

    ax.axhline(60, color=C["warn"],   lw=0.9, linestyle="--", alpha=0.7, label="60% moderate")
    ax.axhline(80, color=C["danger"], lw=0.9, linestyle="--", alpha=0.7, label="80% high")
    ax.set_ylim(0, 112)

    for bar, val in zip(bars, humidity):
        ax.text(bar.get_x() + bar.get_width() / 2, val + 1.8,
                f"{val}%", ha="center", color=C["text"], fontsize=8)

    ax.set_ylabel("Humidity (%)")
    ax.set_title("Humidity - Next 24 Hours", pad=10, fontsize=11, fontweight="bold")
    ax.legend(facecolor=C["panel"], edgecolor=C["border"],
              labelcolor=C["text"], fontsize=8)
    fig.tight_layout(pad=1.5)
    return fig


def make_wind_chart(slots: list) -> plt.Figure:
    """Area chart: wind speed for next 24 hours."""
    data  = slots[:8]
    times = [s["dt_txt"].split()[1][:5] for s in data]
    winds = [s["wind"]["speed"] for s in data]
    xs    = list(range(len(times)))

    fig, ax = _make_fig()
    ax.fill_between(xs, winds, alpha=0.20, color=C["green"])
    ax.plot(xs, winds, color=C["green"], lw=2.2, marker="^", ms=6, zorder=3)

    for x, w in zip(xs, winds):
        ax.annotate(f"{w:.1f}", (x, w), textcoords="offset points",
                    xytext=(0, 7), ha="center", color=C["green"], fontsize=7)

    ax.set_xticks(xs)
    ax.set_xticklabels(times)
    ax.set_ylabel("Wind Speed (m/s)")
    ax.set_title("Wind Speed - Next 24 Hours", pad=10, fontsize=11, fontweight="bold")
    fig.tight_layout(pad=1.5)
    return fig


def make_5day_chart(slots: list) -> plt.Figure:
    """Grouped bar chart: daily min/max temperature for 5 days."""
    daily: dict = defaultdict(list)
    for s in slots:
        daily[s["dt_txt"].split()[0]].append(s["main"]["temp"])

    dates  = list(daily.keys())[:5]
    t_min  = [min(daily[d]) for d in dates]
    t_max  = [max(daily[d]) for d in dates]
    labels = [datetime.strptime(d, "%Y-%m-%d").strftime("%a\n%d %b") for d in dates]
    xs     = list(range(len(dates)))

    fig, ax = _make_fig()
    b1 = ax.bar([x - 0.22 for x in xs], t_min, 0.40,
                label="Min", color=C["cold"], alpha=0.85, zorder=3)
    b2 = ax.bar([x + 0.22 for x in xs], t_max, 0.40,
                label="Max", color=C["hot"],  alpha=0.85, zorder=3)

    for bar, val in zip(b1, t_min):
        ax.text(bar.get_x() + bar.get_width() / 2, val - 1.5,
                f"{val:.0f}", ha="center", va="top", color=C["cold"], fontsize=8)
    for bar, val in zip(b2, t_max):
        ax.text(bar.get_x() + bar.get_width() / 2, val + 0.5,
                f"{val:.0f}", ha="center", va="bottom", color=C["hot"], fontsize=8)

    ax.set_xticks(xs)
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylabel("Temperature")
    ax.set_title("5-Day Forecast - Daily Min / Max", pad=10, fontsize=11, fontweight="bold")
    ax.legend(facecolor=C["panel"], edgecolor=C["border"],
              labelcolor=C["text"], fontsize=8)
    fig.tight_layout(pad=1.5)
    return fig


#  MAIN APPLICATION

# Tab key constants — defined once, used everywhere
TAB_TEMP  = "  Temperature  "
TAB_HUM   = "  Humidity     "
TAB_WIND  = "  Wind Speed   "
TAB_5DAY  = "  5-Day        "
ALL_TABS  = (TAB_TEMP, TAB_HUM, TAB_WIND, TAB_5DAY)


class WeatherApp(tk.Tk):
    """
    Main Tkinter application for the Weather Dashboard.
    Inherits from tk.Tk so the instance IS the root window.
    """

    def __init__(self):
        super().__init__()
        self.title("Weather Dashboard  |  M602")
        self.geometry("1160x740")
        self.minsize(960, 620)
        self.configure(bg=C["bg"])

        # Core objects
        self._api      = WeatherAPI(API_KEY)
        self._favs     = FavouritesManager()

        # Tkinter variables
        self._city_var = tk.StringVar()
        self._unit_var = tk.StringVar(value="metric")
        self._status   = tk.StringVar(value="Enter a city name and press Search.")

        # State
        self._loading       = False
        self._last_curr     = None
        self._last_fore     = None
        self._chart_frames  = {}   

        # Build UI
        self._style_ttk()
        self._build_topbar()
        self._build_body()
        self._build_statusbar()
        self._refresh_favlist()

    # ── TTK styling ────

    def _style_ttk(self):
        s = ttk.Style(self)
        s.theme_use("default")

        s.configure("TNotebook",
                    background=C["bg"], borderwidth=0, tabmargins=[0, 0, 0, 0])
        s.configure("TNotebook.Tab",
                    background=C["panel"], foreground=C["muted"],
                    padding=[16, 8], font=FONT_SMALL, borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", C["accent"])],
              foreground=[("selected", "#ffffff")])

        s.configure("Treeview",
                    background=C["panel"], fieldbackground=C["panel"],
                    foreground=C["text"], rowheight=26, font=FONT_SMALL)
        s.configure("Treeview.Heading",
                    background=C["border"], foreground=C["muted"],
                    font=FONT_SMALL, relief="flat")
        s.map("Treeview", background=[("selected", C["accent"])])

    # ── Top bar ───────

    def _build_topbar(self):
        bar = tk.Frame(self, bg=C["panel"])
        bar.pack(fill="x", side="top")

        tk.Label(bar, text="  Weather Dashboard",
                 font=("Segoe UI", 15, "bold"),
                 bg=C["panel"], fg=C["text"]).pack(side="left", padx=16, pady=11)

        # Right-side controls
        ctrl = tk.Frame(bar, bg=C["panel"])
        ctrl.pack(side="right", padx=14, pady=8)

        # Unit toggle
        tk.Label(ctrl, text="Units:", bg=C["panel"],
                 fg=C["muted"], font=FONT_SMALL).pack(side="left", padx=(0, 4))
        for txt, val in [("Celsius", "metric"), ("Fahrenheit", "imperial")]:
            tk.Radiobutton(ctrl, text=txt, variable=self._unit_var, value=val,
                           bg=C["panel"], fg=C["text"], selectcolor=C["bg"],
                           activebackground=C["panel"], font=FONT_SMALL,
                           command=self._on_unit_toggle).pack(side="left", padx=2)

        self._sep(ctrl)

        # Export CSV
        self._mk_btn(ctrl, "Export CSV", C["green"],
                     self._export_csv).pack(side="left", padx=3)

        self._sep(ctrl)

        # Search entry + buttons
        self._entry = tk.Entry(ctrl, textvariable=self._city_var,
                               font=FONT_BODY, bg="#21262D", fg=C["text"],
                               insertbackground=C["text"], relief="flat",
                               width=22, bd=0)
        self._entry.pack(side="left", ipady=6, padx=(0, 6))
        self._entry.bind("<Return>", lambda _e: self._search())

        self._mk_btn(ctrl, "Search", C["accent"],
                     self._search).pack(side="left", padx=2)
        self._mk_btn(ctrl, "Save City", C["warn"],
                     self._save_city).pack(side="left", padx=2)

    # ── Body ─────────

    def _build_body(self):
        body = tk.Frame(self, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=10, pady=6)

        self._build_sidebar(body)

        right = tk.Frame(body, bg=C["bg"])
        right.pack(side="left", fill="both", expand=True)

        self._build_current_card(right)
        self._build_notebook(right)

    # Sidebar ────────────

    def _build_sidebar(self, parent):
        side = tk.Frame(parent, bg=C["panel"], width=170)
        side.pack(side="left", fill="y", padx=(0, 8))
        side.pack_propagate(False)

        tk.Label(side, text="Favourites", font=FONT_H3,
                 bg=C["panel"], fg=C["text"]
                 ).pack(anchor="w", padx=12, pady=(14, 6))

        self._favlist = tk.Listbox(
            side, font=FONT_SMALL,
            bg="#21262D", fg=C["text"],
            selectbackground=C["accent"],
            selectforeground="white",
            relief="flat", bd=0,
            highlightthickness=0,
            activestyle="none",
        )
        self._favlist.pack(fill="both", expand=True, padx=8)
        self._favlist.bind("<Double-Button-1>", lambda _e: self._load_fav())

        row = tk.Frame(side, bg=C["panel"])
        row.pack(fill="x", padx=8, pady=8)
        self._mk_btn(row, "Load",   C["green"],  self._load_fav
                     ).pack(side="left", fill="x", expand=True, padx=(0, 3))
        self._mk_btn(row, "Delete", C["danger"], self._del_fav
                     ).pack(side="left", fill="x", expand=True)

    # Current weather card ─────────
    def _build_current_card(self, parent):
        card = tk.Frame(parent, bg=C["panel"], padx=16, pady=12)
        card.pack(fill="x", pady=(0, 8))

        # Left: city + temperature + description
        left = tk.Frame(card, bg=C["panel"])
        left.pack(side="left", anchor="w")

        self._lbl_city = tk.Label(left, text="—",
                                  font=("Segoe UI", 17, "bold"),
                                  bg=C["panel"], fg=C["text"])
        self._lbl_city.pack(anchor="w")

        self._lbl_temp = tk.Label(left, text="—",
                                  font=FONT_BIG,
                                  bg=C["panel"], fg=C["accent"])
        self._lbl_temp.pack(anchor="w")

        self._lbl_desc = tk.Label(left, text="",
                                  font=FONT_BODY,
                                  bg=C["panel"], fg=C["muted"])
        self._lbl_desc.pack(anchor="w", pady=(2, 0))

        # Right: stat tiles
        grid = tk.Frame(card, bg=C["panel"])
        grid.pack(side="right", anchor="ne")

        self._stats: dict = {}
        stat_defs = [
            ("feels",    "Feels Like"),
            ("humidity", "Humidity"),
            ("wind",     "Wind"),
            ("pressure", "Pressure"),
            ("vis",      "Visibility"),
            ("sunrise",  "Sunrise"),
            ("sunset",   "Sunset"),
            ("updated",  "Updated"),
        ]
        for idx, (key, title) in enumerate(stat_defs):
            col = (idx % 4) * 2
            row = idx // 4
            tk.Label(grid, text=title, font=FONT_SMALL,
                     bg=C["panel"], fg=C["muted"]
                     ).grid(row=row * 2, column=col, padx=(10, 0), sticky="w")
            val_lbl = tk.Label(grid, text="—", font=FONT_H3,
                               bg=C["panel"], fg=C["text"])
            val_lbl.grid(row=row * 2 + 1, column=col, padx=(10, 16), sticky="w")
            self._stats[key] = val_lbl

    # Chart notebook ────────────
    def _build_notebook(self, parent):
        self._nb = ttk.Notebook(parent)
        self._nb.pack(fill="both", expand=True)

        for tab_name in ALL_TABS:
            # Outer frame added to notebook
            outer = tk.Frame(self._nb, bg=C["bg"])
            self._nb.add(outer, text=tab_name)
            # Inner frame is where the chart canvas will be packed
            inner = tk.Frame(outer, bg=C["bg"])
            inner.pack(fill="both", expand=True)
            self._chart_frames[tab_name] = inner

        # Placeholder labels so tabs look populated before first search
        for tab_name, frame in self._chart_frames.items():
            tk.Label(frame, text="Search for a city to see this chart.",
                     bg=C["bg"], fg=C["muted"], font=FONT_BODY
                     ).pack(expand=True)

    # Status bar ────────

    def _build_statusbar(self):
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")
        tk.Label(self, textvariable=self._status,
                 font=FONT_MONO, bg=C["panel"], fg=C["muted"],
                 anchor="w", padx=14).pack(fill="x", side="bottom")

    # ── Widget helpers ──────

    @staticmethod
    def _mk_btn(parent, text: str, bg: str, cmd) -> tk.Button:
        return tk.Button(parent, text=text, font=FONT_SMALL,
                         bg=bg, fg="white", relief="flat",
                         padx=10, pady=5, cursor="hand2",
                         activebackground=bg, command=cmd)

    @staticmethod
    def _sep(parent):
        tk.Frame(parent, bg=C["border"], width=1, height=24
                 ).pack(side="left", padx=8)

    def _set_status(self, msg: str):
        self._status.set(msg)

    # ── Search & fetch ───────

    def _search(self):
        city = self._city_var.get().strip()
        if not city:
            messagebox.showwarning("Empty Input", "Please enter a city name.")
            return
        if self._loading:
            return
        self._loading = True
        self._set_status(f"Fetching data for '{city}' ...")
        threading.Thread(
            target=self._fetch_data, args=(city,), daemon=True).start()

    def _fetch_data(self, city: str):
        """Runs in a background thread. Posts result back to main thread."""
        unit = self._unit_var.get()
        try:
            curr = self._api.get_current(city, unit)
            fore = self._api.get_forecast(city, unit)
            self.after(0, lambda: self._render(curr, fore))
        except WeatherAPIError as exc:
            self.after(0, lambda: self._on_error(str(exc)))
        except Exception as exc:
            self.after(0, lambda: self._on_error(f"Unexpected error: {exc}"))
        finally:
            self._loading = False

    def _on_error(self, msg: str):
        messagebox.showerror("Error", msg)
        self._set_status(f"Error: {msg}")

    def _on_unit_toggle(self):
        if self._city_var.get().strip():
            self._search()

    # ── Render results ─────────

    def _render(self, curr: dict, fore: dict):
        """Update the current-weather card and all four charts."""
        self._last_curr = curr
        self._last_fore = fore

        unit_sym = "C" if self._unit_var.get() == "metric" else "F"
        main     = curr["main"]
        wind     = curr["wind"]
        sys_info = curr["sys"]
        desc     = curr["weather"][0]["description"].capitalize()
        cond     = weather_label(desc)

        # ── Current card ────────
        self._lbl_city.config(
            text=f"{cond}  {curr['name']}, {sys_info['country']}")
        self._lbl_temp.config(
            text=f"{main['temp']:.1f} {chr(176)}{unit_sym}")
        self._lbl_desc.config(text=desc)

        vis = curr.get("visibility", "N/A")
        self._stats["feels"   ].config(text=f"{main['feels_like']:.1f}{chr(176)}{unit_sym}")
        self._stats["humidity"].config(text=f"{main['humidity']} %")
        self._stats["wind"    ].config(text=f"{wind['speed']} m/s")
        self._stats["pressure"].config(text=f"{main['pressure']} hPa")
        self._stats["vis"     ].config(text=f"{vis} m" if isinstance(vis, int) else "N/A")
        self._stats["sunrise" ].config(
            text=datetime.fromtimestamp(sys_info["sunrise"]).strftime("%H:%M"))
        self._stats["sunset"  ].config(
            text=datetime.fromtimestamp(sys_info["sunset"]).strftime("%H:%M"))
        self._stats["updated" ].config(
            text=datetime.now().strftime("%H:%M:%S"))

        # ── Charts ───────────
        flist = fore["list"]
        self._draw_chart(TAB_TEMP,  make_temperature_chart(flist))
        self._draw_chart(TAB_HUM,   make_humidity_chart(flist))
        self._draw_chart(TAB_WIND,  make_wind_chart(flist))
        self._draw_chart(TAB_5DAY,  make_5day_chart(flist))

        self._set_status(
            f"OK  |  {curr['name']}, {sys_info['country']}  |  "
            f"{desc}  |  Updated {datetime.now().strftime('%H:%M:%S')}")

    def _draw_chart(self, tab_name: str, fig: plt.Figure):
        """Clear a chart frame and embed a fresh matplotlib figure into it."""
        frame = self._chart_frames[tab_name]

        for widget in frame.winfo_children():
            widget.destroy()

        # Create and pack the new canvas
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        widget = canvas.get_tk_widget()
        widget.configure(bg=C["bg"])
        widget.pack(fill="both", expand=True, padx=4, pady=4)

        # Close the figure
        plt.close(fig)

    # ── Favourites ───────

    def _refresh_favlist(self):
        self._favlist.delete(0, tk.END)
        for city in self._favs.cities:
            self._favlist.insert(tk.END, city)

    def _save_city(self):
        city = self._city_var.get().strip()
        if not city:
            messagebox.showwarning("Empty", "Search for a city first.")
            return
        added = self._favs.add(city)
        self._refresh_favlist()
        self._set_status(
            f"Saved '{city}' to favourites." if added
            else f"'{city}' is already in your favourites.")

    def _load_fav(self):
        sel = self._favlist.curselection()
        if not sel:
            return
        city = self._favlist.get(sel[0])
        self._city_var.set(city)
        self._search()

    def _del_fav(self):
        sel = self._favlist.curselection()
        if not sel:
            return
        city = self._favlist.get(sel[0])
        self._favs.remove(city)
        self._refresh_favlist()
        self._set_status(f"Removed '{city}' from favourites.")

    # ── Export CSV ─────────

    def _export_csv(self):
        if self._last_fore is None:
            messagebox.showinfo("No Data", "Please search for a city first.")
            return
        unit_sym = "C" if self._unit_var.get() == "metric" else "F"
        try:
            with open(EXPORT_FILE, "w", newline="", encoding="utf-8") as fh:
                writer = csv.writer(fh)
                writer.writerow([
                    "DateTime",
                    f"Temp (deg {unit_sym})",
                    f"Feels Like (deg {unit_sym})",
                    "Humidity (%)",
                    "Wind (m/s)",
                    "Pressure (hPa)",
                    "Description",
                ])
                for slot in self._last_fore["list"]:
                    writer.writerow([
                        slot["dt_txt"],
                        slot["main"]["temp"],
                        slot["main"]["feels_like"],
                        slot["main"]["humidity"],
                        slot["wind"]["speed"],
                        slot["main"]["pressure"],
                        slot["weather"][0]["description"],
                    ])
            path = os.path.abspath(EXPORT_FILE)
            messagebox.showinfo("Exported", f"Report saved to:\n{path}")
            self._set_status(f"CSV exported -> {path}")
        except OSError as exc:
            messagebox.showerror("Export Error", str(exc))


#  ENTRY POINT

if __name__ == "__main__":
    app = WeatherApp()
    app.mainloop()


# In[ ]:




