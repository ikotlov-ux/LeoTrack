#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
LeoTrack.py — воспроизводимый конвейер анализа GPS-треков переднеазиатского леопарда
(Panthera pardus tulliana / ciscaucasica) на Северном Кавказе.

Запуск без аргументов:
    python LeoTrack.py
    -> окно Проводника для выбора одного или нескольких файлов треков (XLSX/CSV/GPX/GPKG/Movebank CSV),
    -> выбор папки результатов,
    -> окно метаданных животных (пол, дата выпуска, регион, статус) с сохранением animals_metadata.csv,
    -> окно параметров Google Earth Engine (Cloud Project ID и способ входа) с проверкой подключения,
    -> полный конвейер: QC -> шаги -> время суток -> ДЗЗ (GEE) -> обогащение -> анализ -> таблицы,
       рисунки, интерактивная карта, HTML-отчёт.

Командная строка (этапы ТЗ):
    python LeoTrack.py --stage validate|preprocess|fetch|extract|analyze|report|run-all
                       [--files f1.xlsx f2.xlsx ...] [--out DIR] [--config config.yaml]
                       [--no-gui] [--no-gee] [--gee-project ID] [--auth-mode localhost|notebook|gcloud|service_account]
                       [--sa-key key.json] [--public] [--overwrite]
    python LeoTrack.py --write-config config.yaml     # выгрузить конфигурацию по умолчанию с комментариями
    python LeoTrack.py --selftest                     # встроенные тесты на синтетических геометриях

Принципы (из ТЗ):
  * исходные записи не удаляются: qc_flag / qc_reason / included_analysis;
  * шаги не пересекают разрывы трека (burst_id);
  * скорость — «расчётная скорость перемещения между GPS-фиксациями» (минимальная средняя скорость);
  * облачный пиксель не трактуется как отсутствие снега или низкий NDVI (маскируется, считается отдельно);
  * использование отделено от предпочтения (iSSF), потенциальный охотничий/кормовой кластер
    (hunting_candidate) — от подтверждённого места добычи;
  * все пороги — в YAML-конфигурации; все случайные операции — с seed.

Зависимости (Python 3.10+): pandas numpy scipy pyproj shapely matplotlib openpyxl pyyaml
  рекомендуемые: earthengine-api rasterio statsmodels scikit-learn folium plotly
  pip install pandas numpy scipy pyproj shapely matplotlib openpyxl pyyaml tzdata earthengine-api rasterio statsmodels scikit-learn folium plotly
"""
from __future__ import annotations

__version__ = "1.0.0"

import argparse
import base64
import copy
import gc
import datetime as dt
import hashlib
import io
import json
import logging
import math
import os
import pickle
import platform
import re
import shutil
import sys
import time
import traceback
import warnings
import xml.etree.ElementTree as ET
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=RuntimeWarning)
try:
    import pandas as _pd_w
    warnings.filterwarnings("ignore", category=_pd_w.errors.PerformanceWarning)
except Exception:  # noqa: BLE001
    pass

# ----------------------------------------------------------------------------------------------
# Необязательные зависимости: модуль отключается с предупреждением, остальные этапы продолжаются
# ----------------------------------------------------------------------------------------------
OPTIONAL: Dict[str, Any] = {}


IMPORT_ERRORS: Dict[str, str] = {}


def _try_import(name: str, alias: Optional[str] = None):
    try:
        mod = __import__(name, fromlist=["*"])
        OPTIONAL[alias or name] = mod
        return mod
    except Exception as e:  # noqa: BLE001
        OPTIONAL[alias or name] = None
        IMPORT_ERRORS[alias or name] = f"{type(e).__name__}: {e}"
        return None


yaml = _try_import("yaml")
pyproj = _try_import("pyproj")
_try_import("shapely")
_try_import("shapely.geometry", "shapely_geometry")
scipy = _try_import("scipy")
_try_import("scipy.stats", "scipy_stats")
_try_import("scipy.special", "scipy_special")
_try_import("scipy.signal", "scipy_signal")
_try_import("scipy.ndimage", "scipy_ndimage")
_try_import("scipy.cluster.hierarchy", "scipy_hier")
_try_import("scipy.spatial.distance", "scipy_dist")
_try_import("scipy.optimize", "scipy_optimize")
rasterio = _try_import("rasterio")
_try_import("statsmodels.api", "sm")
_try_import("statsmodels.formula.api", "smf")
_try_import("statsmodels.discrete.conditional_models", "sm_cond")
_try_import("sklearn.metrics", "sk_metrics")
_try_import("sklearn.decomposition", "sk_decomp")
folium = _try_import("folium")
_try_import("folium.plugins", "folium_plugins")
_try_import("plotly.graph_objects", "plotly_go")
ee = _try_import("ee")
_try_import("pyarrow")

import matplotlib  # noqa: E402

if os.environ.get("DISPLAY", "") == "" and platform.system() != "Windows" and platform.system() != "Darwin":
    matplotlib.use("Agg")
else:
    try:
        matplotlib.use("Agg")
    except Exception:  # noqa: BLE001
        pass
import matplotlib.pyplot as plt  # noqa: E402

LOG = logging.getLogger("LeoTrack")
WARNINGS_FOR_REPORT: List[str] = []


def warn(msg: str) -> None:
    """Предупреждение в лог и в раздел ограничений отчёта."""
    LOG.warning(msg)
    WARNINGS_FOR_REPORT.append(msg)


# ==============================================================================================
# КОНФИГУРАЦИЯ ПО УМОЛЧАНИЮ (YAML с комментариями). [ЗООЛОГИ] — значения, требующие согласования.
# ==============================================================================================
DEFAULT_CONFIG_YAML = r"""
project:
  name: "LeoTrack"
  language: "ru"                 # язык отчёта (поддерживается ru)
  overwrite: false               # true — писать результаты в results/ поверх; false — новая папка results_<время>
  seed: 42                       # seed для всех случайных операций
  cache_dir: null                # кэш тайлов GEE и этапов; null — <папка входных файлов>/LeoTrack_cache (общий для всех папок результатов)

input:
  # Соответствие колонок входных файлов стандартной схеме (поиск без учёта регистра/пробелов/знаков)
  column_map:
    animal_id: ["ID", "animal_id", "individual-local-identifier", "tag-local-identifier", "name", "animal"]
    latitude: ["LAT", "latitude", "location-lat", "lat", "y_wgs84"]
    longitude: ["LON", "longitude", "location-long", "lon", "long", "lng", "x_wgs84"]
    timestamp: ["timestamp", "datetime", "date_time", "study-local-timestamp", "time_utc"]
    date: ["DATE", "date", "fix_date"]
    time: ["TIME (UTC +0)", "TIME", "time", "fix_time"]
    sex: ["sex", "SEX", "animal-sex", "пол"]
    gps_error_m: ["gps_error_m", "location-error-numerical", "error_m"]
    hdop: ["hdop", "gps:hdop", "HDOP"]
    pdop: ["pdop", "gps:pdop", "PDOP"]
    satellites: ["satellites", "gps:satellite-count", "sats", "nsats"]
    fix_type: ["fix_type", "gps:fix-type", "fix", "nav"]
    collar_id: ["collar_id", "tag-local-identifier", "tag_id", "collar"]
    age_class: ["age_class", "age"]
    release_date: ["release_date", "release"]
    speed_collar: ["ground-speed", "speed", "speed_collar"]
  source_timezone: "UTC"         # если в имени колонки времени есть «UTC +N», смещение берётся из имени
  source_crs: "EPSG:4326"
  metadata_file: null            # animals_metadata.csv (animal_id, sex, release_datetime_utc, region, status, ...)

time:
  local_timezone: "Europe/Moscow"   # местное гражданское время (UTC+3)

crs:
  work_crs: "EPSG:32638"         # метрическая CRS (UTM 38N покрывает 40–46° в.д. с искажением масштаба <0,2 %)

aoi:
  bbox: [36.0, 41.0, 49.5, 45.8] # lon_min, lat_min, lon_max, lat_max — ожидаемая область Северного Кавказа
  polygon_file: null             # необязательный полигон AOI (GPKG/SHP/GeoJSON)
  buffer_m: 3000                 # технический буфер вокруг треков для загрузки статичных слоёв

regions:
  method: "longitude_split"      # longitude_split | metadata
  split_lon: 42.0                # западнее — Западный Кавказ, восточнее — Центральный Кавказ [ЗООЛОГИ]
  west_name: "West_Caucasus"
  central_name: "Central_Caucasus"

qc:
  max_speed_km_h: 8.0            # максимальная правдоподобная расчётная скорость между фиксациями [ЗООЛОГИ]
  max_speed_min_dt_h: 0.5        # критерий скорости применяется при интервале >= этого значения
  spike_speed_km_h: 3.0          # «выброс-возврат»: скорость входа и выхода выше порога...
  spike_angle_deg: 165           # ...и угол разворота больше порога
  max_hdop: null                 # порог HDOP (если поле есть)
  min_satellites: null           # минимальное число спутников (если поле есть)
  min_interval_s: 60             # фиксации ближе по времени считаются дубликатами
  duplicate_policy: "exclude_conflicting"   # exclude_conflicting — исключать все записи с одним временем и разными координатами; keep_first
  burst_max_gap_h: 12            # разрыв непрерывного burst [ЗООЛОГИ]
  qc_iterations: 5

steps:
  min_step_for_angle_m: 15       # при более коротком шаге угол поворота не определён (шум GPS)
  path_sample_spacing_m: 30      # шаг дискретизации линии шага для path_* признаков
  path_max_samples: 400

regularization:
  enabled: true
  interval_h: 2.0                # общий интервал для сопоставимости (Leo ~1 ч, Victoria ~2 ч) [ЗООЛОГИ]
  tolerance_min: 20              # допустимое отклонение фиксации от узла сетки; интерполяции нет

solar:
  civil_twilight_deg: -6.0
  sunrise_deg: -0.833
  periods_A:                     # рабочие периоды зоологов — проверяемая гипотеза [ЗООЛОГИ]
    - {name: "morning_peak",    start: "03:00", end: "09:00"}
    - {name: "day_rest",        start: "09:00", end: "20:00"}
    - {name: "evening_peak",    start: "20:00", end: "01:00"}
    - {name: "deep_night_rest", start: "01:00", end: "03:00"}

seasons:
  primary: "calendar"            # calendar (таблица seasons.csv / календарь ниже) | objective (MODIS снег + NDVI)
  table: null                    # seasons.csv: region,year,season,start,end (start/end: MM-DD или YYYY-MM-DD)
  default_calendar:              # используется, если нет таблицы [ЗООЛОГИ]
    - {season: "snow_period",        start: "12-01", end: "03-15"}
    - {season: "leaf_off_snow_free", start: "03-16", end: "04-30"}
    - {season: "active_vegetation",  start: "05-01", end: "10-15"}
    - {season: "leaf_off_snow_free", start: "10-16", end: "11-30"}
  objective:
    snow_cover_threshold: 40     # MOD10A1 NDSI_Snow_Cover (0–100), выше — снег
    snow_window_days: 7
    snow_fraction_min: 0.5       # доля снежных дней среди валидных за окно
    min_valid_days: 2
    ndvi_active_threshold: 0.55  # MOD13Q1 NDVI для активной вегетации [ЗООЛОГИ]
    smooth_days: 7
    transition_days: 3

post_release:
  bins_days: [0, 7, 14, 21, 28, 42, 100000]
  bin_labels: ["0-7", "8-14", "15-21", "22-28", "29-42", ">42"]
  key_metrics: ["daily_distance_km", "net_disp_release_km", "straightness", "explored_area_km2", "cluster_time_frac", "mean_speed_m_h"]
  confirm_days: 14
  sensitivity_confirm_days: [7, 14, 21]
  min_segment_days: 7
  bootstrap: 200
  max_days: 400

clusters:
  radius_m: 150
  containment: 0.95              # p95 расстояний до центра <= radius_m
  min_duration_h: 24
  allowed_gap_h: 6               # максимальный пропуск внутри кластера
  max_excursion_h: 8             # выход из радиуса с возвращением, не дольше
  min_fix_fraction: 0.5          # минимальная доля ожидаемых фиксаций
  merge_distance_m: 150          # объединение соседних эпизодов с близкими центрами
  merge_gap_h: 6
  bootstrap: 500
  sensitivity_radii: [100, 150, 200]
  sensitivity_durations: [12, 24, 36]

daily:
  min_coverage: 0.75             # доля суток, покрытая валидными шагами
  min_fix_fraction: 0.6          # доля ожидаемых фиксаций

behavior:
  rest_max_speed_m_h: 50         # простой режим: покой/медленно [ЗООЛОГИ]
  directed_min_speed_m_h: 500    # направленное перемещение
  directed_max_abs_turn_deg: 90
  hmm:
    enabled: true
    n_states: [2, 3]
    primary_n_states: 3          # число состояний задаёт исследователь, AIC/BIC — только справочно
    min_steps: 200
    max_iter: 100
    tol: 1.0e-6

terrain:
  tpi_small_radius_m: 150
  tpi_large_radius_m: 750
  flat_slope_deg: 5              # для классификации форм рельефа
  aspect_flat_slope_deg: 3       # экспозиция flat
  elevation_band_m: 500
  buffer_radii_m: [30, 100, 250, 500, 1000]
  buffer_continuous: ["elevation", "slope_deg", "tri", "tpi_large"]

habitat:
  reclass_table: null            # habitat_reclass.csv: code,habitat[,min_elev,max_elev]
  regional_raster: null          # пользовательская региональная карта растительности (приоритет выше WorldCover)
  regional_reclass: null         # CSV: code,habitat

issf:
  enabled: true
  n_available: 10
  covariates: ["elevation_z", "slope_z", "tri_z", "tpi_large_z", "northness", "eastness", "hab_forest", "hab_open"]
  min_strata: 100
  folds: 5
  max_vif: 10

gee:
  enabled: true
  project: null                  # Google Cloud Project ID, зарегистрированный для Earth Engine
  auth_mode: "localhost"         # localhost | notebook | gcloud | service_account
  service_account_key: null
  chunk_size: 400                # максимум точек в одном запросе (порция = подряд идущие сутки одного животного)
  chunk_days: 10                 # максимум суток в порции; при ошибке GEE порция делится пополам
  threads: 4
  max_retries: 5
  dynamic_points: "all"          # all | regular (только регуляризованные фиксации)
  tile_size_m: 10000
  tile_overlap_m: 2500           # >= tpi_large_radius_m + max(buffer_radii_m)
  dem_scale_m: 30
  wc_scale_m: 10
  windows_days: [3, 7, 15, 30]   # симметричные окна ±N суток для S2/Landsat NDVI/NDSI
  dw_windows_days: [7, 15]       # окна Dynamic World
  snow_windows_days: [1, 3, 7, 14, 30]   # ретроспективные окна MOD10A1
  snow_lookback_days: 60
  landsat_before: "2019-01-01"   # Landsat C2 L2 для дат до этой (ранний период без S2 L2A)
  sources:
    dem: "COPERNICUS/DEM/GLO30_2024_1"
    dem_band: "DEM"
    dem_fallback: "NASA/NASADEM_HGT/001"
    worldcover: "ESA/WorldCover/v200"
    dynamic_world: "GOOGLE/DYNAMICWORLD/V1"
    s2: "COPERNICUS/S2_SR_HARMONIZED"
    landsat: ["LANDSAT/LT05/C02/T1_L2", "LANDSAT/LE07/C02/T1_L2", "LANDSAT/LC08/C02/T1_L2", "LANDSAT/LC09/C02/T1_L2"]
    modis_snow: "MODIS/061/MOD10A1"
    modis_ndvi: "MODIS/061/MOD13Q1"
    era5: "ECMWF/ERA5_LAND/DAILY_AGGR"   # суточные средние snow_depth, SWE, temperature_2m

privacy:
  public_mode: false             # огрубление координат и маскирование мест выпуска и кластеров
  round_m: 1000
  mask_release_radius_m: 3000
  mask_clusters: true

stats:
  bootstrap: 1000
  min_animals_group: 2

extra_vectors: {}                # name: путь к слою (дороги, поселения, ООПТ) -> dist_<name>_m

report:
  dpi: 130
  max_table_rows: 40
  map_max_points: 8000
"""


def deep_update(base: dict, upd: dict) -> dict:
    for k, v in (upd or {}).items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            deep_update(base[k], v)
        else:
            base[k] = v
    return base


def load_config(path: Optional[str]) -> dict:
    if yaml is None:
        raise SystemExit("Не установлен pyyaml: pip install pyyaml")
    cfg = yaml.safe_load(DEFAULT_CONFIG_YAML)
    if path:
        with open(path, "r", encoding="utf-8") as fh:
            user = yaml.safe_load(fh) or {}
        deep_update(cfg, user)
    return cfg


def validate_config(cfg: dict) -> List[str]:
    """Проверка конфигурации; возвращает список понятных ошибок (пустой — всё в порядке)."""
    errs: List[str] = []

    def num(path, lo=None, hi=None):
        cur = cfg
        for p in path.split("."):
            if not isinstance(cur, dict) or p not in cur:
                errs.append(f"нет параметра {path}")
                return
            cur = cur[p]
        if cur is None:
            return
        try:
            v = float(cur)
        except Exception:  # noqa: BLE001
            errs.append(f"{path}: ожидается число, получено {cur!r}")
            return
        if lo is not None and v < lo:
            errs.append(f"{path}: значение {v} меньше {lo}")
        if hi is not None and v > hi:
            errs.append(f"{path}: значение {v} больше {hi}")

    num("qc.max_speed_km_h", 0.1, 200)
    num("qc.burst_max_gap_h", 0.1, 10000)
    num("clusters.radius_m", 1, 10000)
    num("clusters.containment", 0.5, 1.0)
    num("clusters.min_duration_h", 0.5, 1000)
    num("regularization.interval_h", 0.05, 48)
    num("daily.min_coverage", 0, 1)
    num("gee.chunk_size", 1, 5000)
    num("issf.n_available", 1, 200)
    if cfg["gee"]["tile_overlap_m"] < cfg["terrain"]["tpi_large_radius_m"] + max(cfg["terrain"]["buffer_radii_m"]):
        errs.append("gee.tile_overlap_m должно быть >= terrain.tpi_large_radius_m + max(terrain.buffer_radii_m)")
    if cfg["seasons"]["primary"] not in ("calendar", "objective"):
        errs.append("seasons.primary: calendar | objective")
    if cfg["gee"]["auth_mode"] not in ("localhost", "notebook", "gcloud", "service_account", "appdefault"):
        errs.append("gee.auth_mode: localhost | notebook | gcloud | service_account | appdefault")
    for p in cfg["solar"]["periods_A"]:
        for k in ("start", "end"):
            if not re.match(r"^\d{1,2}:\d{2}$", str(p.get(k, ""))):
                errs.append(f"solar.periods_A {p.get('name')}: {k} в формате ЧЧ:ММ")
    try:
        if pyproj is not None:
            pyproj.CRS.from_user_input(cfg["crs"]["work_crs"])
    except Exception:  # noqa: BLE001
        errs.append(f"crs.work_crs: неизвестная CRS {cfg['crs']['work_crs']}")
    try:
        from zoneinfo import ZoneInfo
        ZoneInfo(cfg["time"]["local_timezone"])
    except Exception:  # noqa: BLE001
        errs.append(f"time.local_timezone: неизвестный часовой пояс {cfg['time']['local_timezone']} (на Windows: pip install tzdata)")
    return errs


# ----------------------------------------------------------------------------------------------
# Утилиты: хэши, кэш, журнал
# ----------------------------------------------------------------------------------------------
def sha256_file(path: str, block: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        while True:
            b = fh.read(block)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def hash_obj(obj: Any) -> str:
    s = json.dumps(obj, sort_keys=True, default=str, ensure_ascii=False)
    return hashlib.sha256(s.encode("utf-8")).hexdigest()[:16]


def script_hash() -> str:
    try:
        return sha256_file(os.path.abspath(__file__))[:16]
    except Exception:  # noqa: BLE001
        return "unknown"


class Cache:
    """Простой файловый кэш pickle по ключу (этап + хэш входов и параметров)."""

    def __init__(self, root: Path):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)

    def path(self, name: str, key: str) -> Path:
        return self.root / f"{name}_{key}.pkl"

    def get(self, name: str, key: str):
        p = self.path(name, key)
        if p.exists():
            try:
                with open(p, "rb") as fh:
                    LOG.info(f"Кэш: используется {p.name}")
                    return pickle.load(fh)
            except Exception:  # noqa: BLE001
                return None
        return None

    def put(self, name: str, key: str, obj: Any) -> None:
        p = self.path(name, key)
        tmp = p.with_suffix(".tmp")
        with open(tmp, "wb") as fh:
            pickle.dump(obj, fh, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, p)


def setup_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"leotrack_{dt.datetime.now():%Y%m%d_%H%M%S}.log"
    LOG.setLevel(logging.INFO)
    LOG.handlers.clear()
    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%H:%M:%S")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    LOG.addHandler(fh)
    LOG.addHandler(sh)
    return log_path


def check_disk(path: Path, min_gb: float = 2.0) -> None:
    try:
        free = shutil.disk_usage(str(path)).free / 1e9
        if free < min_gb:
            warn(f"Мало места на диске: {free:.1f} ГБ свободно в {path}. Загрузка растров может прерваться.")
    except Exception:  # noqa: BLE001
        pass


def wrap180(a):
    return (np.asarray(a) + 180.0) % 360.0 - 180.0


def safe_stage(name: str):
    """Декоратор: при сбое модуля продолжаем независимые этапы, ошибка попадает в лог и отчёт."""

    def deco(fn):
        def inner(*a, **k):
            try:
                return fn(*a, **k)
            except Exception as e:  # noqa: BLE001
                warn(f"Модуль «{name}» не выполнен: {type(e).__name__}: {e}")
                LOG.debug(traceback.format_exc())
                LOG.error(traceback.format_exc())
                return None

        inner.__name__ = fn.__name__
        return inner

    return deco


def mean_ci(values, conf: float = 0.95) -> Tuple[float, float, float, int]:
    """Среднее и t-доверительный интервал (по умолчанию 95%). Возвращает (mean, low, high, n)."""
    v = pd.to_numeric(pd.Series(values), errors="coerce").dropna().to_numpy(float)
    n = len(v)
    if n == 0:
        return np.nan, np.nan, np.nan, 0
    m = float(v.mean())
    if n < 2:
        return m, np.nan, np.nan, n
    from scipy.stats import t as _t
    half = float(_t.ppf(0.5 + conf / 2, n - 1) * v.std(ddof=1) / math.sqrt(n))
    return m, m - half, m + half, n


# ==============================================================================================
# ИНТЕРАКТИВНЫЙ ИНТЕРФЕЙС (tkinter): выбор файлов, папки, метаданных и параметров GEE
# ==============================================================================================
def _tk_available() -> bool:
    try:
        import tkinter  # noqa: F401
        if platform.system() == "Linux" and not os.environ.get("DISPLAY"):
            return False
        return True
    except Exception:  # noqa: BLE001
        return False


def _tk_root():
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:  # noqa: BLE001
        pass
    return root


def gui_select_files(initialdir: Optional[str] = None) -> List[str]:
    """Окно Проводника для выбора нескольких файлов треков."""
    if not _tk_available():
        s = input("Пути к файлам треков через ';': ").strip()
        return [p.strip().strip('"') for p in s.split(";") if p.strip()]
    from tkinter import filedialog
    root = _tk_root()
    files = filedialog.askopenfilenames(
        parent=root,
        title="LeoTrack: выберите один или несколько файлов GPS-треков",
        initialdir=initialdir or os.getcwd(),
        filetypes=[("Треки (XLSX, CSV, GPX, GPKG)", "*.xlsx *.xls *.csv *.txt *.gpx *.gpkg"),
                   ("Excel", "*.xlsx *.xls"), ("CSV / Movebank", "*.csv *.txt"), ("GPX", "*.gpx"),
                   ("GeoPackage", "*.gpkg"), ("Все файлы", "*.*")])
    root.destroy()
    return list(files)


def gui_select_dir(initialdir: Optional[str] = None, title: str = "LeoTrack: папка для результатов") -> Optional[str]:
    if not _tk_available():
        s = input(f"{title} [{initialdir}]: ").strip().strip('"')
        return s or initialdir
    from tkinter import filedialog
    root = _tk_root()
    d = filedialog.askdirectory(parent=root, title=title, initialdir=initialdir or os.getcwd(), mustexist=False)
    root.destroy()
    return d or initialdir


def gui_ask_config(initialdir: Optional[str] = None) -> Optional[str]:
    if not _tk_available():
        s = input("Файл конфигурации YAML (Enter - по умолчанию): ").strip().strip('"')
        return s or None
    from tkinter import filedialog, messagebox
    root = _tk_root()
    use = messagebox.askyesno("LeoTrack", "Загрузить собственный файл конфигурации YAML?\n\n"
                                          "«Нет» - параметры по умолчанию (будут сохранены в config_used.yaml "
                                          "в папке результатов для редактирования).", parent=root)
    path = None
    if use:
        path = filedialog.askopenfilename(parent=root, title="Файл конфигурации LeoTrack", initialdir=initialdir,
                                          filetypes=[("YAML", "*.yaml *.yml"), ("Все файлы", "*.*")]) or None
    root.destroy()
    return path


def gui_metadata(meta: pd.DataFrame) -> pd.DataFrame:
    """Редактируемая таблица метаданных животных: пол, дата выпуска (UTC), регион, статус, ошейник."""
    if not _tk_available():
        print("\nМетаданные животных (Enter - оставить значение):")
        for i, r in meta.iterrows():
            for col, prompt in (("sex", "пол M/F/unknown"), ("release_datetime_utc", "выпуск UTC ГГГГ-ММ-ДД ЧЧ:ММ"),
                                ("region", "регион"), ("status", "статус released/wild/unknown")):
                v = input(f"  {r['animal_id']}: {prompt} [{r[col]}]: ").strip()
                if v:
                    meta.at[i, col] = v
        return meta
    import tkinter as tk
    from tkinter import ttk, messagebox
    root = tk.Tk()
    root.title("LeoTrack: метаданные животных")
    try:
        root.attributes("-topmost", True)
    except Exception:  # noqa: BLE001
        pass
    frm = ttk.Frame(root, padding=10)
    frm.grid(sticky="nsew")
    ttk.Label(frm, text=("Проверьте и заполните сведения. Пол обязателен для сравнения самцов и самок.\n"
                         "Дата выпуска по умолчанию = первая фиксация (release_source=first_fix_assumed); "
                         "при наличии подтверждённой даты укажите её (UTC).\nРегион определён по долготе - проверьте."),
              justify="left").grid(row=0, column=0, columnspan=7, sticky="w", pady=(0, 8))
    heads = ["Животное", "Пол", "Выпуск (UTC)", "Регион", "Статус", "Широта выпуска", "Долгота выпуска"]
    for j, h in enumerate(heads):
        ttk.Label(frm, text=h, font=("TkDefaultFont", 9, "bold")).grid(row=1, column=j, padx=3, sticky="w")
    widgets = []
    regions = sorted(set(meta["region"].astype(str)) | {"West_Caucasus", "Central_Caucasus"})
    for i, (_, r) in enumerate(meta.iterrows()):
        row = i + 2
        ttk.Label(frm, text=str(r["animal_id"])).grid(row=row, column=0, sticky="w")
        v_sex = tk.StringVar(value=str(r["sex"]))
        ttk.Combobox(frm, textvariable=v_sex, values=["M", "F", "unknown"], width=9).grid(row=row, column=1, padx=3)
        v_rel = tk.StringVar(value=str(r["release_datetime_utc"]))
        ttk.Entry(frm, textvariable=v_rel, width=20).grid(row=row, column=2, padx=3)
        v_reg = tk.StringVar(value=str(r["region"]))
        ttk.Combobox(frm, textvariable=v_reg, values=regions, width=18).grid(row=row, column=3, padx=3)
        v_st = tk.StringVar(value=str(r["status"]))
        ttk.Combobox(frm, textvariable=v_st, values=["released", "wild", "unknown"], width=10).grid(row=row, column=4, padx=3)
        v_la = tk.StringVar(value=str(r["release_lat"]))
        ttk.Entry(frm, textvariable=v_la, width=11).grid(row=row, column=5, padx=3)
        v_lo = tk.StringVar(value=str(r["release_lon"]))
        ttk.Entry(frm, textvariable=v_lo, width=11).grid(row=row, column=6, padx=3)
        widgets.append((v_sex, v_rel, v_reg, v_st, v_la, v_lo))
    result = {"ok": False}

    def _num(txt):
        txt = str(txt).strip().replace(",", ".")
        if txt == "" or txt.lower() == "nan":
            return np.nan
        return float(txt)

    def on_ok():
        bad = []
        for i, ws in enumerate(widgets):
            a = meta.at[meta.index[i], "animal_id"]
            try:
                v = _num(ws[4].get()), _num(ws[5].get())
                if not (np.isnan(v[0]) or -90 <= v[0] <= 90) or not (np.isnan(v[1]) or -180 <= v[1] <= 180):
                    raise ValueError
            except ValueError:
                bad.append(f"{a}: координаты выпуска")
            rel = ws[1].get().strip()
            if rel and pd.isna(pd.to_datetime(rel, errors="coerce")):
                bad.append(f"{a}: дата выпуска (формат ГГГГ-ММ-ДД ЧЧ:ММ)")
        if bad:
            messagebox.showerror("LeoTrack", "Исправьте значения:\n" + "\n".join(bad), parent=root)
            return
        for c in ("release_lat", "release_lon"):
            meta[c] = pd.to_numeric(meta[c], errors="coerce").astype(float)
        for c in ("sex", "release_datetime_utc", "release_source", "region", "status"):
            meta[c] = meta[c].astype(object)
        for i, ws in enumerate(widgets):
            idx = meta.index[i]
            sex = ws[0].get().strip().upper()
            meta.at[idx, "sex"] = {"M": "M", "F": "F", "М": "M", "Ж": "F", "MALE": "M", "FEMALE": "F"}.get(sex, "unknown")
            rel = ws[1].get().strip()
            if rel != str(meta.at[idx, "release_datetime_utc"]):
                meta.at[idx, "release_source"] = "user_input"
            meta.at[idx, "release_datetime_utc"] = rel
            meta.at[idx, "region"] = ws[2].get().strip()
            meta.at[idx, "status"] = ws[3].get().strip()
            meta.at[idx, "release_lat"] = _num(ws[4].get())
            meta.at[idx, "release_lon"] = _num(ws[5].get())
        if (meta["sex"] == "unknown").any():
            if not messagebox.askyesno("LeoTrack", "Для части животных пол не указан (unknown).\n"
                                                   "Межполовые сравнения будут неполными. Продолжить?", parent=root):
                return
        result["ok"] = True
        root.destroy()

    ttk.Button(frm, text="Сохранить и продолжить", command=on_ok).grid(row=len(widgets) + 2, column=0, columnspan=7, pady=10)
    root.protocol("WM_DELETE_WINDOW", on_ok)
    root.mainloop()
    return meta


GEE_SETTINGS_FILE = Path.home() / ".leotrack" / "gee_settings.json"


def gui_gee_settings(cfg: dict) -> dict:
    """Окно параметров Earth Engine. Пароль Google в скрипт не вводится: вход выполняется на странице Google,
    токен сохраняется самим earthengine-api; для сервисного аккаунта указывается JSON-ключ."""
    saved = {}
    if GEE_SETTINGS_FILE.exists():
        try:
            saved = json.loads(GEE_SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            saved = {}
    g = cfg["gee"]
    s = {"enabled": g.get("enabled", True),
         "project": g.get("project") or saved.get("project") or os.environ.get("EARTHENGINE_PROJECT", ""),
         "auth_mode": saved.get("auth_mode", g.get("auth_mode", "localhost")),
         "service_account_key": g.get("service_account_key") or saved.get("service_account_key", ""),
         "force_reauth": False,
         "public_mode": cfg["privacy"]["public_mode"],
         "overwrite": cfg["project"]["overwrite"]}
    if not _tk_available():
        use = input("Использовать Google Earth Engine? [Y/n]: ").strip().lower()
        s["enabled"] = use not in ("n", "no", "н", "нет")
        if s["enabled"]:
            s["project"] = input(f"Google Cloud Project ID для Earth Engine [{s['project']}]: ").strip() or s["project"]
            m = input(f"Способ входа localhost/notebook/gcloud/service_account [{s['auth_mode']}]: ").strip()
            s["auth_mode"] = m or s["auth_mode"]
            if s["auth_mode"] == "service_account":
                s["service_account_key"] = input("Путь к JSON-ключу сервисного аккаунта: ").strip().strip('"')
        return s
    import tkinter as tk
    from tkinter import ttk, filedialog
    root = tk.Tk()
    root.title("LeoTrack: Google Earth Engine и параметры запуска")
    try:
        root.attributes("-topmost", True)
    except Exception:  # noqa: BLE001
        pass
    f = ttk.Frame(root, padding=12)
    f.grid(sticky="nsew")
    v_en = tk.BooleanVar(value=s["enabled"])
    ttk.Checkbutton(f, text="Загружать данные ДЗЗ из Google Earth Engine (DEM, WorldCover, Dynamic World, "
                            "Sentinel-2, Landsat, MODIS, ERA5-Land)", variable=v_en).grid(row=0, column=0, columnspan=3, sticky="w")
    ttk.Label(f, text="Google Cloud Project ID:").grid(row=1, column=0, sticky="w", pady=4)
    v_pr = tk.StringVar(value=s["project"])
    ttk.Entry(f, textvariable=v_pr, width=40).grid(row=1, column=1, columnspan=2, sticky="w")
    ttk.Label(f, text="Способ входа:").grid(row=2, column=0, sticky="nw", pady=4)
    v_am = tk.StringVar(value=s["auth_mode"])
    modes = [("localhost", "Браузер: вход в Google-аккаунт (рекомендуется для ПК)"),
             ("notebook", "Ссылка + код подтверждения (если браузер не открывается)"),
             ("gcloud", "Google Cloud SDK (gcloud уже установлен и настроен)"),
             ("service_account", "Сервисный аккаунт (JSON-ключ)")]
    for i, (m, txt) in enumerate(modes):
        ttk.Radiobutton(f, text=txt, value=m, variable=v_am).grid(row=2 + i, column=1, columnspan=2, sticky="w")
    ttk.Label(f, text="JSON-ключ сервисного аккаунта:").grid(row=6, column=0, sticky="w", pady=4)
    v_key = tk.StringVar(value=s["service_account_key"])
    ttk.Entry(f, textvariable=v_key, width=40).grid(row=6, column=1, sticky="w")

    def pick_key():
        p = filedialog.askopenfilename(parent=root, title="JSON-ключ сервисного аккаунта",
                                       filetypes=[("JSON", "*.json"), ("Все файлы", "*.*")])
        if p:
            v_key.set(p)

    ttk.Button(f, text="Обзор...", command=pick_key).grid(row=6, column=2, sticky="w")
    v_force = tk.BooleanVar(value=False)
    ttk.Checkbutton(f, text="Войти заново (сбросить сохранённый токен)", variable=v_force).grid(row=7, column=1, columnspan=2, sticky="w")
    ttk.Separator(f).grid(row=8, column=0, columnspan=3, sticky="ew", pady=8)
    v_pub = tk.BooleanVar(value=s["public_mode"])
    ttk.Checkbutton(f, text="Публичный режим отчёта (огрубление координат, маскирование мест выпуска и кластеров)",
                    variable=v_pub).grid(row=9, column=0, columnspan=3, sticky="w")
    v_ow = tk.BooleanVar(value=s["overwrite"])
    ttk.Checkbutton(f, text="Перезаписать предыдущие результаты (иначе - новая папка results_<время>)",
                    variable=v_ow).grid(row=10, column=0, columnspan=3, sticky="w")
    ttk.Label(f, foreground="#555", justify="left",
              text=("Пароль Google в скрипт не вводится: откроется страница входа Google, токен сохранится в профиле\n"
                    "пользователя (~/.config/earthengine). Проект должен быть зарегистрирован для Earth Engine:\n"
                    "https://code.earthengine.google.com/register")).grid(row=11, column=0, columnspan=3, sticky="w", pady=6)

    def ok():
        s.update(enabled=v_en.get(), project=v_pr.get().strip(), auth_mode=v_am.get(),
                 service_account_key=v_key.get().strip(), force_reauth=v_force.get(),
                 public_mode=v_pub.get(), overwrite=v_ow.get())
        root.destroy()

    ttk.Button(f, text="Продолжить", command=ok).grid(row=12, column=0, columnspan=3, pady=8)
    root.protocol("WM_DELETE_WINDOW", ok)
    root.mainloop()
    try:
        GEE_SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        GEE_SETTINGS_FILE.write_text(json.dumps({k: s[k] for k in ("project", "auth_mode", "service_account_key")},
                                                ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    return s


def gui_message(text: str, title: str = "LeoTrack", error: bool = False, ask: bool = False) -> bool:
    if not _tk_available():
        print(f"\n[{title}] {text}")
        if ask:
            return input("Продолжить? [Y/n]: ").strip().lower() not in ("n", "no", "нет")
        return True
    from tkinter import messagebox
    root = _tk_root()
    if ask:
        r = messagebox.askyesno(title, text, parent=root)
    elif error:
        messagebox.showerror(title, text, parent=root)
        r = True
    else:
        messagebox.showinfo(title, text, parent=root)
        r = True
    root.destroy()
    return r


# ==============================================================================================
# GEE: аутентификация
# ==============================================================================================
GEE_HELP = """Не удалось подключиться к Google Earth Engine.
Проверьте:
 1) pip install earthengine-api
 2) Проект Google Cloud зарегистрирован для Earth Engine: https://code.earthengine.google.com/register
    (некоммерческое/научное использование бесплатно), включён Earth Engine API.
 3) Указан правильный Project ID (не имя, а идентификатор, например my-project-123).
 4) Вход: в командной строке можно выполнить  earthengine authenticate  и повторить запуск.
 5) Для сервисного аккаунта: аккаунт добавлен в проект с ролью Earth Engine Resource Viewer/Writer
    и Service Usage Consumer, указан JSON-ключ.
Без GEE будут выполнены только модули, не требующие ДЗЗ (QC, шаги, время суток, кластеры, суточный ход и т.д.)."""


GEE_PACKAGES = [("ee", "earthengine-api"), ("rasterio", "rasterio")]


def ensure_gee_packages(interactive: bool = True) -> bool:
    """Проверяет earthengine-api и rasterio в ТЕКУЩЕМ интерпретаторе; при отсутствии предлагает установить через pip."""
    global ee, rasterio
    missing = [(m, p) for m, p in GEE_PACKAGES if OPTIONAL.get(m) is None]
    if not missing:
        return True
    det = "\n".join(f"  {p}: {IMPORT_ERRORS.get(m, 'не найден')}" for m, p in missing)
    pkgs = " ".join(p for _, p in missing)
    cmd = f'"{sys.executable}" -m pip install {pkgs}'
    LOG.info(f"Для GEE не хватает пакетов в интерпретаторе {sys.executable}:\n{det}")
    if interactive and gui_message(f"Для загрузки данных ДЗЗ не хватает пакетов:\n{det}\n\n"
                                   f"Интерпретатор: {sys.executable}\n\nУстановить сейчас командой\n{cmd} ?",
                                   ask=True):
        import subprocess
        LOG.info(f"Установка: {cmd}")
        r = subprocess.run([sys.executable, "-m", "pip", "install", *[p for _, p in missing]],
                           capture_output=True, text=True)
        LOG.info((r.stdout or "")[-1500:])
        if r.returncode != 0:
            warn(f"pip завершился с ошибкой:\n{(r.stderr or '')[-1500:]}")
        import importlib
        importlib.invalidate_caches()
        for m, _ in missing:
            mod = _try_import(m)
            if m == "ee":
                ee = mod
            elif m == "rasterio":
                rasterio = mod
    still = [p for m, p in GEE_PACKAGES if OPTIONAL.get(m) is None]
    if still:
        warn(f"Не установлены: {', '.join(still)} — модули ДЗЗ отключены. Установите в ЭТОМ окружении:\n  {cmd}\n"
             "(в JupyterLab: %pip install " + " ".join(still) + ", затем перезапустите ядро)")
        return False
    return True


def gee_connect(settings: dict, interactive: bool = True) -> bool:
    """Инициализация earthengine-api. Возвращает True при успешном тестовом запросе."""
    if not settings.get("enabled"):
        LOG.info("GEE отключён пользователем.")
        return False
    if not ensure_gee_packages(interactive):
        return False
    project = settings.get("project") or None
    mode = settings.get("auth_mode", "localhost")
    for attempt in range(3):
        try:
            if mode == "service_account":
                key = settings.get("service_account_key")
                if not key or not os.path.exists(key):
                    raise RuntimeError("Не указан JSON-ключ сервисного аккаунта")
                email = json.loads(Path(key).read_text(encoding="utf-8"))["client_email"]
                creds = ee.ServiceAccountCredentials(email, key)
                ee.Initialize(creds, project=project)
            else:
                if settings.get("force_reauth") and attempt == 0:
                    ee.Authenticate(auth_mode=mode, force=True)
                try:
                    ee.Initialize(project=project)
                except Exception:  # noqa: BLE001
                    LOG.info(f"Авторизация Earth Engine (режим {mode}) - следуйте инструкциям в браузере/консоли...")
                    try:
                        ee.Authenticate(auth_mode=mode)
                    except TypeError:
                        ee.Authenticate()
                    ee.Initialize(project=project)
            v = ee.Number(1).add(1).getInfo()
            if v == 2:
                LOG.info(f"GEE подключён (project={project}, auth_mode={mode}).")
                return True
        except Exception as e:  # noqa: BLE001
            LOG.error(f"GEE: ошибка подключения: {e}")
            if not interactive:
                break
            if not gui_message(f"{GEE_HELP}\n\nОшибка: {e}\n\nПовторить попытку подключения?", ask=True):
                break
            settings = gui_gee_settings({"gee": settings, "privacy": {"public_mode": settings.get("public_mode", False)},
                                         "project": {"overwrite": settings.get("overwrite", False)}})
            project = settings.get("project") or None
            mode = settings.get("auth_mode", "localhost")
            if not settings.get("enabled"):
                break
    warn("GEE недоступен: модули рельефа, местообитаний, NDVI и снега пропущены. " + GEE_HELP.splitlines()[0])
    return False


# ==============================================================================================
# ВВОД ДАННЫХ
# ==============================================================================================
STD_FIELDS = ["animal_id", "sex", "timestamp_utc", "timestamp_local", "longitude", "latitude", "source_crs",
              "gps_error_m", "hdop", "pdop", "satellites", "fix_type", "release_date", "age_class", "collar_id",
              "speed_collar"]


def _norm(s: str) -> str:
    return re.sub(r"[^a-zа-я0-9+]", "", str(s).lower())


def _find_col(cols: Sequence[str], candidates: Sequence[str]) -> Optional[str]:
    nm = {_norm(c): c for c in cols}
    for cand in candidates:
        if _norm(cand) in nm:
            return nm[_norm(cand)]
    return None


def _parse_time_values(v: pd.Series) -> pd.Series:
    """Время из Excel: datetime.time, строка ЧЧ:ММ:СС, доля суток (float) или datetime."""
    def one(x):
        if pd.isna(x):
            return pd.NaT
        if isinstance(x, dt.time):
            return pd.Timedelta(hours=x.hour, minutes=x.minute, seconds=x.second, microseconds=x.microsecond)
        if isinstance(x, (dt.datetime, pd.Timestamp)):
            return pd.Timedelta(hours=x.hour, minutes=x.minute, seconds=x.second)
        if isinstance(x, (float, int, np.floating, np.integer)):
            return pd.Timedelta(days=float(x) % 1.0)
        try:
            return pd.to_timedelta(str(x).strip())
        except Exception:  # noqa: BLE001
            return pd.NaT
    return v.map(one)


def read_gpx(path: str) -> pd.DataFrame:
    tree = ET.parse(path)
    rows = []
    for el in tree.iter():
        tag = el.tag.split("}")[-1]
        if tag in ("trkpt", "wpt", "rtept"):
            t = None
            for ch in el:
                if ch.tag.split("}")[-1] == "time":
                    t = ch.text
            rows.append({"latitude": float(el.attrib["lat"]), "longitude": float(el.attrib["lon"]), "timestamp": t})
    df = pd.DataFrame(rows)
    df["animal_id"] = Path(path).stem
    return df


def read_one_file(path: str, cfg: dict) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        sheets = pd.read_excel(path, sheet_name=None)
        frames = []
        for sh, d in sheets.items():
            if d.shape[0] == 0:
                continue
            d = d.copy()
            d["source_sheet"] = sh
            frames.append(d)
        df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    elif ext in (".csv", ".txt"):
        df = pd.read_csv(path, sep=None, engine="python")
    elif ext == ".gpx":
        df = read_gpx(path)
    elif ext in (".gpkg", ".shp", ".geojson"):
        try:
            import geopandas as gpd
        except Exception as e:  # noqa: BLE001
            raise RuntimeError("Для GPKG/SHP нужен geopandas: pip install geopandas") from e
        g = gpd.read_file(path).to_crs(4326)
        df = pd.DataFrame(g.drop(columns="geometry"))
        df["longitude"] = g.geometry.x
        df["latitude"] = g.geometry.y
    else:
        raise RuntimeError(f"Неподдерживаемый формат: {ext}")
    df["source_file"] = Path(path).name
    df["source_row"] = np.arange(len(df)) + 2
    return df


def standardize(df: pd.DataFrame, cfg: dict, path: str) -> pd.DataFrame:
    cm = cfg["input"]["column_map"]
    cols = list(df.columns)
    out = pd.DataFrame(index=df.index)
    c_id = _find_col(cols, cm["animal_id"])
    out["animal_id"] = df[c_id].astype(str).str.strip() if c_id else Path(path).stem
    c_lat, c_lon = _find_col(cols, cm["latitude"]), _find_col(cols, cm["longitude"])
    if not c_lat or not c_lon:
        raise RuntimeError(f"{Path(path).name}: не найдены колонки координат (ожидались {cm['latitude']} / {cm['longitude']})")
    out["latitude"] = pd.to_numeric(df[c_lat], errors="coerce")
    out["longitude"] = pd.to_numeric(df[c_lon], errors="coerce")
    tz_offset_h = None
    c_ts = _find_col(cols, cm["timestamp"])
    c_d, c_t = _find_col(cols, cm["date"]), _find_col(cols, cm["time"])
    if c_ts is not None and c_ts not in (c_d, c_t):
        ts = pd.to_datetime(df[c_ts], errors="coerce", utc=False)
        time_col_name = c_ts
    elif c_d is not None:
        d = pd.to_datetime(df[c_d], errors="coerce", dayfirst=True)
        if c_t is not None:
            ts = d.dt.normalize() + _parse_time_values(df[c_t])
        else:
            ts = d
        time_col_name = c_t or c_d
    else:
        raise RuntimeError(f"{Path(path).name}: не найдены колонки даты/времени")
    m = re.search(r"UTC\s*([+-−]\s*\d+(?:[.,]\d+)?)", str(time_col_name), re.I)
    if m:
        tz_offset_h = float(m.group(1).replace(" ", "").replace("−", "-").replace(",", "."))
    if getattr(ts.dt, "tz", None) is not None:
        ts_utc = ts.dt.tz_convert("UTC")
        tz_note = "aware"
    elif tz_offset_h is not None:
        ts_utc = (ts - pd.to_timedelta(tz_offset_h, unit="h")).dt.tz_localize("UTC")
        tz_note = f"UTC{tz_offset_h:+g} (из имени колонки «{time_col_name}»)"
    else:
        src_tz = cfg["input"]["source_timezone"]
        ts_utc = ts.dt.tz_localize(src_tz, ambiguous="NaT", nonexistent="NaT").dt.tz_convert("UTC")
        tz_note = f"{src_tz} (из конфигурации)"
    out["timestamp_utc"] = ts_utc
    for f in ("sex", "gps_error_m", "hdop", "pdop", "satellites", "fix_type", "collar_id", "age_class",
              "release_date", "speed_collar"):
        c = _find_col(cols, cm.get(f, []))
        if c is not None and c not in (c_id, c_lat, c_lon, c_d, c_t, c_ts):
            out[f] = df[c]
        else:
            out[f] = np.nan
    out["source_crs"] = cfg["input"]["source_crs"]
    out["source_file"] = df["source_file"]
    out["source_sheet"] = df.get("source_sheet", "")
    out["source_row"] = df["source_row"]
    out["source_time_zone_rule"] = tz_note
    return out


def load_inputs(files: List[str], cfg: dict) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Читает все файлы; возвращает (стандартизованная таблица, неизменная копия исходных записей, сведения о файлах)."""
    frames, raws, info = [], [], []
    skipped = []
    for p in files:
        LOG.info(f"Чтение {p}")
        try:
            raw = read_one_file(p, cfg)
            std = standardize(raw, cfg, p)
        except Exception as e:  # noqa: BLE001
            warn(f"Файл пропущен (не похож на GPS-трек): {Path(p).name} - {e}")
            skipped.append(Path(p).name)
            continue
        frames.append(std)
        r = raw.copy()
        r.columns = [str(c) for c in r.columns]
        raws.append(r)
        info.append({"file": Path(p).name, "path": str(Path(p).resolve()), "sha256": sha256_file(p),
                     "n_records": len(raw), "columns": "; ".join(map(str, raw.columns)),
                     "time_rule": std["source_time_zone_rule"].iloc[0] if len(std) else "",
                     "animals": "; ".join(sorted(std["animal_id"].unique()))})
    if not frames:
        raise RuntimeError("Ни один из выбранных файлов не распознан как GPS-трек: " + ", ".join(skipped))
    pts = pd.concat(frames, ignore_index=True)
    pts.insert(0, "record_id", np.arange(len(pts)))
    raw_all = pd.concat(raws, ignore_index=True, sort=False)
    return pts, raw_all, pd.DataFrame(info)


def default_metadata(pts: pd.DataFrame, cfg: dict, existing: Optional[pd.DataFrame]) -> pd.DataFrame:
    rows = []
    for a, g in pts.dropna(subset=["timestamp_utc", "latitude", "longitude"]).sort_values("timestamp_utc").groupby("animal_id"):
        first = g.iloc[0]
        lon_med = g["longitude"].median()
        reg = cfg["regions"]["west_name"] if lon_med < cfg["regions"]["split_lon"] else cfg["regions"]["central_name"]
        sex = g["sex"].dropna().astype(str).iloc[0] if g["sex"].notna().any() else "unknown"
        rows.append({"animal_id": a, "sex": sex,
                     "release_datetime_utc": first["timestamp_utc"].strftime("%Y-%m-%d %H:%M"),
                     "release_source": "first_fix_assumed", "region": reg, "status": "released",
                     "release_lat": round(first["latitude"], 6), "release_lon": round(first["longitude"], 6),
                     "collar_id": "", "age_class": "unknown"})
    meta = pd.DataFrame(rows).astype(object)
    if existing is not None and len(existing):
        ex = existing.astype({"animal_id": str}).drop_duplicates("animal_id").set_index("animal_id")
        for i, r in meta.iterrows():
            if r["animal_id"] in ex.index:
                for c in ex.columns:
                    if c in meta.columns and pd.notna(ex.at[r["animal_id"], c]) and str(ex.at[r["animal_id"], c]) != "":
                        meta.at[i, c] = ex.at[r["animal_id"], c]
    for c in ("release_lat", "release_lon"):
        meta[c] = pd.to_numeric(meta[c], errors="coerce")
    for c in ("animal_id", "sex", "release_datetime_utc", "release_source", "region", "status", "collar_id", "age_class"):
        meta[c] = meta[c].fillna("").astype(str)
    return meta


# ==============================================================================================
# КООРДИНАТЫ, СОЛНЦЕ, СЕЗОНЫ
# ==============================================================================================
_GEOD = pyproj.Geod(ellps="WGS84") if pyproj is not None else None
_TRANSFORMERS: Dict[str, Any] = {}


def to_work_xy(lon, lat, crs: str):
    if crs not in _TRANSFORMERS:
        _TRANSFORMERS[crs] = pyproj.Transformer.from_crs("EPSG:4326", crs, always_xy=True)
    return _TRANSFORMERS[crs].transform(np.asarray(lon, float), np.asarray(lat, float))


def from_work_xy(x, y, crs: str):
    key = crs + "_inv"
    if key not in _TRANSFORMERS:
        _TRANSFORMERS[key] = pyproj.Transformer.from_crs(crs, "EPSG:4326", always_xy=True)
    return _TRANSFORMERS[key].transform(np.asarray(x, float), np.asarray(y, float))


def solar_position(ts_utc: pd.Series, lat, lon) -> pd.DataFrame:
    """Алгоритм NOAA (точность ~1 мин): высота Солнца, часовой угол, восход/закат и гражданские сумерки (UTC)."""
    t = pd.to_datetime(ts_utc).dt.tz_convert("UTC").dt.tz_localize(None)
    lat = np.asarray(lat, float)
    lon = np.asarray(lon, float)
    jd = (t.values.astype("datetime64[ns]").astype("int64") / 86400e9) + 2440587.5
    jc = (jd - 2451545.0) / 36525.0
    L = np.mod(280.46646 + jc * (36000.76983 + jc * 0.0003032), 360)
    M = 357.52911 + jc * (35999.05029 - 0.0001537 * jc)
    e = 0.016708634 - jc * (0.000042037 + 0.0000001267 * jc)
    r = np.radians
    C = (np.sin(r(M)) * (1.914602 - jc * (0.004817 + 0.000014 * jc)) + np.sin(r(2 * M)) * (0.019993 - 0.000101 * jc)
         + np.sin(r(3 * M)) * 0.000289)
    app = L + C - 0.00569 - 0.00478 * np.sin(r(125.04 - 1934.136 * jc))
    eps0 = 23 + (26 + (21.448 - jc * (46.815 + jc * (0.00059 - jc * 0.001813))) / 60) / 60
    eps = eps0 + 0.00256 * np.cos(r(125.04 - 1934.136 * jc))
    decl = np.degrees(np.arcsin(np.sin(r(eps)) * np.sin(r(app))))
    y = np.tan(r(eps / 2)) ** 2
    eot = 4 * np.degrees(y * np.sin(2 * r(L)) - 2 * e * np.sin(r(M)) + 4 * e * y * np.sin(r(M)) * np.cos(2 * r(L))
                         - 0.5 * y * y * np.sin(4 * r(L)) - 1.25 * e * e * np.sin(2 * r(M)))
    minutes = (t.dt.hour * 60 + t.dt.minute + t.dt.second / 60).values
    tst = np.mod(minutes + eot + 4 * lon, 1440)
    ha = tst / 4 - 180
    cosz = np.sin(r(lat)) * np.sin(r(decl)) + np.cos(r(lat)) * np.cos(r(decl)) * np.cos(r(ha))
    elev = 90 - np.degrees(np.arccos(np.clip(cosz, -1, 1)))
    noon = 720 - 4 * lon - eot

    def ha_for(z):
        c = np.cos(r(z)) / (np.cos(r(lat)) * np.cos(r(decl))) - np.tan(r(lat)) * np.tan(r(decl))
        return np.degrees(np.arccos(np.clip(c, -1, 1)))

    day0 = t.dt.normalize()
    out = pd.DataFrame({"sun_elevation_deg": elev, "hour_angle_deg": ha, "solar_hour": tst / 60.0}, index=ts_utc.index)
    for name, z in (("sunrise", 90.833), ("civil_dawn", 96.0)):
        h = ha_for(z)
        out[name + "_utc"] = (day0 + pd.to_timedelta(noon - 4 * h, unit="m")).dt.tz_localize("UTC").values
        nm2 = "sunset" if name == "sunrise" else "civil_dusk"
        out[nm2 + "_utc"] = (day0 + pd.to_timedelta(noon + 4 * h, unit="m")).dt.tz_localize("UTC").values
    out["solar_noon_utc"] = (day0 + pd.to_timedelta(noon, unit="m")).dt.tz_localize("UTC").values
    return out


def solar_period_B(elev, ha, cfg) -> np.ndarray:
    s = cfg["solar"]
    res = np.where(elev > s["sunrise_deg"], "day", np.where(elev <= s["civil_twilight_deg"], "night",
                                                            np.where(ha < 0, "dawn_civil_twilight", "dusk_civil_twilight")))
    return res.astype(object)


def _hm(s: str) -> int:
    h, m = s.split(":")
    return int(h) * 60 + int(m)


def solar_period_A(ts_local: pd.Series, cfg) -> np.ndarray:
    mins = (ts_local.dt.hour * 60 + ts_local.dt.minute).values
    res = np.full(len(mins), "unassigned", dtype=object)
    for p in cfg["solar"]["periods_A"]:
        a, b = _hm(p["start"]), _hm(p["end"])
        sel = ((mins >= a) & (mins < b)) if a < b else ((mins >= a) | (mins < b))
        res[sel & (res == "unassigned")] = p["name"]
    return res


def add_time_fields(df: pd.DataFrame, tcol: str, latcol: str, loncol: str, cfg: dict, prefix: str = "") -> pd.DataFrame:
    tz = cfg["time"]["local_timezone"]
    t = df[tcol]
    loc = t.dt.tz_convert(tz)
    sp = solar_position(t, df[latcol].values, df[loncol].values)
    df[prefix + "solar_period_A"] = solar_period_A(loc, cfg)
    df[prefix + "solar_period_B"] = solar_period_B(sp["sun_elevation_deg"].values, sp["hour_angle_deg"].values, cfg)
    if prefix == "":
        df["timestamp_local"] = loc
        df["local_date"] = loc.dt.date
        df["local_hour"] = loc.dt.hour + loc.dt.minute / 60.0
        df["solar_hour"] = sp["solar_hour"].values
        df["sun_elevation_deg"] = sp["sun_elevation_deg"].values
        df["sin_hour"] = np.sin(2 * np.pi * df["local_hour"] / 24)
        df["cos_hour"] = np.cos(2 * np.pi * df["local_hour"] / 24)
        for c in ("sunrise_utc", "sunset_utc", "civil_dawn_utc", "civil_dusk_utc"):
            df[c.replace("_utc", "_local")] = pd.to_datetime(sp[c]).dt.tz_localize("UTC").dt.tz_convert(tz).dt.strftime("%H:%M").values
    return df


def load_season_table(cfg: dict) -> pd.DataFrame:
    p = cfg["seasons"].get("table")
    if p and os.path.exists(p):
        t = pd.read_csv(p)
        t.columns = [c.strip().lower() for c in t.columns]
        for c in ("region", "year"):
            if c not in t.columns:
                t[c] = "all"
        t["source"] = f"table:{Path(p).name}"
        return t
    rows = [dict(r, region="all", year="all", source="default_calendar_needs_confirmation") for r in cfg["seasons"]["default_calendar"]]
    return pd.DataFrame(rows)


def assign_calendar_season(dates: pd.Series, regions: pd.Series, table: pd.DataFrame) -> Tuple[np.ndarray, np.ndarray]:
    d = pd.to_datetime(pd.Series(dates).astype(str))
    md = d.dt.month * 100 + d.dt.day
    yr = d.dt.year.astype(str)
    res = np.full(len(d), "unassigned", dtype=object)
    src = np.full(len(d), "", dtype=object)
    reg = pd.Series(regions).astype(str).values
    for _, r in table.iterrows():
        rm = (r["region"] == "all") | (reg == str(r["region"]))
        ym = (str(r["year"]) == "all") | (yr.values == str(r["year"]))
        s, e = str(r["start"]), str(r["end"])
        if len(s) > 5:  # полные даты
            sel = (d >= pd.Timestamp(s)).values & (d <= pd.Timestamp(e)).values
        else:
            a = int(s[:2]) * 100 + int(s[3:5])
            b = int(e[:2]) * 100 + int(e[3:5])
            sel = ((md >= a) & (md <= b)).values if a <= b else ((md >= a) | (md <= b)).values
        m = sel & rm & ym & (res == "unassigned")
        res[m] = r["season"]
        src[m] = r["source"]
    return res, src


# ==============================================================================================
# КОНТРОЛЬ КАЧЕСТВА
# ==============================================================================================
def _add_reason(reasons: np.ndarray, mask, text: str) -> None:
    idx = np.where(np.asarray(mask))[0]
    for i in idx:
        reasons[i] = text if reasons[i] == "" else reasons[i] + ";" + text


def run_qc(pts: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    q = cfg["qc"]
    df = pts.sort_values(["animal_id", "timestamp_utc", "record_id"]).reset_index(drop=True)
    n = len(df)
    reasons = np.full(n, "", dtype=object)
    _add_reason(reasons, df[["latitude", "longitude"]].isna().any(axis=1) | df["timestamp_utc"].isna(), "missing_coords_or_time")
    bad = ~df["latitude"].between(-90, 90) | ~df["longitude"].between(-180, 180) | ((df["latitude"] == 0) & (df["longitude"] == 0))
    _add_reason(reasons, bad & df["latitude"].notna(), "invalid_coordinates")
    b = cfg["aoi"]["bbox"]
    out_aoi = ~(df["longitude"].between(b[0], b[2]) & df["latitude"].between(b[1], b[3]))
    _add_reason(reasons, out_aoi & ~bad & df["latitude"].notna(), "outside_aoi_bbox")
    if cfg["aoi"].get("polygon_file") and OPTIONAL.get("shapely"):
        try:
            import geopandas as gpd
            from shapely.prepared import prep
            from shapely.geometry import Point
            poly = prep(gpd.read_file(cfg["aoi"]["polygon_file"]).to_crs(4326).unary_union)
            inside = np.array([poly.contains(Point(x, y)) if np.isfinite(x) and np.isfinite(y) else False
                               for x, y in zip(df["longitude"], df["latitude"])])
            _add_reason(reasons, ~inside & ~out_aoi, "outside_aoi_polygon")
        except Exception as e:  # noqa: BLE001
            warn(f"AOI-полигон не применён: {e}")
    if q.get("max_hdop") is not None and df["hdop"].notna().any():
        _add_reason(reasons, pd.to_numeric(df["hdop"], errors="coerce") > q["max_hdop"], "hdop_above_max")
    if q.get("min_satellites") is not None and df["satellites"].notna().any():
        _add_reason(reasons, pd.to_numeric(df["satellites"], errors="coerce") < q["min_satellites"], "satellites_below_min")
    # дубликаты
    dup_exact = df.duplicated(["animal_id", "timestamp_utc", "latitude", "longitude"], keep="first")
    _add_reason(reasons, dup_exact, "duplicate_exact")
    nd = df[~dup_exact]
    grp = nd.groupby(["animal_id", "timestamp_utc"])["record_id"].transform("count")
    conflict_idx = nd.index[grp > 1]
    if q["duplicate_policy"] == "exclude_conflicting":
        _add_reason(reasons, df.index.isin(conflict_idx), "conflicting_timestamp")
    else:
        keep_first = nd.loc[conflict_idx].duplicated(["animal_id", "timestamp_utc"], keep="first")
        _add_reason(reasons, df.index.isin(keep_first[keep_first].index), "conflicting_timestamp_not_first")
    # слишком короткий интервал
    valid = reasons == ""
    sub = df[valid]
    dtp = sub.groupby("animal_id")["timestamp_utc"].diff().dt.total_seconds()
    _add_reason(reasons, df.index.isin(sub.index[(dtp < q["min_interval_s"]).values]), "interval_below_min")
    # скорость и «выброс-возврат» (итеративно)
    for it in range(int(q["qc_iterations"])):
        valid = reasons == ""
        sub = df[valid]
        if len(sub) < 3:
            break
        lon, lat = sub["longitude"].values, sub["latitude"].values
        a = sub["animal_id"].values
        same_prev = np.r_[False, a[1:] == a[:-1]]
        same_next = np.r_[a[1:] == a[:-1], False]
        az, _, dist = _GEOD.inv(lon[:-1], lat[:-1], lon[1:], lat[1:])
        dist = np.r_[np.nan, dist]
        az = np.r_[np.nan, az]
        dts = np.r_[np.nan, np.diff(sub["timestamp_utc"].values).astype("timedelta64[s]").astype(float)]
        dist[~same_prev] = np.nan
        dts[~same_prev] = np.nan
        sp_in = dist / np.where(dts > 0, dts, np.nan) * 3.6
        sp_out = np.r_[sp_in[1:], np.nan]
        dt_in = dts / 3600
        dt_out = np.r_[dts[1:], np.nan] / 3600
        turn = np.abs(wrap180(np.r_[az[1:], np.nan] - az))
        mx = q["max_speed_km_h"]
        mind = q["max_speed_min_dt_h"]
        hi_in = (sp_in > mx) & (dt_in >= mind)
        hi_out = (sp_out > mx) & (dt_out >= mind)
        sp_flag = (hi_in & (hi_out | ~same_next)) | (hi_out & ~same_prev)
        spike = (sp_in > q["spike_speed_km_h"]) & (sp_out > q["spike_speed_km_h"]) & (turn > q["spike_angle_deg"]) & same_prev & same_next
        new = sp_flag | spike
        if not new.any():
            break
        _add_reason(reasons, df.index.isin(sub.index[sp_flag]), "speed_exceeds_max")
        _add_reason(reasons, df.index.isin(sub.index[spike & ~sp_flag]), "spike_return")
    df["qc_reason"] = reasons
    df["included_analysis"] = reasons == ""
    df["qc_flag"] = np.where(df["included_analysis"], "ok", "excluded")
    return df


def assign_bursts(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    gap = cfg["qc"]["burst_max_gap_h"] * 3600
    df["burst_id"] = None
    df["dt_prev_s"] = np.nan
    inc = df.index[df["included_analysis"]]
    sub = df.loc[inc]
    for a, g in sub.groupby("animal_id"):
        dts = g["timestamp_utc"].diff().dt.total_seconds()
        collar = g["collar_id"].fillna("na").astype(str)
        new = (dts > gap) | dts.isna() | (collar != collar.shift().fillna(collar.iloc[0]))
        bid = new.cumsum()
        df.loc[g.index, "burst_id"] = [f"{a}_b{int(k):03d}" for k in bid]
        d2 = dts.copy()
        d2[new] = np.nan
        df.loc[g.index, "dt_prev_s"] = d2.values
    return df


def nominal_intervals(df: pd.DataFrame) -> Dict[str, float]:
    res = {}
    for a, g in df[df["included_analysis"]].groupby("animal_id"):
        d = g["dt_prev_s"].dropna() / 3600
        d = d[(d > 0) & (d < 6)]
        if len(d) == 0:
            res[a] = np.nan
            continue
        r = (d * 2).round() / 2
        res[a] = float(r.mode().iloc[0]) if len(r.mode()) else float(d.median())
        if res[a] == 0:
            res[a] = float(d.median())
    return res


def make_steps(p: pd.DataFrame, cfg: dict, burst_col: str = "burst_id", tcol: str = "timestamp_utc") -> pd.DataFrame:
    """Шаги между последовательными валидными фиксациями одного burst. Длина — геодезическая (WGS84)."""
    p = p.sort_values(["animal_id", tcol]).reset_index(drop=True)
    b = p[burst_col].values
    same = np.r_[b[1:] == b[:-1]]
    i0 = np.where(same)[0]
    i1 = i0 + 1
    s = pd.DataFrame({
        "animal_id": p["animal_id"].values[i0],
        "burst_id": b[i0],
        "rec_start": p["record_id"].values[i0], "rec_end": p["record_id"].values[i1],
        "t_start": p[tcol].values[i0], "t_end": p[tcol].values[i1],
        "lon_start": p["longitude"].values[i0], "lat_start": p["latitude"].values[i0],
        "lon_end": p["longitude"].values[i1], "lat_end": p["latitude"].values[i1],
        "x_start": p["x"].values[i0], "y_start": p["y"].values[i0],
        "x_end": p["x"].values[i1], "y_end": p["y"].values[i1],
    })
    for c in ("sex", "region"):
        if c in p.columns:
            s[c] = p[c].values[i0]
    if "speed_collar" in p.columns:
        s["speed_collar"] = (pd.to_numeric(p["speed_collar"], errors="coerce").values[i0]
                             + pd.to_numeric(p["speed_collar"], errors="coerce").values[i1]) / 2
    for c in ("t_start", "t_end"):
        s[c] = pd.to_datetime(s[c], utc=True)
    s["delta_t_s"] = (s["t_end"] - s["t_start"]).dt.total_seconds()
    s = s[s["delta_t_s"] > 0].reset_index(drop=True)
    s["t_mid"] = s["t_start"] + pd.to_timedelta(s["delta_t_s"] / 2, unit="s")
    az, _, dist = _GEOD.inv(s["lon_start"].values, s["lat_start"].values, s["lon_end"].values, s["lat_end"].values)
    s["step_length_m"] = dist
    s["bearing_deg"] = np.mod(az, 360)
    s["speed_m_s"] = s["step_length_m"] / s["delta_t_s"]
    s["speed_km_h"] = s["speed_m_s"] * 3.6
    s["speed_m_h"] = s["speed_m_s"] * 3600
    mn = cfg["steps"]["min_step_for_angle_m"]
    prev_same = np.r_[False, s["burst_id"].values[1:] == s["burst_id"].values[:-1]]
    prev_same &= np.r_[False, s["rec_end"].values[:-1] == s["rec_start"].values[1:]]
    prev_b = np.r_[np.nan, s["bearing_deg"].values[:-1]]
    prev_l = np.r_[np.nan, s["step_length_m"].values[:-1]]
    ta = wrap180(s["bearing_deg"].values - prev_b)
    ok = prev_same & (prev_l >= mn) & (s["step_length_m"].values >= mn)
    s["turn_angle_deg"] = np.where(ok, ta, np.nan)
    s["mid_lon"] = (s["lon_start"] + s["lon_end"]) / 2
    s["mid_lat"] = (s["lat_start"] + s["lat_end"]) / 2
    s.insert(0, "step_id", [f"{a}_s{i:06d}" for a, i in zip(s["animal_id"], range(len(s)))])
    return s


def regularize(p: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Выбор фиксаций, ближайших к узлам регулярной сетки (без интерполяции); пропуск узла рвёт reg-burst."""
    r = cfg["regularization"]
    step = pd.Timedelta(hours=r["interval_h"])
    tol = pd.Timedelta(minutes=r["tolerance_min"])
    out = []
    for bid, g in p[p["included_analysis"]].groupby("burst_id"):
        t = g["timestamp_utc"]
        start = t.iloc[0].ceil(step) if t.iloc[0] != t.iloc[0].floor(step) else t.iloc[0]
        start = start - step
        grid = pd.date_range(start, t.iloc[-1] + step, freq=step)
        tv = t.values.astype("datetime64[ns]").astype("int64")
        gv = grid.values.astype("datetime64[ns]").astype("int64") if grid.tz is None else grid.tz_convert("UTC").tz_localize(None).values.astype("datetime64[ns]").astype("int64")
        pos = np.searchsorted(tv, gv)
        pos0 = np.clip(pos - 1, 0, len(tv) - 1)
        pos1 = np.clip(pos, 0, len(tv) - 1)
        d0 = np.abs(tv[pos0] - gv)
        d1 = np.abs(tv[pos1] - gv)
        best = np.where(d0 <= d1, pos0, pos1)
        dbest = np.minimum(d0, d1)
        ok = dbest <= tol.value
        sel_idx = best[ok]
        grid_ok = grid[ok]
        if len(sel_idx) == 0:
            continue
        uniq, first = np.unique(sel_idx, return_index=True)
        sel = g.iloc[sel_idx[first]].copy()
        sel["reg_time"] = grid_ok[first]
        k = np.round((sel["reg_time"] - grid_ok[first][0]) / step).astype(int).values
        brk = np.r_[True, np.diff(k) != 1]
        sel["reg_burst_id"] = [f"{bid}_r{int(c):03d}" for c in np.cumsum(brk)]
        out.append(sel)
    if not out:
        return pd.DataFrame()
    return pd.concat(out).sort_values(["animal_id", "timestamp_utc"]).reset_index(drop=True)


# ==============================================================================================
# ДАННЫЕ ДЗЗ: СТАТИЧНЫЕ ТАЙЛЫ (DEM, WorldCover) И ДИНАМИЧЕСКИЕ ЗНАЧЕНИЯ В ТОЧКАХ (S2, Landsat, DW, MODIS, ERA5)
# ==============================================================================================
PROVENANCE: List[dict] = []


def add_provenance(**kw) -> None:
    kw.setdefault("access_date_utc", dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M"))
    PROVENANCE.append(kw)


def with_retries(fn: Callable, retries: int = 5, what: str = "запрос GEE"):
    delay = 2.0
    for k in range(retries):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            msg = str(e)
            if any(s in msg for s in ("Too many concurrent", "Quota", "429", "rate", "deadline", "timed out",
                                      "Internal error", "503", "502", "Connection", "temporarily")) and k < retries - 1:
                LOG.info(f"{what}: {msg[:120]} — повтор через {delay:.0f} с")
                time.sleep(delay)
                delay = min(delay * 2, 120)
                continue
            raise


def tile_grid(x: np.ndarray, y: np.ndarray, cfg: dict) -> List[Tuple[int, int]]:
    g = cfg["gee"]
    T = g["tile_size_m"]
    buf = cfg["aoi"]["buffer_m"]
    keys = set()
    for dx in (-buf, 0, buf):
        for dy in (-buf, 0, buf):
            ii = np.floor((x + dx) / T).astype(int)
            jj = np.floor((y + dy) / T).astype(int)
            keys.update(zip(ii.tolist(), jj.tolist()))
    return sorted(keys)


def _ee_dem(cfg):
    s = cfg["gee"]["sources"]
    cop = ee.ImageCollection(s["dem"]).select(s["dem_band"]).map(lambda i: i.resample("bilinear").toFloat()).mosaic()
    nas = ee.Image(s["dem_fallback"]).select("elevation").resample("bilinear").toFloat()
    return cop.unmask(nas).rename("elevation").toFloat()


def _download_tile(img, crs: str, scale: float, x0: float, y0: float, w: int, h: int, out: Path, retries: int) -> None:
    req = {"expression": img, "fileFormat": "GEO_TIFF",
           "grid": {"dimensions": {"width": int(w), "height": int(h)},
                    "affineTransform": {"scaleX": scale, "shearX": 0, "translateX": x0,
                                        "shearY": 0, "scaleY": -scale, "translateY": y0},
                    "crsCode": crs}}
    data = with_retries(lambda: ee.data.computePixels(req), retries, f"тайл {out.name}")
    tmp = out.with_suffix(".part")
    with open(tmp, "wb") as fh:
        fh.write(data)
    os.replace(tmp, out)


def fetch_static_tiles(x: np.ndarray, y: np.ndarray, cfg: dict, cache_dir: Path) -> Dict[str, Any]:
    """Скачивает DEM (30 м) и WorldCover (10 м) тайлами в рабочей CRS с перекрытием; уже скачанные тайлы не повторяются."""
    g = cfg["gee"]
    crs = cfg["crs"]["work_crs"]
    T, O = g["tile_size_m"], g["tile_overlap_m"]
    keys = tile_grid(x, y, cfg)
    tdir = cache_dir / "tiles"
    tdir.mkdir(parents=True, exist_ok=True)
    check_disk(tdir)
    dem = _ee_dem(cfg)
    wc = ee.ImageCollection(g["sources"]["worldcover"]).first().select("Map").rename("worldcover").toUint8()
    tag_dem = hash_obj([g["sources"]["dem"], g["sources"]["dem_fallback"], crs, g["dem_scale_m"], T, O])[:8]
    tag_wc = hash_obj([g["sources"]["worldcover"], crs, g["wc_scale_m"], T, O])[:8]
    jobs = []
    for (i, j) in keys:
        x0, y1 = i * T - O, (j + 1) * T + O
        for kind, img, sc, tag in (("dem", dem, g["dem_scale_m"], tag_dem), ("wc", wc, g["wc_scale_m"], tag_wc)):
            out = tdir / f"{kind}_{tag}_{i}_{j}.tif"
            if not out.exists():
                n = int(round((T + 2 * O) / sc))
                jobs.append((img, sc, x0, y1, n, out))
    LOG.info(f"Статичные тайлы: {len(keys)} ячеек, к загрузке {len(jobs)} файлов (остальное — из кэша)")
    done = 0
    with ThreadPoolExecutor(max_workers=int(g["threads"])) as ex:
        futs = {ex.submit(_download_tile, im, crs, sc, x0, y1, n, n, out, int(g["max_retries"])): out
                for im, sc, x0, y1, n, out in jobs}
        for f in as_completed(futs):
            done += 1
            try:
                f.result()
            except Exception as e:  # noqa: BLE001
                warn(f"Не загружен тайл {futs[f].name}: {e}")
            if done % 20 == 0:
                LOG.info(f"  загружено {done}/{len(jobs)}")
    add_provenance(component="elevation", asset_id=g["sources"]["dem"] + f" (band {g['sources']['dem_band']}), пропуски — "
                   + g["sources"]["dem_fallback"], scale_m=g["dem_scale_m"], crs=crs,
                   temporal="DSM; NASADEM — Feb 2000", processing="mosaic, bilinear resample; производные рассчитаны локально (Horn 1981, Riley 1999, Weiss 2001)")
    add_provenance(component="land_cover_static", asset_id=g["sources"]["worldcover"], scale_m=g["wc_scale_m"], crs=crs,
                   temporal="2021 (v200)", processing="Map band, nearest")
    return {"keys": keys, "dir": str(tdir), "tag_dem": tag_dem, "tag_wc": tag_wc, "T": T, "O": O, "crs": crs,
            "dem_scale": g["dem_scale_m"], "wc_scale": g["wc_scale_m"]}


# -------------------------- локальные производные рельефа -------------------------------------
def _disk_kernel(radius_px: float) -> np.ndarray:
    r = int(math.ceil(radius_px))
    yy, xx = np.mgrid[-r:r + 1, -r:r + 1]
    return (xx ** 2 + yy ** 2 <= radius_px ** 2).astype(np.float32)


def _focal_mean(a: np.ndarray, k: np.ndarray) -> np.ndarray:
    from scipy.signal import fftconvolve
    m = np.isfinite(a).astype(np.float32)
    v = np.where(np.isfinite(a), a, 0).astype(np.float64)
    s = fftconvolve(v, k, mode="same")
    c = fftconvolve(m, k, mode="same")
    with np.errstate(invalid="ignore", divide="ignore"):
        out = s / c
    out[c < 0.5 * k.sum()] = np.nan
    return out.astype(np.float32)


def terrain_derivatives(z: np.ndarray, cs: float, cfg: dict) -> Dict[str, np.ndarray]:
    """Производные DEM: уклон (°/%) и экспозиция (Horn 1981), northness/eastness, TRI (Riley 1999), roughness,
    TPI малого/крупного радиуса (z − средняя высота в круге), кривизны Zevenbergen–Thorne (1/100 м)."""
    zp = np.pad(z, 1, mode="edge")
    a, b, c = zp[:-2, :-2], zp[:-2, 1:-1], zp[:-2, 2:]
    d, e, f = zp[1:-1, :-2], zp[1:-1, 1:-1], zp[1:-1, 2:]
    g, h, i = zp[2:, :-2], zp[2:, 1:-1], zp[2:, 2:]
    dzdx = ((c + 2 * f + i) - (a + 2 * d + g)) / (8 * cs)
    dzdy = ((g + 2 * h + i) - (a + 2 * b + c)) / (8 * cs)  # к югу положительно
    slope = np.degrees(np.arctan(np.hypot(dzdx, dzdy)))
    aspect = np.mod(np.degrees(np.arctan2(-dzdx, dzdy)), 360)
    nb = np.stack([a, b, c, d, f, g, h, i])
    tri = np.sqrt(np.nansum((nb - e) ** 2, axis=0))
    rough = np.nanmax(np.stack([nb.max(0), e]), 0) - np.nanmin(np.stack([nb.min(0), e]), 0)
    L = cs
    D = ((d + f) / 2 - e) / L ** 2
    E = ((b + h) / 2 - e) / L ** 2
    F = (-a + c + g - i) / (4 * L ** 2)
    G = (-d + f) / (2 * L)
    H = (b - h) / (2 * L)
    gg = G ** 2 + H ** 2
    with np.errstate(invalid="ignore", divide="ignore"):
        prof = np.where(gg > 0, 2 * (D * G ** 2 + E * H ** 2 + F * G * H) / gg, 0) * 100
        plan = np.where(gg > 0, -2 * (D * H ** 2 + E * G ** 2 - F * G * H) / gg, 0) * 100
    curv = -2 * (D + E) * 100
    t = cfg["terrain"]
    tpi_s = z - _focal_mean(z, _disk_kernel(t["tpi_small_radius_m"] / cs))
    tpi_l = z - _focal_mean(z, _disk_kernel(t["tpi_large_radius_m"] / cs))
    return {"elevation": z, "slope_deg": slope, "slope_pct": np.tan(np.radians(slope)) * 100, "aspect_deg": aspect,
            "northness": np.cos(np.radians(aspect)), "eastness": np.sin(np.radians(aspect)),
            "tri": tri, "roughness": rough, "tpi_small": tpi_s, "tpi_large": tpi_l,
            "curvature": curv, "profile_curvature": prof, "plan_curvature": plan}


TERRAIN_BANDS = ["elevation", "slope_deg", "slope_pct", "aspect_deg", "northness", "eastness", "tri", "roughness",
                 "tpi_small", "tpi_large", "curvature", "profile_curvature", "plan_curvature"]
TERRAIN_UNITS = {"elevation": "м (DSM)", "slope_deg": "°, Horn 3×3", "slope_pct": "%", "aspect_deg": "° от севера по часовой",
                 "northness": "cos(aspect)", "eastness": "sin(aspect)", "tri": "м, Riley 3×3", "roughness": "м, max−min 3×3",
                 "tpi_small": "м, z − среднее в круге r=tpi_small_radius_m", "tpi_large": "м, z − среднее в круге r=tpi_large_radius_m",
                 "curvature": "1/100 м, Zevenbergen–Thorne", "profile_curvature": "1/100 м", "plan_curvature": "1/100 м"}


class TileStore:
    """Доступ к локальным тайлам (рельеф 30 м + WorldCover 10 м) с LRU-кэшем в памяти."""

    def __init__(self, meta: dict, cfg: dict, max_tiles: int = 24):
        self.m = meta
        self.cfg = cfg
        self.max = max_tiles
        self.cache: "OrderedDict[Tuple[str,int,int], Any]" = OrderedDict()
        self.dir = Path(meta["dir"])
        self.T, self.O = meta["T"], meta["O"]
        self.tpi_sd = None

    def _load(self, kind: str, i: int, j: int):
        key = (kind, i, j)
        if key in self.cache:
            self.cache.move_to_end(key)
            return self.cache[key]
        val = None
        if kind == "terrain":
            dp = self.dir / f"terrain_{self.m['tag_dem']}_{hash_obj(self.cfg['terrain'])[:6]}_{i}_{j}.npz"
            if dp.exists():
                with np.load(dp) as z:
                    val = {k: z[k] for k in z.files}
            else:
                src = self.dir / f"dem_{self.m['tag_dem']}_{i}_{j}.tif"
                if src.exists():
                    with rasterio.open(src) as r:
                        zz = r.read(1).astype(np.float32)
                        if r.nodata is not None:
                            zz[zz == r.nodata] = np.nan
                    zz[(zz < -500) | (zz > 9000)] = np.nan
                    val = terrain_derivatives(zz, self.m["dem_scale"], self.cfg)
                    val = {k: v.astype(np.float32) for k, v in val.items()}
                    np.savez_compressed(dp, **val)
        else:
            src = self.dir / f"wc_{self.m['tag_wc']}_{i}_{j}.tif"
            if src.exists():
                with rasterio.open(src) as r:
                    val = {"worldcover": r.read(1)}
        self.cache[key] = val
        if len(self.cache) > self.max:
            self.cache.popitem(last=False)
        return val

    def _ij(self, x, y):
        return np.floor(x / self.T).astype(int), np.floor(y / self.T).astype(int)

    def _rc(self, x, y, i, j, scale):
        x0 = i * self.T - self.O
        y1 = (j + 1) * self.T + self.O
        return np.floor((y1 - y) / scale).astype(int), np.floor((x - x0) / scale).astype(int)

    def sample(self, x: np.ndarray, y: np.ndarray, bands: Optional[List[str]] = None) -> pd.DataFrame:
        """Значения пикселов (nearest) в точках рабочей CRS."""
        x = np.asarray(x, float)
        y = np.asarray(y, float)
        bands = bands or TERRAIN_BANDS + ["worldcover"]
        out = {b: np.full(len(x), np.nan, dtype=np.float32) for b in bands}
        ok = np.isfinite(x) & np.isfinite(y)
        ii, jj = np.zeros(len(x), int), np.zeros(len(x), int)
        ii[ok], jj[ok] = self._ij(x[ok], y[ok])
        tb = [b for b in bands if b != "worldcover"]
        for (i, j) in sorted(set(zip(ii[ok].tolist(), jj[ok].tolist()))):
            sel = np.where(ok & (ii == i) & (jj == j))[0]
            if tb:
                t = self._load("terrain", i, j)
                if t is not None:
                    r, c = self._rc(x[sel], y[sel], i, j, self.m["dem_scale"])
                    n = t["elevation"].shape
                    v = (r >= 0) & (r < n[0]) & (c >= 0) & (c < n[1])
                    for b in tb:
                        out[b][sel[v]] = t[b][r[v], c[v]]
            if "worldcover" in bands:
                w = self._load("wc", i, j)
                if w is not None:
                    r, c = self._rc(x[sel], y[sel], i, j, self.m["wc_scale"])
                    n = w["worldcover"].shape
                    v = (r >= 0) & (r < n[0]) & (c >= 0) & (c < n[1])
                    out["worldcover"][sel[v]] = w["worldcover"][r[v], c[v]]
        df = pd.DataFrame(out)
        if "worldcover" in df:
            df["worldcover"] = df["worldcover"].where(df["worldcover"] > 0)
        return df

    def window(self, kind: str, x: float, y: float, radius: float):
        i, j = int(math.floor(x / self.T)), int(math.floor(y / self.T))
        t = self._load("terrain" if kind == "terrain" else "wc", i, j)
        if t is None:
            return None, None
        scale = self.m["dem_scale"] if kind == "terrain" else self.m["wc_scale"]
        r, c = self._rc(np.array([x]), np.array([y]), i, j, scale)
        r, c = int(r[0]), int(c[0])
        rp = int(math.ceil(radius / scale))
        arrs = t
        anyarr = next(iter(arrs.values()))
        r0, r1, c0, c1 = max(r - rp, 0), min(r + rp + 1, anyarr.shape[0]), max(c - rp, 0), min(c + rp + 1, anyarr.shape[1])
        yy, xx = np.mgrid[r0:r1, c0:c1]
        mask = ((yy - r) ** 2 + (xx - c) ** 2) * scale ** 2 <= max(radius, scale / 2) ** 2
        return {k: v[r0:r1, c0:c1] for k, v in arrs.items()}, mask

    def estimate_tpi_sd(self, n_tiles: int = 40) -> float:
        rng = np.random.default_rng(self.cfg["project"]["seed"])
        vals = []
        keys = self.m["keys"]
        for k in rng.permutation(len(keys))[:n_tiles]:
            i, j = keys[k]
            t = self._load("terrain", i, j)
            if t is None:
                continue
            o = int(self.O / self.m["dem_scale"])
            core = t["tpi_large"][o:-o, o:-o].ravel()
            core = core[np.isfinite(core)]
            if len(core):
                vals.append(rng.choice(core, min(5000, len(core)), replace=False))
        self.tpi_sd = float(np.std(np.concatenate(vals))) if vals else np.nan
        return self.tpi_sd


def landform_class(tpi_l, slope, sd, cfg) -> np.ndarray:
    """Формы рельефа по стандартизованному TPI крупного радиуса и уклону (Weiss 2001, 6 классов)."""
    z = np.asarray(tpi_l, float) / sd if sd and np.isfinite(sd) and sd > 0 else np.full(np.shape(tpi_l), np.nan)
    s = np.asarray(slope, float)
    flat = cfg["terrain"]["flat_slope_deg"]
    res = np.full(z.shape, None, dtype=object)
    res[z <= -1] = "valley"
    res[(z > -1) & (z <= -0.5)] = "lower_slope"
    res[(z > -0.5) & (z < 0.5) & (s <= flat)] = "flat"
    res[(z > -0.5) & (z < 0.5) & (s > flat)] = "mid_slope"
    res[(z >= 0.5) & (z < 1)] = "upper_slope"
    res[z >= 1] = "ridge"
    return res


def aspect_class(aspect, slope, cfg) -> np.ndarray:
    names = np.array(["N", "NE", "E", "SE", "S", "SW", "W", "NW"], dtype=object)
    a = np.asarray(aspect, float)
    k = (np.floor(np.mod(a + 22.5, 360) / 45).astype("float"))
    res = np.full(a.shape, None, dtype=object)
    ok = np.isfinite(a)
    res[ok] = names[k[ok].astype(int)]
    res[np.asarray(slope, float) < cfg["terrain"]["aspect_flat_slope_deg"]] = "flat"
    return res


WC_NAMES = {10: "tree_cover", 20: "shrubland", 30: "grassland", 40: "cropland", 50: "built_up", 60: "bare_sparse",
            70: "snow_ice", 80: "water", 90: "herbaceous_wetland", 95: "mangroves", 100: "moss_lichen"}
DEFAULT_RECLASS = pd.DataFrame({"code": list(WC_NAMES.keys()),
                                "habitat": ["forest", "shrubland", "grassland_meadow", "agriculture", "built", "rock_scree_sparse",
                                            "snow_ice", "water", "wetland", "other", "moss_lichen_alpine"]})
HAB_FOREST = {"forest"}
HAB_OPEN = {"grassland_meadow", "rock_scree_sparse", "moss_lichen_alpine", "shrubland", "snow_ice"}


def load_reclass(cfg: dict) -> pd.DataFrame:
    p = cfg["habitat"].get("reclass_table")
    if p and os.path.exists(p):
        t = pd.read_csv(p)
        t.columns = [c.strip().lower() for c in t.columns]
        return t
    return DEFAULT_RECLASS.copy()


def reclass_habitat(codes, elev, table: pd.DataFrame) -> np.ndarray:
    codes = np.asarray(codes, float)
    elev = np.asarray(elev, float)
    res = np.full(codes.shape, None, dtype=object)
    for _, r in table.iterrows():
        m = codes == float(r["code"])
        if "min_elev" in table.columns and pd.notna(r.get("min_elev")):
            m &= elev >= float(r["min_elev"])
        if "max_elev" in table.columns and pd.notna(r.get("max_elev")):
            m &= elev < float(r["max_elev"])
        res[m & (res == None)] = r["habitat"]  # noqa: E711
    return res


# -------------------------- динамические значения в точках (GEE) ------------------------------
DW_BANDS = ["water", "trees", "grass", "flooded_vegetation", "crops", "shrub_and_scrub", "built", "bare", "snow_and_ice"]


def _s2_prep(img):
    scl = img.select("SCL")
    clear = scl.eq(4).Or(scl.eq(5)).Or(scl.eq(6)).Or(scl.eq(11))
    nosnow = clear.And(scl.neq(11))
    ndvi = img.normalizedDifference(["B8", "B4"]).rename("ndvi").updateMask(nosnow)
    ndsi = img.normalizedDifference(["B3", "B11"]).rename("ndsi").updateMask(clear)
    return ee.Image(ndvi.addBands(ndsi).copyProperties(img, ["system:time_start"]))


def _ls_prep(img):
    sat = ee.String(img.get("SPACECRAFT_ID"))
    is_oli = sat.compareTo("LANDSAT_8").eq(0).Or(sat.compareTo("LANDSAT_9").eq(0))
    red = ee.String(ee.Algorithms.If(is_oli, "SR_B4", "SR_B3"))
    nir = ee.String(ee.Algorithms.If(is_oli, "SR_B5", "SR_B4"))
    grn = ee.String(ee.Algorithms.If(is_oli, "SR_B3", "SR_B2"))
    sw1 = ee.String(ee.Algorithms.If(is_oli, "SR_B6", "SR_B5"))
    qa = img.select("QA_PIXEL")
    cloud = qa.bitwiseAnd(1 << 1).neq(0).Or(qa.bitwiseAnd(1 << 3).neq(0)).Or(qa.bitwiseAnd(1 << 4).neq(0)).Or(qa.bitwiseAnd(1 << 2).neq(0))
    snow = qa.bitwiseAnd(1 << 5).neq(0)
    sr = img.select([red, nir, grn, sw1], ["red", "nir", "green", "swir1"]).multiply(0.0000275).add(-0.2)
    ndvi = sr.normalizedDifference(["nir", "red"]).rename("ndvi").updateMask(cloud.Not().And(snow.Not()))
    ndsi = sr.normalizedDifference(["green", "swir1"]).rename("ndsi").updateMask(cloud.Not())
    return ee.Image(ndvi.addBands(ndsi).copyProperties(img, ["system:time_start"]))


def _empty():
    return ee.Image().select([])


def build_day_image(cfg: dict, cols: dict, day_ms):
    """Мультиканальное изображение ковариат для одних UTC-суток (центр окна — полдень UTC).
    Коллекции cols уже отфильтрованы по охвату и интервалу дат порции, поэтому серверная нагрузка мала."""
    g = cfg["gee"]
    W = sorted(g["windows_days"])
    DWW = sorted(g["dw_windows_days"])
    SW = sorted(g["snow_windows_days"])
    look = max(g["snow_lookback_days"], SW[-1])
    thr = cfg["seasons"]["objective"]["snow_cover_threshold"]
    day0 = ee.Date(day_ms)
    t0 = day0.advance(12, "hour")
    imgs, sizes = [], {}

    def optical(col, prefix):
        def addq(i):
            q = ee.Image.constant(ee.Number(i.get("system:time_start")).subtract(t0.millis()).abs().divide(-86400000)).toFloat()
            return i.addBands(q.updateMask(i.select("ndvi").mask()).rename("q_ndvi")) \
                    .addBands(q.updateMask(i.select("ndsi").mask()).rename("q_ndsi"))
        big = col.filterDate(t0.advance(-W[-1], "day"), t0.advance(W[-1], "day")).map(addq)
        nn_v = big.select(["ndvi", "q_ndvi"]).qualityMosaic("q_ndvi").rename([f"{prefix}ndvi_nearest", f"{prefix}ndvi_nearest_dtdays"])
        nn_s = big.select(["ndsi", "q_ndsi"]).qualityMosaic("q_ndsi").rename([f"{prefix}ndsi_nearest", f"{prefix}ndsi_nearest_dtdays"])
        out = [ee.Image(ee.Algorithms.If(big.size().gt(0), nn_v.addBands(nn_s), _empty()))]
        red = ee.Reducer.median().combine(ee.Reducer.stdDev(), sharedInputs=True).combine(ee.Reducer.count(), sharedInputs=True)
        for w in W:
            sub = big.filterDate(t0.advance(-w, "day"), t0.advance(w, "day"))
            comp = sub.select(["ndvi", "ndsi"]).reduce(red)
            comp = comp.rename(comp.bandNames().map(lambda b: ee.String(prefix).cat(ee.String(b)).cat(f"_w{w}")))
            out.append(ee.Image(ee.Algorithms.If(sub.size().gt(0), comp, _empty())))
            sizes[f"{prefix}nscenes_w{w}"] = sub.size()
        return out

    imgs += optical(cols["s2"], "s2_")
    if cols.get("ls") is not None:
        imgs += optical(cols["ls"], "ls_")
    for w in DWW:
        sub = cols["dw"].filterDate(t0.advance(-w, "day"), t0.advance(w, "day"))
        mp = sub.select(DW_BANDS).mean().rename([f"dw_{b}_w{w}" for b in DW_BANDS])
        lab = sub.select("label").mode().rename(f"dw_label_mode_w{w}")
        cnt = sub.select("label").count().rename(f"dw_nvalid_w{w}")
        imgs.append(ee.Image(ee.Algorithms.If(sub.size().gt(0), mp.addBands(lab).addBands(cnt), _empty())))
        sizes[f"dw_nscenes_w{w}"] = sub.size()
    day_end = day0.advance(1, "day")

    def snow_prep(i):
        qa = i.select("NDSI_Snow_Cover_Basic_QA")
        sc = i.select("NDSI_Snow_Cover").updateMask(qa.lte(2))
        cloud = i.select("NDSI_Snow_Cover_Class").eq(250).unmask(0).toFloat().rename("cloud")
        valid = sc.mask().rename("valid").toFloat()
        snow = sc.gte(thr).rename("snowday").toFloat().updateMask(sc.mask())
        q = ee.Image.constant(ee.Number(i.get("system:time_start")).subtract(day_end.millis()).divide(86400000)).toFloat()
        return ee.Image(ee.Image.cat([sc.rename("ndsi_sc").toFloat(), snow, cloud, valid,
                                      q.updateMask(sc.gte(thr)).rename("q_snow"), q.updateMask(sc.lt(thr)).rename("q_nosnow")])
                        .copyProperties(i, ["system:time_start"]))
    mp = cols["snow"].filterDate(day_end.advance(-look, "day"), day_end).map(snow_prep)
    for w in SW:
        sub = mp.filterDate(day_end.advance(-w, "day"), day_end)
        comp = sub.select("ndsi_sc").mean().rename(f"modis_snowcover_mean_d{w}") \
            .addBands(sub.select("ndsi_sc").max().rename(f"modis_snowcover_max_d{w}")) \
            .addBands(sub.select("snowday").sum().rename(f"modis_snowdays_d{w}")) \
            .addBands(sub.select("valid").sum().rename(f"modis_validdays_d{w}")) \
            .addBands(sub.select("cloud").sum().rename(f"modis_clouddays_d{w}"))
        imgs.append(ee.Image(ee.Algorithms.If(sub.size().gt(0), comp, _empty())))
        sizes[f"modis_ndays_d{w}"] = sub.size()
    last = mp.select("q_snow").max().rename("modis_days_since_snow").addBands(mp.select("q_nosnow").max().rename("modis_days_since_nosnow"))
    imgs.append(ee.Image(ee.Algorithms.If(mp.size().gt(0), last, _empty())))
    mn = cols["mndvi"].filterDate(day0.advance(-16, "day"), day_end)
    mn_img = mn.map(lambda i: ee.Image(i.select("NDVI").updateMask(i.select("SummaryQA").lte(1)).multiply(0.0001)
                                       .copyProperties(i, ["system:time_start"]))).sort("system:time_start").mosaic().rename("modis_ndvi_16d")
    imgs.append(ee.Image(ee.Algorithms.If(mn.size().gt(0), mn_img, _empty())))
    er = cols["era"].filterDate(day0, day_end)
    imgs.append(ee.Image(ee.Algorithms.If(er.size().gt(0),
                                          ee.Image(er.first()).select(["snow_depth", "snow_depth_water_equivalent", "temperature_2m"],
                                                                      ["era5_snow_depth_m", "era5_swe_m", "era5_t2m_K"]), _empty())))
    return ee.Image.cat(imgs), ee.Dictionary(sizes)


def _split_dynamic_chunks(pts: pd.DataFrame, max_points: int, max_days: int) -> List[pd.DataFrame]:
    """Порции = последовательные сутки одного животного (компактный охват и интервал дат)."""
    out = []
    for _, g in pts.sort_values(["animal_id", "t_ms"]).groupby("animal_id", sort=False):
        cur, days = [], []
        for d, gd in g.groupby("day_ms", sort=True):
            if cur and (sum(len(x) for x in cur) + len(gd) > max_points or len(days) >= max_days):
                out.append(pd.concat(cur))
                cur, days = [], []
            cur.append(gd)
            days.append(d)
        if cur:
            out.append(pd.concat(cur))
    return out


def fetch_dynamic(points: pd.DataFrame, cfg: dict, cache_dir: Path) -> pd.DataFrame:
    """Динамические ковариаты в точках. Один набор изображений строится на UTC-сутки и применяется ко всем фиксациям
    этих суток через reduceRegions; коллекции заранее фильтруются по охвату и датам порции.
    Порции кэшируются; при ошибке GEE (память/таймаут) порция делится пополам."""
    g = cfg["gee"]
    src = g["sources"]
    ddir = cache_dir / "gee_dynamic"
    ddir.mkdir(parents=True, exist_ok=True)
    gkey = hash_obj(["v2_daily", src, g["windows_days"], g["dw_windows_days"], g["snow_windows_days"], g["snow_lookback_days"],
                     g["landsat_before"], cfg["seasons"]["objective"]["snow_cover_threshold"]])[:10]
    pts = points[["record_id", "animal_id", "longitude", "latitude", "timestamp_utc"]].copy()
    pts["t_ms"] = pts["timestamp_utc"].values.astype("datetime64[ms]").astype("int64")
    pts["day_ms"] = (pts["t_ms"] // 86400000) * 86400000
    W = max(g["windows_days"] + g["dw_windows_days"])
    look = max(g["snow_lookback_days"], max(g["snow_windows_days"]), 17)
    ls_before = int(pd.Timestamp(g["landsat_before"], tz="UTC").value // 10 ** 6)
    s2_all = ee.ImageCollection(src["s2"])
    ls_all = ee.ImageCollection(src["landsat"][0])
    for c in src["landsat"][1:]:
        ls_all = ls_all.merge(ee.ImageCollection(c))
    pieces = _split_dynamic_chunks(pts, int(g["chunk_size"]), int(g.get("chunk_days", 10)))

    def run_piece(df: pd.DataFrame, depth: int = 0) -> List[dict]:
        key = hash_obj([gkey, df["record_id"].tolist(), df["t_ms"].tolist(), df["longitude"].round(7).tolist(), df["latitude"].round(7).tolist()])
        fp = ddir / f"dyn_{gkey}_{key}.json"
        if fp.exists():
            return json.loads(fp.read_text(encoding="utf-8"))
        d0, d1 = int(df["day_ms"].min()), int(df["day_ms"].max()) + 86400000
        pad = 0.01
        region = ee.Geometry.Rectangle([float(df["longitude"].min()) - pad, float(df["latitude"].min()) - pad,
                                        float(df["longitude"].max()) + pad, float(df["latitude"].max()) + pad], None, False)
        a, b = ee.Date(d0 - W * 86400000), ee.Date(d1 + W * 86400000)
        cols = {"s2": s2_all.filterBounds(region).filterDate(a, b).map(_s2_prep),
                "ls": ls_all.filterBounds(region).filterDate(a, b).map(_ls_prep) if d0 < ls_before else None,
                "dw": ee.ImageCollection(src["dynamic_world"]).filterBounds(region).filterDate(a, b),
                "snow": ee.ImageCollection(src["modis_snow"]).filterDate(ee.Date(d0 - look * 86400000), ee.Date(d1)),
                "mndvi": ee.ImageCollection(src["modis_ndvi"]).filterDate(ee.Date(d0 - 17 * 86400000), ee.Date(d1)),
                "era": ee.ImageCollection(src["era5"]).filterDate(ee.Date(d0), ee.Date(d1))}
        parts = []
        for d, gd in df.groupby("day_ms"):
            fc = ee.FeatureCollection([ee.Feature(ee.Geometry.Point([float(r.longitude), float(r.latitude)]), {"pid": int(r.record_id)})
                                       for r in gd.itertuples()])
            img, sizes = build_day_image(cfg, cols, int(d))
            parts.append(img.reduceRegions(collection=fc, reducer=ee.Reducer.first(), scale=10)
                         .map(lambda f, s=sizes: f.set(s)))
        allfc = parts[0]
        for p_ in parts[1:]:
            allfc = allfc.merge(p_)
        try:
            res = with_retries(lambda: allfc.getInfo(), int(g["max_retries"]), "динамические ковариаты")
            rows = [dict(f["properties"]) for f in res["features"]]
            fp.write_text(json.dumps(rows), encoding="utf-8")
            return rows
        except Exception as e:  # noqa: BLE001
            ndays = df["day_ms"].nunique()
            if (len(df) > 5 or ndays > 1) and depth < 8:
                LOG.info(f"GEE: порция {len(df)} точек / {ndays} сут не выполнена ({str(e)[:80]}), делю пополам")
                if ndays > 1:
                    days = sorted(df["day_ms"].unique())
                    h = days[len(days) // 2]
                    return run_piece(df[df["day_ms"] < h], depth + 1) + run_piece(df[df["day_ms"] >= h], depth + 1)
                h = len(df) // 2
                return run_piece(df.iloc[:h], depth + 1) + run_piece(df.iloc[h:], depth + 1)
            warn(f"GEE: не получены динамические ковариаты для {len(df)} точек: {str(e)[:200]}")
            return []

    rows: List[dict] = []
    n_cached = sum(1 for p in pieces if (ddir / f"dyn_{gkey}_{hash_obj([gkey, p['record_id'].tolist(), p['t_ms'].tolist(), p['longitude'].round(7).tolist(), p['latitude'].round(7).tolist()])}.json").exists())
    LOG.info(f"Динамические ковариаты GEE: {len(pts)} точек, {pts.groupby(['animal_id', 'day_ms']).ngroups} животно-суток, "
             f"{len(pieces)} порций (до {g['chunk_size']} точек / {g.get('chunk_days', 10)} сут), из кэша {n_cached}")
    t_start = time.time()
    with ThreadPoolExecutor(max_workers=int(g["threads"])) as ex:
        futs = [ex.submit(run_piece, p) for p in pieces]
        for k, f in enumerate(as_completed(futs)):
            rows.extend(f.result())
            if (k + 1) % 10 == 0 or k + 1 == len(futs):
                el = time.time() - t_start
                LOG.info(f"  порций готово: {k + 1}/{len(pieces)}; прошло {el / 60:.1f} мин, осталось ~{el / (k + 1) * (len(pieces) - k - 1) / 60:.0f} мин")
    dyn = pd.DataFrame(rows)
    if len(dyn) == 0:
        return pd.DataFrame({"record_id": []})
    dyn = dyn.rename(columns={"pid": "record_id"})
    for c in dyn.columns:
        if c != "record_id":
            dyn[c] = pd.to_numeric(dyn[c], errors="coerce")
    for comp, aid, sc, tmp, proc in (
            ("ndvi_ndsi_s2", g["sources"]["s2"], "10–20 м", "2017–", "SCL: оставлены 4,5,6 (+11 для NDSI); NDVI=(B8−B4)/(B8+B4), NDSI=(B3−B11)/(B3+B11); ближайшее валидное наблюдение к полудню UTC суток фиксации (qualityMosaic по |Δt|) и медиана/SD/число в окнах ±" + ",".join(map(str, g["windows_days"])) + " сут"),
            ("ndvi_ndsi_landsat", ", ".join(g["sources"]["landsat"]), "30 м", "1984–", f"QA_PIXEL биты 1–5 (облако, тень, цирус, снег); SR×0.0000275−0.2; только для дат до {g['landsat_before']}; поле sensor"),
            ("land_cover_dynamic", g["sources"]["dynamic_world"], "10 м", "2015-06–", "средние вероятности 9 классов, мода label, число валидных наблюдений в окнах ±" + ",".join(map(str, g["dw_windows_days"])) + " сут"),
            ("snow_modis", g["sources"]["modis_snow"], "500 м", "2000–", f"NDSI_Snow_Cover при Basic_QA≤2; снег ≥{cfg['seasons']['objective']['snow_cover_threshold']}; облака (класс 250) считаются отдельно, не как отсутствие снега; ретроспективные окна " + ",".join(map(str, g["snow_windows_days"])) + " сут"),
            ("ndvi_modis", g["sources"]["modis_ndvi"], "250 м", "2000–", "NDVI×0.0001 при SummaryQA≤1, последний 16-суточный композит"),
            ("snow_era5", g["sources"]["era5"], "~9–11 км", "1950–", "суточные средние snow_depth, snow_depth_water_equivalent, temperature_2m (UTC-сутки фиксации); региональный фон")):
        add_provenance(component=comp, asset_id=aid, scale_m=sc, crs="EPSG:4326 (выборка в точке, scale=10)", temporal=tmp, processing=proc)
    return dyn


def postprocess_dynamic(pts: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Сводные поля NDVI/снега/DW с указанием источника и качества."""
    pts = pts.copy()
    g = cfg["gee"]
    W = sorted(g["windows_days"])
    # NDVI: Sentinel-2 при наличии валидного наблюдения в ±W[-1], иначе Landsat
    for kind in ("ndvi", "ndsi"):
        s2 = pts.get(f"s2_{kind}_nearest")
        ls = pts.get(f"ls_{kind}_nearest")
        if s2 is None and ls is None:
            continue
        s2 = s2 if s2 is not None else pd.Series(np.nan, index=pts.index)
        ls = ls if ls is not None else pd.Series(np.nan, index=pts.index)
        pts[f"{kind}_nearest"] = s2.where(s2.notna(), ls)
        pts[f"{kind}_sensor"] = np.where(s2.notna(), "Sentinel-2", np.where(ls.notna(), "Landsat", "none"))
        dts2 = pts.get(f"s2_{kind}_nearest_dtdays", pd.Series(np.nan, index=pts.index))
        dtls = pts.get(f"ls_{kind}_nearest_dtdays", pd.Series(np.nan, index=pts.index))
        pts[f"{kind}_obs_dt_days"] = (-dts2).where(s2.notna(), -dtls)
        for w in W:
            a = pts.get(f"s2_{kind}_median_w{w}", pd.Series(np.nan, index=pts.index))
            b = pts.get(f"ls_{kind}_median_w{w}", pd.Series(np.nan, index=pts.index))
            pts[f"{kind}_median_w{w}"] = a.where(a.notna(), b)
            ca = pts.get(f"s2_{kind}_count_w{w}", pd.Series(np.nan, index=pts.index))
            cb = pts.get(f"ls_{kind}_count_w{w}", pd.Series(np.nan, index=pts.index))
            pts[f"{kind}_nvalid_w{w}"] = ca.where(a.notna(), cb)
            na = pts.get(f"s2_nscenes_w{w}", pd.Series(np.nan, index=pts.index))
            nb = pts.get(f"ls_nscenes_w{w}", pd.Series(np.nan, index=pts.index))
            pts[f"{kind}_nscenes_w{w}"] = na.where(a.notna(), nb)
            with np.errstate(invalid="ignore", divide="ignore"):
                pts[f"{kind}_valid_frac_w{w}"] = pts[f"{kind}_nvalid_w{w}"] / pts[f"{kind}_nscenes_w{w}"]
        pts[f"{kind}_quality_flag"] = np.where(pts[f"{kind}_nearest"].isna(), "no_valid_obs",
                                               np.where(pts[f"{kind}_obs_dt_days"].abs() > 7, "nearest_obs_gt7d", "ok"))
    # Снег
    if "modis_snowcover_mean_d1" in pts:
        sc1 = pts["modis_snowcover_mean_d1"]
        thr = cfg["seasons"]["objective"]["snow_cover_threshold"]
        pts["snow_source"] = np.where(sc1.notna(), "MOD10A1_day", "none")
        pts["snow_fraction_or_ndsi"] = sc1
        pts["snow_present"] = np.where(sc1.isna(), "unknown", np.where(sc1 >= thr, "snow", "no_snow"))
        for w in g["snow_windows_days"]:
            v = pts.get(f"modis_validdays_d{w}")
            n = pts.get(f"modis_ndays_d{w}")
            if v is not None and n is not None:
                with np.errstate(invalid="ignore", divide="ignore"):
                    pts[f"modis_valid_frac_d{w}"] = v / n
                    pts[f"modis_cloud_frac_d{w}"] = pts.get(f"modis_clouddays_d{w}") / n
                    pts[f"modis_snowday_frac_d{w}"] = pts.get(f"modis_snowdays_d{w}") / v.where(v > 0)
        pts["cloud_fraction"] = pts.get("modis_cloud_frac_d1")
        pts["valid_observation_count"] = pts.get("modis_validdays_d7")
        pts["obs_time_difference"] = "same_day"
        pts["snow_quality_flag"] = np.where(sc1.notna(), "valid_qa_le2",
                                            np.where(pts.get("modis_clouddays_d1", 0) > 0, "cloud", "no_valid_obs"))
        # уточнение по S2 NDSI, если MODIS не дал значения
        if "ndsi_nearest" in pts:
            m = (pts["snow_present"] == "unknown") & pts["ndsi_nearest"].notna() & (pts["ndsi_obs_dt_days"].abs() <= 3)
            pts.loc[m, "snow_present"] = np.where(pts.loc[m, "ndsi_nearest"] > 0.4, "snow", "no_snow")
            pts.loc[m, "snow_source"] = "S2_NDSI_le3d"
    for c in ("modis_days_since_snow", "modis_days_since_nosnow"):
        if c in pts:
            pts[c] = -pts[c]
    if "era5_snow_depth_m" in pts:
        pts["snow_depth_m"] = pts["era5_snow_depth_m"]
        pts["snow_water_equivalent_m"] = pts["era5_swe_m"]
    # Dynamic World
    w = sorted(g["dw_windows_days"])[0]
    cols = [f"dw_{b}_w{w}" for b in DW_BANDS]
    if all(c in pts for c in cols):
        P = pts[cols].values.astype(float)
        has = np.isfinite(P).all(1)
        top = np.full(len(pts), None, dtype=object)
        mx = np.full(len(pts), np.nan)
        if has.any():
            top[has] = np.array(DW_BANDS, dtype=object)[np.nanargmax(P[has], 1)]
            mx[has] = np.nanmax(P[has], 1)
        pts["dw_top_class"] = top
        pts["dw_max_prob"] = mx
        pts["dw_low_confidence"] = np.where(np.isfinite(mx), mx < 0.5, True)
        pts["dw_nvalid"] = pts.get(f"dw_nvalid_w{w}")
    return pts


# ==============================================================================================
# ОБОГАЩЕНИЕ ТОЧЕК И ШАГОВ: значения в точке, буферы, вдоль линии шага
# ==============================================================================================
def enrich_points_static(pts: pd.DataFrame, store: TileStore, cfg: dict, reclass: pd.DataFrame) -> pd.DataFrame:
    v = store.sample(pts["x"].values, pts["y"].values)
    for c in v.columns:
        pts[c] = v[c].values
    sd = store.tpi_sd if store.tpi_sd is not None else store.estimate_tpi_sd()
    pts["landform"] = landform_class(pts["tpi_large"], pts["slope_deg"], sd, cfg)
    pts["aspect_class"] = aspect_class(pts["aspect_deg"], pts["slope_deg"], cfg)
    pts["worldcover_name"] = pts["worldcover"].map(WC_NAMES)
    pts["habitat"] = reclass_habitat(pts["worldcover"], pts["elevation"], reclass)
    pts["habitat_source"] = np.where(pd.notna(pts["habitat"]), "WorldCover2021+reclass", "none")
    band = cfg["terrain"]["elevation_band_m"]
    pts["elevation_band"] = (np.floor(pts["elevation"] / band) * band).map(lambda v: f"{int(v)}-{int(v + band)}" if np.isfinite(v) else None)
    return pts


def apply_regional_raster(df: pd.DataFrame, cfg: dict, lon="longitude", lat="latitude", prefix="") -> pd.DataFrame:
    p = cfg["habitat"].get("regional_raster")
    if not p or rasterio is None or not os.path.exists(p):
        return df
    with rasterio.open(p) as r:
        tr = pyproj.Transformer.from_crs("EPSG:4326", r.crs, always_xy=True)
        xx, yy = tr.transform(df[lon].values, df[lat].values)
        vals = np.array([v[0] for v in r.sample(zip(xx, yy))], dtype=float)
        if r.nodata is not None:
            vals[vals == r.nodata] = np.nan
    rc = cfg["habitat"].get("regional_reclass")
    mp = {}
    if rc and os.path.exists(rc):
        t = pd.read_csv(rc)
        mp = dict(zip(t.iloc[:, 0].astype(float), t.iloc[:, 1]))
    reg = pd.Series(vals).map(mp) if mp else pd.Series(vals).map(lambda v: f"class_{int(v)}" if np.isfinite(v) else None)
    m = reg.notna().values
    df.loc[m, prefix + "habitat"] = reg[m].values
    df.loc[m, prefix + "habitat_source"] = "regional_raster"
    return df


def buffer_stats(pts: pd.DataFrame, store: TileStore, cfg: dict, reclass: pd.DataFrame) -> pd.DataFrame:
    """Пространственные окна вокруг точки: непрерывные (mean, median, sd, p10, p90, range) и категориальные
    (доли местообитаний, доминирующий класс, разнообразие Шеннона, плотность границ классов)."""
    radii = cfg["terrain"]["buffer_radii_m"]
    bands = cfg["terrain"]["buffer_continuous"]
    codes = np.array(sorted(WC_NAMES.keys()))
    rc = dict(zip(reclass["code"].astype(int), reclass["habitat"]))
    habs = sorted(set(rc.values()))
    hab_arr = np.array(habs, dtype=object)
    lut = np.full(256, -1, dtype=np.int16)
    for code, h in rc.items():
        if 0 <= int(code) < 256:
            lut[int(code)] = habs.index(h)
    X, Y, RID = pts["x"].to_numpy(float), pts["y"].to_numpy(float), pts["record_id"].to_numpy()
    order = np.lexsort((np.floor(Y / store.T), np.floor(X / store.T)))
    rows = [None] * len(pts)
    t0 = time.time()
    for n, k in enumerate(order):
        x, y = X[k], Y[k]
        rec: Dict[str, float] = {"record_id": RID[k]}
        if not (np.isfinite(x) and np.isfinite(y)):
            rows[k] = rec
            continue
        for R in radii:
            w, m = store.window("terrain", x, y, R)
            if w is not None:
                for b in bands:
                    a = w[b][m]
                    a = a[np.isfinite(a)]
                    if len(a):
                        q10, q50, q90 = np.percentile(a, [10, 50, 90])
                        rec.update({f"{b}_r{R}_mean": float(a.mean()), f"{b}_r{R}_median": float(q50), f"{b}_r{R}_sd": float(a.std()),
                                    f"{b}_r{R}_p10": float(q10), f"{b}_r{R}_p90": float(q90), f"{b}_r{R}_range": float(a.max() - a.min())})
            w2, m2 = store.window("wc", x, y, R)
            if w2 is not None:
                arr = w2["worldcover"]
                a = arr[m2]
                hi = lut[np.clip(a.astype(np.int64), 0, 255)]
                hi = hi[(a > 0) & (hi >= 0)]
                if len(hi):
                    cnt = np.bincount(hi, minlength=len(habs)).astype(float)
                    fr = cnt / cnt.sum()
                    for j, h in enumerate(habs):
                        rec[f"hab_{h}_r{R}_frac"] = float(fr[j])
                    rec[f"hab_dominant_r{R}"] = hab_arr[int(np.argmax(cnt))]
                    p = fr[fr > 0]
                    rec[f"hab_shannon_r{R}"] = float(-(p * np.log(p)).sum())
                    hv = arr.astype(np.int16)
                    mm = m2
                    e = ((hv[:, 1:] != hv[:, :-1]) & mm[:, 1:] & mm[:, :-1]).sum() + ((hv[1:, :] != hv[:-1, :]) & mm[1:, :] & mm[:-1, :]).sum()
                    area_ha = mm.sum() * store.m["wc_scale"] ** 2 / 1e4
                    rec[f"wc_edge_density_r{R}_m_ha"] = float(e * store.m["wc_scale"] / area_ha) if area_ha > 0 else np.nan
        rows[k] = rec
        if (n + 1) % 5000 == 0:
            LOG.info(f"  буферы: {n + 1}/{len(pts)} точек ({time.time() - t0:.0f} с)")
    return pd.DataFrame(rows)


def path_stats(steps: pd.DataFrame, store: TileStore, cfg: dict, reclass: pd.DataFrame) -> pd.DataFrame:
    """Признаки вдоль линии шага: дискретизация с шагом path_sample_spacing_m."""
    sp = cfg["steps"]["path_sample_spacing_m"]
    mx = cfg["steps"]["path_max_samples"]
    L = steps["step_length_m"].fillna(0).values
    nseg = np.clip(np.ceil(L / sp).astype(int), 1, mx)
    idx = np.repeat(np.arange(len(steps)), nseg + 1)
    frac = np.concatenate([np.linspace(0, 1, k + 1) for k in nseg])
    xs = steps["x_start"].values[idx] + frac * (steps["x_end"].values[idx] - steps["x_start"].values[idx])
    ys = steps["y_start"].values[idx] + frac * (steps["y_end"].values[idx] - steps["y_start"].values[idx])
    v = store.sample(xs, ys, ["elevation", "slope_deg", "tri", "tpi_large", "worldcover"])
    v["sid"] = idx
    v["habitat"] = reclass_habitat(v["worldcover"], v["elevation"], reclass)
    v["landform"] = landform_class(v["tpi_large"], v["slope_deg"], store.tpi_sd, cfg)
    g = v.groupby("sid")
    out = pd.DataFrame(index=np.arange(len(steps)))
    for b in ("elevation", "slope_deg", "tri", "tpi_large"):
        agg = g[b].agg(["mean", "median", "min", "max", "std"])
        for c in agg.columns:
            out[f"path_{b}_{c}"] = agg[c]
    # подъём/спуск вдоль профиля
    def climb(s):
        d = np.diff(s.values)
        return pd.Series({"path_ascent_m": np.nansum(d[d > 0]), "path_descent_m": -np.nansum(d[d < 0])})
    cl = g["elevation"].apply(climb).unstack()
    out = out.join(cl)
    for cat in ("habitat", "landform"):
        ct = pd.crosstab(v["sid"], v[cat].fillna("na"), normalize="index")
        for c in ct.columns:
            if c == "na":
                continue
            out[f"path_{cat}_{c}_frac"] = ct[c]
        dom = ct.drop(columns=[c for c in ct.columns if c == "na"], errors="ignore")
        out[f"path_{cat}_dominant"] = dom.idxmax(axis=1).where(dom.sum(axis=1) > 0) if dom.shape[1] else None
        chg = v.groupby("sid")[cat].apply(lambda s: int((s.fillna("na").values[1:] != s.fillna("na").values[:-1]).sum()))
        out[f"path_{cat}_boundaries_n"] = chg
    out["path_n_samples"] = g.size()
    out["step_id"] = steps["step_id"].values
    return out.reset_index(drop=True)


def distance_to_vectors(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    ev = cfg.get("extra_vectors") or {}
    if not ev:
        return df
    try:
        import geopandas as gpd
        from shapely.strtree import STRtree
        from shapely.geometry import Point
    except Exception:  # noqa: BLE001
        warn("extra_vectors: нужен geopandas — расстояния до дорог/поселений не рассчитаны")
        return df
    for name, path in ev.items():
        try:
            g = gpd.read_file(path).to_crs(cfg["crs"]["work_crs"])
            geoms = list(g.geometry.values)
            tree = STRtree(geoms)
            pts = [Point(x, y) for x, y in zip(df["x"], df["y"])]
            idx = tree.nearest(pts)
            df[f"dist_{name}_m"] = [p.distance(geoms[i]) for p, i in zip(pts, idx)]
        except Exception as e:  # noqa: BLE001
            warn(f"extra_vectors {name}: {e}")
    return df


def attach_point_covariates_to_steps(steps: pd.DataFrame, pts: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    p = pts.set_index("record_id")
    cols = [c for c in cols if c in p.columns]
    st = p.loc[steps["rec_start"].values, cols].reset_index(drop=True)
    en = p.loc[steps["rec_end"].values, cols].reset_index(drop=True)
    st.columns = ["start_" + c for c in cols]
    en.columns = ["end_" + c for c in cols]
    steps = pd.concat([steps.reset_index(drop=True), st, en], axis=1)
    for c in cols:
        a, b = steps["start_" + c], steps["end_" + c]
        if pd.api.types.is_numeric_dtype(a) and pd.api.types.is_numeric_dtype(b):
            steps["delta_" + c] = b - a
        else:
            steps["transition_" + c] = a.astype(str) + " → " + b.astype(str)
    if "start_elevation" in steps:
        steps["vertical_change_m"] = steps["end_elevation"] - steps["start_elevation"]
        steps["vertical_speed_m_s"] = steps["vertical_change_m"] / steps["delta_t_s"]
    return steps


# ==============================================================================================
# ПОВЕДЕНЧЕСКИЕ СОСТОЯНИЯ: простой пороговый режим и HMM (гамма × фон Мизес)
# ==============================================================================================
def simple_behavior(steps: pd.DataFrame, cfg: dict) -> np.ndarray:
    b = cfg["behavior"]
    sp = steps["speed_m_h"].values
    ta = np.abs(steps["turn_angle_deg"].values)
    res = np.where(sp < b["rest_max_speed_m_h"], "rest_slow",
                   np.where((sp >= b["directed_min_speed_m_h"]) & ((ta <= b["directed_max_abs_turn_deg"]) | np.isnan(ta)),
                            "directed_movement", "area_restricted_search"))
    return res.astype(object)


def _vm_kappa(R: float) -> float:
    if R < 0.53:
        return 2 * R + R ** 3 + 5 * R ** 5 / 6
    if R < 0.85:
        return -0.4 + 1.39 * R + 0.43 / (1 - R)
    return 1 / max(R ** 3 - 4 * R ** 2 + 3 * R, 1e-6)


def _gamma_wmle(x: np.ndarray, w: np.ndarray) -> Tuple[float, float]:
    from scipy.special import digamma, polygamma
    W = w.sum()
    if W <= 0:
        return 1.0, float(np.mean(x))
    m = (w * x).sum() / W
    lm = (w * np.log(x)).sum() / W
    s = math.log(m) - lm
    k = (3 - s + math.sqrt((s - 3) ** 2 + 24 * s)) / (12 * s) if s > 0 else 50.0
    for _ in range(50):
        num = math.log(k) - digamma(k) - s
        den = 1 / k - polygamma(1, k)
        k_new = k - num / den
        if not np.isfinite(k_new) or k_new <= 0:
            break
        if abs(k_new - k) < 1e-8:
            k = k_new
            break
        k = k_new
    k = float(np.clip(k, 0.05, 200))
    return k, m / k


def fit_hmm(seqs: List[Tuple[np.ndarray, np.ndarray]], K: int, cfg: dict, seed: int) -> Optional[dict]:
    """Баум–Велч для HMM с гамма-распределением длины шага и фон Мизес для угла поворота.
    seqs — список (длины шагов > 0, углы в радианах c NaN). Пропущенный угол не входит в правдоподобие."""
    from scipy.stats import gamma as G, vonmises as VM
    h = cfg["behavior"]["hmm"]
    allx = np.concatenate([s[0] for s in seqs])
    qs = np.quantile(allx, np.linspace(0.15, 0.85, K))
    rng = np.random.default_rng(seed)
    best = None
    for start in range(3):
        mu = np.sort(qs * (1 + 0.2 * rng.standard_normal(K)).clip(0.5, 1.5))
        shape = np.full(K, 1.0)
        scale = mu / shape
        vmu = np.zeros(K)
        kap = np.linspace(0.3, 2.0, K)
        A = np.full((K, K), 0.1 / max(K - 1, 1))
        np.fill_diagonal(A, 0.9)
        pi = np.full(K, 1 / K)
        prev_ll = -np.inf
        for it in range(h["max_iter"]):
            ll = 0.0
            gam_all, xi_sum, g0 = [], np.zeros((K, K)), np.zeros(K)
            for x, ang in seqs:
                T = len(x)
                B = np.column_stack([G.pdf(x, shape[k], scale=scale[k]) for k in range(K)])
                ok = np.isfinite(ang)
                if ok.any():
                    B[ok] *= np.column_stack([VM.pdf(ang[ok], kap[k], loc=vmu[k]) for k in range(K)])
                B = np.clip(B, 1e-300, None)
                al = np.zeros((T, K))
                c = np.zeros(T)
                al[0] = pi * B[0]
                c[0] = al[0].sum()
                al[0] /= c[0]
                for t in range(1, T):
                    al[t] = (al[t - 1] @ A) * B[t]
                    c[t] = al[t].sum()
                    al[t] /= c[t]
                be = np.ones((T, K))
                for t in range(T - 2, -1, -1):
                    be[t] = (A @ (B[t + 1] * be[t + 1])) / c[t + 1]
                gm = al * be
                gm /= gm.sum(1, keepdims=True)
                if T > 1:
                    xi_sum += A * (al[:-1].T @ ((B[1:] * be[1:]) / c[1:, None]))
                g0 += gm[0]
                gam_all.append(gm)
                ll += np.log(c).sum()
            # M-шаг
            pi = g0 / g0.sum()
            A = xi_sum / xi_sum.sum(1, keepdims=True)
            GM = np.vstack(gam_all)
            ANG = np.concatenate([s[1] for s in seqs])
            ok = np.isfinite(ANG)
            for k in range(K):
                shape[k], scale[k] = _gamma_wmle(allx, GM[:, k])
                w = GM[ok, k]
                C, S = (w * np.cos(ANG[ok])).sum(), (w * np.sin(ANG[ok])).sum()
                vmu[k] = math.atan2(S, C)
                R = math.hypot(C, S) / max(w.sum(), 1e-9)
                kap[k] = float(np.clip(_vm_kappa(R), 1e-3, 100))
            if abs(ll - prev_ll) < h["tol"] * max(1, abs(ll)):
                break
            prev_ll = ll
        if best is None or ll > best["loglik"]:
            best = {"loglik": ll, "shape": shape.copy(), "scale": scale.copy(), "vm_mu": vmu.copy(), "kappa": kap.copy(),
                    "A": A.copy(), "pi": pi.copy(), "iter": it + 1}
    # упорядочить состояния по средней длине шага
    order = np.argsort(best["shape"] * best["scale"])
    for k in ("shape", "scale", "vm_mu", "kappa", "pi"):
        best[k] = best[k][order]
    best["A"] = best["A"][np.ix_(order, order)]
    n_par = K * 4 + K * (K - 1) + (K - 1)
    n = len(allx)
    best["AIC"] = -2 * best["loglik"] + 2 * n_par
    best["BIC"] = -2 * best["loglik"] + n_par * math.log(n)
    best["K"] = K
    best["n_steps"] = n
    return best


def hmm_decode(seqs, m: dict) -> List[Tuple[np.ndarray, np.ndarray]]:
    from scipy.stats import gamma as G, vonmises as VM
    K = m["K"]
    out = []
    for x, ang in seqs:
        T = len(x)
        B = np.column_stack([G.pdf(x, m["shape"][k], scale=m["scale"][k]) for k in range(K)])
        ok = np.isfinite(ang)
        if ok.any():
            B[ok] *= np.column_stack([VM.pdf(ang[ok], m["kappa"][k], loc=m["vm_mu"][k]) for k in range(K)])
        B = np.clip(B, 1e-300, None)
        al = np.zeros((T, K))
        c = np.zeros(T)
        al[0] = m["pi"] * B[0]
        c[0] = al[0].sum()
        al[0] /= c[0]
        for t in range(1, T):
            al[t] = (al[t - 1] @ m["A"]) * B[t]
            c[t] = al[t].sum()
            al[t] /= c[t]
        be = np.ones((T, K))
        for t in range(T - 2, -1, -1):
            be[t] = (m["A"] @ (B[t + 1] * be[t + 1])) / c[t + 1]
        post = al * be
        post /= post.sum(1, keepdims=True)
        # Витерби
        lA = np.log(m["A"])
        d = np.log(m["pi"]) + np.log(B[0])
        bp = np.zeros((T, K), int)
        for t in range(1, T):
            s = d[:, None] + lA
            bp[t] = s.argmax(0)
            d = s.max(0) + np.log(B[t])
        path = np.zeros(T, int)
        path[-1] = d.argmax()
        for t in range(T - 1, 0, -1):
            path[t - 1] = bp[t, path[t]]
        out.append((path, post))
    return out


HMM_NAMES = {2: ["rest_slow", "directed_movement"], 3: ["rest_slow", "area_restricted_search", "directed_movement"]}


@safe_stage("HMM")
def run_hmm(reg_steps: pd.DataFrame, cfg: dict) -> Tuple[pd.DataFrame, pd.DataFrame]:
    h = cfg["behavior"]["hmm"]
    if not h["enabled"] or len(reg_steps) == 0:
        return pd.DataFrame(), pd.DataFrame()
    diag, labels = [], []
    for a, g in reg_steps.groupby("animal_id"):
        g = g.sort_values("t_start")
        if len(g) < h["min_steps"]:
            warn(f"HMM: у {a} {len(g)} регулярных шагов < {h['min_steps']} — модуль для животного отключён")
            continue
        seqs, idxs = [], []
        for _, gb in g.groupby("burst_id"):
            if len(gb) < 5:
                continue
            x = np.clip(gb["step_length_m"].values, 1.0, None)
            ang = np.radians(gb["turn_angle_deg"].values)
            seqs.append((x, ang))
            idxs.append(gb.index.values)
        if not seqs:
            continue
        for K in h["n_states"]:
            m = fit_hmm(seqs, K, cfg, cfg["project"]["seed"])
            if m is None:
                continue
            names = HMM_NAMES.get(K, [f"state{k}" for k in range(K)])
            for k in range(K):
                diag.append({"animal_id": a, "K": K, "state": names[k], "mean_step_m": m["shape"][k] * m["scale"][k],
                             "gamma_shape": m["shape"][k], "gamma_scale": m["scale"][k], "vm_mu_rad": m["vm_mu"][k],
                             "vm_kappa": m["kappa"][k], "stay_prob": m["A"][k, k], "loglik": m["loglik"], "AIC": m["AIC"],
                             "BIC": m["BIC"], "n_steps": m["n_steps"], "primary": K == h["primary_n_states"]})
            if K == h["primary_n_states"]:
                dec = hmm_decode(seqs, m)
                for ix, (path, post) in zip(idxs, dec):
                    for r, (s, p) in enumerate(zip(path, post)):
                        rec = {"step_index": ix[r], "hmm_state": names[s]}
                        for k in range(K):
                            rec[f"hmm_p_{names[k]}"] = p[k]
                        labels.append(rec)
    return pd.DataFrame(diag), pd.DataFrame(labels)


# ==============================================================================================
# ПОТЕНЦИАЛЬНЫЕ ОХОТНИЧЬИ / КОРМОВЫЕ КЛАСТЕРЫ (hunting_candidate)
# ==============================================================================================
def _medoid(x: np.ndarray, y: np.ndarray) -> int:
    if len(x) > 400:
        cx, cy = np.median(x), np.median(y)
        return int(np.argmin((x - cx) ** 2 + (y - cy) ** 2))
    d = np.hypot(x[:, None] - x[None, :], y[:, None] - y[None, :])
    return int(np.argmin(d.sum(1)))


def detect_clusters(p: pd.DataFrame, cfg: dict, radius: float = None, min_dur_h: float = None,
                    nominal_h: Dict[str, float] = None, record_windows: bool = True, bootstrap: bool = True):
    """Скользящее по времени окно: от каждой стартовой фиксации окно расширяется, пока p95 расстояний до
    центра (покоординатная медиана) <= radius и пропуски <= allowed_gap_h; выход из радиуса допускается при
    возвращении не позднее max_excursion_h. Принятые эпизоды не перекрываются (точка входит не более чем в
    один кластер); соседние эпизоды с близкими центрами в одном burst объединяются."""
    c = cfg["clusters"]
    R = radius if radius is not None else c["radius_m"]
    D = (min_dur_h if min_dur_h is not None else c["min_duration_h"]) * 3600
    q = c["containment"]
    gap = c["allowed_gap_h"] * 3600
    exc = c["max_excursion_h"] * 3600
    rng = np.random.default_rng(cfg["project"]["seed"])
    clusters, windows = [], []
    for (a, bid), g in p.groupby(["animal_id", "burst_id"], sort=False):
        g = g.sort_values("timestamp_utc")
        x, y = g["x"].values, g["y"].values
        t = g["timestamp_utc"].values.astype("datetime64[s]").astype(np.int64).astype(float)
        rid = g["record_id"].values
        n = len(g)
        nom = (nominal_h or {}).get(a, np.nan)
        i = 0
        while i < n - 1:
            j = i
            last_in = i
            reason = None
            while j + 1 < n:
                if t[j + 1] - t[j] > gap:
                    reason = "gap"
                    break
                xs, ys = x[i:j + 2], y[i:j + 2]
                cx, cy = np.median(xs), np.median(ys)
                dd = np.hypot(xs - cx, ys - cy)
                if dd[-1] <= R:
                    if np.quantile(dd, q) <= R:
                        j += 1
                        last_in = j
                        continue
                    reason = "containment"
                    break
                # точка вне радиуса — допускаем экскурсию, если животное вернётся
                k = j + 1
                ret = None
                while k < n and t[k] - t[last_in] <= exc and t[k] - t[k - 1] <= gap:
                    if np.hypot(x[k] - cx, y[k] - cy) <= R:
                        ret = k
                        break
                    k += 1
                if ret is None:
                    reason = "left_radius"
                    break
                xs2, ys2 = x[i:ret + 1], y[i:ret + 1]
                cx2, cy2 = np.median(xs2), np.median(ys2)
                if np.quantile(np.hypot(xs2 - cx2, ys2 - cy2), q) <= R:
                    j = ret
                    last_in = ret
                else:
                    reason = "containment_after_excursion"
                    break
            if reason is None:
                reason = "end_of_burst"
            j = last_in
            dur = t[j] - t[i]
            nfix = j - i + 1
            exp_fix = dur / 3600 / nom + 1 if np.isfinite(nom) and nom > 0 else np.nan
            frac = nfix / exp_fix if np.isfinite(exp_fix) else np.nan
            accepted = dur >= D and nfix >= 3 and (not np.isfinite(frac) or frac >= c["min_fix_fraction"])
            if record_windows and (dur >= 0.25 * D or accepted):
                windows.append({"animal_id": a, "burst_id": bid, "rec_first": rid[i], "rec_last": rid[j],
                                "t_first": pd.Timestamp(t[i], unit="s", tz="UTC"), "t_last": pd.Timestamp(t[j], unit="s", tz="UTC"),
                                "duration_h": dur / 3600, "n_fixes": nfix, "expected_fixes": exp_fix, "fix_fraction": frac,
                                "stop_reason": reason, "accepted": accepted,
                                "reject_reason": "" if accepted else ("duration_lt_min" if dur < D else ("too_few_fixes" if nfix < 3 else "low_fix_fraction"))})
            if accepted:
                clusters.append({"animal_id": a, "burst_id": bid, "i0": i, "i1": j, "idx": g.index.values[i:j + 1]})
                i = j + 1
            else:
                i += 1
    # объединение соседних эпизодов
    merged = []
    for cl in clusters:
        if merged and merged[-1]["burst_id"] == cl["burst_id"]:
            pm = p.loc[merged[-1]["idx"]]
            pc = p.loc[cl["idx"]]
            d = math.hypot(np.median(pm["x"]) - np.median(pc["x"]), np.median(pm["y"]) - np.median(pc["y"]))
            g_h = (pc["timestamp_utc"].min() - pm["timestamp_utc"].max()).total_seconds() / 3600
            if d <= c["merge_distance_m"] and g_h <= c["merge_gap_h"]:
                merged[-1]["idx"] = np.r_[merged[-1]["idx"], cl["idx"]]
                continue
        merged.append(dict(cl))
    rows, pt_lab = [], {}
    for n_c, cl in enumerate(merged):
        g = p.loc[cl["idx"]].sort_values("timestamp_utc")
        cid = f"{cl['animal_id']}_c{n_c:04d}"
        x, y = g["x"].values, g["y"].values
        m = _medoid(x, y)
        mx, my = x[m], y[m]
        cx, cy = x.mean(), y.mean()
        dd = np.hypot(x - mx, y - my)
        rec = {"cluster_id": cid, "animal_id": cl["animal_id"], "burst_id": cl["burst_id"],
               "t_start": g["timestamp_utc"].iloc[0], "t_end": g["timestamp_utc"].iloc[-1],
               "duration_h": (g["timestamp_utc"].iloc[-1] - g["timestamp_utc"].iloc[0]).total_seconds() / 3600,
               "n_fixes": len(g), "medoid_x": mx, "medoid_y": my, "centroid_x": cx, "centroid_y": cy,
               "p50_distance_m": float(np.median(dd)), "p95_distance_m": float(np.quantile(dd, 0.95)), "max_distance_m": float(dd.max()),
               "medoid_record_id": g["record_id"].iloc[m]}
        nom = (nominal_h or {}).get(cl["animal_id"], np.nan)
        rec["expected_fixes"] = rec["duration_h"] / nom + 1 if np.isfinite(nom) and nom > 0 else np.nan
        rec["completeness"] = rec["n_fixes"] / rec["expected_fixes"] if np.isfinite(rec["expected_fixes"]) else np.nan
        outside = dd > R
        runs, cur = [], None
        for k_, o in enumerate(outside):
            if o and cur is None:
                cur = [k_, k_]
            elif o:
                cur[1] = k_
            elif cur is not None:
                runs.append(cur)
                cur = None
        if cur is not None:
            runs.append(cur)
        tt = g["timestamp_utc"].values
        exc_rows = []
        for r0, r1 in runs:
            t0 = tt[r0 - 1] if r0 > 0 else tt[r0]
            t1 = tt[r1 + 1] if r1 + 1 < len(tt) else tt[r1]
            exc_rows.append({"cluster_id": cid, "excursion_start": pd.Timestamp(tt[r0]), "excursion_end": pd.Timestamp(tt[r1]),
                             "duration_h": (t1 - t0) / np.timedelta64(1, "h"), "n_fixes": r1 - r0 + 1,
                             "max_distance_m": float(dd[r0:r1 + 1].max())})
        rec["n_excursions"] = len(runs)
        rec["n_returns"] = sum(1 for r0, r1 in runs if r1 + 1 < len(tt))
        rec["excursion_max_distance_m"] = max([e["max_distance_m"] for e in exc_rows], default=0.0)
        rec["_excursions"] = exc_rows
        if bootstrap and cfg["clusters"]["bootstrap"] > 0:
            B = int(cfg["clusters"]["bootstrap"])
            ii = rng.integers(0, len(x), size=(B, len(x)))
            bx, by = np.median(x[ii], axis=1), np.median(y[ii], axis=1)
            rec.update({"center_ci95_x_low": np.quantile(bx, 0.025), "center_ci95_x_high": np.quantile(bx, 0.975),
                        "center_ci95_y_low": np.quantile(by, 0.025), "center_ci95_y_high": np.quantile(by, 0.975)})
            cov = np.cov(np.vstack([bx, by]))
            rec["center_ci95_ellipse_area_m2"] = float(math.pi * 5.991464547 * math.sqrt(max(np.linalg.det(cov), 0)))
        rows.append(rec)
        for r_ in g["record_id"].values:
            pt_lab[r_] = cid
    cl_df = pd.DataFrame(rows)
    exc_df = pd.DataFrame([e for r in rows for e in r["_excursions"]])
    if len(cl_df):
        cl_df = cl_df.drop(columns=["_excursions"])
    return cl_df, pd.DataFrame(windows), exc_df, pt_lab


def label_modes(pts: pd.DataFrame, steps: pd.DataFrame, pt_lab: Dict[int, str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    pts["cluster_id"] = pts["record_id"].map(pt_lab)
    pts["behavior_mode"] = np.where(pts["cluster_id"].notna(), "hunting_candidate", "movement")
    pts.loc[~pts["included_analysis"], "behavior_mode"] = "excluded"
    a = steps["rec_start"].map(pt_lab)
    b = steps["rec_end"].map(pt_lab)
    steps["cluster_start"] = a
    steps["cluster_end"] = b
    steps["behavior_mode"] = np.where(a.notna() & b.notna() & (a == b), "hunting_candidate",
                                      np.where(a.isna() & b.isna(), "movement", "transition"))
    return pts, steps


# ==============================================================================================
# СУТОЧНЫЙ ХОД, СЕЗОНЫ, ВРЕМЯ ПОСЛЕ ВЫПУСКА
# ==============================================================================================
def split_steps_by_local_day(steps: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Шаг, пересекающий местную полночь, делится пропорционально времени (только для суточного суммирования)."""
    tz = cfg["time"]["local_timezone"]
    ls = steps["t_start"].dt.tz_convert(tz)
    le = steps["t_end"].dt.tz_convert(tz)
    pieces = []
    cur = ls.copy()
    remaining = pd.Series(True, index=steps.index)
    for _ in range(4):
        if not remaining.any():
            break
        nxt_mid = cur.dt.normalize() + pd.Timedelta(days=1)
        end = le.where(le <= nxt_mid, nxt_mid)
        sel = remaining & (end > cur)
        frac = (end - cur).dt.total_seconds() / steps["delta_t_s"]
        pieces.append(pd.DataFrame({"step_index": steps.index[sel], "animal_id": steps["animal_id"][sel],
                                    "local_date": cur[sel].dt.date, "dt_s": (end - cur)[sel].dt.total_seconds(),
                                    "length_m": steps["step_length_m"][sel] * frac[sel],
                                    "behavior_mode": steps["behavior_mode"][sel] if "behavior_mode" in steps else "movement"}))
        remaining = remaining & (le > nxt_mid)
        cur = nxt_mid.where(remaining, cur)
    return pd.concat(pieces, ignore_index=True)


def objective_seasons(pts: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Объективный сезон по дням: снег (MOD10A1) и NDVI (MOD13Q1) в точках животного."""
    o = cfg["seasons"]["objective"]
    w = o["snow_window_days"]
    sc, vc = f"modis_snowdays_d{w}", f"modis_validdays_d{w}"
    if sc not in pts or "modis_ndvi_16d" not in pts:
        return pd.DataFrame()
    d = pts[pts["included_analysis"]].groupby(["animal_id", "local_date"]).agg(
        snowdays=(sc, "median"), validdays=(vc, "median"), ndvi=("modis_ndvi_16d", "median")).reset_index()
    out = []
    for a, g in d.groupby("animal_id"):
        g = g.sort_values("local_date").copy()
        frac = g["snowdays"] / g["validdays"].where(g["validdays"] >= o["min_valid_days"])
        raw = np.where(frac >= o["snow_fraction_min"], "snow_period",
                       np.where(frac.notna() & (g["ndvi"] >= o["ndvi_active_threshold"]), "active_vegetation",
                                np.where(frac.notna() & g["ndvi"].notna(), "leaf_off_snow_free", "unknown")))
        g["season_obj_raw"] = raw
        k = int(o["smooth_days"])
        codes = pd.Series(pd.Categorical(raw))
        cat = codes.cat.categories
        cc = codes.cat.codes.astype(float).replace(-1, np.nan)
        sm = cc.rolling(k, center=True, min_periods=1).apply(lambda v: pd.Series(v).mode().iloc[0] if pd.Series(v).notna().any() else np.nan, raw=True)
        g["season_objective"] = [cat[int(v)] if np.isfinite(v) else "unknown" for v in sm]
        g["season_obj_confidence"] = cc.rolling(k, center=True, min_periods=1).apply(
            lambda v: np.mean(v == pd.Series(v).mode().iloc[0]) if pd.Series(v).notna().any() else np.nan, raw=True).values
        chg = g["season_objective"].ne(g["season_objective"].shift()).to_numpy(copy=True)
        chg[0] = False
        trans = np.zeros(len(g), bool)
        for i_ in np.where(chg)[0]:
            trans[max(0, i_ - o["transition_days"]): i_ + o["transition_days"]] = True
        g["season_obj_transition"] = trans
        g["season_obj_features"] = [f"snow_frac_{w}d={f:.2f}; ndvi={n:.2f}" if np.isfinite(f) and np.isfinite(n) else "insufficient"
                                    for f, n in zip(frac, g["ndvi"])]
        out.append(g)
    return pd.concat(out, ignore_index=True)


def daily_movement(steps: pd.DataFrame, pts: pd.DataFrame, cfg: dict, nominal: Dict[str, float], meta: pd.DataFrame) -> pd.DataFrame:
    pcs = split_steps_by_local_day(steps, cfg)
    g = pcs.groupby(["animal_id", "local_date"])
    day = pd.DataFrame({"daily_distance_m": g["length_m"].sum(), "covered_s": g["dt_s"].sum()})
    for mode in ("movement", "hunting_candidate", "transition"):
        sub = pcs[pcs["behavior_mode"] == mode].groupby(["animal_id", "local_date"])
        day[f"distance_{mode}_m"] = sub["length_m"].sum()
        day[f"time_frac_{mode}"] = sub["dt_s"].sum() / 86400
    day = day.fillna({c: 0 for c in day.columns if c.startswith(("distance_", "time_frac_"))}).reset_index()
    inc = pts[pts["included_analysis"]]
    pg = inc.sort_values("timestamp_utc").groupby(["animal_id", "local_date"])
    first = pg.first()
    last = pg.last()
    fx = pd.DataFrame({"n_fixes": pg.size(), "n_fixes_hunting": pg["behavior_mode"].apply(lambda s: (s == "hunting_candidate").sum())})
    _, _, net = _GEOD.inv(first["longitude"].values, first["latitude"].values, last["longitude"].values, last["latitude"].values)
    fx["net_displacement_m"] = net
    mxd = []
    for (a, d), gg in pg:
        xx, yy = gg["x"].values, gg["y"].values
        mxd.append(float(np.max(np.hypot(xx - xx[0], yy - yy[0]))) if len(xx) else np.nan)
    fx["max_distance_from_first_m"] = mxd
    fx["x_last"], fx["y_last"] = last["x"].values, last["y"].values
    fx["lon_last"], fx["lat_last"] = last["longitude"].values, last["latitude"].values
    fx = fx.reset_index()
    day = day.merge(fx, on=["animal_id", "local_date"], how="outer")
    st = steps.copy()
    st["local_date"] = st["t_start"].dt.tz_convert(cfg["time"]["local_timezone"]).dt.date
    sg = st.groupby(["animal_id", "local_date"])
    sp = pd.DataFrame({"mean_speed_m_h": sg["speed_m_h"].mean(), "median_speed_m_h": sg["speed_m_h"].median(),
                       "max_gap_h": sg["delta_t_s"].max() / 3600,
                       "directed_frac": sg["behavior_simple"].apply(lambda s: (s == "directed_movement").mean()) if "behavior_simple" in st else np.nan,
                       }).reset_index()
    day = day.merge(sp, on=["animal_id", "local_date"], how="left")
    day["nominal_interval_h"] = day["animal_id"].map(nominal)
    day["expected_fixes"] = 24 / day["nominal_interval_h"]
    day["fix_fraction"] = day["n_fixes"] / day["expected_fixes"]
    day["coverage"] = day["covered_s"].fillna(0) / 86400
    dcfg = cfg["daily"]
    day["complete_day"] = (day["coverage"] >= dcfg["min_coverage"]) & (day["fix_fraction"] >= dcfg["min_fix_fraction"])
    day["daily_distance_km"] = day["daily_distance_m"] / 1000
    day["daily_distance_effort_adj_km"] = day["daily_distance_km"] / day["coverage"].where(day["coverage"] > 0.25)
    day["straightness"] = (day["net_displacement_m"] / day["daily_distance_m"]).where(day["daily_distance_m"] > 0)
    day["fix_frac_hunting"] = day["n_fixes_hunting"] / day["n_fixes"]
    m = meta.set_index("animal_id")
    day["sex"] = day["animal_id"].map(m["sex"])
    day["region"] = day["animal_id"].map(m["region"])
    rel = pd.to_datetime(m["release_datetime_utc"], utc=True, errors="coerce").dt.tz_convert(cfg["time"]["local_timezone"]).dt.date
    day["days_since_release"] = [(d - rel[a]).days if pd.notna(d) and a in rel.index and pd.notna(rel[a]) else np.nan
                                 for a, d in zip(day["animal_id"], day["local_date"])]
    return day.sort_values(["animal_id", "local_date"]).reset_index(drop=True)


def bin_days(days: pd.Series, cfg: dict) -> pd.Series:
    pr = cfg["post_release"]
    return pd.cut(days, bins=[-0.5] + [b + 0.5 for b in pr["bins_days"][1:]], labels=pr["bin_labels"]).astype(str)


def _summ(v: pd.Series) -> dict:
    v = v.dropna()
    if len(v) == 0:
        return {"n": 0}
    return {"n": len(v), "mean": v.mean(), "median": v.median(), "sd": v.std(), "q25": v.quantile(.25), "q75": v.quantile(.75),
            "iqr": v.quantile(.75) - v.quantile(.25)}


def animal_bootstrap_ci(df: pd.DataFrame, col: str, B: int, seed: int) -> Tuple[float, float, str]:
    """95% ДИ среднего: бутстреп по животным (двухэтапный), при одном животном — по суткам."""
    rng = np.random.default_rng(seed)
    am = df.groupby("animal_id")[col].mean().dropna()
    if len(am) >= 2:
        bs = [rng.choice(am.values, len(am), replace=True).mean() for _ in range(B)]
        return float(np.quantile(bs, .025)), float(np.quantile(bs, .975)), "bootstrap_animals"
    v = df[col].dropna().values
    if len(v) >= 3:
        bs = [rng.choice(v, len(v), replace=True).mean() for _ in range(B)]
        return float(np.quantile(bs, .025)), float(np.quantile(bs, .975)), "bootstrap_days_single_animal"
    return np.nan, np.nan, "insufficient"


def daily_summary(day: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    d = day[day["complete_day"]].copy()
    rows = []
    B, seed = int(cfg["stats"]["bootstrap"]), cfg["project"]["seed"]
    for mode, col in (("all", "daily_distance_km"), ("movement", "distance_movement_m"), ("hunting_candidate", "distance_hunting_candidate_m")):
        if col.endswith("_m"):
            d[col + "_km"] = d[col] / 1000
            col = col + "_km"
        for keys, g in d.groupby(["sex", "region", "season"]):
            s = _summ(g[col])
            lo, hi, how = animal_bootstrap_ci(g, col, B, seed)
            ind = g.groupby("animal_id")[col].agg(["mean", "count"])
            rows.append({"sex": keys[0], "region": keys[1], "season": keys[2], "mode": mode, "n_animals": g["animal_id"].nunique(),
                         "n_animal_days": len(g), **{k: v for k, v in s.items() if k != "n"}, "ci95_low": lo, "ci95_high": hi, "ci_method": how,
                         "individual_values": "; ".join(f"{a}: {r['mean']:.2f} км (n={int(r['count'])})" for a, r in ind.iterrows())})
    return pd.DataFrame(rows)


def binseg(y: np.ndarray, min_seg: int, max_cp: int = 5) -> List[int]:
    """Бинарная сегментация по среднему с BIC-штрафом."""
    y = np.asarray(y, float)
    n = len(y)
    if n < 2 * min_seg:
        return []
    sig2 = (np.median(np.abs(np.diff(y) - np.median(np.diff(y)))) * 1.4826) ** 2 / 2
    sig2 = sig2 if sig2 > 0 else np.var(y) + 1e-9
    pen = 2 * math.log(n) * sig2

    def best_split(a, b):
        seg = y[a:b]
        m = len(seg)
        if m < 2 * min_seg:
            return None, 0
        cs = np.cumsum(seg)
        cs2 = np.cumsum(seg ** 2)
        tot = cs2[-1] - cs[-1] ** 2 / m
        k = np.arange(min_seg, m - min_seg + 1)
        left = cs2[k - 1] - cs[k - 1] ** 2 / k
        right = (cs2[-1] - cs2[k - 1]) - (cs[-1] - cs[k - 1]) ** 2 / (m - k)
        gain = tot - left - right
        j = int(np.argmax(gain))
        return a + k[j], gain[j]

    cps, segs = [], [(0, n)]
    while len(cps) < max_cp:
        cands = [(best_split(a, b), (a, b)) for a, b in segs]
        cands = [(s, g, ab) for (s, g), ab in cands if s is not None]
        if not cands:
            break
        s, g, ab = max(cands, key=lambda c: c[1])
        if g < pen:
            break
        cps.append(s)
        segs.remove(ab)
        segs += [(ab[0], s), (s, ab[1])]
    return sorted(cps)


def stabilization_day(days: np.ndarray, y: np.ndarray, cps: List[int], confirm: int) -> float:
    from scipy.stats import linregress
    if len(y) < confirm + 3:
        return np.nan
    start = cps[-1] if cps else 0
    post = y[start:]
    lo, hi = np.nanquantile(post, 0.05), np.nanquantile(post, 0.95)
    sm = pd.Series(y).rolling(3, center=True, min_periods=1).median().values
    for i in range(start, len(y) - confirm + 1):
        w = slice(i, i + confirm)
        if np.all((sm[w] >= lo) & (sm[w] <= hi)):
            r = linregress(days[w], y[w])
            if r.pvalue > 0.05:
                return float(days[i])
    return np.nan


@safe_stage("Время после выпуска и стабилизация")
def post_release_analysis(day: pd.DataFrame, pts: pd.DataFrame, meta: pd.DataFrame, cfg: dict) -> Tuple[pd.DataFrame, pd.DataFrame]:
    pr = cfg["post_release"]
    rng = np.random.default_rng(cfg["project"]["seed"])
    from shapely.geometry import MultiPoint
    m = meta.set_index("animal_id")
    rows = []
    for a, g in day.groupby("animal_id"):
        g = g.sort_values("local_date").copy()
        g = g[(g["days_since_release"] >= 0) & (g["days_since_release"] <= pr["max_days"])]
        if len(g) == 0:
            continue
        try:
            rlat, rlon = float(m.at[a, "release_lat"]), float(m.at[a, "release_lon"])
        except Exception:  # noqa: BLE001
            p0 = pts[pts["animal_id"] == a].iloc[0]
            rlat, rlon = p0["latitude"], p0["longitude"]
        _, _, nd = _GEOD.inv(np.full(len(g), rlon), np.full(len(g), rlat), g["lon_last"].values, g["lat_last"].values)
        g["net_disp_release_km"] = nd / 1000
        pa = pts[(pts["animal_id"] == a) & pts["included_analysis"]].sort_values("timestamp_utc")
        area = []
        for d in g["local_date"]:
            sub = pa[pa["local_date"] <= d]
            area.append(MultiPoint(list(zip(sub["x"], sub["y"]))).convex_hull.area / 1e6 if len(sub) >= 3 else 0.0)
        g["explored_area_km2"] = area
        g["area_growth_km2_per_day"] = pd.Series(area, index=g.index).diff(7) / 7
        g["cluster_time_frac"] = g["time_frac_hunting_candidate"]
        dd = pa.groupby("local_date")
        for col, src in (("median_elevation_m", "elevation"), ("forest_frac", "habitat"), ("snow_frac", "snow_present")):
            if src in pa:
                if src == "habitat":
                    v = dd[src].apply(lambda s: (s == "forest").mean())
                elif src == "snow_present":
                    v = dd[src].apply(lambda s: (s == "snow").sum() / max((s != "unknown").sum(), 1))
                else:
                    v = dd[src].median()
                g[col] = g["local_date"].map(v)
        rows.append(g)
    if not rows:
        return pd.DataFrame(), pd.DataFrame()
    prm = pd.concat(rows, ignore_index=True)
    prm["post_release_bin"] = bin_days(prm["days_since_release"], cfg)
    prm["weeks_since_release"] = np.floor(prm["days_since_release"] / 7).astype("Int64")
    est = []
    for a, g in prm[prm["complete_day"]].groupby("animal_id"):
        g = g.sort_values("days_since_release")
        days = g["days_since_release"].values.astype(float)
        for metric in pr["key_metrics"]:
            if metric not in g:
                continue
            y = g[metric].values.astype(float)
            ok = np.isfinite(y)
            if ok.sum() < 2 * pr["min_segment_days"] + pr["confirm_days"]:
                est.append({"animal_id": a, "metric": metric, "confirm_days": pr["confirm_days"], "n_days": int(ok.sum()),
                            "status": "insufficient_days"})
                continue
            dy, yy = days[ok], y[ok]
            cps = binseg(yy, pr["min_segment_days"])
            for conf in pr["sensitivity_confirm_days"]:
                s_day = stabilization_day(dy, yy, cps, conf)
                # бутстреп блоков остатков
                fit = np.zeros_like(yy)
                bounds = [0] + cps + [len(yy)]
                for b0, b1 in zip(bounds[:-1], bounds[1:]):
                    fit[b0:b1] = yy[b0:b1].mean()
                res = yy - fit
                bl = 7
                bs = []
                for _ in range(int(pr["bootstrap"])):
                    nb = int(math.ceil(len(res) / bl))
                    starts = rng.integers(0, max(len(res) - bl, 1), nb)
                    rr = np.concatenate([res[s:s + bl] for s in starts])[:len(res)]
                    yb = fit + rr
                    bs.append(stabilization_day(dy, yb, binseg(yb, pr["min_segment_days"]), conf))
                bs = np.array(bs, float)
                est.append({"animal_id": a, "metric": metric, "confirm_days": conf, "n_days": int(ok.sum()),
                            "change_points_day": ";".join(str(int(dy[c])) for c in cps),
                            "stabilization_day": s_day,
                            "ci95_low": np.nanquantile(bs, .025) if np.isfinite(bs).any() else np.nan,
                            "ci95_high": np.nanquantile(bs, .975) if np.isfinite(bs).any() else np.nan,
                            "bootstrap_success_frac": float(np.isfinite(bs).mean()),
                            "status": "ok" if np.isfinite(s_day) else "not_stabilized_within_series"})
    est = pd.DataFrame(est)
    if len(est):
        cons = []
        for a, g in est[(est.get("confirm_days") == pr["confirm_days"])].groupby("animal_id"):
            v = g["stabilization_day"].dropna()
            if len(v) == 0:
                cons.append({"animal_id": a, "metric": "CONSENSUS", "confirm_days": pr["confirm_days"], "status": "no_estimates"})
                continue
            med = v.median()
            agree = (np.abs(v - med) <= 14).sum()
            ok = agree > len(g) / 2
            cons.append({"animal_id": a, "metric": "CONSENSUS", "confirm_days": pr["confirm_days"],
                         "stabilization_day": med if ok else np.nan, "n_days": int(g["n_days"].max()),
                         "ci95_low": g["ci95_low"].median(), "ci95_high": g["ci95_high"].median(),
                         "status": f"consensus_{agree}_of_{len(g)}_metrics" if ok else f"no_majority_{agree}_of_{len(g)}"})
        est = pd.concat([est, pd.DataFrame(cons)], ignore_index=True)
    return prm, est


# ==============================================================================================
# ПЕРЕХОДЫ МЕЖДУ УСЛОВИЯМИ
# ==============================================================================================
TRANSITION_VARS = ["habitat", "landform", "aspect_class", "elevation_band", "dw_top_class", "snow_present", "worldcover_name"]


def transitions(steps: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    cnt_rows, prob_rows, spd_rows = [], [], []
    groupers = [("all", None), ("animal_id", "animal_id"), ("sex", "sex"), ("season", "season"),
                ("solar_period_A", "start_solar_period_A"), ("behavior_mode", "behavior_mode")]
    for v in TRANSITION_VARS:
        a, b = "start_" + v, "end_" + v
        if a not in steps or b not in steps:
            continue
        s = steps[steps[a].notna() & steps[b].notna() & (steps[a].astype(str) != "unknown") & (steps[b].astype(str) != "unknown")]
        if len(s) == 0:
            continue
        for gname, gcol in groupers:
            if gcol is not None and gcol not in s:
                continue
            for gval, g in ([("all", s)] if gcol is None else s.groupby(gcol)):
                ct = g.groupby([a, b]).agg(n=("step_id", "size"), n_animals=("animal_id", "nunique"),
                                           median_speed_m_h=("speed_m_h", "median"), mean_speed_m_h=("speed_m_h", "mean")).reset_index()
                tot = ct.groupby(a)["n"].transform("sum")
                ct["prob_from"] = ct["n"] / tot
                for r in ct.itertuples(index=False):
                    base = {"variable": v, "group_by": gname, "group": gval, "from": r[0], "to": r[1]}
                    cnt_rows.append({**base, "n_steps": r.n, "n_animals": r.n_animals})
                    prob_rows.append({**base, "prob_from": r.prob_from, "n_steps": r.n})
                    if gname == "all":
                        spd_rows.append({**base, "n_steps": r.n, "n_animals": r.n_animals, "median_speed_m_h": r.median_speed_m_h,
                                         "mean_speed_m_h": r.mean_speed_m_h})
    pr = pd.DataFrame(prob_rows)
    if len(pr):
        pr["n_from"] = pr.groupby(["variable", "group_by", "group", "from"])["n_steps"].transform("sum")
    return pd.DataFrame(cnt_rows), pr, pd.DataFrame(spd_rows)


# ==============================================================================================
# ДОСТУПНОСТЬ И ПРЕДПОЧТЕНИЕ: коэффициенты избирательности (MCP) и iSSF
# ==============================================================================================
@safe_stage("Использование/доступность местообитаний (MCP)")
def selection_ratios(pts: pd.DataFrame, store: TileStore, reclass: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    from shapely.geometry import MultiPoint, Point
    from shapely.prepared import prep
    rng = np.random.default_rng(cfg["project"]["seed"])
    rows = []
    for a, g in pts[pts["included_analysis"] & pts["habitat"].notna()].groupby("animal_id"):
        hull = MultiPoint(list(zip(g["x"], g["y"]))).convex_hull.buffer(250)
        ph = prep(hull)
        minx, miny, maxx, maxy = hull.bounds
        n = min(max(5 * len(g), 2000), 20000)
        xs, ys = [], []
        while len(xs) < n:
            cx = rng.uniform(minx, maxx, n)
            cy = rng.uniform(miny, maxy, n)
            ins = [ph.contains(Point(u, v)) for u, v in zip(cx, cy)]
            xs += list(cx[ins])
            ys += list(cy[ins])
        xs, ys = np.array(xs[:n]), np.array(ys[:n])
        av = store.sample(xs, ys, ["elevation", "worldcover"])
        ah = pd.Series(reclass_habitat(av["worldcover"], av["elevation"], reclass)).dropna()
        u = g["habitat"].value_counts()
        A = ah.value_counts(normalize=True)
        nu = u.sum()
        for h in sorted(set(u.index) | set(A.index)):
            o = u.get(h, 0) / nu
            pi = A.get(h, 0)
            w = o / pi if pi > 0 else np.nan
            se = math.sqrt(o * (1 - o) / (nu * pi ** 2)) if pi > 0 else np.nan
            rows.append({"animal_id": a, "sex": g["sex"].iloc[0], "habitat": h, "n_used": int(u.get(h, 0)), "used_frac": o,
                         "available_frac": pi, "n_available_points": len(ah), "selection_ratio": w,
                         "sr_ci95_low": w - 1.96 * se if np.isfinite(se) else np.nan, "sr_ci95_high": w + 1.96 * se if np.isfinite(se) else np.nan,
                         "availability_definition": "MCP всех валидных точек + 250 м"})
    return pd.DataFrame(rows)


def _vif(X: pd.DataFrame) -> pd.Series:
    Xc = X.dropna()
    out = {}
    for c in X.columns:
        y = Xc[c].values
        Z = np.column_stack([np.ones(len(Xc)), Xc.drop(columns=c).values])
        beta, *_ = np.linalg.lstsq(Z, y, rcond=None)
        r2 = 1 - ((y - Z @ beta) ** 2).sum() / max(((y - y.mean()) ** 2).sum(), 1e-12)
        out[c] = 1 / max(1 - r2, 1e-6)
    return pd.Series(out)


@safe_stage("iSSF")
def run_issf(reg_steps: pd.DataFrame, store: TileStore, reclass: pd.DataFrame, cfg: dict):
    from scipy.stats import gamma as G, vonmises as VM, spearmanr
    sm_cond = OPTIONAL.get("sm_cond")
    if sm_cond is None:
        warn("iSSF: statsmodels не установлен — модуль отключён")
        return None
    ic = cfg["issf"]
    rng = np.random.default_rng(cfg["project"]["seed"])
    all_avail, coefs, diags = [], [], []
    for a, g in reg_steps.groupby("animal_id"):
        g = g[g["step_length_m"] > 0].copy()
        if len(g) < ic["min_strata"]:
            warn(f"iSSF: у {a} {len(g)} регулярных шагов < {ic['min_strata']} — модель не строилась")
            continue
        sl = np.clip(g["step_length_m"].values, 1, None)
        k, _, sc = G.fit(sl, floc=0)
        ta = np.radians(g["turn_angle_deg"].dropna().values)
        kap = _vm_kappa(math.hypot(np.cos(ta).mean(), np.sin(ta).mean())) if len(ta) > 10 else 0.0
        n = int(ic["n_available"])
        rows = []
        prev_bear = g["bearing_deg"].values - np.nan_to_num(g["turn_angle_deg"].values)
        for idx, r in enumerate(g.itertuples()):
            L = G.rvs(k, scale=sc, size=n, random_state=rng)
            if np.isfinite(r.turn_angle_deg):
                ang = np.degrees(VM.rvs(max(kap, 1e-3), size=n, random_state=rng)) + prev_bear[idx]
            else:
                ang = rng.uniform(0, 360, n)
            ex = r.x_start + L * np.sin(np.radians(ang))
            ey = r.y_start + L * np.cos(np.radians(ang))
            ta_av = wrap180(ang - prev_bear[idx]) if np.isfinite(r.turn_angle_deg) else np.full(n, np.nan)
            rows.append(pd.DataFrame({"step_id": r.step_id, "animal_id": a, "case": 0, "x_end": ex, "y_end": ey,
                                      "sl": L, "ta": ta_av, "t_start": r.t_start}))
            rows.append(pd.DataFrame({"step_id": [r.step_id], "animal_id": a, "case": 1, "x_end": r.x_end, "y_end": r.y_end,
                                      "sl": [max(r.step_length_m, 1)], "ta": [r.turn_angle_deg], "t_start": [r.t_start]}))
        d = pd.concat(rows, ignore_index=True)
        cov = store.sample(d["x_end"].values, d["y_end"].values, ["elevation", "slope_deg", "tri", "tpi_large", "northness", "eastness", "worldcover"])
        d = pd.concat([d, cov], axis=1)
        hab = reclass_habitat(d["worldcover"], d["elevation"], reclass)
        d["habitat"] = hab
        d["hab_forest"] = np.isin(hab, list(HAB_FOREST)).astype(float)
        d["hab_open"] = np.isin(hab, list(HAB_OPEN)).astype(float)
        for c in ("elevation", "slope_deg", "tri", "tpi_large"):
            z = c.replace("_deg", "") + "_z"
            d[z] = (d[c] - d[c].mean()) / d[c].std()
        d["log_sl"] = np.log(d["sl"])
        d["cos_ta"] = np.cos(np.radians(d["ta"])).fillna(0)
        all_avail.append(d)
        covs = [c for c in ic["covariates"] if c in d and d[c].std() > 0]
        dd = d.dropna(subset=covs + ["sl", "log_sl"])
        good = dd.groupby("step_id")["case"].transform("sum") == 1
        dd = dd[good]
        vif = _vif(dd[covs])
        dropped = []
        while len(covs) > 1 and vif.max() > ic["max_vif"]:
            worst = vif.idxmax()
            dropped.append(f"{worst}(VIF={vif.max():.1f})")
            covs.remove(worst)
            vif = _vif(dd[covs])
        X = dd[covs + ["sl", "log_sl", "cos_ta"]]
        try:
            mod = sm_cond.ConditionalLogit(dd["case"].values, X.values, groups=dd["step_id"].values)
            res = mod.fit(disp=False, method="bfgs", maxiter=500)
            ci = res.conf_int()
            for j, name in enumerate(X.columns):
                coefs.append({"animal_id": a, "term": name, "coef": res.params[j], "se": res.bse[j], "ci95_low": ci[j][0],
                              "ci95_high": ci[j][1], "p_value": res.pvalues[j], "rss_per_1unit": math.exp(res.params[j]),
                              "n_strata": dd["step_id"].nunique(), "n_available_per_stratum": n})
            # блочная по времени k-fold валидация (Fortin et al. 2009)
            strata = dd.drop_duplicates("step_id").sort_values("t_start")["step_id"].values
            folds = np.array_split(strata, int(ic["folds"]))
            rhos = []
            for fi in range(len(folds)):
                test = set(folds[fi])
                tr = dd[~dd["step_id"].isin(test)]
                te = dd[dd["step_id"].isin(test)].copy()
                try:
                    r2 = sm_cond.ConditionalLogit(tr["case"].values, tr[X.columns].values, groups=tr["step_id"].values).fit(disp=False, method="bfgs", maxiter=300)
                    te["score"] = te[X.columns].values @ r2.params
                    te["rank"] = te.groupby("step_id")["score"].rank(ascending=False, method="average")
                    ranks = te[te["case"] == 1]["rank"].round().astype(int)
                    freq = np.bincount(ranks, minlength=n + 2)[1:n + 2]
                    rhos.append(spearmanr(np.arange(1, n + 2), freq).correlation)
                except Exception:  # noqa: BLE001
                    continue
            diags.append({"animal_id": a, "model": "iSSF_conditional_logit", "n_strata": dd["step_id"].nunique(),
                          "terms": ", ".join(X.columns), "dropped_collinear": "; ".join(dropped),
                          "max_vif": float(vif.max()) if len(vif) else np.nan, "loglik": res.llf,
                          "kfold_spearman_rho_mean": float(np.nanmean(rhos)) if rhos else np.nan,
                          "kfold_note": "rho<0 ожидается при хорошей модели (используемые шаги получают ранг 1); ~0 — случайно",
                          "gamma_shape": k, "gamma_scale": sc, "vm_kappa": kap})
        except Exception as e:  # noqa: BLE001
            warn(f"iSSF {a}: модель не сошлась: {e}")
    coef = pd.DataFrame(coefs)
    pooled = []
    if len(coef):
        for term, g in coef.groupby("term"):
            g = g[np.isfinite(g["se"]) & (g["se"] > 0)]
            if len(g) == 0:
                continue
            w = 1 / g["se"] ** 2
            fe = (w * g["coef"]).sum() / w.sum()
            Q = (w * (g["coef"] - fe) ** 2).sum()
            tau2 = max(0, (Q - (len(g) - 1)) / (w.sum() - (w ** 2).sum() / w.sum())) if len(g) > 1 else 0
            wr = 1 / (g["se"] ** 2 + tau2)
            re_ = (wr * g["coef"]).sum() / wr.sum()
            se_r = math.sqrt(1 / wr.sum())
            pooled.append({"animal_id": "POOLED_random_effects", "term": term, "coef": re_, "se": se_r, "ci95_low": re_ - 1.96 * se_r,
                           "ci95_high": re_ + 1.96 * se_r, "rss_per_1unit": math.exp(re_), "n_animals": len(g), "tau2": tau2,
                           "n_strata": int(g["n_strata"].sum())})
    coef = pd.concat([coef, pd.DataFrame(pooled)], ignore_index=True)
    avail = pd.concat(all_avail, ignore_index=True) if all_avail else pd.DataFrame()
    return avail, coef, pd.DataFrame(diags)


# ==============================================================================================
# МОДЕЛИ СКОРОСТИ, СВОДКИ, ТИПОЛОГИЯ
# ==============================================================================================
@safe_stage("Модели скорости")
def speed_models(steps: pd.DataFrame, cfg: dict) -> Tuple[pd.DataFrame, pd.DataFrame]:
    smf = OPTIONAL.get("smf")
    sm = OPTIONAL.get("sm")
    if smf is None:
        warn("Модели скорости: statsmodels не установлен")
        return pd.DataFrame(), pd.DataFrame()
    d = steps.copy()
    d["dt_h"] = d["delta_t_s"] / 3600
    terms = ["C(start_solar_period_A)", "dt_h"]
    if d["season"].nunique() > 1:
        terms.append("C(season)")
    if d["sex"].isin(["M", "F"]).sum() > 0 and d.loc[d["sex"].isin(["M", "F"]), "sex"].nunique() > 1:
        d = d[d["sex"].isin(["M", "F"])]
        terms.append("C(sex)")
    for c, z in (("start_elevation", "elev_z"), ("start_slope_deg", "slope_z"), ("start_ndvi_nearest", "ndvi_z")):
        if c in d and d[c].notna().mean() > 0.5:
            d[z] = (d[c] - d[c].mean()) / d[c].std()
            terms.append(z)
    if "start_snow_present" in d and (d["start_snow_present"] != "unknown").mean() > 0.3:
        terms.append("C(start_snow_present)")
    if "start_habitat" in d and d["start_habitat"].notna().mean() > 0.5:
        top = d["start_habitat"].value_counts().index[:5]
        d["hab5"] = np.where(d["start_habitat"].isin(top), d["start_habitat"], "other")
        terms.append("C(hab5)")
    form = "speed_m_h ~ " + " + ".join(terms)
    vars_ = [t.replace("C(", "").replace(")", "") for t in terms]
    d = d.dropna(subset=["speed_m_h"] + vars_)
    rows, diag = [], []
    # Средняя скорость в абсолютных единицах (м/ч), без преобразования отклика: GEE (гауссово семейство, тождественная связь),
    # кластеры = животные, обменная корреляция, робастные (сэндвич) SE — корректно при асимметричном распределении скорости.
    try:
        res = smf.gee(form, "animal_id", d, family=sm.families.Gaussian(), cov_struct=sm.cov_struct.Exchangeable()).fit()
        ci = res.conf_int()
        for t in res.params.index:
            rows.append({"model": "GEE mean speed, м/ч (identity link, кластер = животное, robust SE)", "term": t,
                         "coef": res.params[t], "unit": "м/ч (разность средних относительно базового уровня)" if t != "Intercept" else "м/ч",
                         "se": res.bse[t], "ci95_low": ci.loc[t, 0], "ci95_high": ci.loc[t, 1], "p_value": res.pvalues[t]})
        diag.append({"model": "speed_gee_identity", "formula": form, "n_steps": len(d), "n_animals": d["animal_id"].nunique()})
    except Exception as e:  # noqa: BLE001
        warn(f"GEE скорости не выполнена ({e}); используется OLS с кластерными ошибками по животному")
        res = smf.ols(form, d).fit(cov_type="cluster", cov_kwds={"groups": pd.factorize(d["animal_id"])[0]})
        ci = res.conf_int()
        for t in res.params.index:
            rows.append({"model": "OLS mean speed, м/ч, cluster-robust SE", "term": t, "coef": res.params[t], "unit": "м/ч",
                         "se": res.bse[t], "ci95_low": ci.loc[t, 0], "ci95_high": ci.loc[t, 1], "p_value": res.pvalues[t]})
        diag.append({"model": "speed_ols_cluster", "formula": form, "n_steps": len(d), "n_animals": d["animal_id"].nunique()})
    # почти нулевые перемещения: GEE-логит
    try:
        d["near_zero"] = (d["speed_m_h"] < cfg["behavior"]["rest_max_speed_m_h"]).astype(int)
        f2 = "near_zero ~ " + " + ".join(terms)
        r2 = smf.gee(f2, "animal_id", d, family=sm.families.Binomial(), cov_struct=sm.cov_struct.Exchangeable()).fit()
        ci = r2.conf_int()
        for t in r2.params.index:
            rows.append({"model": "GEE logit P(speed<rest_max)", "term": t, "coef": r2.params[t], "se": r2.bse[t],
                         "ci95_low": ci.loc[t, 0], "ci95_high": ci.loc[t, 1], "p_value": r2.pvalues[t], "unit": "лог-шансы",
                         "odds_ratio": math.exp(r2.params[t])})
        diag.append({"model": "near_zero_gee", "formula": f2, "n_steps": len(d), "n_animals": d["animal_id"].nunique()})
    except Exception as e:  # noqa: BLE001
        warn(f"Модель почти нулевых перемещений: {e}")
    return pd.DataFrame(rows), pd.DataFrame(diag)


def daily_mixed_model(day: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    smf = OPTIONAL.get("smf")
    if smf is None:
        return pd.DataFrame()
    d = day[day["complete_day"]].dropna(subset=["daily_distance_km", "days_since_release"]).copy()
    if d["sex"].isin(["M", "F"]).groupby(d["animal_id"]).any().sum() >= 2 * cfg["stats"]["min_animals_group"]:
        d = d[d["sex"].isin(["M", "F"])]  # пол известен у достаточного числа животных — модель с полом
    else:
        d["sex"] = "unknown"  # пол не задан — модель без фактора пола
    if d["animal_id"].nunique() < 3:
        return pd.DataFrame()
    bins = [-0.1, 30, 90, 180, 1e6]
    d["release_phase"] = pd.cut(d["days_since_release"], bins, labels=["0-30 сут", "31-90 сут", "91-180 сут", ">180 сут"]).astype(str)
    terms = [f"C({c})" for c in ("sex", "region", "season") if d[c].nunique() > 1]
    if d["release_phase"].nunique() > 1:
        ref = next(l for l in ["0-30 сут", "31-90 сут", "91-180 сут", ">180 сут"] if l in set(d["release_phase"]))
        terms.append(f'C(release_phase, Treatment("{ref}"))')
    for a_, b_ in (("sex", "season"), ("region", "season"), ("sex", "region")):
        cell = d.groupby([a_, b_])["animal_id"].nunique()
        if d[a_].nunique() > 1 and d[b_].nunique() > 1 and (cell >= cfg["stats"]["min_animals_group"]).all() and len(cell) == d[a_].nunique() * d[b_].nunique():
            terms.append(f"C({a_}):C({b_})")
    form = "daily_distance_km ~ " + (" + ".join(terms) if terms else "1")
    try:
        sm = OPTIONAL.get("sm")
        r = smf.gee(form, "animal_id", d, family=sm.families.Gaussian(), cov_struct=sm.cov_struct.Exchangeable()).fit()
        ci = r.conf_int()
        return pd.DataFrame([{"model": "GEE суточный ход, км/сут (identity link, кластер = животное, robust SE)", "formula": form,
                              "term": t, "coef": r.params[t], "unit": "км/сут" + (" (разность средних)" if t != "Intercept" else ""),
                              "se": r.bse.get(t, np.nan), "ci95_low": ci.loc[t, 0], "ci95_high": ci.loc[t, 1],
                              "p_value": r.pvalues.get(t, np.nan), "n_days": len(d), "n_animals": d["animal_id"].nunique()} for t in r.params.index])
    except Exception as e:  # noqa: BLE001
        warn(f"Модель суточного хода: {e}")
        return pd.DataFrame()


def group_comparisons(day: pd.DataFrame, steps: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Сравнения на уровне животных (средние животного), а не тысяч точек."""
    from scipy.stats import mannwhitneyu, kruskal
    rows = []
    dd = day[day["complete_day"]]
    metrics = {"daily_distance_km": dd, "time_frac_hunting_candidate": dd}
    for metric, src in metrics.items():
        for fac in ("sex", "region", "season"):
            am = src.groupby(["animal_id", fac])[metric].mean().reset_index()
            groups = {k: g[metric].dropna().values for k, g in am.groupby(fac) if k not in ("unknown", "unassigned")}
            groups = {k: v for k, v in groups.items() if len(v) > 0}
            rec = {"metric": metric, "factor": fac, "unit": "animal mean",
                   "groups": "; ".join(f"{k}: n={len(v)}, mean={np.mean(v):.3f} [{mean_ci(v)[1]:.3f}; {mean_ci(v)[2]:.3f}]" for k, v in groups.items())}
            try:
                if len(groups) == 2 and all(len(v) >= 2 for v in groups.values()):
                    v = list(groups.values())
                    rec.update(test="Mann–Whitney U", stat=mannwhitneyu(*v).statistic, p_value=mannwhitneyu(*v).pvalue)
                elif len(groups) > 2 and all(len(v) >= 2 for v in groups.values()):
                    rec.update(test="Kruskal–Wallis", stat=kruskal(*groups.values()).statistic, p_value=kruskal(*groups.values()).pvalue)
                else:
                    rec.update(test="not tested (n<2 животных в группе)")
            except Exception as e:  # noqa: BLE001
                rec.update(test=f"error: {e}")
            rec["interpretation"] = "описательно/предварительно (малое число животных)" if min([len(v) for v in groups.values()] or [0]) < 5 else ""
            rows.append(rec)
    for fac in ("start_solar_period_A", "start_solar_period_B", "behavior_mode"):
        if fac not in steps:
            continue
        am = steps.groupby(["animal_id", fac])["speed_m_h"].mean().reset_index()
        groups = {k: g["speed_m_h"].values for k, g in am.groupby(fac)}
        rec = {"metric": "mean_speed_m_h", "factor": fac, "unit": "animal mean",
               "groups": "; ".join(f"{k}: n={len(v)}, mean={np.mean(v):.1f} [{mean_ci(v)[1]:.1f}; {mean_ci(v)[2]:.1f}]" for k, v in groups.items())}
        try:
            if len(groups) >= 2 and all(len(v) >= 2 for v in groups.values()):
                k_ = kruskal(*groups.values())
                rec.update(test="Kruskal–Wallis", stat=k_.statistic, p_value=k_.pvalue)
        except Exception as e:  # noqa: BLE001
            rec.update(test=f"error: {e}")
        rows.append(rec)
    return pd.DataFrame(rows)


def individual_features(day: pd.DataFrame, steps: pd.DataFrame, pts: pd.DataFrame, clusters: pd.DataFrame,
                        stab: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    rows = []
    for a, d in day[day["complete_day"]].groupby("animal_id"):
        st = steps[steps["animal_id"] == a]
        p = pts[(pts["animal_id"] == a) & pts["included_analysis"]]
        cl = clusters[clusters["animal_id"] == a] if len(clusters) else pd.DataFrame()
        ndays = max((p["timestamp_utc"].max() - p["timestamp_utc"].min()).total_seconds() / 86400, 1)
        r = {"animal_id": a, "sex": d["sex"].iloc[0], "region": d["region"].iloc[0], "n_complete_days": len(d),
             "daily_km_mean": d["daily_distance_km"].mean(), "daily_km_q90": d["daily_distance_km"].quantile(.9),
             "speed_m_h_mean": st["speed_m_h"].mean(), "speed_m_h_q90": st["speed_m_h"].quantile(.9),
             "directed_frac": (st.get("behavior_simple") == "directed_movement").mean() if "behavior_simple" in st else np.nan,
             "localized_frac": (st.get("behavior_simple") == "rest_slow").mean() if "behavior_simple" in st else np.nan,
             "straightness_median": d["straightness"].median(),
             "turn_concentration": math.hypot(np.cos(np.radians(st["turn_angle_deg"].dropna())).mean(),
                                              np.sin(np.radians(st["turn_angle_deg"].dropna())).mean()) if st["turn_angle_deg"].notna().any() else np.nan,
             "clusters_per_30d": len(cl) / ndays * 30, "cluster_duration_h_median": cl["duration_h"].median() if len(cl) else np.nan,
             "hunting_time_frac": d["time_frac_hunting_candidate"].mean()}
        if len(cl) > 1:
            cs = cl.sort_values("t_start")
            r["intercluster_interval_h_median"] = ((cs["t_start"].values[1:] - cs["t_end"].values[:-1]) / np.timedelta64(1, "h")).astype(float).mean()
        for per in [p_["name"] for p_ in cfg["solar"]["periods_A"]]:
            r[f"speed_{per}"] = st.loc[st["start_solar_period_A"] == per, "speed_m_h"].mean()
        for c, nm in (("elevation", "elev"), ("slope_deg", "slope"), ("tri", "tri")):
            if c in p:
                r[f"{nm}_median"] = p[c].median()
                r[f"{nm}_range_p5_p95"] = p[c].quantile(.95) - p[c].quantile(.05)
        if "habitat" in p and p["habitat"].notna().any():
            fr = p["habitat"].value_counts(normalize=True)
            for h in ("forest", "shrubland", "grassland_meadow", "rock_scree_sparse"):
                r[f"hab_{h}_frac"] = fr.get(h, 0.0)
        if "start_snow_present" in st and (st["start_snow_present"] == "snow").sum() >= 20 and (st["start_snow_present"] == "no_snow").sum() >= 20:
            r["speed_ratio_snow_vs_nosnow"] = st.loc[st["start_snow_present"] == "snow", "speed_m_h"].mean() / max(
                st.loc[st["start_snow_present"] == "no_snow", "speed_m_h"].mean(), 1e-6)
        if "start_ndvi_nearest" in st and st["start_ndvi_nearest"].notna().sum() > 30:
            r["speed_ndvi_spearman"] = st[["speed_m_h", "start_ndvi_nearest"]].corr("spearman").iloc[0, 1]
        sv = d.groupby("season")["daily_distance_km"].median()
        r["seasonal_cv_daily_km"] = sv.std() / sv.mean() if len(sv) > 1 and sv.mean() > 0 else np.nan
        if stab is not None and len(stab):
            s_ = stab[(stab["animal_id"] == a) & (stab["metric"] == "CONSENSUS")]
            r["stabilization_day"] = s_["stabilization_day"].iloc[0] if len(s_) else np.nan
        rows.append(r)
    return pd.DataFrame(rows)


def gower(X: pd.DataFrame) -> np.ndarray:
    A = X.values.astype(float)
    rng_ = np.nanmax(A, 0) - np.nanmin(A, 0)
    rng_[rng_ == 0] = 1
    n = len(A)
    D = np.zeros((n, n))
    for i in range(n):
        diff = np.abs(A - A[i]) / rng_
        D[i] = np.nanmean(diff, 1)
    return D


@safe_stage("Типология животных")
def typology(feat: pd.DataFrame, day: pd.DataFrame, cfg: dict):
    from scipy.cluster.hierarchy import linkage, fcluster
    from scipy.spatial.distance import squareform
    if len(feat) < 4:
        warn(f"Типология: животных {len(feat)} < 4 — кластеризация не выполнялась")
        return None
    num = feat.drop(columns=["animal_id", "sex", "region"]).select_dtypes("number")
    num = num.loc[:, num.notna().mean() >= 0.7]
    corr = num.corr().abs()
    drop = set()
    for i, c1 in enumerate(corr.columns):
        for c2 in corr.columns[i + 1:]:
            if c1 not in drop and c2 not in drop and corr.loc[c1, c2] > 0.9:
                drop.add(c2)
    num = num.drop(columns=list(drop))
    Z = (num - num.mean()) / num.std()
    D = gower(Z)
    Dc = squareform((D + D.T) / 2, checks=False)
    L = linkage(Dc, "average")
    sil = {}
    skm = OPTIONAL.get("sk_metrics")
    for k in range(2, min(6, len(feat))):
        lab = fcluster(L, k, "maxclust")
        if len(set(lab)) > 1 and skm is not None:
            sil[k] = skm.silhouette_score(D, lab, metric="precomputed")
    best_k = max(sil, key=sil.get) if sil else 1
    weak = (not sil) or sil[best_k] < 0.25
    lab = fcluster(L, best_k, "maxclust") if not weak else np.ones(len(feat), int)
    # устойчивость: бутстреп по признакам (80%)
    rng = np.random.default_rng(cfg["project"]["seed"])
    co = np.zeros((len(feat), len(feat)))
    B = 200
    for _ in range(B):
        cols = rng.choice(Z.columns, max(2, int(0.8 * Z.shape[1])), replace=False)
        Db = gower(Z[cols])
        lb = fcluster(linkage(squareform((Db + Db.T) / 2, checks=False), "average"), max(best_k, 2), "maxclust")
        co += (lb[:, None] == lb[None, :])
    co /= B
    # PCA
    Zf = Z.fillna(0).values
    U, S, Vt = np.linalg.svd(Zf - Zf.mean(0), full_matrices=False)
    pcs = U[:, :2] * S[:2]
    expl = (S ** 2 / (S ** 2).sum())[:2]
    clus = pd.DataFrame({"animal_id": feat["animal_id"], "sex": feat["sex"], "region": feat["region"], "group": lab,
                         "PC1": pcs[:, 0], "PC2": pcs[:, 1] if pcs.shape[1] > 1 else 0})
    clus["structure"] = "weak (silhouette<0.25) — разделение не форсируется" if weak else f"k={best_k}, silhouette={sil[best_k]:.2f}"
    names = {}
    if not weak:
        for gid, g in Z.assign(group=lab).groupby("group"):
            m = g.drop(columns="group").mean().sort_values()
            names[gid] = f"высокие: {', '.join(m.index[-2:])}; низкие: {', '.join(m.index[:2])}"
    clus["group_profile"] = clus["group"].map(names).fillna("")
    sim = pd.DataFrame(1 - D, index=feat["animal_id"], columns=feat["animal_id"])
    coas = pd.DataFrame(co, index=feat["animal_id"], columns=feat["animal_id"])
    load = pd.DataFrame(Vt[:2].T, index=Z.columns, columns=["PC1_loading", "PC2_loading"][:Vt[:2].shape[0]])
    info = {"silhouette": sil, "best_k": best_k, "weak": weak, "explained": expl.tolist(), "dropped_corr": sorted(drop),
            "linkage": L, "features_used": list(Z.columns)}
    return clus, sim, coas, load, info


# ==============================================================================================
# ТАБЛИЦЫ, СЛОВАРЬ ДАННЫХ, РИСУНКИ, КАРТА, HTML-ОТЧЁТ
# ==============================================================================================
FIELD_DOC = {
    "record_id": ("—", "Идентификатор исходной записи (порядок чтения)"),
    "animal_id": ("—", "Животное"), "sex": ("M/F/unknown", "Пол"), "region": ("—", "Регион (Западный/Центральный Кавказ)"),
    "timestamp_utc": ("UTC", "Время фиксации"), "timestamp_local": ("местное", "Время фиксации в time.local_timezone"),
    "longitude": ("° WGS84", "Долгота"), "latitude": ("° WGS84", "Широта"), "x": ("м, work_crs", "Восточная координата"),
    "y": ("м, work_crs", "Северная координата"), "qc_flag": ("ok/excluded", "Итог контроля качества"),
    "qc_reason": ("—", "Причины исключения через ';'"), "included_analysis": ("bool", "Используется в анализе"),
    "burst_id": ("—", "Непрерывный фрагмент трека (разрыв > qc.burst_max_gap_h)"), "dt_prev_s": ("с", "Интервал от предыдущей валидной фиксации burst"),
    "step_id": ("—", "Идентификатор шага"), "step_length_m": ("м", "Геодезическая длина шага (WGS84)"),
    "delta_t_s": ("с", "Интервал шага"), "speed_m_s": ("м/с", "Расчётная скорость перемещения между GPS-фиксациями (минимальная средняя)"),
    "speed_km_h": ("км/ч", "То же, км/ч (для порогов QC)"), "speed_m_h": ("м/ч", "Расчётная скорость перемещения между фиксациями, м/ч (основная единица)"), "bearing_deg": ("°", "Азимут шага"),
    "turn_angle_deg": ("°, −180…180", "Угол поворота относительно предыдущего шага"), "vertical_change_m": ("м", "Изменение высоты (конец − начало)"),
    "solar_period_A": ("—", "Рабочие периоды зоологов (местное время)"), "solar_period_B": ("—", "Астрономический период (высота Солнца)"),
    "sun_elevation_deg": ("°", "Высота Солнца (NOAA)"), "solar_hour": ("ч", "Истинное солнечное время"),
    "season": ("—", "Основной сезон (seasons.primary)"), "season_source": ("—", "Источник сезона"),
    "behavior_mode": ("movement/hunting_candidate/transition", "Операционный режим (кластер ≥24 ч, p95 ≤150 м)"),
    "behavior_simple": ("—", "Пороговый поведенческий режим"), "hmm_state": ("—", "Состояние HMM (Витерби)"),
    "cluster_id": ("—", "Потенциальный охотничий/кормовой кластер (не подтверждённое место добычи)"),
    "elevation": ("м", "Высота Copernicus DEM GLO-30 (DSM)"), "slope_deg": ("°", "Уклон"), "aspect_deg": ("°", "Экспозиция"),
    "tri": ("м", "Terrain Ruggedness Index"), "tpi_large": ("м", "TPI крупного радиуса"), "tpi_small": ("м", "TPI малого радиуса"),
    "landform": ("—", "Форма рельефа (Weiss 2001)"), "aspect_class": ("—", "Экспозиция по 8 румбам + flat"),
    "worldcover": ("код", "ESA WorldCover 2021"), "habitat": ("—", "Переклассифицированное местообитание"),
    "ndvi_nearest": ("—", "NDVI ближайшего валидного наблюдения"), "ndvi_sensor": ("—", "Сенсор NDVI"),
    "ndvi_obs_dt_days": ("сут", "|Δt| GPS — снимок"), "snow_present": ("snow/no_snow/unknown", "Снег (облака → unknown)"),
    "snow_depth_m": ("м", "ERA5-Land snow_depth (региональный фон)"), "snow_water_equivalent_m": ("м в.э.", "ERA5-Land SWE"),
    "dw_top_class": ("—", "Dynamic World: класс с макс. средней вероятностью"), "dw_max_prob": ("0–1", "Dynamic World: макс. вероятность"),
    "daily_distance_km": ("км", "Суточный ход (сумма шагов, разрезанных по местной полуночи)"), "coverage": ("0–1", "Доля суток, покрытая валидными шагами"),
    "complete_day": ("bool", "Сутки достаточной полноты"), "days_since_release": ("сут", "Время после выпуска"),
    "duration_h": ("ч", "Продолжительность"), "p95_distance_m": ("м", "95-й процентиль расстояний до медоида"),
    "center_ci95_ellipse_area_m2": ("м²", "Площадь 95% доверительного эллипса центра (бутстреп)"),
    "selection_ratio": ("—", "Коэффициент избирательности Манли (используемая/доступная доля)"),
    "coef": ("—", "Коэффициент модели"), "rss_per_1unit": ("—", "Относительная сила выбора exp(β) на единицу ковариаты"),
}

TABLE_KEYS = {
    "points_cleaned": "record_id", "steps_enriched": "step_id", "individual_summary": "animal_id", "sex_summary": "sex",
    "season_summary": "season", "solar_period_summary": "classification+period", "habitat_use": "animal_id+habitat",
    "terrain_use": "animal_id+variable+class", "ndvi_snow_summary": "animal_id+season", "transition_counts": "variable+group_by+group+from+to",
    "transition_probabilities": "variable+group_by+group+from+to", "hunting_clusters": "cluster_id", "cluster_candidate_windows": "animal_id+rec_first",
    "cluster_excursions": "cluster_id+excursion_start", "behavior_labeled_points": "record_id", "behavior_labeled_steps": "step_id",
    "daily_movement": "animal_id+local_date", "daily_movement_summary": "sex+region+season+mode", "post_release_metrics": "animal_id+local_date",
    "stabilization_estimates": "animal_id+metric+confirm_days", "individual_features": "animal_id", "individual_similarity": "animal_id×animal_id",
    "individual_clusters": "animal_id", "speed_by_conditions": "variable+class", "available_steps": "step_id+case+row",
    "issf_coefficients": "animal_id+term", "model_diagnostics": "model+animal_id", "data_provenance": "component",
    "excluded_records": "record_id", "cluster_sensitivity": "radius_m+min_duration_h+animal_id", "hmm_parameters": "animal_id+K+state",
    "selection_ratios": "animal_id+habitat", "group_comparisons": "metric+factor", "input_files": "file", "field_completeness": "field",
    "speed_models": "model+term", "daily_models": "term", "seasons_objective": "animal_id+local_date", "animals_metadata": "animal_id",
}


def _excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    for c in d.columns:
        if isinstance(d[c].dtype, pd.DatetimeTZDtype):
            d[c] = d[c].dt.strftime("%Y-%m-%d %H:%M:%S%z")
        elif d[c].dtype == object:
            d[c] = d[c].map(lambda v: v.isoformat() if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)) else v)
    return d


def write_tables(tables: Dict[str, pd.DataFrame], out: Path, n_source: int) -> pd.DataFrame:
    tdir = out / "tables"
    tdir.mkdir(parents=True, exist_ok=True)
    dict_rows = []
    xlsx = tdir / "LeoTrack_tables.xlsx"
    big = {}
    for name, df in tables.items():
        if df is None or not isinstance(df, pd.DataFrame):
            continue
        df.to_csv(tdir / f"{name}.csv", index=False, encoding="utf-8-sig")
        if len(df) <= 200_000 and len(df) * max(len(df.columns), 1) <= 1_500_000:
            big[name] = df
        for c in df.columns:
            u, dsc = FIELD_DOC.get(c, ("", ""))
            if not dsc:
                for pref in ("start_", "end_", "path_", "delta_"):
                    if c.startswith(pref) and c[len(pref):] in FIELD_DOC:
                        u, dsc = FIELD_DOC[c[len(pref):]]
                        dsc = {"start_": "в начале шага: ", "end_": "в конце шага: ", "path_": "вдоль шага: ", "delta_": "изменение (конец − начало): "}[pref] + dsc
            dict_rows.append({"table": name, "field": c, "unit": u, "description": dsc, "dtype": str(df[c].dtype),
                              "non_null_frac": float(df[c].notna().mean()) if len(df) else np.nan})
    ddict = pd.DataFrame(dict_rows)
    ddict.to_csv(tdir / "data_dictionary.csv", index=False, encoding="utf-8-sig")
    meta = pd.DataFrame([{"table": n, "primary_key": TABLE_KEYS.get(n, ""), "n_rows": len(d), "n_columns": len(d.columns),
                          "n_source_records": n_source, "in_xlsx": n in big, "csv_file": f"{n}.csv"}
                         for n, d in tables.items() if isinstance(d, pd.DataFrame)])
    meta.to_csv(tdir / "tables_index.csv", index=False, encoding="utf-8-sig")
    try:
        with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
            meta.to_excel(xw, sheet_name="_index", index=False)
            ddict.to_excel(xw, sheet_name="_data_dictionary", index=False)
            for name, df in big.items():
                _excel_safe(df).to_excel(xw, sheet_name=name[:31], index=False)
        gc.collect()
    except Exception as e:  # noqa: BLE001
        warn(f"XLSX не записан: {e}")
    if OPTIONAL.get("pyarrow") is not None:
        for name in ("points_cleaned", "steps_enriched"):
            if name in tables:
                try:
                    _excel_safe(tables[name]).to_parquet(tdir / f"{name}.parquet")
                except Exception:  # noqa: BLE001
                    pass
    return meta


# -------------------------------------- рисунки ----------------------------------------------
FIGS: List[Tuple[str, str, str]] = []  # (раздел, файл, подпись)
COLORS = plt.get_cmap("tab20").colors


def savefig(fig, out: Path, name: str, section: str, caption: str, dpi: int) -> None:
    p = out / "figures" / f"{name}.png"
    p.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(p, dpi=dpi, bbox_inches="tight")
    plt.close(fig)
    FIGS.append((section, str(p), caption))


def _fig_guard(fn):
    def inner(*a, **k):
        try:
            return fn(*a, **k)
        except Exception as e:  # noqa: BLE001
            warn(f"Рисунок {fn.__name__} не построен: {type(e).__name__}: {e}")
            plt.close("all")
    return inner


def _acolor(animals):
    return {a: COLORS[i % len(COLORS)] for i, a in enumerate(sorted(animals))}


@_fig_guard
def fig_qc(pts, out, dpi):
    fig, ax = plt.subplots(1, 2, figsize=(13, 4.5))
    c = pts.groupby(["animal_id", "qc_flag"]).size().unstack(fill_value=0)
    c.plot.barh(stacked=True, ax=ax[0], color={"ok": "#4c9f70", "excluded": "#d1495b"})
    ax[0].set_title(f"Фиксации по животным (всего {len(pts)})")
    ax[0].set_xlabel("число записей")
    d = pts.loc[pts["included_analysis"], ["animal_id", "dt_prev_s"]].dropna()
    for a, g in d.groupby("animal_id"):
        ax[1].hist(np.log10(g["dt_prev_s"] / 3600), bins=60, histtype="step", label=a)
    ax[1].set_xlabel("log10(интервал, ч)")
    ax[1].set_ylabel("число интервалов")
    ax[1].set_title("Распределение интервалов между фиксациями")
    ax[1].legend(fontsize=7, ncol=2)
    savefig(fig, out, "qc_overview", "qc", "Контроль качества: исключённые записи и интервалы фиксаций.", dpi)
    r = pts.loc[~pts["included_analysis"], "qc_reason"].str.split(";").explode().value_counts()
    if len(r):
        fig, ax = plt.subplots(figsize=(8, 0.4 * len(r) + 1.5))
        r.plot.barh(ax=ax, color="#d1495b")
        ax.set_title("Причины исключения записей")
        savefig(fig, out, "qc_reasons", "qc", "Причины исключения (одна запись может иметь несколько причин).", dpi)


@_fig_guard
def fig_overview(pts, clusters, meta, out, dpi, public, cfg):
    inc = pts[pts["included_analysis"]]
    col = _acolor(inc["animal_id"].unique())
    fig, ax = plt.subplots(figsize=(11, 7))
    if "elevation" in inc and inc["elevation"].notna().any():
        sc = ax.scatter(inc["longitude"], inc["latitude"], c=inc["elevation"], s=1, cmap="terrain", alpha=0.25, zorder=1)
        plt.colorbar(sc, ax=ax, label="высота, м", shrink=0.7)
    for a, g in inc.groupby("animal_id"):
        for _, gb in g.groupby("burst_id"):
            ax.plot(gb["longitude"], gb["latitude"], lw=0.6, color=col[a], zorder=2)
        ax.plot([], [], color=col[a], label=f"{a} ({len(g)} фикс.)")
    if not public:
        m = meta.copy()
        ax.scatter(pd.to_numeric(m["release_lon"], errors="coerce"), pd.to_numeric(m["release_lat"], errors="coerce"),
                   marker="*", s=120, c="k", zorder=3, label="место выпуска")
        if len(clusters):
            lon, lat = from_work_xy(clusters["medoid_x"], clusters["medoid_y"], cfg["crs"]["work_crs"])
            ax.scatter(lon, lat, s=8, c="red", zorder=4, label=f"hunting_candidate (n={len(clusters)})")
    ax.set_xlabel("долгота")
    ax.set_ylabel("широта")
    ax.set_title(f"Треки: {inc['animal_id'].nunique()} животных, {len(inc)} валидных фиксаций")
    ax.legend(fontsize=7, loc="best", ncol=2)
    ax.set_aspect(1 / math.cos(math.radians(inc["latitude"].mean())))
    savefig(fig, out, "overview_map", "data", "Обзорная карта треков (линии не соединяют фиксации через разрывы burst).", dpi)


def _mean_ci_plot(ax, labels, animal_vals, pooled_vals, color="k"):
    """Точка — среднее средних животных, усы — 95% ДИ (t, единица — животное); если животное одно — 95% ДИ по шагам/суткам.
    Красные точки — средние отдельных животных."""
    rng = np.random.default_rng(0)
    for i, (am, pv) in enumerate(zip(animal_vals, pooled_vals)):
        am = np.asarray(am, float)
        am = am[np.isfinite(am)]
        if len(am) >= 2:
            m, lo, hi, _ = mean_ci(am)
            mk = "o"
        else:
            m, lo, hi, _ = mean_ci(pv)
            mk = "s"
        if np.isfinite(m):
            yerr = None if not np.isfinite(lo) else [[m - lo], [hi - m]]
            ax.errorbar(i, m, yerr=yerr, fmt=mk, c=color, ms=7, capsize=5, lw=1.6, zorder=4)
        if len(am):
            ax.scatter(i + rng.uniform(-0.18, 0.18, len(am)), am, s=14, c="crimson", alpha=0.8, zorder=3)
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=7)
    ax.set_xlim(-0.6, len(labels) - 0.4)
    ax.grid(axis="y", alpha=0.3)


@_fig_guard
def fig_speed(steps, out, dpi):
    fig, ax = plt.subplots(1, 2, figsize=(13, 4.5))
    xmax = float(np.nanquantile(steps["speed_m_h"], 0.99))
    bins = np.linspace(0, xmax, 60)
    for a, g in steps.groupby("animal_id"):
        ax[0].hist(g["speed_m_h"].clip(upper=xmax), bins=bins, histtype="step", density=True, label=a)
        s = np.sort(g["speed_m_h"].values)
        ax[1].plot(s, np.arange(1, len(s) + 1) / len(s), lw=1, label=a)
    ax[0].set_xlabel("скорость, м/ч")
    ax[0].set_ylabel("плотность")
    ax[0].set_title(f"Расчётная скорость между фиксациями (n шагов={len(steps)}; ось до 99-го перцентиля)")
    ax[1].set_xlim(0, xmax)
    ax[1].set_xlabel("скорость, м/ч")
    ax[1].set_ylabel("ECDF")
    ax[1].legend(fontsize=7, ncol=2)
    savefig(fig, out, "speed_hist_ecdf", "speed", "Распределение и ECDF расчётной скорости (м/ч, линейная шкала; значения выше 99-го перцентиля "
            "собраны в последний интервал гистограммы).", dpi)


@_fig_guard
def fig_hourly(steps, out, dpi, cfg):
    s = steps.copy()
    s["hour"] = s["t_mid"].dt.tz_convert(cfg["time"]["local_timezone"]).dt.hour
    fig, ax = plt.subplots(1, 3, figsize=(16, 4.5))
    for k, col in enumerate(("speed_m_h", "step_length_m")):
        ah = s.groupby(["animal_id", "hour"])[col].mean().unstack()
        for a, row in ah.iterrows():
            ax[k].plot(row.index, row.values, lw=0.8, alpha=0.5)
        st_ = [mean_ci(ah[h_]) for h_ in ah.columns]
        ax[k].fill_between(ah.columns, [v[1] for v in st_], [v[2] for v in st_], color="k", alpha=0.18, lw=0, label="95% ДИ (по животным)")
        ax[k].plot(ah.columns, [v[0] for v in st_], lw=2.5, c="k", label="среднее животных")
        ax[k].set_ylim(bottom=0)
    ax[0].legend(fontsize=7)
    for p in cfg["solar"]["periods_A"]:
        ax[0].axvline(_hm(p["start"]) / 60, ls=":", c="grey")
    ax[0].set_title(f"Средняя скорость по местному часу (n животных={s['animal_id'].nunique()})")
    ax[0].set_xlabel("час (середина шага)")
    ax[0].set_ylabel("м/ч")
    ax[1].set_title("Средняя длина шага")
    ax[1].set_xlabel("час")
    ax[1].set_ylabel("м")
    fr = s.groupby("hour")["behavior_mode"].value_counts(normalize=True).unstack(fill_value=0)
    fr.plot.area(ax=ax[2], alpha=0.7)
    ax[2].set_title("Доля шагов по режиму")
    ax[2].set_xlabel("час")
    savefig(fig, out, "hourly_activity", "speed", "Средние скорость (м/ч) и длина шага по местному часу: тонкие линии — средние отдельных животных, чёрная — среднее животных с 95% ДИ (t, единица — животное); пунктир — границы рабочих периодов A; справа — доли режимов.", dpi)
    fig = plt.figure(figsize=(12, 5))
    ax1 = fig.add_subplot(121, projection="polar")
    w = s.groupby("hour")["step_length_m"].sum()
    th = np.radians(w.index * 15)
    ax1.bar(th, w.values / w.sum(), width=np.radians(15), alpha=0.7)
    ax1.set_theta_zero_location("N")
    ax1.set_theta_direction(-1)
    ax1.set_xticks(np.radians(np.arange(0, 360, 45)))
    ax1.set_xticklabels([f"{h}:00" for h in range(0, 24, 3)])
    ax1.set_title("Циркулярная плотность активности\n(доля суммарной дистанции по часу)")
    ax2 = fig.add_subplot(122, projection="polar")
    for per, g in s.groupby("start_solar_period_A"):
        h_, e_ = np.histogram(np.radians(g["bearing_deg"]), bins=16, range=(0, 2 * np.pi))
        ax2.plot(np.r_[e_[:-1], e_[0]], np.r_[h_, h_[0]] / max(h_.sum(), 1), label=per)
    ax2.set_theta_zero_location("N")
    ax2.set_theta_direction(-1)
    ax2.set_title("Направления шагов по периодам A")
    ax2.legend(fontsize=7, loc="lower left", bbox_to_anchor=(1, 0))
    savefig(fig, out, "circular_activity", "speed", "Эмпирическая циркулярная плотность активности (без заданных границ) и направления движения.", dpi)
    if "season" in s and s["season"].nunique() > 0:
        fig, ax = plt.subplots(1, 2, figsize=(16, 3.8))
        fig.subplots_adjust(wspace=0.45)
        for k, col in enumerate(("speed_m_h", "step_length_m")):
            hm = s.pivot_table(index="season", columns="hour", values=col, aggfunc="mean")
            im = ax[k].imshow(hm.values, aspect="auto", cmap="viridis")
            ax[k].set_yticks(range(len(hm.index)))
            ax[k].set_yticklabels(hm.index)
            ax[k].set_xticks(range(0, 24, 3))
            ax[k].set_xlabel("местный час")
            ax[k].set_title("средняя скорость, м/ч" if col == "speed_m_h" else "средняя длина шага, м")
            plt.colorbar(im, ax=ax[k])
        savefig(fig, out, "heatmap_hour_season", "speed", "Heatmap «час × сезон».", dpi)
    ct = pd.crosstab(s["start_solar_period_A"], s["start_solar_period_B"], normalize="index")
    fig, ax = plt.subplots(figsize=(7, 3.5))
    im = ax.imshow(ct.values, cmap="Blues", vmin=0, vmax=1)
    ax.set_xticks(range(ct.shape[1]))
    ax.set_xticklabels(ct.columns, rotation=30, ha="right")
    ax.set_yticks(range(ct.shape[0]))
    ax.set_yticklabels(ct.index)
    for i in range(ct.shape[0]):
        for j in range(ct.shape[1]):
            ax.text(j, i, f"{ct.values[i, j]:.2f}", ha="center", va="center", fontsize=8)
    ax.set_title("Рабочие периоды A × астрономические периоды B")
    savefig(fig, out, "periods_A_vs_B", "speed", "Сопоставление рабочих (A) и астрономических (B) периодов: доли шагов по строкам.", dpi)


@_fig_guard
def fig_box(steps, out, dpi, var, title, section, fname):
    if var not in steps or steps[var].notna().sum() == 0:
        return
    s = steps[steps[var].notna() & (steps[var].astype(str) != "unknown")]
    if len(s) == 0:
        return
    cats = s.groupby(var)["speed_m_h"].mean().sort_values().index.tolist()
    if len(cats) > 15:
        cats = s[var].value_counts().index[:15].tolist()
    am = s.groupby([var, "animal_id"])["speed_m_h"].mean()
    fig, ax = plt.subplots(figsize=(max(6, 0.85 * len(cats) + 2), 4.5))
    _mean_ci_plot(ax, [f"{c}\nшагов={int((s[var] == c).sum())}, жив.={s.loc[s[var] == c, 'animal_id'].nunique()}" for c in cats],
                  [am.loc[c].values if c in am.index.get_level_values(0) else [] for c in cats],
                  [s.loc[s[var] == c, "speed_m_h"].values for c in cats])
    ax.set_ylabel("средняя скорость, м/ч")
    ax.set_ylim(bottom=0)
    ax.set_title(title + ": среднее ± 95% ДИ")
    savefig(fig, out, fname, section, title + ": точка — среднее средних животных, усы — 95% ДИ (t, единица — животное; квадрат — "
            "одно животное, ДИ по шагам), красные точки — средние отдельных животных; м/ч.", dpi)


@_fig_guard
def fig_speed_vs(steps, out, dpi):
    vars_ = [(c, n) for c, n in (("start_elevation", "высота, м"), ("start_slope_deg", "уклон, °"), ("start_ndvi_nearest", "NDVI"),
                                  ("start_snow_depth_m", "глубина снега ERA5, м")) if c in steps and steps[c].notna().sum() > 30]
    if not vars_:
        return
    fig, ax = plt.subplots(1, len(vars_), figsize=(4.5 * len(vars_), 4), squeeze=False)
    for k, (c, n) in enumerate(vars_):
        s = steps[[c, "speed_m_h", "animal_id"]].dropna()
        bins = np.unique(np.nanquantile(s[c], np.linspace(0, 1, 13)))
        s["bin"] = pd.cut(s[c], bins, include_lowest=True)
        for a, g in s.groupby("animal_id"):
            m = g.groupby("bin", observed=True)[["speed_m_h", c]].mean()
            ax[0, k].plot(m[c], m["speed_m_h"], lw=0.8, alpha=0.6)
        ab = s.groupby(["bin", "animal_id"], observed=True)["speed_m_h"].mean()
        xc = s.groupby("bin", observed=True)[c].mean()
        st_ = [mean_ci(ab.loc[b_]) if b_ in ab.index.get_level_values(0) else (np.nan,) * 4 for b_ in xc.index]
        ax[0, k].fill_between(xc.values, [v[1] for v in st_], [v[2] for v in st_], color="k", alpha=0.18, lw=0)
        ax[0, k].plot(xc.values, [v[0] for v in st_], lw=2.5, c="k")
        ax[0, k].set_ylim(bottom=0)
        ax[0, k].set_xlabel(n)
        ax[0, k].set_ylabel("средняя скорость, м/ч")
        ax[0, k].set_title(f"n={len(s)}, жив.={s['animal_id'].nunique()}")
    savefig(fig, out, "speed_vs_environment", "speed", "Средняя скорость (м/ч) по квантильным классам высоты, уклона, NDVI и снега: тонкие линии — животные, чёрная — среднее животных с 95% ДИ.", dpi)


@_fig_guard
def fig_rose_aspect(pts, out, dpi):
    if "aspect_class" not in pts or pts["aspect_class"].notna().sum() == 0:
        return
    order = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"]
    fig = plt.figure(figsize=(6, 6))
    ax = fig.add_subplot(111, projection="polar")
    for a, g in pts[pts["included_analysis"]].groupby("animal_id"):
        fr = g["aspect_class"].value_counts(normalize=True).reindex(order).fillna(0)
        th = np.radians(np.arange(0, 360, 45))
        ax.plot(np.r_[th, th[0]], np.r_[fr.values, fr.values[0]], lw=1, label=a)
    ax.set_theta_zero_location("N")
    ax.set_theta_direction(-1)
    ax.set_xticks(np.radians(np.arange(0, 360, 45)))
    ax.set_xticklabels(order)
    ax.legend(fontsize=6, loc="lower left", bbox_to_anchor=(1, 0))
    ax.set_title("Экспозиции склонов в местах фиксаций (использование, не предпочтение)")
    savefig(fig, out, "rose_aspect", "terrain", "Роза экспозиций склонов по животным.", dpi)


@_fig_guard
def fig_composition(pts, var, out, dpi, section, title, fname):
    if var not in pts or pts[var].notna().sum() == 0:
        return
    ct = pd.crosstab(pts.loc[pts["included_analysis"], "animal_id"], pts.loc[pts["included_analysis"], var], normalize="index")
    fig, ax = plt.subplots(figsize=(10, 0.35 * len(ct) + 2))
    ct.plot.barh(stacked=True, ax=ax, colormap="tab20", width=0.85)
    ax.legend(fontsize=7, bbox_to_anchor=(1, 1))
    ax.set_xlabel("доля фиксаций")
    ax.set_title(title)
    savefig(fig, out, fname, section, title + " (использование).", dpi)


@_fig_guard
def fig_profile(pts, steps, out, dpi):
    if "elevation" not in pts or pts["elevation"].notna().sum() == 0:
        return
    an = sorted(pts["animal_id"].unique())
    fig, axs = plt.subplots(len(an), 1, figsize=(13, 2.6 * len(an)), squeeze=False)
    fig.subplots_adjust(hspace=0.6)
    for k, a in enumerate(an):
        s = steps[steps["animal_id"] == a].sort_values("t_start")
        cum = s["step_length_m"].cumsum() / 1000
        ax = axs[k, 0]
        ax.plot(cum, s["start_elevation"], lw=0.6, c="sienna")
        ax.set_ylabel("м", color="sienna")
        ax2 = ax.twinx()
        ax2.plot(cum, s["speed_m_h"].rolling(12, min_periods=1).mean(), lw=0.5, c="steelblue")
        ax2.set_ylabel("м/ч", color="steelblue")
        ax2.set_ylim(bottom=0)
        ax.set_title(f"{a}: дистанция по треку — высота — скользящее среднее скорости (12 шагов)", fontsize=9, loc="left")
    axs[-1, 0].set_xlabel("накопленная дистанция, км")
    savefig(fig, out, "longitudinal_profiles", "terrain", "Продольные профили треков.", dpi)


@_fig_guard
def fig_transitions(tc, out, dpi, var="habitat"):
    if tc is None or len(tc) == 0:
        return
    d = tc[(tc["variable"] == var) & (tc["group_by"] == "all")]
    if len(d) == 0:
        return
    m = d.pivot_table(index="from", columns="to", values="n_steps", aggfunc="sum", fill_value=0)
    p = m.div(m.sum(1), axis=0)
    fig, ax = plt.subplots(figsize=(1 + 0.7 * m.shape[1], 1 + 0.6 * m.shape[0]))
    im = ax.imshow(p.values, cmap="magma_r", vmin=0, vmax=1)
    ax.set_xticks(range(m.shape[1]))
    ax.set_xticklabels(m.columns, rotation=40, ha="right", fontsize=8)
    ax.set_yticks(range(m.shape[0]))
    ax.set_yticklabels(m.index, fontsize=8)
    for i in range(m.shape[0]):
        for j in range(m.shape[1]):
            if m.values[i, j]:
                ax.text(j, i, f"{p.values[i, j]:.2f}\n({m.values[i, j]})", ha="center", va="center", fontsize=6,
                        color="white" if p.values[i, j] > 0.5 else "black")
    ax.set_xlabel("в конце шага")
    ax.set_ylabel("в начале шага")
    ax.set_title(f"Переходы {var}: P(to | from), (n шагов)")
    plt.colorbar(im, ax=ax, shrink=0.7)
    savefig(fig, out, f"transitions_{var}", "transitions", f"Матрица переходов «{var}» между началом и концом шага; частота перехода не равна предпочтению.", dpi)


def sankey_html(tc, out, var="habitat") -> Optional[str]:
    go = OPTIONAL.get("plotly_go")
    if go is None or tc is None or len(tc) == 0:
        return None
    d = tc[(tc["variable"] == var) & (tc["group_by"] == "all")].sort_values("n_steps", ascending=False).head(30)
    if len(d) == 0:
        return None
    labs_from = [f"{v} (start)" for v in d["from"]]
    labs_to = [f"{v} (end)" for v in d["to"]]
    labels = list(dict.fromkeys(labs_from + labs_to))
    idx = {l: i for i, l in enumerate(labels)}
    fig = go.Figure(go.Sankey(node=dict(label=labels, pad=12), link=dict(source=[idx[l] for l in labs_from],
                                                                          target=[idx[l] for l in labs_to], value=d["n_steps"].tolist())))
    fig.update_layout(title=f"Sankey: переходы {var} (30 наиболее частых)", font_size=11)
    p = out / "figures" / f"sankey_{var}.html"
    fig.write_html(str(p), include_plotlyjs="cdn")
    return str(p)


@_fig_guard
def fig_clusters(clusters, windows, sens, pts, out, dpi):
    if clusters is None or len(clusters) == 0:
        return
    fig, ax = plt.subplots(1, 3, figsize=(16, 4.2))
    for a, g in clusters.groupby("animal_id"):
        ax[0].hist(g["duration_h"], bins=np.arange(24, max(clusters["duration_h"].max(), 48) + 12, 12), histtype="step", label=a)
    ax[0].set_xlabel("продолжительность, ч")
    ax[0].set_title(f"Длительность кластеров (n={len(clusters)})")
    ax[0].legend(fontsize=6)
    iv = []
    for a, g in clusters.sort_values("t_start").groupby("animal_id"):
        iv += list(((g["t_start"].values[1:] - g["t_end"].values[:-1]) / np.timedelta64(1, "D")).astype(float))
    ax[1].hist(iv, bins=40, color="slateblue")
    ax[1].set_xlabel("интервал между кластерами, сут")
    ax[1].set_title("Интервалы между кластерами")
    an = sorted(clusters["animal_id"].unique())
    for k, a in enumerate(an):
        g = clusters[clusters["animal_id"] == a]
        for r in g.itertuples():
            ax[2].plot([r.t_start, r.t_end], [k, k], lw=6, c="crimson")
        pa = pts[(pts["animal_id"] == a) & pts["included_analysis"]]
        ax[2].plot([pa["timestamp_utc"].min(), pa["timestamp_utc"].max()], [k, k], lw=1, c="grey")
    ax[2].set_yticks(range(len(an)))
    ax[2].set_yticklabels(an, fontsize=7)
    ax[2].set_title("Временная диаграмма кластеров")
    fig.autofmt_xdate()
    savefig(fig, out, "clusters_overview", "clusters", "Потенциальные охотничьи/кормовые кластеры: длительность, интервалы, временная диаграмма.", dpi)
    if sens is not None and len(sens):
        tot = sens.groupby(["radius_m", "min_duration_h"])["n_clusters"].sum().unstack()
        fig, ax = plt.subplots(figsize=(5, 3.5))
        im = ax.imshow(tot.values, cmap="YlOrRd")
        ax.set_xticks(range(tot.shape[1]))
        ax.set_xticklabels([f"{int(v)} ч" for v in tot.columns])
        ax.set_yticks(range(tot.shape[0]))
        ax.set_yticklabels([f"{int(v)} м" for v in tot.index])
        for i in range(tot.shape[0]):
            for j in range(tot.shape[1]):
                ax.text(j, i, int(tot.values[i, j]), ha="center", va="center")
        ax.set_title("Чувствительность: число кластеров")
        savefig(fig, out, "clusters_sensitivity", "clusters", "Анализ чувствительности: радиус × минимальная длительность.", dpi)


@_fig_guard
def fig_daily(day, out, dpi):
    d = day[day["complete_day"]]
    if len(d) == 0:
        return
    keys = d.groupby(["sex", "region", "season"]).size().index.tolist()
    labels, av, pv = [], [], []
    for k in keys:
        g = d[(d["sex"] == k[0]) & (d["region"] == k[1]) & (d["season"] == k[2])]
        labels.append(f"{k[0]}\n{k[1]}\n{k[2]}\nжив.={g['animal_id'].nunique()}, сут.={len(g)}")
        av.append(g.groupby("animal_id")["daily_distance_km"].mean().values)
        pv.append(g["daily_distance_km"].values)
    fig, ax = plt.subplots(figsize=(max(8, 1.2 * len(keys)), 5))
    _mean_ci_plot(ax, labels, av, pv)
    ax.set_xticklabels(labels, rotation=0, ha="center", fontsize=7)
    ax.set_ylabel("суточный ход, км")
    ax.set_ylim(bottom=0)
    ax.set_title("Суточный ход по полу × региону × сезону (полные сутки): среднее ± 95% ДИ")
    savefig(fig, out, "daily_distance_groups", "daily", "Суточный ход: точка — среднее средних животных, усы — 95% ДИ (t, единица — животное; "
            "квадрат — одно животное, ДИ по суткам), красные точки — средние животных.", dpi)


@_fig_guard
def fig_post_release(prm, est, cfg, out, dpi):
    if prm is None or len(prm) == 0:
        return
    mets = [m for m in cfg["post_release"]["key_metrics"] if m in prm]
    fig, axs = plt.subplots(len(mets), 1, figsize=(12, 2.6 * len(mets)), squeeze=False)
    fig.subplots_adjust(hspace=0.35)
    for k, m in enumerate(mets):
        ax = axs[k, 0]
        for a, g in prm.groupby("animal_id"):
            g = g.assign(_d=np.floor(pd.to_numeric(g["days_since_release"], errors="coerce"))).dropna(subset=["_d"])
            v = g[m].where(g["complete_day"].astype(bool)) if m != "explored_area_km2" else g[m]
            ser = pd.Series(v.values, index=g["_d"].astype(int).values).groupby(level=0).mean()
            if len(ser) == 0:
                continue
            ser = ser.reindex(range(int(ser.index.min()), int(ser.index.max()) + 1))
            l, = ax.plot(ser.index, ser.rolling(7, min_periods=3).mean(), lw=1, label=a)
            if est is not None and len(est):
                e = est[(est["animal_id"] == a) & (est["metric"] == m) & (est["confirm_days"] == cfg["post_release"]["confirm_days"])]
                if len(e) and np.isfinite(e["stabilization_day"].iloc[0]):
                    ax.axvline(e["stabilization_day"].iloc[0], c=l.get_color(), ls="--", lw=0.8)
        for b in cfg["post_release"]["bins_days"][1:-1]:
            ax.axvline(b, c="grey", ls=":", lw=0.6)
        ax.set_ylabel(m, fontsize=8)
        ax.set_xlim(0, min(cfg["post_release"]["max_days"], prm["days_since_release"].max()))
    axs[0, 0].legend(fontsize=6, ncol=6)
    axs[-1, 0].set_xlabel("сутки после выпуска")
    savefig(fig, out, "post_release_curves", "post_release", "Ежедневные метрики после выпуска (скользящее среднее 7 сут); штрих — оценка стабилизации, точки — границы недельных интервалов.", dpi)


@_fig_guard
def fig_typology(typo, out, dpi):
    if typo is None:
        return
    clus, sim, coas, load, info = typo
    from scipy.cluster.hierarchy import dendrogram
    fig, ax = plt.subplots(1, 3, figsize=(17, 5))
    dendrogram(info["linkage"], labels=clus["animal_id"].tolist(), ax=ax[0], leaf_rotation=90)
    ax[0].set_title("Дендрограмма (Gower, average)")
    im = ax[1].imshow(sim.values, cmap="viridis")
    ax[1].set_xticks(range(len(sim)))
    ax[1].set_xticklabels(sim.columns, rotation=90, fontsize=7)
    ax[1].set_yticks(range(len(sim)))
    ax[1].set_yticklabels(sim.index, fontsize=7)
    ax[1].set_title("Сходство животных (1 − Gower)")
    plt.colorbar(im, ax=ax[1], shrink=0.7)
    for r in clus.itertuples():
        ax[2].scatter(r.PC1, r.PC2, c=[COLORS[int(r.group) % 20]], s=50, marker="o" if r.sex == "F" else "s")
        ax[2].annotate(r.animal_id, (r.PC1, r.PC2), fontsize=7)
    ev = info["explained"]
    ax[2].set_xlabel(f"PC1 ({ev[0] * 100:.0f}%)")
    ax[2].set_ylabel(f"PC2 ({ev[1] * 100:.0f}%)" if len(ev) > 1 else "PC2")
    ax[2].set_title("PCA индивидуальных признаков (○ самки, □ самцы/прочие)")
    savefig(fig, out, "individual_typology", "typology", "Исследовательская типология животных (единица — животное).", dpi)


@_fig_guard
def fig_timeseries(day, pts, out, dpi):
    cols = [c for c in ("modis_ndvi_16d", "ndvi_nearest", "modis_snowcover_mean_d1", "snow_depth_m") if c in pts]
    if not cols:
        return
    an = sorted(day["animal_id"].unique())
    fig, axs = plt.subplots(len(an), 1, figsize=(13, 2.6 * len(an)), squeeze=False)
    fig.subplots_adjust(hspace=0.6)
    for k, a in enumerate(an):
        ax = axs[k, 0]
        p = pts[(pts["animal_id"] == a) & pts["included_analysis"]].groupby("local_date")[cols].mean()
        d = day[day["animal_id"] == a].set_index("local_date")
        x = pd.to_datetime(pd.Series(p.index.astype(str)))
        if "modis_ndvi_16d" in p:
            ax.plot(x, p["modis_ndvi_16d"], c="green", lw=1, label="NDVI MODIS")
        if "modis_snowcover_mean_d1" in p:
            ax.plot(x, p["modis_snowcover_mean_d1"] / 100, c="deepskyblue", lw=1, label="снег MODIS/100")
        ax2 = ax.twinx()
        ax2.plot(pd.to_datetime(pd.Series(d.index.astype(str))), d["daily_distance_km"].rolling(7, min_periods=2).mean(), c="k", lw=0.8, label="суточный ход, км")
        ax.set_title(a, fontsize=9, loc="left", fontweight="bold")
        if k == 0:
            ax.legend(fontsize=6, loc="upper left")
            ax2.legend(fontsize=6, loc="upper right")
    savefig(fig, out, "timeseries_ndvi_snow_speed", "ndvi", "Временные ряды NDVI, снега и суточного хода по животным.", dpi)


@_fig_guard
def fig_coefs(coef, out, dpi, title, fname, section, term_col="term", xlabel=None):
    if coef is None or len(coef) == 0:
        return
    c = coef.copy()
    if "animal_id" in c:
        c = c[c["animal_id"].astype(str).str.startswith("POOLED")] if (c["animal_id"].astype(str).str.startswith("POOLED")).any() else c
    if "model" in c:
        c = c[c["model"] == c["model"].iloc[0]]
    c = c[c[term_col] != "Intercept"]
    fig, ax = plt.subplots(figsize=(7, 0.35 * len(c) + 1.5))
    y = np.arange(len(c))
    ax.errorbar(c["coef"], y, xerr=[c["coef"] - c["ci95_low"], c["ci95_high"] - c["coef"]], fmt="o", c="k")
    ax.axvline(0, c="grey", ls=":")
    ax.set_yticks(y)
    ax.set_yticklabels(c[term_col], fontsize=7)
    if xlabel:
        ax.set_xlabel(xlabel)
    ax.set_title(title)
    savefig(fig, out, fname, section, title + " (95% ДИ).", dpi)


# -------------------------------------- карта Folium ------------------------------------------
def _coarsen(lon, lat, cfg):
    x, y = to_work_xy(lon, lat, cfg["crs"]["work_crs"])
    r = cfg["privacy"]["round_m"]
    return from_work_xy(np.round(np.asarray(x) / r) * r, np.round(np.asarray(y) / r) * r, cfg["crs"]["work_crs"])


def build_map(pts, steps, clusters, meta, cfg, out: Path, public: bool) -> Optional[str]:
    if folium is None:
        warn("folium не установлен — интерактивная карта не создана (pip install folium)")
        return None
    inc = pts[pts["included_analysis"]].copy()
    if public:
        lo, la = _coarsen(inc["longitude"].values, inc["latitude"].values, cfg)
        inc["longitude"], inc["latitude"] = lo, la
        try:
            rx, ry = to_work_xy(pd.to_numeric(meta["release_lon"], errors="coerce"), pd.to_numeric(meta["release_lat"], errors="coerce"), cfg["crs"]["work_crs"])
            mask = np.zeros(len(inc), bool)
            for a, xx, yy in zip(meta["animal_id"], rx, ry):
                sel = (inc["animal_id"] == a).values & (np.hypot(inc["x"].values - xx, inc["y"].values - yy) < cfg["privacy"]["mask_release_radius_m"])
                mask |= sel
            if cfg["privacy"]["mask_clusters"]:
                mask |= (inc["behavior_mode"] == "hunting_candidate").values
            inc = inc[~mask]
        except Exception:  # noqa: BLE001
            pass
    m = folium.Map(location=[inc["latitude"].mean(), inc["longitude"].mean()], zoom_start=9, tiles=None, control_scale=True)
    folium.TileLayer("OpenStreetMap", name="OSM").add_to(m)
    folium.TileLayer("https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png", attr="OpenTopoMap", name="Рельеф (OpenTopoMap)").add_to(m)
    folium.TileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", attr="Esri", name="Снимки Esri").add_to(m)
    col = _acolor(inc["animal_id"].unique())
    hexc = {a: matplotlib.colors.to_hex(c) for a, c in col.items()}
    for a, g in inc.groupby("animal_id"):
        fg = folium.FeatureGroup(name=f"Трек: {a}", show=True)
        for _, gb in g.groupby("burst_id"):
            if len(gb) > 1:
                folium.PolyLine(list(zip(gb["latitude"].round(5), gb["longitude"].round(5))), color=hexc[a], weight=1.5, opacity=0.8).add_to(fg)
        fg.add_to(m)
    nmax = cfg["report"]["map_max_points"]
    sp = inc.sample(min(nmax, len(inc)), random_state=cfg["project"]["seed"]) if len(inc) > nmax else inc
    def _fmt(v, f="{:.0f}"):
        try:
            return f.format(float(v)) if np.isfinite(float(v)) else ""
        except (TypeError, ValueError):
            return "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v)

    props_cols = [("animal_id", "животное", None), ("t_local", "время (местн.)", None), ("season", "сезон", None),
                  ("behavior_mode", "режим", None), ("solar_period_A", "период A", None), ("elevation", "высота, м", "{:.0f}"),
                  ("habitat", "местообитание", None), ("ndvi_nearest", "NDVI", "{:.2f}"), ("snow_present", "снег", None)]
    sp = sp.copy()
    sp["t_local"] = pd.to_datetime(sp["timestamp_local"]).dt.strftime("%Y-%m-%d %H:%M") if "timestamp_local" in sp else ""
    for lay, colname in (("Сезон", "season"), ("Режим", "behavior_mode"), ("Период суток A", "solar_period_A")):
        if colname not in sp:
            continue
        for val, g in sp.groupby(colname):
            feats = []
            for r in g.to_dict("records"):
                pr = {c: _fmt(r.get(c), f) if f else _fmt(r.get(c)) for c, lab, f in props_cols if c in r}
                pr["_c"] = hexc[r["animal_id"]]
                feats.append({"type": "Feature", "properties": pr,
                              "geometry": {"type": "Point", "coordinates": [round(float(r["longitude"]), 5), round(float(r["latitude"]), 5)]}})
            flds = [c for c, lab, f in props_cols if c in sp]
            labs = [lab for c, lab, f in props_cols if c in sp]
            gj = folium.GeoJson({"type": "FeatureCollection", "features": feats}, name=f"{lay}: {val} (n={len(g)})", show=False,
                                marker=folium.CircleMarker(radius=2.5, fill=True, weight=0, fill_opacity=0.85),
                                style_function=lambda f: {"fillColor": f["properties"]["_c"], "color": f["properties"]["_c"]},
                                popup=folium.GeoJsonPopup(fields=flds, aliases=labs, labels=True))
            gj.add_to(m)
    if not public and clusters is not None and len(clusters):
        lon, lat = from_work_xy(clusters["medoid_x"], clusters["medoid_y"], cfg["crs"]["work_crs"])
        fg = folium.FeatureGroup(name=f"Потенциальные охотничьи/кормовые кластеры (n={len(clusters)})", show=True)
        for r, lo_, la_ in zip(clusters.itertuples(), lon, lat):
            folium.Circle([la_, lo_], radius=cfg["clusters"]["radius_m"], color="red", fill=True, fill_opacity=0.3,
                          popup=folium.Popup(f"<b>{r.cluster_id}</b><br>{r.t_start:%Y-%m-%d %H:%M} UTC<br>{r.duration_h:.0f} ч, "
                                             f"{r.n_fixes} фикс.<br>p95={r.p95_distance_m:.0f} м<br>не подтверждён полевой проверкой", max_width=250)).add_to(fg)
        fg.add_to(m)
        fg2 = folium.FeatureGroup(name="Места выпуска", show=True)
        for r in meta.itertuples():
            try:
                folium.Marker([float(r.release_lat), float(r.release_lon)], tooltip=f"Выпуск: {r.animal_id} {r.release_datetime_utc}",
                              icon=folium.Icon(color="black", icon="star")).add_to(fg2)
            except Exception:  # noqa: BLE001
                pass
        fg2.add_to(m)
    folium.LayerControl(collapsed=False).add_to(m)
    p = out / "maps" / ("leotrack_map_public.html" if public else "leotrack_map.html")
    p.parent.mkdir(parents=True, exist_ok=True)
    m.save(str(p))
    return str(p)


# -------------------------------------- HTML-отчёт --------------------------------------------
def _tbl(df: Optional[pd.DataFrame], n: int, cols: Optional[List[str]] = None) -> str:
    if df is None or not isinstance(df, pd.DataFrame) or len(df) == 0:
        return "<p class='na'>Нет данных для таблицы.</p>"
    d = df[cols] if cols else df
    d = _excel_safe(d.head(n))
    more = f"<p class='na'>Показано {min(n, len(df))} из {len(df)} строк; полная таблица — в tables/.</p>" if len(df) > n else ""
    return d.to_html(index=False, float_format=lambda v: f"{v:.3g}", na_rep="", border=0, classes="t") + more


def _img(path: str) -> str:
    with open(path, "rb") as fh:
        b = base64.b64encode(fh.read()).decode()
    return f"<img src='data:image/png;base64,{b}'/>"


def build_report(ctx: dict, out: Path, cfg: dict) -> str:
    T = ctx["tables"]
    pts = T["points_cleaned"]
    inc = pts[pts["included_analysis"]]
    steps = T["steps_enriched"]
    n_an = inc["animal_id"].nunique()
    nmax = cfg["report"]["max_table_rows"]
    figs_by = {}
    for sec, p, cap in FIGS:
        figs_by.setdefault(sec, []).append((p, cap))

    def figs(sec):
        return "".join(f"<figure>{_img(p)}<figcaption>{cap}</figcaption></figure>" for p, cap in figs_by.get(sec, []))

    sx = ctx["meta"]["sex"].value_counts().to_dict()
    gee_on = ctx.get("gee_ok", False)
    secs = []
    secs.append(("1. Цели и ограничения анализа", f"""
<p>Комплексная характеристика перемещений переднеазиатского леопарда по GPS-трекам: скорость и активность, рельеф,
местообитания, NDVI, снег, переходы между условиями, потенциальные охотничьи/кормовые кластеры, суточный ход, адаптация после выпуска,
доступность и предпочтение (iSSF), индивидуальная типология.</p>
<ul><li><b>Наблюдаемый паттерн</b> — описательные статистики использования; <b>результат модели</b> — iSSF, смешанные модели;
<b>биологическая интерпретация</b> — только с учётом размера выборки и без причинных выводов из корреляций.</li>
<li>Скорость — <i>расчётная скорость перемещения между GPS-фиксациями</i>, т.е. минимальная средняя скорость прямолинейного перемещения.</li>
<li>Скорость приводится в абсолютных единицах (м/ч), без логарифмирования. Основная сводная статистика — среднее; 95% ДИ рассчитывается
по t-распределению, единица повторности — животное (среднее средних животных); если в группе одно животное, ДИ дан по шагам/суткам
и носит описательный характер. Медианы оставлены в таблицах как дополнительная характеристика.</li>
<li><code>hunting_candidate</code> — операционная метка: пребывание ≥{cfg['clusters']['min_duration_h']} ч, p{int(cfg['clusters']['containment'] * 100)} расстояний до центра ≤{cfg['clusters']['radius_m']} м.
Без полевой проверки это <b>потенциальный охотничий или кормовой кластер</b>, а не подтверждённое место добычи (кластеры могут отражать отдых, лёжки и др.).</li>
<li>Частота фиксаций в классе ≠ предпочтение; предпочтение оценивается сравнением с доступными альтернативами (раздел 15).</li>
<li>Животных: {n_an}; пол: {sx}. При малом числе особей сравнения полов и регионов — описательные/предварительные.</li></ul>"""))
    fc = T.get("field_completeness")
    secs.append(("2. Данные и выборка", f"""<p>Файлов: {len(T['input_files'])}; записей: {len(pts)}; валидных фиксаций: {len(inc)};
шагов: {len(steps)}; животно-суток: {len(T['daily_movement'])}, из них полных: {int(T['daily_movement']['complete_day'].sum())}.</p>
{_tbl(T['input_files'], nmax, ['file', 'n_records', 'animals', 'time_rule', 'sha256'])}
<h4>Метаданные животных</h4>{_tbl(T['animals_metadata'], nmax)}
<h4>Сводка по животным</h4>{_tbl(T['individual_summary'], nmax)}<h4>Полнота полей</h4>{_tbl(fc, nmax)}{figs('data')}"""))
    secs.append(("3. Контроль качества", f"""<p>Исходные записи не удалялись: каждая имеет <code>qc_flag</code>, <code>qc_reason</code>,
<code>included_analysis</code>. Исключено записей: {int((~pts['included_analysis']).sum())} ({(~pts['included_analysis']).mean() * 100:.1f}%).
Порог скорости {cfg['qc']['max_speed_km_h']} км/ч (при интервале ≥{cfg['qc']['max_speed_min_dt_h']} ч), «выброс-возврат»: >{cfg['qc']['spike_speed_km_h']} км/ч
и разворот >{cfg['qc']['spike_angle_deg']}°; разрыв burst >{cfg['qc']['burst_max_gap_h']} ч; записи с одинаковым временем и разными
координатами — политика <code>{cfg['qc']['duplicate_policy']}</code>.</p>
{_tbl(T.get('qc_summary'), nmax)}{figs('qc')}"""))
    secs.append(("4. Пространственно-временная характеристика движения", f"""{_tbl(T['individual_summary'], nmax)}
<p>Номинальные интервалы фиксации (мода, ч): {ctx['nominal']}. Регуляризация для сопоставимости: {cfg['regularization']['interval_h']} ч,
допуск ±{cfg['regularization']['tolerance_min']} мин, без интерполяции; регулярных шагов: {len(ctx.get('reg_steps', []))}.</p>"""))
    secs.append(("5. Скорость и активность", f"""<p>Две параллельные классификации времени суток: A — рабочие периоды зоологов
({', '.join(p['name'] + ' ' + p['start'] + '–' + p['end'] for p in cfg['solar']['periods_A'])}) как проверяемая гипотеза;
B — астрономическая (высота Солнца: день > {cfg['solar']['sunrise_deg']}°, гражданские сумерки до {cfg['solar']['civil_twilight_deg']}°, ночь).</p>
{_tbl(T['solar_period_summary'], nmax)}{figs('speed')}<h4>Скорость по условиям</h4>{_tbl(T.get('speed_by_conditions'), nmax)}
<h4>Модели скорости</h4><p>Средняя скорость моделируется в м/ч без преобразования отклика: GEE (тождественная связь, кластер — животное,
обменная корреляция, робастные SE); коэффициенты — разности средних скоростей (м/ч) относительно базового уровня.</p>{_tbl(T.get('speed_models'), nmax)}"""))
    secs.append(("6. Рельеф", (f"""<p>Copernicus DEM GLO-30 (DSM, пропуски — NASADEM), рабочая CRS {cfg['crs']['work_crs']}, 30 м. Уклон и экспозиция —
Horn (3×3); TRI — Riley; TPI — z минус средняя высота в круге r={cfg['terrain']['tpi_small_radius_m']} и {cfg['terrain']['tpi_large_radius_m']} м;
формы рельефа — Weiss (2001) по стандартизованному TPI крупного радиуса (SD={ctx.get('tpi_sd', float('nan')):.1f} м) и уклону ≤{cfg['terrain']['flat_slope_deg']}°.
Значения извлекаются в начале, конце и вдоль шага (шаг дискретизации {cfg['steps']['path_sample_spacing_m']} м), а также в буферах {cfg['terrain']['buffer_radii_m']} м.</p>
{_tbl(T.get('terrain_use'), nmax)}{figs('terrain')}""" if gee_on else "<p class='na'>Модуль не выполнен: нет подключения к GEE.</p>")))
    secs.append(("7. Местообитания", (f"""<p>Два независимых уровня: стабильная карта ESA WorldCover 2021 (10 м) с переклассификацией по таблице
<code>habitat_reclass</code> и динамический покров Dynamic World V1 (вероятности 9 классов в окнах ±{cfg['gee']['dw_windows_days']} сут).
Глобальные классы не выдаются за региональную типологию растительности; при наличии региональной карты она имеет приоритет.</p>
{_tbl(T.get('habitat_use'), nmax)}{figs('habitat')}""" if gee_on else "<p class='na'>Модуль не выполнен: нет подключения к GEE.</p>")))
    secs.append(("8. NDVI", (f"""<p>Sentinel-2 SR Harmonized (маска SCL), для дат до {cfg['gee']['landsat_before']} — Landsat C2 L2 (поле <code>ndvi_sensor</code>);
MODIS MOD13Q1 — согласованный многолетний ряд. Облачные пиксели маскируются, а не превращаются в низкий NDVI.</p>
{_tbl(T.get('ndvi_snow_summary'), nmax)}{figs('ndvi')}""" if gee_on else "<p class='na'>Модуль не выполнен: нет подключения к GEE.</p>")))
    secs.append(("9. Снег", ("""<p>MOD10A1.061 (500 м, QA≤2; облачный пиксель → <code>unknown</code>, не «нет снега»), Sentinel-2 NDSI (при отсутствии MODIS ≤3 сут),
ERA5-Land (глубина и SWE ~9–11 км — только региональный фон).</p>""" + figs("snow") if gee_on else "<p class='na'>Модуль не выполнен: нет подключения к GEE.</p>")))
    secs.append(("10. Переходы между условиями", f"""<p>Матрицы «начало → конец шага»; вероятности нормированы по исходному состоянию.
Частые переходы не называются предпочитаемыми без оценки доступности.</p>{figs('transitions')}
{('<p>Sankey: <a href="../figures/' + Path(ctx['sankey']).name + '">интерактивная диаграмма</a></p>') if ctx.get('sankey') else ''}
{_tbl(T.get('transition_counts'), nmax)}"""))
    cl = T.get("hunting_clusters")
    secs.append(("11. Потенциальные охотничьи кластеры и чувствительность", f"""<p>Критерий включения: p{int(cfg['clusters']['containment'] * 100)} расстояний
фиксаций до центра ≤ {cfg['clusters']['radius_m']} м, длительность ≥ {cfg['clusters']['min_duration_h']} ч, пропуск ≤ {cfg['clusters']['allowed_gap_h']} ч,
экскурсия с возвращением ≤ {cfg['clusters']['max_excursion_h']} ч. Неопределённость центра — бутстреп ({cfg['clusters']['bootstrap']} повторов),
95% ДИ координат и площадь 95% доверительного эллипса (это характеристика надёжности центра, а не радиус кластера).
Найдено кластеров: {0 if cl is None else len(cl)}.</p>{_tbl(cl, nmax)}<h4>Чувствительность (радиус × длительность; влияние интервала фиксации)</h4>
{_tbl(T.get('cluster_sensitivity'), nmax)}{figs('clusters')}"""))
    bm = steps["behavior_mode"].value_counts().to_dict()
    secs.append(("12. Разделение трека на movement и hunting_candidate", f"""<p>Шаги: {bm}. Шаг между кластером и движением — <code>transition</code>
(не входит ни в одну «чистую» группу).</p>{_tbl(T.get('hmm_parameters'), nmax)}"""))
    secs.append(("13. Суточный ход по полу, региону и сезону", f"""<p>Суточный ход — сумма валидных шагов, разрезанных по местной полуночи; полные
сутки: покрытие ≥ {cfg['daily']['min_coverage']} и доля ожидаемых фиксаций ≥ {cfg['daily']['min_fix_fraction']}. ДИ — бутстреп по животным.
Сезоны: {ctx.get('season_note', '')}.</p>{_tbl(T.get('daily_movement_summary'), nmax)}<h4>Смешанная модель</h4>{_tbl(T.get('daily_models'), nmax)}{figs('daily')}"""))
    secs.append(("14. Изменения после выпуска и стабилизация", f"""<p>Интервалы: {cfg['post_release']['bin_labels']} сут + непрерывный анализ.
Стабилизация — первая дата после последней точки изменения (бинарная сегментация, BIC), после которой метрика не имеет значимого тренда
и остаётся в устойчивом диапазоне ≥{cfg['post_release']['confirm_days']} сут; 95% интервал — блочный бутстреп; чувствительность {cfg['post_release']['sensitivity_confirm_days']} сут.
Даты выпуска с <code>release_source=first_fix_assumed</code> требуют подтверждения.</p>{_tbl(T.get('stabilization_estimates'), nmax)}{figs('post_release')}"""))
    secs.append(("15. Доступность и предпочтение", f"""<p>Основной метод — iSSF (условная логистическая регрессия, страта = шаг, {cfg['issf']['n_available']}
доступных шагов из той же точки по гамма/фон Мизес; члены движения sl, log(sl), cos(ta)); индивидуальные модели и двухэтапное объединение
(DerSimonian–Laird). Валидация — блоки по времени. Дополнительно — коэффициенты избирательности Манли (доступность = MCP + 250 м).</p>
{_tbl(T.get('issf_coefficients'), nmax)}{figs('issf')}<h4>Коэффициенты избирательности</h4>{_tbl(T.get('selection_ratios'), nmax)}"""))
    secs.append(("16. Различия между группами", f"""<p>Сравнения на уровне средних животного (не точек).</p>{_tbl(T.get('group_comparisons'), nmax)}"""))
    secs.append(("17. Индивидуальные профили и типология", f"""<p>{ctx.get('typology_note', '')}</p>{_tbl(T.get('individual_features'), nmax)}
{_tbl(T.get('individual_clusters'), nmax)}{figs('typology')}"""))
    secs.append(("18. Неопределённость, смещения и ограничения", "<ul>" + "".join(f"<li>{w}</li>" for w in dict.fromkeys(WARNINGS_FOR_REPORT)) +
                 "<li>Интервалы фиксации различаются между животными; для сравнений используйте регуляризованные шаги и анализ чувствительности.</li>"
                 "<li>Copernicus DEM — модель поверхности (DSM), в лесу завышает высоту земли на высоту полога.</li></ul>"))
    secs.append(("19. Техническое приложение", f"""<p>LeoTrack {__version__} (sha256 скрипта {script_hash()}), Python {platform.python_version()},
{platform.platform()}; хэш конфигурации {hash_obj(cfg)}; запуск {ctx['run_time']}.</p>{_tbl(T.get('data_provenance'), 100)}
<h4>Конфигурация</h4><pre>{yaml.safe_dump(cfg, allow_unicode=True, sort_keys=False)}</pre>"""))
    css = """body{font-family:Segoe UI,Arial,sans-serif;max-width:1250px;margin:auto;padding:20px;color:#222;line-height:1.45}
h1{color:#5a3e1b}h2{border-bottom:2px solid #c8a46b;padding-bottom:4px;margin-top:36px;color:#5a3e1b}
table.t{border-collapse:collapse;font-size:11px;margin:8px 0;display:block;overflow-x:auto}table.t th{background:#f1e7d6;position:sticky;top:0}
table.t td,table.t th{border:1px solid #ddd;padding:3px 6px;white-space:nowrap}img{max-width:100%}figure{margin:14px 0}
figcaption{font-size:12px;color:#555}.na{color:#888;font-size:12px}pre{font-size:10px;background:#f7f7f7;padding:8px;overflow-x:auto}
nav a{display:block;font-size:13px}"""
    nav = "".join(f"<a href='#s{i}'>{t}</a>" for i, (t, _) in enumerate(secs))
    body = "".join(f"<h2 id='s{i}'>{t}</h2>{c}" for i, (t, c) in enumerate(secs))
    mp = f"<p>Интерактивная карта: <a href='../maps/{Path(ctx['map']).name}'>{Path(ctx['map']).name}</a></p>" if ctx.get("map") else ""
    html = f"""<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>LeoTrack — отчёт</title><style>{css}</style></head><body>
<h1>Анализ GPS-треков переднеазиатского леопарда — LeoTrack</h1><p>Сформировано {ctx['run_time']}. Животных: {n_an}; валидных фиксаций: {len(inc)};
шагов: {len(steps)}.{' <b>Публичный режим: координаты огрублены до ' + str(cfg['privacy']['round_m']) + ' м.</b>' if cfg['privacy']['public_mode'] else ''}</p>
{mp}<nav>{nav}</nav>{body}</body></html>"""
    p = out / "reports" / "LeoTrack_report.html"
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(html, encoding="utf-8")
    return str(p)


# ==============================================================================================
# СВОДНЫЕ ТАБЛИЦЫ
# ==============================================================================================
def build_summaries(ctx: dict, cfg: dict) -> Dict[str, pd.DataFrame]:
    pts, steps, day, meta = ctx["pts"], ctx["steps"], ctx["day"], ctx["meta"]
    inc = pts[pts["included_analysis"]]
    cl = ctx.get("clusters", pd.DataFrame())
    S: Dict[str, pd.DataFrame] = {}
    rows = []
    for a, g in pts.groupby("animal_id"):
        gi = g[g["included_analysis"]]
        st = steps[steps["animal_id"] == a]
        d = day[day["animal_id"] == a]
        m = meta.set_index("animal_id").loc[a]
        rows.append({"animal_id": a, "sex": m["sex"], "region": m["region"], "status": m["status"],
                     "release_datetime_utc": m["release_datetime_utc"], "release_source": m["release_source"],
                     "first_fix_utc": gi["timestamp_utc"].min(), "last_fix_utc": gi["timestamp_utc"].max(),
                     "tracking_days": (gi["timestamp_utc"].max() - gi["timestamp_utc"].min()).total_seconds() / 86400,
                     "n_records": len(g), "n_valid": len(gi), "n_excluded": int((~g["included_analysis"]).sum()),
                     "n_bursts": gi["burst_id"].nunique(), "nominal_interval_h": ctx["nominal"].get(a),
                     "n_steps": len(st), "total_distance_km": st["step_length_m"].sum() / 1000,
                     **dict(zip(("mean_speed_m_h", "mean_speed_ci95_low", "mean_speed_ci95_high"), mean_ci(st["speed_m_h"])[:3])),
                     "median_speed_m_h": st["speed_m_h"].median(), "q90_speed_m_h": st["speed_m_h"].quantile(.9),
                     "n_days": len(d), "n_complete_days": int(d["complete_day"].sum()),
                     **dict(zip(("mean_daily_km_complete", "mean_daily_km_ci95_low", "mean_daily_km_ci95_high"),
                                mean_ci(d.loc[d["complete_day"], "daily_distance_km"])[:3])),
                     "median_daily_km_complete": d.loc[d["complete_day"], "daily_distance_km"].median(),
                     "n_hunting_candidate_clusters": int((cl["animal_id"] == a).sum()) if len(cl) else 0,
                     "elevation_median_m": gi["elevation"].median() if "elevation" in gi else np.nan})
    S["individual_summary"] = pd.DataFrame(rows)
    ind = S["individual_summary"]
    S["sex_summary"] = ind.groupby("sex").agg(n_animals=("animal_id", "nunique"), n_valid_fixes=("n_valid", "sum"), n_steps=("n_steps", "sum"),
                                              n_complete_days=("n_complete_days", "sum"),
                                              mean_of_animal_mean_daily_km=("mean_daily_km_complete", "mean"),
                                              mean_of_animal_mean_speed_m_h=("mean_speed_m_h", "mean"),
                                              animals=("animal_id", lambda s: ", ".join(s))).reset_index()
    dd = day[day["complete_day"]]
    rows = []
    for sea, g in dd.groupby("season"):
        am = g.groupby("animal_id")["daily_distance_km"].mean()
        m, lo, hi, n = mean_ci(am)
        sg = steps[steps["season"] == sea]
        sm_, slo, shi, _ = mean_ci(sg.groupby("animal_id")["speed_m_h"].mean())
        rows.append({"season": sea, "n_animals": n, "n_animal_days": len(g), "mean_daily_km": m, "mean_daily_km_ci95_low": lo,
                     "mean_daily_km_ci95_high": hi, "n_steps": len(sg), "mean_speed_m_h": sm_, "mean_speed_ci95_low": slo,
                     "mean_speed_ci95_high": shi, "ci_unit": "животное (среднее животного)"})
    S["season_summary"] = pd.DataFrame(rows)
    rows = []
    for cls, col in (("A_working", "start_solar_period_A"), ("B_astronomical", "start_solar_period_B")):
        for per, g in steps.groupby(col):
            m, lo, hi, _ = mean_ci(g.groupby("animal_id")["speed_m_h"].mean())
            rows.append({"classification": cls, "period": per, "n_animals": g["animal_id"].nunique(), "n_steps": len(g),
                         "mean_speed_m_h": m, "mean_speed_ci95_low": lo, "mean_speed_ci95_high": hi,
                         "median_speed_m_h": g["speed_m_h"].median(), "mean_step_m": g["step_length_m"].mean(),
                         "frac_hunting_candidate": (g["behavior_mode"] == "hunting_candidate").mean(),
                         "frac_directed": (g["behavior_simple"] == "directed_movement").mean()})
    S["solar_period_summary"] = pd.DataFrame(rows)
    if "habitat" in inc and inc["habitat"].notna().any():
        hu = inc.groupby(["animal_id", "habitat"]).size().rename("n_fixes").reset_index()
        hu["frac"] = hu["n_fixes"] / hu.groupby("animal_id")["n_fixes"].transform("sum")
        allh = hu.groupby("habitat").agg(n_fixes=("n_fixes", "sum"), n_animals=("animal_id", "nunique"), mean_animal_frac=("frac", "mean")).reset_index()
        allh["animal_id"] = "ALL (среднее долей животных)"
        S["habitat_use"] = pd.concat([hu, allh], ignore_index=True)
        rows = []
        for var in ("landform", "aspect_class", "elevation_band"):
            if var in inc:
                t = inc.groupby(["animal_id", var]).size().rename("n_fixes").reset_index()
                t["frac"] = t["n_fixes"] / t.groupby("animal_id")["n_fixes"].transform("sum")
                t = t.rename(columns={var: "class"})
                t["variable"] = var
                rows.append(t)
        cont = []
        for a, g in inc.groupby("animal_id"):
            for v in ("elevation", "slope_deg", "tri", "tpi_large", "curvature"):
                if v in g:
                    cont.append({"animal_id": a, "variable": v, "class": "continuous", "n_fixes": int(g[v].notna().sum()),
                                 "median": g[v].median(), "q25": g[v].quantile(.25), "q75": g[v].quantile(.75)})
        S["terrain_use"] = pd.concat(rows + [pd.DataFrame(cont)], ignore_index=True)
    if "ndvi_nearest" in inc or "snow_present" in inc:
        rows = []
        for (a, s_), g in inc.groupby(["animal_id", "season"]):
            r = {"animal_id": a, "season": s_, "n_fixes": len(g)}
            for c in ("ndvi_nearest", "modis_ndvi_16d", "ndsi_nearest", "snow_depth_m", "modis_snowcover_mean_d7"):
                if c in g:
                    r[c + "_median"] = g[c].median()
                    r[c + "_valid_frac"] = g[c].notna().mean()
            if "snow_present" in g:
                vc = g["snow_present"].value_counts(normalize=True)
                for k in ("snow", "no_snow", "unknown"):
                    r[f"snow_{k}_frac"] = vc.get(k, 0.0)
            if "ndvi_sensor" in g:
                r["ndvi_sensors"] = ", ".join(f"{k}:{v}" for k, v in g["ndvi_sensor"].value_counts().items())
            rows.append(r)
        S["ndvi_snow_summary"] = pd.DataFrame(rows)
    rows = []
    for var in ("start_habitat", "start_landform", "start_aspect_class", "start_snow_present", "start_dw_top_class", "season",
                "start_solar_period_A", "start_solar_period_B", "behavior_mode", "behavior_simple", "sex", "region"):
        if var not in steps:
            continue
        for k, g in steps.groupby(var):
            m, lo, hi, na = mean_ci(g.groupby("animal_id")["speed_m_h"].mean())
            ms, slo, shi, _ = mean_ci(g["speed_m_h"])
            rows.append({"variable": var, "class": k, "n_steps": len(g), "n_animals": na,
                         "mean_speed_m_h": m, "mean_speed_ci95_low": lo, "mean_speed_ci95_high": hi,
                         "pooled_mean_speed_m_h": ms, "pooled_ci95_low": slo, "pooled_ci95_high": shi,
                         "median_speed_m_h": g["speed_m_h"].median(), "mean_step_m": g["step_length_m"].mean()})
    S["speed_by_conditions"] = pd.DataFrame(rows)
    rows = []
    for a, g in pts.groupby("animal_id"):
        gi = g[g["included_analysis"]]
        span_h = (gi["timestamp_utc"].max() - gi["timestamp_utc"].min()).total_seconds() / 3600
        nom = ctx["nominal"].get(a, np.nan)
        rr = g.loc[~g["included_analysis"], "qc_reason"].str.split(";").explode().value_counts()
        rows.append({"animal_id": a, "n_records": len(g), "n_valid": len(gi), "excluded_pct": 100 * (1 - len(gi) / max(len(g), 1)),
                     "fix_success_rate": len(gi) / (span_h / nom + 1) if nom and np.isfinite(nom) else np.nan,
                     "median_interval_h": gi["dt_prev_s"].median() / 3600, "p95_interval_h": gi["dt_prev_s"].quantile(.95) / 3600,
                     "n_gaps_gt_burst": gi["burst_id"].nunique() - 1, **{f"reason_{k}": v for k, v in rr.items()}})
    S["qc_summary"] = pd.DataFrame(rows)
    return S


# ==============================================================================================
# КОНВЕЙЕР
# ==============================================================================================
def _cfg_part(cfg, *keys):
    return {k: cfg[k] for k in keys}


def stage_preprocess(ctx: dict, cfg: dict) -> dict:
    LOG.info("=== Этап preprocess: импорт, QC, шаги, время суток ===")
    pts, raw, finfo = load_inputs(ctx["files"], cfg)
    ctx["raw"], ctx["input_files"] = raw, finfo
    meta = ctx["meta"]
    m = meta.set_index("animal_id")
    pts["sex"] = pts["animal_id"].map(m["sex"]).fillna(pts["sex"]).fillna("unknown")
    pts["region"] = pts["animal_id"].map(m["region"]).fillna("unknown")
    pts["release_date"] = pts["animal_id"].map(m["release_datetime_utc"])
    pts["age_class"] = pts["animal_id"].map(m["age_class"]).fillna("unknown")
    pts = run_qc(pts, cfg)
    ok = pts["latitude"].notna() & pts["longitude"].notna()
    pts["x"], pts["y"] = np.nan, np.nan
    xx, yy = to_work_xy(pts.loc[ok, "longitude"], pts.loc[ok, "latitude"], cfg["crs"]["work_crs"])
    pts.loc[ok, "x"], pts.loc[ok, "y"] = xx, yy
    pts = assign_bursts(pts, cfg)
    pts = add_time_fields(pts, "timestamp_utc", "latitude", "longitude", cfg)
    rel = pd.to_datetime(pts["release_date"], utc=True, errors="coerce")
    pts["days_since_release"] = (pts["timestamp_utc"] - rel).dt.total_seconds() / 86400
    pts["weeks_since_release"] = np.floor(pts["days_since_release"] / 7)
    pts["post_release_bin"] = bin_days(pts["days_since_release"], cfg)
    pts["fix_number_since_release"] = pts[pts["included_analysis"]].groupby("animal_id").cumcount()
    table = load_season_table(cfg)
    s, src = assign_calendar_season(pts["local_date"], pts["region"], table)
    pts["season_calendar"], pts["season_calendar_source"] = s, src
    pts["season"], pts["season_source"] = s, src
    ctx["season_note"] = ("таблица " + cfg["seasons"]["table"]) if cfg["seasons"].get("table") else \
        "календарь по умолчанию (snow_period 01.12–15.03, межсезонье 16.03–30.04 и 16.10–30.11, вегетация 01.05–15.10) — требует подтверждения зоологами"
    ctx["nominal"] = nominal_intervals(pts)
    inc = pts[pts["included_analysis"]].copy()
    steps = make_steps(inc, cfg)
    reg = regularize(pts, cfg) if cfg["regularization"]["enabled"] else pd.DataFrame()
    reg_steps = make_steps(reg.drop(columns=["burst_id"]).rename(columns={"reg_burst_id": "burst_id"}), cfg) if len(reg) else pd.DataFrame()
    LOG.info(f"Записей {len(pts)}, валидных {len(inc)}, шагов {len(steps)}, регулярных шагов {len(reg_steps)}")
    for s_ in (steps, reg_steps):
        if len(s_):
            s_["behavior_simple"] = simple_behavior(s_, cfg)
    ctx.update(pts=pts, steps=steps, reg_points=reg, reg_steps=reg_steps)
    return ctx


def stage_fetch(ctx: dict, cfg: dict) -> dict:
    LOG.info("=== Этап fetch: загрузка данных ДЗЗ (GEE) ===")
    if not ctx.get("gee_ok"):
        LOG.info("GEE недоступен — этап fetch пропущен")
        return ctx
    inc = ctx["pts"][ctx["pts"]["included_analysis"]]
    cache_dir = Path(ctx["cache_dir"])
    ctx["tiles"] = fetch_static_tiles(inc["x"].values, inc["y"].values, cfg, cache_dir)
    sel = inc
    if cfg["gee"]["dynamic_points"] == "regular" and len(ctx["reg_points"]):
        sel = inc[inc["record_id"].isin(ctx["reg_points"]["record_id"])]
    ctx["dyn"] = fetch_dynamic(sel, cfg, cache_dir)
    return ctx


def stage_extract(ctx: dict, cfg: dict) -> dict:
    LOG.info("=== Этап extract: обогащение точек и шагов ===")
    pts, steps = ctx["pts"], ctx["steps"]
    reclass = load_reclass(cfg)
    ctx["reclass"] = reclass
    store = None
    if ctx.get("tiles") and rasterio is not None:
        store = TileStore(ctx["tiles"], cfg)
        store.estimate_tpi_sd()
        ctx["tpi_sd"] = store.tpi_sd
        inc_idx = pts.index[pts["included_analysis"]]
        sub = enrich_points_static(pts.loc[inc_idx].copy(), store, cfg, reclass)
        newc = [c for c in sub.columns if c not in pts.columns]
        pts = pts.join(sub[newc])
        pts = apply_regional_raster(pts, cfg)
        LOG.info("Буферные окна вокруг точек…")
        bs = buffer_stats(pts.loc[inc_idx], store, cfg, reclass)
        ctx["buffers"] = bs
        pts = pts.merge(bs, on="record_id", how="left")
        pts = distance_to_vectors(pts, cfg) if cfg.get("extra_vectors") else pts
    elif ctx.get("tiles"):
        warn("rasterio не установлен — статичные слои не прочитаны (pip install rasterio)")
    if ctx.get("dyn") is not None and len(ctx["dyn"]):
        pts = pts.merge(ctx["dyn"], on="record_id", how="left")
        pts = postprocess_dynamic(pts, cfg)
    # сезоны: объективный вариант
    obj = objective_seasons(pts, cfg) if "modis_snowdays_d7" in pts or f"modis_snowdays_d{cfg['seasons']['objective']['snow_window_days']}" in pts else pd.DataFrame()
    ctx["seasons_objective"] = obj
    if len(obj):
        pts = pts.merge(obj[["animal_id", "local_date", "season_objective", "season_obj_confidence", "season_obj_transition", "season_obj_features"]],
                        on=["animal_id", "local_date"], how="left")
        if cfg["seasons"]["primary"] == "objective":
            m = pts["season_objective"].notna() & (pts["season_objective"] != "unknown")
            pts.loc[m, "season"] = pts.loc[m, "season_objective"]
            pts.loc[m, "season_source"] = "objective_MODIS_snow_NDVI"
            ctx["season_note"] = "объективные сезоны по MOD10A1 и MOD13Q1 (при отсутствии данных — календарь)"
    elif cfg["seasons"]["primary"] == "objective":
        warn("seasons.primary=objective, но данных MODIS нет — используется календарь")
    ctx["pts"] = pts
    cols = ["elevation", "slope_deg", "aspect_deg", "northness", "eastness", "tri", "tpi_small", "tpi_large", "curvature",
            "landform", "aspect_class", "habitat", "worldcover_name", "elevation_band", "ndvi_nearest", "ndvi_sensor",
            "ndvi_obs_dt_days", "ndsi_nearest", "modis_ndvi_16d", "snow_present", "snow_fraction_or_ndsi", "snow_depth_m",
            "snow_water_equivalent_m", "dw_top_class", "dw_max_prob", "solar_period_A", "solar_period_B", "season",
            "sun_elevation_deg", "days_since_release", "post_release_bin"]
    for key in ("steps", "reg_steps"):
        s = ctx[key]
        if len(s) == 0:
            continue
        s = s[[c for c in s.columns if c == "delta_t_s" or (not c.startswith(("start_", "end_", "delta_", "transition_", "path_"))
                                                            and c not in ("season", "vertical_change_m", "vertical_speed_m_s"))]]
        s = attach_point_covariates_to_steps(s, pts, cols)
        s["season"] = s["start_season"]
        tz = cfg["time"]["local_timezone"]
        mid = pd.DataFrame({"t": s["t_mid"], "lat": s["mid_lat"], "lon": s["mid_lon"]})
        mid = add_time_fields(mid, "t", "lat", "lon", cfg, prefix="mid_")
        s["mid_solar_period_A"], s["mid_solar_period_B"] = mid["mid_solar_period_A"].values, mid["mid_solar_period_B"].values
        if store is not None and key == "steps":
            LOG.info("Признаки вдоль линий шагов…")
            ps = path_stats(s, store, cfg, reclass)
            s = s.merge(ps, on="step_id", how="left")
        ctx[key] = s
    ctx["store_meta"] = ctx.get("tiles")
    return ctx


def stage_analyze(ctx: dict, cfg: dict) -> dict:
    LOG.info("=== Этап analyze ===")
    pts, steps = ctx["pts"], ctx["steps"]
    inc = pts[pts["included_analysis"]]
    LOG.info("Потенциальные охотничьи кластеры…")
    cl, win, exc, lab = detect_clusters(inc, cfg, nominal_h=ctx["nominal"])
    if len(cl):
        p = pts.set_index("record_id")
        env_cols = [c for c in ("elevation", "slope_deg", "habitat", "landform", "ndvi_nearest", "snow_present", "season",
                                "solar_period_A", "solar_period_B", "days_since_release", "sex", "region") if c in p]
        e = p.loc[cl["medoid_record_id"], env_cols].reset_index(drop=True)
        e.columns = ["medoid_" + c if c not in ("sex", "region", "season", "days_since_release") else c for c in e.columns]
        cl = pd.concat([cl.reset_index(drop=True), e], axis=1)
        cl["start_solar_period_A"] = solar_period_A(cl["t_start"].dt.tz_convert(cfg["time"]["local_timezone"]), cfg)
        lon, lat = from_work_xy(cl["medoid_x"], cl["medoid_y"], cfg["crs"]["work_crs"])
        cl["medoid_lon"], cl["medoid_lat"] = lon, lat
        cl["cluster_validation"] = "not_checked"
    ctx["clusters"], ctx["cluster_windows"], ctx["cluster_excursions"] = cl, win, exc
    pts, steps = label_modes(pts, steps, lab)
    if len(ctx["reg_steps"]):
        _, ctx["reg_steps"] = label_modes(pts.copy(), ctx["reg_steps"], lab)
    # чувствительность
    sens = []
    for R in cfg["clusters"]["sensitivity_radii"]:
        for D in cfg["clusters"]["sensitivity_durations"]:
            c2, _, _, _ = detect_clusters(inc, cfg, radius=R, min_dur_h=D, nominal_h=ctx["nominal"], record_windows=False, bootstrap=False)
            for a in inc["animal_id"].unique():
                g = c2[c2["animal_id"] == a] if len(c2) else pd.DataFrame()
                sens.append({"radius_m": R, "min_duration_h": D, "animal_id": a, "fix_regime": "original",
                             "nominal_interval_h": ctx["nominal"].get(a), "n_clusters": len(g),
                             "total_cluster_h": g["duration_h"].sum() if len(g) else 0.0})
    rp = ctx["reg_points"]
    if len(rp):
        nom2 = {a: cfg["regularization"]["interval_h"] for a in rp["animal_id"].unique()}
        c3, _, _, _ = detect_clusters(rp, cfg, nominal_h=nom2, record_windows=False, bootstrap=False)
        for a in rp["animal_id"].unique():
            g = c3[c3["animal_id"] == a] if len(c3) else pd.DataFrame()
            sens.append({"radius_m": cfg["clusters"]["radius_m"], "min_duration_h": cfg["clusters"]["min_duration_h"], "animal_id": a,
                         "fix_regime": f"regularized_{cfg['regularization']['interval_h']}h", "nominal_interval_h": cfg["regularization"]["interval_h"],
                         "n_clusters": len(g), "total_cluster_h": g["duration_h"].sum() if len(g) else 0.0})
    ctx["cluster_sensitivity"] = pd.DataFrame(sens)
    # HMM
    LOG.info("HMM…")
    res = run_hmm(ctx["reg_steps"], cfg)
    if res is not None:
        hdiag, hlab = res
        ctx["hmm_parameters"] = hdiag
        if len(hlab):
            rs = ctx["reg_steps"]
            hl = hlab.set_index("step_index")
            for c in hl.columns:
                rs[c] = hl[c].reindex(rs.index).values
            ctx["reg_steps"] = rs
            m = rs.set_index("rec_start")["hmm_state"].dropna()
            steps["hmm_state"] = steps["rec_start"].map(m)
    # суточный ход
    LOG.info("Суточный ход…")
    day = daily_movement(steps, pts, cfg, ctx["nominal"], ctx["meta"])
    ds = pts[pts["included_analysis"]].groupby(["animal_id", "local_date"])["season"].agg(lambda s: s.mode().iloc[0] if len(s.mode()) else "unassigned")
    day = day.merge(ds.rename("season").reset_index(), on=["animal_id", "local_date"], how="left")
    if "season_obj_transition" in pts:
        tr = pts.groupby(["animal_id", "local_date"])["season_obj_transition"].max()
        day = day.merge(tr.rename("season_transition").reset_index(), on=["animal_id", "local_date"], how="left")
    day["post_release_bin"] = bin_days(day["days_since_release"], cfg)
    ctx["day"] = day
    ctx["daily_summary"] = daily_summary(day, cfg)
    ctx["daily_models"] = daily_mixed_model(day, cfg)
    pr = post_release_analysis(day, pts, ctx["meta"], cfg)
    ctx["post_release"], ctx["stabilization"] = pr if pr is not None else (pd.DataFrame(), pd.DataFrame())
    # переходы
    tc, tp, tsp = transitions(steps)
    ctx["transition_counts"], ctx["transition_probabilities"], ctx["transition_speed"] = tc, tp, tsp
    # доступность
    if ctx.get("tiles") and rasterio is not None:
        store = TileStore(ctx["tiles"], cfg)
        store.tpi_sd = ctx.get("tpi_sd")
        ctx["selection_ratios"] = selection_ratios(pts, store, ctx["reclass"], cfg) if "habitat" in pts else None
        if cfg["issf"]["enabled"] and len(ctx["reg_steps"]):
            LOG.info("iSSF…")
            r = run_issf(ctx["reg_steps"], store, ctx["reclass"], cfg)
            if r is not None:
                ctx["available_steps"], ctx["issf_coefficients"], ctx["issf_diag"] = r
    else:
        warn("iSSF и коэффициенты избирательности не рассчитаны: нет слоёв среды (GEE)")
    sm_res = speed_models(steps, cfg)
    if sm_res is not None:
        ctx["speed_models"], ctx["speed_diag"] = sm_res
    ctx["group_comparisons"] = group_comparisons(day, steps, cfg)
    feat = individual_features(day, steps, pts, cl, ctx.get("stabilization"), cfg)
    ctx["individual_features"] = feat
    ty = typology(feat, day, cfg)
    ctx["typology"] = ty
    if ty is not None:
        ctx["typology_note"] = (f"Признаки: {len(ty[4]['features_used'])} (исключены коррелированные |r|>0.9: {ty[4]['dropped_corr']}). "
                                f"Силуэт: { {k: round(v, 2) for k, v in ty[4]['silhouette'].items()} }. "
                                + ("Структура слабая — группы не выделяются; дендрограмма и матрица сходства — исследовательский результат."
                                   if ty[4]["weak"] else f"Выбрано k={ty[4]['best_k']}; устойчивость — матрица соприсвоения (бутстреп признаков)."))
    else:
        ctx["typology_note"] = "Типология не выполнена (мало животных или ошибка модуля)."
    ctx["pts"], ctx["steps"] = pts, steps
    ctx.update(build_summaries(ctx, cfg))
    return ctx


def assemble_tables(ctx: dict) -> Dict[str, pd.DataFrame]:
    pts, steps = ctx["pts"], ctx["steps"]
    diag = [d for d in (ctx.get("issf_diag"), ctx.get("speed_diag")) if isinstance(d, pd.DataFrame) and len(d)]
    fields = [c for c in STD_FIELDS if c in pts]
    fc = pd.DataFrame({"field": fields, "non_null_frac": [float(pts[c].notna().mean() if c not in ("sex",) else (pts[c] != "unknown").mean()) for c in fields]})
    T = {
        "points_cleaned": pts, "steps_enriched": steps, "individual_summary": ctx.get("individual_summary"),
        "sex_summary": ctx.get("sex_summary"), "season_summary": ctx.get("season_summary"),
        "solar_period_summary": ctx.get("solar_period_summary"), "habitat_use": ctx.get("habitat_use"),
        "terrain_use": ctx.get("terrain_use"), "ndvi_snow_summary": ctx.get("ndvi_snow_summary"),
        "transition_counts": ctx.get("transition_counts"), "transition_probabilities": ctx.get("transition_probabilities"),
        "hunting_clusters": ctx.get("clusters"), "cluster_candidate_windows": ctx.get("cluster_windows"),
        "cluster_excursions": ctx.get("cluster_excursions"), "cluster_sensitivity": ctx.get("cluster_sensitivity"),
        "behavior_labeled_points": pts[[c for c in ("record_id", "animal_id", "timestamp_utc", "longitude", "latitude", "behavior_mode", "cluster_id") if c in pts]],
        "behavior_labeled_steps": steps[[c for c in ("step_id", "animal_id", "t_start", "t_end", "behavior_mode", "behavior_simple", "hmm_state", "cluster_start", "cluster_end") if c in steps]],
        "daily_movement": ctx.get("day"), "daily_movement_summary": ctx.get("daily_summary"), "daily_models": ctx.get("daily_models"),
        "post_release_metrics": ctx.get("post_release"), "stabilization_estimates": ctx.get("stabilization"),
        "individual_features": ctx.get("individual_features"), "speed_by_conditions": ctx.get("speed_by_conditions"),
        "speed_models": ctx.get("speed_models"), "available_steps": ctx.get("available_steps"),
        "issf_coefficients": ctx.get("issf_coefficients"), "selection_ratios": ctx.get("selection_ratios"),
        "hmm_parameters": ctx.get("hmm_parameters"), "group_comparisons": ctx.get("group_comparisons"),
        "model_diagnostics": pd.concat(diag, ignore_index=True) if diag else pd.DataFrame(),
        "data_provenance": pd.DataFrame(PROVENANCE).drop_duplicates() if PROVENANCE else pd.DataFrame(),
        "excluded_records": pts[~pts["included_analysis"]][["record_id", "animal_id", "timestamp_utc", "longitude", "latitude", "qc_reason", "source_file", "source_row"]],
        "input_files": ctx.get("input_files"), "field_completeness": fc, "animals_metadata": ctx["meta"],
        "seasons_objective": ctx.get("seasons_objective"), "qc_summary": ctx.get("qc_summary"),
    }
    if ctx.get("typology") is not None:
        clus, sim, coas, load, info = ctx["typology"]
        T["individual_clusters"] = clus
        T["individual_similarity"] = sim.reset_index()
        T["individual_coassignment"] = coas.reset_index()
        T["individual_pca_loadings"] = load.reset_index().rename(columns={"index": "feature"})
    return {k: v for k, v in T.items() if isinstance(v, pd.DataFrame)}


def stage_report(ctx: dict, cfg: dict, out: Path) -> dict:
    LOG.info("=== Этап report: таблицы, рисунки, карта, HTML ===")
    FIGS.clear()
    PROVENANCE.extend([dict(component="tracks_input", asset_id=r["file"], scale_m="", crs=cfg["input"]["source_crs"],
                            temporal="", processing=f"sha256={r['sha256']}; {r['time_rule']}") for _, r in ctx["input_files"].iterrows()])
    add_provenance(component="software", asset_id=f"LeoTrack {__version__}", scale_m="", crs=cfg["crs"]["work_crs"], temporal="",
                   processing=f"script sha256 {script_hash()}; pandas {pd.__version__}; numpy {np.__version__}; ee {getattr(ee, '__version__', 'n/a')}")
    T = assemble_tables(ctx)
    ctx["tables"] = T
    n_source = len(ctx["pts"])
    raw = ctx["raw"]
    (out / "tables").mkdir(parents=True, exist_ok=True)
    raw.to_csv(out / "tables" / "raw_input_copy.csv", index=False, encoding="utf-8-sig")
    write_tables(T, out, n_source)
    dpi = cfg["report"]["dpi"]
    pts, steps, public = ctx["pts"], ctx["steps"], cfg["privacy"]["public_mode"]
    fig_qc(pts, out, dpi)
    fig_overview(pts, ctx.get("clusters", pd.DataFrame()), ctx["meta"], out, dpi, public, cfg)
    fig_speed(steps, out, dpi)
    fig_hourly(steps, out, dpi, cfg)
    for var, title, fname in (("sex", "Скорость по полу", "box_speed_sex"), ("season", "Скорость по сезонам", "box_speed_season"),
                              ("start_solar_period_A", "Скорость по рабочим периодам суток A", "box_speed_periodA"),
                              ("start_solar_period_B", "Скорость по астрономическим периодам B", "box_speed_periodB"),
                              ("behavior_mode", "Скорость по режиму", "box_speed_mode"), ("region", "Скорость по регионам", "box_speed_region")):
        fig_box(steps, out, dpi, var, title, "speed", fname)
    fig_box(steps, out, dpi, "start_habitat", "Скорость по местообитанию (начало шага)", "habitat", "box_speed_habitat")
    fig_box(steps, out, dpi, "start_landform", "Скорость по формам рельефа (начало шага)", "terrain", "box_speed_landform")
    fig_box(steps, out, dpi, "start_snow_present", "Скорость при снеге/без снега (MOD10A1)", "snow", "box_speed_snow")
    fig_speed_vs(steps, out, dpi)
    fig_rose_aspect(pts, out, dpi)
    fig_composition(pts, "habitat", out, dpi, "habitat", "Состав местообитаний по животным (WorldCover → reclass)", "habitat_composition")
    fig_composition(pts, "dw_top_class", out, dpi, "habitat", "Dynamic World: класс с максимальной вероятностью", "dw_composition")
    fig_composition(pts, "landform", out, dpi, "terrain", "Формы рельефа в местах фиксаций", "landform_composition")
    fig_profile(pts, steps, out, dpi)
    for v in ("habitat", "landform", "elevation_band", "snow_present"):
        fig_transitions(ctx.get("transition_counts"), out, dpi, v)
    ctx["sankey"] = sankey_html(ctx.get("transition_counts"), out)
    fig_clusters(ctx.get("clusters"), ctx.get("cluster_windows"), ctx.get("cluster_sensitivity"), pts, out, dpi)
    fig_daily(ctx["day"], out, dpi)
    fig_post_release(ctx.get("post_release"), ctx.get("stabilization"), cfg, out, dpi)
    fig_typology(ctx.get("typology"), out, dpi)
    fig_timeseries(ctx["day"], pts, out, dpi)
    fig_coefs(ctx.get("issf_coefficients"), out, dpi, "iSSF: объединённые коэффициенты (random effects)", "issf_coefficients", "issf")
    fig_coefs(ctx.get("speed_models"), out, dpi, "Модель скорости: разности средних, м/ч", "speed_model_coefficients", "speed", xlabel="Δ средней скорости, м/ч")
    try:
        ctx["map"] = build_map(pts, steps, ctx.get("clusters"), ctx["meta"], cfg, out, public)
    except Exception as e:  # noqa: BLE001
        warn(f"Интерактивная карта не создана: {e}")
    ctx["report"] = build_report(ctx, out, cfg)
    LOG.info(f"Отчёт: {ctx['report']}")
    return ctx


# ==============================================================================================
# ТЕСТЫ (синтетические геометрии; в отчёт не попадают)
# ==============================================================================================
def run_selftests() -> bool:
    cfg = load_config(None)
    ok = True

    def check(name, cond):
        nonlocal ok
        print(f"[{'PASS' if cond else 'FAIL'}] {name}")
        ok &= bool(cond)

    print("LeoTrack self-test: СИНТЕТИЧЕСКИЕ ГЕОМЕТРИИ (не биологические данные)")
    errs = validate_config(cfg)
    check("конфигурация по умолчанию валидна", not errs)
    # геодезия и азимуты
    _, _, d = _GEOD.inv(43.0, 43.0, 43.0, 43.009)
    check("1 км по меридиану ≈ 1000 м", abs(d - 1000) < 5)
    t0 = pd.Timestamp("2022-07-01 00:00", tz="UTC")
    # синтетический трек: 30 ч на месте (дрожание 40 м), затем прямолинейный уход, разрыв 20 ч, ещё 10 точек
    rng = np.random.default_rng(1)
    rows = []
    for k in range(31):
        rows.append(("SYN", t0 + pd.Timedelta(hours=k), 43.5 + rng.normal(0, 0.0003), 43.0 + rng.normal(0, 0.0003)))
    for k in range(1, 15):
        rows.append(("SYN", t0 + pd.Timedelta(hours=30 + k), 43.5 + 0.01 * k, 43.0))
    for k in range(10):
        rows.append(("SYN", t0 + pd.Timedelta(hours=64 + k), 43.7 + 0.005 * k, 43.0))
    rows.append(("SYN", t0 + pd.Timedelta(hours=74, minutes=30), 47.0, 43.0))  # невозможный скачок
    rows.append(("SYN", t0 + pd.Timedelta(hours=75), 43.75, 43.0))
    p = pd.DataFrame(rows, columns=["animal_id", "timestamp_utc", "longitude", "latitude"])
    for c in ("sex", "hdop", "satellites", "collar_id", "speed_collar"):
        p[c] = np.nan
    p["record_id"] = np.arange(len(p))
    q = run_qc(p, cfg)
    check("выброс скорости помечен, но не удалён", len(q) == len(p) and (q["qc_reason"].str.contains("speed|spike")).sum() >= 1)
    q["x"], q["y"] = to_work_xy(q["longitude"], q["latitude"], cfg["crs"]["work_crs"])
    q = assign_bursts(q, cfg)
    check("разрыв 20 ч > 12 ч разделяет burst", q.loc[q["included_analysis"], "burst_id"].nunique() == 2)
    st = make_steps(q[q["included_analysis"]], cfg)
    check("шаги не пересекают разрыв burst", (st.groupby("burst_id").size().sum() == len(st)) and (st["delta_t_s"] <= 12 * 3600).all())
    check("угол поворота прямолинейного участка ≈ 0", st.loc[st["step_length_m"] > 500, "turn_angle_deg"].abs().median() < 5)
    cl, win, exc, lab = detect_clusters(q[q["included_analysis"]], cfg, nominal_h={"SYN": 1.0})
    check("найден ровно 1 кластер ≥24 ч в радиусе 150 м", len(cl) == 1 and cl["duration_h"].iloc[0] >= 24)
    check("p95 расстояний кластера ≤ 150 м", len(cl) == 1 and cl["p95_distance_m"].iloc[0] <= 150)
    check("95% эллипс центра рассчитан", len(cl) == 1 and cl["center_ci95_ellipse_area_m2"].iloc[0] > 0)
    cl2, *_ = detect_clusters(q[q["included_analysis"]], cfg, min_dur_h=36, nominal_h={"SYN": 1.0}, bootstrap=False)
    check("при 36 ч кластер не принимается", len(cl2) == 0)
    # экспозиция плоскости, поднимающейся к востоку → склон смотрит на запад (270°)
    xx = np.tile(np.arange(50) * 30.0, (50, 1))
    td = terrain_derivatives(xx.astype(np.float32) * 0.5, 30.0, cfg)
    check("экспозиция плоскости = 270° (W)", abs(np.nanmedian(td["aspect_deg"][5:-5, 5:-5]) - 270) < 1)
    check("уклон плоскости = atan(0.5)", abs(np.nanmedian(td["slope_deg"][5:-5, 5:-5]) - math.degrees(math.atan(0.5))) < 0.1)
    # деление шага по местной полуночи
    s1 = pd.DataFrame({"animal_id": ["SYN"], "t_start": [pd.Timestamp("2022-07-01 20:00", tz="UTC")], "t_end": [pd.Timestamp("2022-07-01 22:00", tz="UTC")],
                       "delta_t_s": [7200.0], "step_length_m": [1000.0], "behavior_mode": ["movement"]})
    pcs = split_steps_by_local_day(s1, cfg)  # 23:00–01:00 МСК
    check("шаг через полночь разделён 50/50", len(pcs) == 2 and abs(pcs["length_m"].iloc[0] - 500) < 1)
    # солнце: полдень в Нальчике летом высоко, полночь — ночь
    sp = solar_position(pd.Series([pd.Timestamp("2022-06-21 09:00", tz="UTC"), pd.Timestamp("2022-06-21 21:00", tz="UTC")]), [43.5, 43.5], [43.6, 43.6])
    check("высота Солнца: полдень > 60°, полночь < −6°", sp["sun_elevation_deg"].iloc[0] > 60 and sp["sun_elevation_deg"].iloc[1] < -6)
    per = solar_period_A(pd.Series(pd.to_datetime(["2022-01-01 00:30", "2022-01-01 02:00", "2022-01-01 05:00", "2022-01-01 12:00"])), cfg)
    check("периоды A с переходом через полночь", list(per) == ["evening_peak", "deep_night_rest", "morning_peak", "day_rest"])
    # HMM на синтетике из двух режимов
    x = np.r_[rng.gamma(2, 30, 300), rng.gamma(4, 500, 300)]
    ang = np.r_[rng.vonmises(0, 0.2, 300), rng.vonmises(0, 3, 300)]
    m = fit_hmm([(x, ang)], 2, cfg, 1)
    check("HMM разделяет медленный и быстрый режим", m is not None and m["shape"][0] * m["scale"][0] < 150 < m["shape"][1] * m["scale"][1])
    cps = binseg(np.r_[np.zeros(30), np.ones(40) * 5] + rng.normal(0, 0.5, 70), 7)
    check("точка изменения найдена около 30", len(cps) >= 1 and abs(cps[0] - 30) <= 3)
    print("ИТОГ:", "все тесты пройдены" if ok else "есть ошибки")
    return ok


# ==============================================================================================
# ТОЧКА ВХОДА
# ==============================================================================================
def parse_args(argv=None):
    ap = argparse.ArgumentParser(description="LeoTrack — анализ GPS-треков переднеазиатского леопарда")
    ap.add_argument("--stage", default="run-all", choices=["validate", "preprocess", "fetch", "extract", "analyze", "report", "run-all"])
    ap.add_argument("--files", nargs="*", help="входные файлы (без аргумента — окно выбора)")
    ap.add_argument("--out", help="папка результатов")
    ap.add_argument("--config", help="YAML-конфигурация")
    ap.add_argument("--metadata", help="CSV метаданных животных")
    ap.add_argument("--no-gui", action="store_true", help="без окон (консольный ввод)")
    ap.add_argument("--no-gee", action="store_true", help="не использовать Google Earth Engine")
    ap.add_argument("--gee-project")
    ap.add_argument("--auth-mode", choices=["localhost", "notebook", "gcloud", "service_account", "appdefault"])
    ap.add_argument("--sa-key")
    ap.add_argument("--public", action="store_true", help="публичный режим (огрубление координат)")
    ap.add_argument("--overwrite", action="store_true")
    ap.add_argument("--write-config", metavar="PATH")
    ap.add_argument("--selftest", action="store_true")
    return ap.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    if args.write_config:
        Path(args.write_config).write_text(DEFAULT_CONFIG_YAML.lstrip(), encoding="utf-8")
        print(f"Конфигурация по умолчанию записана: {args.write_config}")
        return 0
    if args.selftest:
        return 0 if run_selftests() else 1
    missing = [n for n in ("yaml", "pyproj", "scipy", "shapely") if OPTIONAL.get(n) is None]
    if missing:
        print("Не установлены обязательные пакеты: " + ", ".join(missing) +
              "\npip install pandas numpy scipy pyproj shapely matplotlib openpyxl pyyaml earthengine-api rasterio statsmodels scikit-learn folium plotly")
        return 2
    gui = not args.no_gui and _tk_available()
    files = args.files or gui_select_files()
    files = [f for f in files if f]
    meta_files = [f for f in files if Path(f).name.lower().startswith("animals_metadata")]
    if meta_files:
        files = [f for f in files if f not in meta_files]
        if not args.metadata:
            args.metadata = meta_files[0]
        LOG.info(f"Файл метаданных среди выбранных: {Path(meta_files[0]).name} - используется как метаданные, а не как трек")
    if not files:
        print("Файлы не выбраны — работа завершена.")
        return 1
    default_out = str(Path(files[0]).resolve().parent / "LeoTrack_results")
    out_root = Path(args.out or (gui_select_dir(default_out) if gui else default_out) or default_out)
    out_root.mkdir(parents=True, exist_ok=True)
    cfg_path = args.config or (gui_ask_config(str(out_root)) if gui and not args.files else None)
    cfg = load_config(cfg_path)
    if args.public:
        cfg["privacy"]["public_mode"] = True
    if args.overwrite:
        cfg["project"]["overwrite"] = True
    errs = validate_config(cfg)
    if errs:
        msg = "Ошибки конфигурации:\n - " + "\n - ".join(errs)
        print(msg)
        if gui:
            gui_message(msg, error=True)
        return 2
    # метаданные животных
    pts0, _, _ = load_inputs(files, cfg)
    meta_path = args.metadata or cfg["input"].get("metadata_file")
    cand = [meta_path, out_root / "animals_metadata.csv", Path(files[0]).parent / "animals_metadata.csv"]
    existing = None
    for c in cand:
        if c and Path(c).exists():
            existing = pd.read_csv(c, dtype=str)
            break
    meta = default_metadata(pts0, cfg, existing)
    if gui or (not args.files and not args.no_gui):
        meta = gui_metadata(meta)
    meta.to_csv(out_root / "animals_metadata.csv", index=False, encoding="utf-8-sig")
    # GEE
    if args.no_gee:
        gs = {"enabled": False}
    elif gui:
        gs = gui_gee_settings(cfg)
    else:
        gs = {"enabled": cfg["gee"]["enabled"], "project": args.gee_project or cfg["gee"]["project"],
              "auth_mode": args.auth_mode or cfg["gee"]["auth_mode"], "service_account_key": args.sa_key or cfg["gee"]["service_account_key"],
              "public_mode": cfg["privacy"]["public_mode"], "overwrite": cfg["project"]["overwrite"]}
    if args.gee_project:
        gs["project"] = args.gee_project
    cfg["privacy"]["public_mode"] = bool(gs.get("public_mode", cfg["privacy"]["public_mode"]))
    cfg["project"]["overwrite"] = bool(gs.get("overwrite", cfg["project"]["overwrite"]))
    cfg["gee"]["project"] = gs.get("project") or cfg["gee"]["project"]
    run_time = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    out = out_root / ("results" if cfg["project"]["overwrite"] else f"results_{dt.datetime.now():%Y%m%d_%H%M%S}")
    for sub in ("tables", "figures", "maps", "rasters", "logs", "reports"):
        (out / sub).mkdir(parents=True, exist_ok=True)
    log_path = setup_logging(out / "logs")
    LOG.info(f"LeoTrack {__version__}; файлов: {len(files)}; результаты: {out}")
    with open(out / "config_used.yaml", "w", encoding="utf-8") as fh:
        yaml.safe_dump(cfg, fh, allow_unicode=True, sort_keys=False)
    if args.stage == "validate":
        LOG.info("Проверка входных данных:")
        LOG.info(f"\n{pts0.groupby('animal_id').agg(n=('record_id', 'size'), t0=('timestamp_utc', 'min'), t1=('timestamp_utc', 'max')).to_string()}")
        LOG.info("Полнота полей:\n" + pts0[[c for c in STD_FIELDS if c in pts0]].notna().mean().round(3).to_string())
        LOG.info("Конфигурация корректна.")
        return 0
    gee_ok = gee_connect(gs, interactive=gui) if gs.get("enabled") else False
    cache_root = Path(cfg["project"].get("cache_dir") or (Path(files[0]).resolve().parent / "LeoTrack_cache"))
    old_cache = out_root / "cache"
    if not cache_root.exists() and old_cache.exists():
        cache_root = old_cache  # совместимость с кэшем прежних версий
    cache_root.mkdir(parents=True, exist_ok=True)
    LOG.info(f"Кэш: {cache_root}")
    cache = Cache(cache_root)
    ctx = {"files": files, "meta": meta, "gee_ok": gee_ok, "cache_dir": str(cache_root), "run_time": run_time}
    in_hash = hash_obj([sha256_file(f) for f in files])
    k_pre = hash_obj([in_hash, script_hash(), meta.to_dict("records"),
                      _cfg_part(cfg, "input", "qc", "steps", "regularization", "solar", "seasons", "time", "crs", "aoi", "regions", "post_release", "behavior")])
    order = ["preprocess", "fetch", "extract", "analyze", "report"]
    last = order.index(args.stage) if args.stage in order else len(order) - 1
    c = cache.get("preprocess", k_pre)
    if c is None:
        ctx = stage_preprocess(ctx, cfg)
        cache.put("preprocess", k_pre, {k: ctx[k] for k in ("raw", "input_files", "pts", "steps", "reg_points", "reg_steps", "nominal", "season_note")})
    else:
        ctx.update(c)
    if last >= 1:
        k_f = hash_obj([k_pre, gee_ok, _cfg_part(cfg, "gee"), cfg["seasons"]["objective"]["snow_cover_threshold"]])
        c = cache.get("fetch", k_f)
        if c is None:
            ctx = stage_fetch(ctx, cfg)
            if gee_ok:
                cache.put("fetch", k_f, {k: ctx.get(k) for k in ("tiles", "dyn")})
        else:
            ctx.update(c)
            PROVENANCE.append({"component": "cache", "asset_id": "GEE fetch", "processing": f"использован кэш {k_f}"})
    if last >= 2:
        k_e = hash_obj([k_pre, gee_ok, _cfg_part(cfg, "gee", "terrain", "habitat", "seasons", "extra_vectors")])
        c = cache.get("extract", k_e)
        if c is None:
            ctx = stage_extract(ctx, cfg)
            cache.put("extract", k_e, {k: ctx.get(k) for k in ("pts", "steps", "reg_steps", "seasons_objective", "tpi_sd", "reclass", "season_note")})
        else:
            ctx.update(c)
        ctx.setdefault("reclass", load_reclass(cfg))
    if last >= 3:
        ctx = stage_analyze(ctx, cfg)
    if last >= 4:
        ctx = stage_report(ctx, cfg, out)
        msg = f"Готово.\nОтчёт: {ctx['report']}\nТаблицы: {out / 'tables'}\nЖурнал: {log_path}"
        LOG.info(msg)
        if gui:
            gui_message(msg)
            try:
                if platform.system() == "Windows":
                    os.startfile(ctx["report"])  # noqa: S606
            except Exception:  # noqa: BLE001
                pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
