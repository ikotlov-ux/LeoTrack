# LeoTrack

Воспроизводимый конвейер анализа GPS-треков переднеазиатского леопарда (*Panthera pardus*) на Северном Кавказе: контроль качества фиксаций, шаги и скорость, время суток (рабочие и астрономические периоды), рельеф, местообитания, NDVI и снег (Google Earth Engine), потенциальные охотничьи/кормовые кластеры, суточный ход, стабилизация после выпуска, iSSF и индивидуальная типология. На выходе — таблицы (CSV/XLSX), рисунки, интерактивная карта и HTML-отчёт.

## Установка

```bash
conda create -n leotrack python=3.11
conda activate leotrack
pip install -r requirements.txt
```

Недостающие пакеты GEE (`earthengine-api`, `rasterio`) скрипт предлагает доустановить при запуске.

## Запуск

```bash
python LeoTrack.py
```

Без аргументов открываются окна выбора файлов треков (XLSX/CSV/GPX/GPKG/Movebank CSV, можно несколько), папки результатов, метаданных животных (пол, дата и место выпуска, регион, статус) и параметров Google Earth Engine (Cloud Project ID, способ авторизации).

Командная строка:

```bash
python LeoTrack.py --stage validate|preprocess|fetch|extract|analyze|report|run-all \
    [--files f1.xlsx f2.xlsx ...] [--out DIR] [--config config.yaml] [--metadata animals_metadata.csv] \
    [--no-gui] [--no-gee] [--gee-project ID] [--auth-mode localhost|notebook|gcloud|service_account|appdefault] \
    [--sa-key key.json] [--public] [--overwrite]
python LeoTrack.py --write-config config.yaml   # конфигурация по умолчанию с комментариями
python LeoTrack.py --selftest                   # встроенные тесты на синтетических данных
```

Все пороги и параметры задаются в YAML (`config.example.yaml`). Загруженные из GEE слои кэшируются в `<папка входных данных>/LeoTrack_cache` и используются повторно.

## Принципы

- исходные записи не удаляются: `qc_flag`, `qc_reason`, `included_analysis`;
- шаги не пересекают разрывы трека (`burst_id`);
- скорость — расчётная скорость перемещения между GPS-фиксациями, м/ч, без логарифмирования;
- основная сводная статистика — среднее с 95% ДИ (t-распределение, единица повторности — животное);
- облачный пиксель не трактуется как отсутствие снега или низкий NDVI;
- использование отделено от предпочтения (iSSF); `hunting_candidate` — потенциальный кластер, а не подтверждённое место добычи;
- режим `--public` огрубляет координаты в выходных материалах.

## Данные

Данные GPS-треков в репозиторий не входят: точные координаты редкого вида не публикуются.

## Лицензия

MIT
