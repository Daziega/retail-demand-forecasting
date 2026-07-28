# Raw data — not tracked in Git

The M5-Forecasting (Walmart) dataset is **not redistributed in this repository**.
It is ~450 MB and is governed by the Kaggle competition's own terms of use.

## How to obtain it

1. Go to the competition page:
   <https://www.kaggle.com/competitions/m5-forecasting-accuracy/data>
2. Accept the competition rules (a free Kaggle account is required).
3. Download and place these three files directly in this folder:

```
data/raw/
├── sales_train_validation.csv    # daily unit sales, 3,049 series × 1,913 days
├── calendar.csv                  # dates, weekday, SNAP flags, event labels
└── sell_prices.csv               # weekly selling price per item-store
```

Using the Kaggle CLI instead:

```bash
pip install kaggle
kaggle competitions download -c m5-forecasting-accuracy -p data/raw
unzip data/raw/m5-forecasting-accuracy.zip -d data/raw
```

## What the pipeline does with it

`notebooks/01_eda.ipynb` through `notebooks/03_feature_engineering.ipynb` read
these three files and produce the engineered feature set. From there, a
stratified subsample of **502 product-store series** (proportional across the
three product categories × three US states, fixed seed 42) is drawn for all
modelling — see `data/processed/subsample_series_ids.csv` for the exact IDs,
which makes the sample reproducible without re-running the sampling step.

## Already-derived outputs are committed

You do **not** need the raw data to inspect the results. Everything in
`data/processed/` — model leaderboards, SHAP importances, the full simulation
grid, and the Power BI feeds — is committed and can be read directly.

The raw data is only required if you want to retrain the models from scratch.
