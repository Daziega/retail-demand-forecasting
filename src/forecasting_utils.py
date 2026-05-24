"""Shared utilities: chronological splits, metrics, walk-forward CV."""
import numpy as np
import pandas as pd

# ── Chronological split per methodology §3.6.2 (adjusted to data window) ──
# Raw M5 ends 2016-04-24, so test window is shortened from the spec.
TRAIN_END   = pd.Timestamp("2015-12-31")
VAL_END     = pd.Timestamp("2016-02-29")
TEST_END    = pd.Timestamp("2016-04-24")

FEATURE_COLS = [
    # Lag features
    "lag_7", "lag_14", "lag_28",
    # Rolling statistics
    "rolling_mean_7", "rolling_mean_28", "rolling_std_7",
    # Temporal
    "day_of_week", "week_of_year", "month", "year", "is_weekend",
    # Event proximity
    "days_since_last_event", "days_to_next_event",
    # Promotional / event
    "is_snap_day", "is_sporting_event", "is_cultural_event",
    "is_national_event", "is_religious_event",
    # Price
    "sell_price", "price_change",
    # Identifiers (label-encoded)
    "item_id", "dept_id", "cat_id", "store_id", "state_id",
]
TARGET_COL = "sales"


def chronological_split(df: pd.DataFrame, date_col: str = "date"):
    """Return (train, val, test) DataFrames split by calendar date."""
    train = df[df[date_col] <= TRAIN_END]
    val   = df[(df[date_col] > TRAIN_END) & (df[date_col] <= VAL_END)]
    test  = df[(df[date_col] > VAL_END)  & (df[date_col] <= TEST_END)]
    return train, val, test


def walk_forward_folds(df: pd.DataFrame, date_col: str = "date"):
    """Three walk-forward folds per methodology §3.6.2."""
    folds = [
        (("2011-01-01", "2013-12-31"), ("2014-01-01", "2014-12-31")),
        (("2011-01-01", "2014-12-31"), ("2015-01-01", "2015-12-31")),
        (("2011-01-01", "2015-12-31"), ("2016-01-01", "2016-03-31")),
    ]
    out = []
    for (tr_s, tr_e), (va_s, va_e) in folds:
        tr = df[(df[date_col] >= tr_s) & (df[date_col] <= tr_e)]
        va = df[(df[date_col] >= va_s) & (df[date_col] <= va_e)]
        out.append((tr, va))
    return out


# ── Evaluation metrics per methodology §3.6.4 ─────────────────────────────
def mae(y_true, y_pred):
    return float(np.mean(np.abs(np.asarray(y_true) - np.asarray(y_pred))))


def rmse(y_true, y_pred):
    return float(np.sqrt(np.mean((np.asarray(y_true) - np.asarray(y_pred)) ** 2)))


def mape(y_true, y_pred, eps: float = 1e-6):
    """MAPE in %. Skips zero-actuals to avoid division blow-ups."""
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    mask = np.abs(y_true) > eps
    if mask.sum() == 0:
        return float("nan")
    return float(np.mean(np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])) * 100)


def wmape(y_true, y_pred):
    """Weighted MAPE in %. Robust to zero-actuals (sum-of-abs in denominator)."""
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    denom = np.sum(np.abs(y_true))
    if denom == 0:
        return float("nan")
    return float(np.sum(np.abs(y_true - y_pred)) / denom * 100)


def pred10(y_true, y_pred, eps: float = 1e-6):
    """% of forecasts within ±10% of the actual value."""
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    mask = np.abs(y_true) > eps
    if mask.sum() == 0:
        return float("nan")
    pct_err = np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])
    return float(np.mean(pct_err <= 0.10) * 100)


def pinball_loss(y_true, y_pred_quantile, alpha: float):
    """Pinball (quantile) loss for evaluating quantile forecasts."""
    y_true = np.asarray(y_true, dtype=float)
    y_pred_quantile = np.asarray(y_pred_quantile, dtype=float)
    diff = y_true - y_pred_quantile
    return float(np.mean(np.maximum(alpha * diff, (alpha - 1) * diff)))


def all_metrics(y_true, y_pred, model_name: str = "") -> dict:
    return {
        "model":      model_name,
        "MAE":        mae(y_true, y_pred),
        "RMSE":       rmse(y_true, y_pred),
        "MAPE_pct":   mape(y_true, y_pred),
        "WMAPE_pct":  wmape(y_true, y_pred),
        "Pred10_pct": pred10(y_true, y_pred),
    }