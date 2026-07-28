"""
Forward inventory simulation for the (R, s, S) Order-Up-To-Level policy.

PURPOSE
-------
An earlier draft of this research reported a ~21% annual cost reduction for
the ML-Empirical-Quantile policy over the Classical baseline. That figure was
NOT simulated: the achieved service level of each policy was *assigned*
analytically (Gaussian -> ~83%, Quantile -> ~95% "by calibration") and then
plugged into  SC = (target_SL - achieved_SL) * D * unit_cost * m.
Because the service levels were inputs rather than measured outcomes, the
quantile policy won by construction, before any demand was ever simulated.

This script replaces that assumption with a real experiment: it runs each
(R, s, S) policy forward day-by-day over the test window using the ACTUAL
observed demand, and COUNTS realized stockout units and realized on-hand
inventory. Costs are computed from what actually happened, not from an
assumed service-level gap.

WHAT IT FOUND
-------------
On the M5 test window all three policies achieve a 100% realized service
level — no policy ever stocks out — so there is no stockout cost for a better
forecast or a smarter safety-stock formula to remove. The 21% saving does not
survive measurement. Chapter 5 of the thesis reports this instead as a
boundary condition: the quantile buffer only pays for itself when the baseline
policy is genuinely exposed to stockout risk. Run falsification_tests.py to
see the engine correctly detect stockouts once that exposure is introduced.

The three policies compared (identical level forecast; they differ ONLY in
how safety stock is computed, so the comparison isolates the SS formulation):

  1. CLASSICAL          : Gaussian SS from classical-forecast residual sigma
  2. ML_GAUSSIAN        : Gaussian SS from ML-forecast residual sigma
  3. ML_EMP_QUANTILE    : SS from the empirical (q95 - q50) quantile spread

HOW THE STUDENTS PLUG IN REAL DATA
----------------------------------
Replace `build_demo_dataset()` with `load_real_dataset()` (a stub is
provided at the bottom). Each series needs, from the existing Python pipeline:
  - train_demand : 1D array of historical daily units (for fitting forecasts)
  - test_demand  : 1D array of actual daily units over the test window
  - level_forecast      : mean daily demand forecast d_bar (ensemble)
  - sigma_classical     : std of Croston-SBA residuals on validation
  - sigma_ml            : std of ensemble residuals on validation
  - q50, q95            : LightGBM quantile predictions (median, P95)
  - price               : unit selling price

Author: prepared for the TFM supervision (Demand Forecasting Using ML)
"""

import numpy as np
from dataclasses import dataclass, field

# ----------------------------------------------------------------------------
# Economic constants (from methodology section 3.7.2)
# ----------------------------------------------------------------------------
CARRYING_RATE   = 0.25     # annual holding cost as fraction of unit cost
ORDERING_COST_K = 50.0     # fixed cost per order ($)
GROSS_MARGIN    = 0.40     # unit_cost = (1 - margin) * price
REVIEW_PERIOD_R = 7        # periodic review every 7 days
TARGET_SL       = 0.95
Z_SCORE         = 1.6449   # z for 95% service level
DAYS_PER_YEAR   = 365


# ----------------------------------------------------------------------------
# Per-series inputs
# ----------------------------------------------------------------------------
@dataclass
class Series:
    train_demand: np.ndarray
    test_demand:  np.ndarray
    price:        float
    # forecast quantities (in real use these come from the trained models;
    # in the demo they are derived from the training window)
    level_forecast:  float = None   # d_bar, mean daily demand
    sigma_classical: float = None   # std of classical-forecast residuals
    sigma_ml:        float = None   # std of ML-forecast residuals
    q50:             float = None   # median daily demand (P50)
    q95:             float = None   # 95th percentile daily demand (P95)


# ----------------------------------------------------------------------------
# Core simulation: run ONE (R,s,S) policy forward over the test window
# ----------------------------------------------------------------------------
def simulate_policy(test_demand, d_bar, safety_stock, price, lead_time):
    """
    Lost-sales (R, s, S) simulation.

    Returns a dict with REALIZED quantities:
      achieved_sl   : unit fill rate actually observed
      stockout_units: total unmet demand actually observed
      avg_on_hand   : average on-hand inventory actually held
      num_orders    : number of replenishment orders actually placed
      annual_HC/OC/SC/TC : annualized cost components
    """
    R, L = REVIEW_PERIOD_R, lead_time
    horizon = len(test_demand)

    unit_cost = (1.0 - GROSS_MARGIN) * price
    h = unit_cost * CARRYING_RATE                      # $/unit/year
    D_annual = max(d_bar, 1e-9) * DAYS_PER_YEAR        # annual demand estimate
    Q = np.sqrt(2.0 * D_annual * ORDERING_COST_K / max(h, 1e-9))  # EOQ

    # Policy parameters
    reorder_point_s = d_bar * (R + L) + safety_stock
    order_up_to_S   = reorder_point_s + Q

    # State
    on_hand = order_up_to_S          # start fully stocked
    in_transit = {}                  # arrival_day -> qty
    stockout_units = 0.0
    served_total = 0.0
    demand_total = 0.0
    on_hand_accum = 0.0
    num_orders = 0

    for t in range(horizon):
        # 1. receive any arrivals scheduled for today
        if t in in_transit:
            on_hand += in_transit.pop(t)

        # 2. meet demand (lost sales: unmet demand is lost, not backordered)
        d = float(test_demand[t])
        demand_total += d
        served = min(on_hand, d)
        served_total += served
        stockout_units += (d - served)
        on_hand -= served
        on_hand_accum += on_hand

        # 3. periodic review every R days -> place order up to S if at/below s
        if t % R == 0:
            inv_position = on_hand + sum(in_transit.values())
            if inv_position <= reorder_point_s:
                order_qty = order_up_to_S - inv_position
                if order_qty > 0:
                    arrival = t + L
                    in_transit[arrival] = in_transit.get(arrival, 0.0) + order_qty
                    num_orders += 1

    avg_on_hand = on_hand_accum / horizon
    achieved_sl = served_total / demand_total if demand_total > 0 else 1.0

    # Annualize (test window is `horizon` days)
    f = DAYS_PER_YEAR / horizon
    annual_HC = avg_on_hand * h                 # on-hand is a level -> already annual rate via h
    annual_OC = num_orders * f * ORDERING_COST_K
    # NOTE: stockout cost multiplier m is applied later (it is a scenario knob)
    annual_stockout_units = stockout_units * f

    return dict(
        achieved_sl=achieved_sl,
        stockout_units=stockout_units,
        annual_stockout_units=annual_stockout_units,
        avg_on_hand=avg_on_hand,
        num_orders=num_orders,
        Q=Q, reorder_point=reorder_point_s, order_up_to=order_up_to_S,
        safety_stock=safety_stock, unit_cost=unit_cost,
        annual_HC=annual_HC, annual_OC=annual_OC,
    )


def safety_stock_gaussian(sigma, lead_time):
    return Z_SCORE * sigma * np.sqrt(REVIEW_PERIOD_R + lead_time)


def safety_stock_empirical(q50, q95, lead_time):
    # replicates the thesis formula (q95 - q50) * sqrt(R+L) for a like-for-like
    # comparison; see note in the report about the sqrt-scaling caveat
    return max(q95 - q50, 0.0) * np.sqrt(REVIEW_PERIOD_R + lead_time)


# ----------------------------------------------------------------------------
# Run all three policies across the whole portfolio at a given (L, m)
# ----------------------------------------------------------------------------
def run_portfolio(series_list, lead_time, stockout_multiplier_m):
    policies = {"CLASSICAL": {}, "ML_GAUSSIAN": {}, "ML_EMP_QUANTILE": {}}
    agg = {p: dict(HC=0.0, OC=0.0, SC=0.0, ss=0.0, su=0.0,
                   sl_weighted=0.0, dem=0.0) for p in policies}

    for s in series_list:
        ss_classical = safety_stock_gaussian(s.sigma_classical, lead_time)
        ss_ml        = safety_stock_gaussian(s.sigma_ml, lead_time)
        ss_emp       = safety_stock_empirical(s.q50, s.q95, lead_time)

        runs = {
            "CLASSICAL":       simulate_policy(s.test_demand, s.level_forecast, ss_classical, s.price, lead_time),
            "ML_GAUSSIAN":     simulate_policy(s.test_demand, s.level_forecast, ss_ml,        s.price, lead_time),
            "ML_EMP_QUANTILE": simulate_policy(s.test_demand, s.level_forecast, ss_emp,       s.price, lead_time),
        }
        dem = float(s.test_demand.sum())
        for p, r in runs.items():
            sc = r["annual_stockout_units"] * r["unit_cost"] * stockout_multiplier_m
            agg[p]["HC"] += r["annual_HC"]
            agg[p]["OC"] += r["annual_OC"]
            agg[p]["SC"] += sc
            agg[p]["ss"] += r["safety_stock"]
            agg[p]["su"] += r["stockout_units"]
            agg[p]["sl_weighted"] += r["achieved_sl"] * dem
            agg[p]["dem"] += dem

    for p in agg:
        a = agg[p]
        a["TC"] = a["HC"] + a["OC"] + a["SC"]
        a["achieved_sl"] = a["sl_weighted"] / a["dem"] if a["dem"] > 0 else 1.0
    return agg


# ----------------------------------------------------------------------------
# DEMO DATA: synthetic intermittent demand that mimics M5 SKU-daily series
# (zero-inflated, right-skewed). Lets us verify the engine and SEE the number
# the simulation produces. Replace with load_real_dataset() for the thesis.
# ----------------------------------------------------------------------------
def build_demo_dataset(n_series=502, seed=42):
    rng = np.random.default_rng(seed)
    series = []
    train_len, test_len = 1800, 56   # ~5 yrs train, 8 wks test (matches thesis)

    for _ in range(n_series):
        # Heterogeneous SKUs: some near-dead, some moderate movers.
        base = rng.gamma(shape=0.7, scale=1.4)          # mean daily demand
        zero_inflation = rng.uniform(0.25, 0.75)        # fraction of dead days
        price = rng.uniform(2.0, 25.0)

        spike_rate = rng.uniform(0.02, 0.06)            # promo/event days
        spike_mult = rng.uniform(6.0, 14.0)             # size of the upper tail

        def draw(n):
            active = rng.random(n) > zero_inflation
            # right-skewed counts on active days (negative-binomial-like)
            lam = rng.gamma(shape=2.0, scale=base, size=n)
            d = rng.poisson(lam) * active
            # inject promotional spikes -> heavy right tail (where SS matters)
            spikes = rng.random(n) < spike_rate
            d = d + spikes * rng.poisson(lam * spike_mult)
            return d.astype(float)

        train_demand = draw(train_len)
        test_demand  = draw(test_len)

        # "Forecasts" — in the real thesis these come from the trained models.
        # Here we derive sensible proxies from the training window:
        d_bar = train_demand.mean()
        # classical (Croston-style) and ML residual sigmas: ML slightly tighter
        resid_std = train_demand.std()
        sigma_classical = resid_std
        sigma_ml        = resid_std * 0.95              # ML ~5% tighter (as in Ch.4)
        q50 = np.quantile(train_demand, 0.50)
        q95 = np.quantile(train_demand, 0.95)

        series.append(Series(
            train_demand=train_demand, test_demand=test_demand, price=price,
            level_forecast=d_bar, sigma_classical=sigma_classical,
            sigma_ml=sigma_ml, q50=q50, q95=q95,
        ))
    return series


def load_real_dataset():
    """
    Build per-series inputs from the trained-pipeline artefacts.

    Sources (all parquet files produced earlier in the project):
      - artifacts/ensemble_test_predictions.parquet : actual sales + ML ensemble forecast
      - artifacts/baseline_test_predictions.parquet : Croston-SBA classical forecast
      - artifacts/quantile_test_predictions.parquet : LightGBM quantile forecasts (q50, q95)
      - data/processed/subsample_features.parquet  : per-series selling price

    Sigma handling: per-series residual standard deviation is computed on the
    test window for BOTH the classical and ML forecasts. This is a consistent
    choice across policies and resolves the §3.7.2 / §5.2 inconsistency the
    supervisor flagged. Strictly, validation residuals would be cleaner; we
    document that limitation in the rewritten Chapter 5.
    """
    import pandas as pd

    ens = pd.read_parquet("artifacts/ensemble_test_predictions.parquet")
    bas = pd.read_parquet("artifacts/baseline_test_predictions.parquet")
    qtl = pd.read_parquet("artifacts/quantile_test_predictions.parquet")
    feat = pd.read_parquet(
        "data/processed/subsample_features.parquet",
        columns=["id", "sell_price"],
    )

    # Align by (id, date)
    ens["date"] = pd.to_datetime(ens["date"])
    bas["date"] = pd.to_datetime(bas["date"])
    qtl["date"] = pd.to_datetime(qtl["date"])

    df = (
        ens[["id", "date", "sales", "pred_ensemble_trees"]]
        .merge(bas[["id", "date", "Croston_SBA"]], on=["id", "date"], how="inner")
        .merge(qtl[["id", "date", "q50", "q95"]], on=["id", "date"], how="inner")
        .sort_values(["id", "date"])
    )

    # Mean selling price per series (sell_price varies weekly per item-store)
    price_per_series = feat.groupby("id", as_index=False)["sell_price"].mean()

    # Build a Series object per id
    series_list = []
    for sid, grp in df.groupby("id", sort=False):
        actual = grp["sales"].to_numpy(dtype=float)
        pred_ml = grp["pred_ensemble_trees"].to_numpy(dtype=float)
        pred_cls = grp["Croston_SBA"].clip(lower=0).to_numpy(dtype=float)

        sigma_ml = float(np.std(actual - pred_ml, ddof=0))
        sigma_classical = float(np.std(actual - pred_cls, ddof=0))

        # Use per-series mean quantile prediction (single SS value per series)
        q50 = float(grp["q50"].mean())
        q95 = float(grp["q95"].mean())

        price_row = price_per_series.loc[price_per_series["id"] == sid, "sell_price"]
        price = float(price_row.iloc[0]) if len(price_row) else 5.0
        if price <= 0 or np.isnan(price):
            price = 5.0

        # Level forecast: mean of the ensemble forecast over the test window
        d_bar = float(np.mean(pred_ml))
        if d_bar <= 0:
            d_bar = float(np.mean(actual)) if np.mean(actual) > 0 else 1e-3

        # We don't need historical demand for the simulation (it doesn't refit),
        # so train_demand is filled with the actuals (placeholder).
        series_list.append(Series(
            train_demand=actual.copy(),
            test_demand=actual,
            price=price,
            level_forecast=d_bar,
            sigma_classical=sigma_classical,
            sigma_ml=sigma_ml,
            q50=q50,
            q95=q95,
        ))

    return series_list


# ----------------------------------------------------------------------------
# Reporting
# ----------------------------------------------------------------------------
def pct_reduction(classical_tc, other_tc):
    return 100.0 * (classical_tc - other_tc) / classical_tc if classical_tc else 0.0


def main():
    import sys
    use_real = "--demo" not in sys.argv
    if use_real:
        print("Loading real M5 stratified-subsample portfolio from artifacts/ ...")
        series = load_real_dataset()
    else:
        series = build_demo_dataset(n_series=502, seed=42)
    total_test_demand = sum(s.test_demand.sum() for s in series)

    print("=" * 78)
    label = "REAL M5 stratified subsample, 502 series" if use_real else "synthetic M5-like demo, 502 series"
    print(f"FORWARD INVENTORY SIMULATION  ({label})")
    print("=" * 78)
    print(f"Total test-window demand across portfolio: {total_test_demand:,.0f} units")
    print(f"Review period R = {REVIEW_PERIOD_R} d | Target service level = {TARGET_SL:.0%}\n")

    # --- Headline table: central case L=14, m=1.0 ---
    print("CENTRAL CASE  (L = 14 days, stockout multiplier m = 1.0)")
    print("-" * 78)
    agg = run_portfolio(series, lead_time=14, stockout_multiplier_m=1.0)
    base_tc = agg["CLASSICAL"]["TC"]
    hdr = f"{'Policy':<18}{'Realized SL':>12}{'SS (units)':>12}{'HC $':>11}{'OC $':>10}{'SC $':>11}{'Total $':>12}{'vs Cls':>9}"
    print(hdr)
    for p in ["CLASSICAL", "ML_GAUSSIAN", "ML_EMP_QUANTILE"]:
        a = agg[p]
        red = pct_reduction(base_tc, a["TC"])
        red_s = "-" if p == "CLASSICAL" else f"{red:+.1f}%"
        print(f"{p:<18}{a['achieved_sl']*100:>11.1f}%{a['ss']:>12,.0f}"
              f"{a['HC']:>11,.0f}{a['OC']:>10,.0f}{a['SC']:>11,.0f}"
              f"{a['TC']:>12,.0f}{red_s:>9}")

    # --- Sensitivity to stockout multiplier m (the knob that actually matters) ---
    print("\nSENSITIVITY TO STOCKOUT MULTIPLIER m   (L = 14, SL = 95%)")
    print("-" * 78)
    print(f"{'m':>5}{'  Interpretation':<34}"
          f"{'Quantile vs Classical':>24}{'ML-Gauss vs Cls':>18}")
    interp = {0.4: "lost margin only (conservative)",
              1.0: "full unit cost (baseline)",
              2.0: "unit cost + reputation (aggr.)"}
    for m in [0.4, 1.0, 2.0]:
        a = run_portfolio(series, lead_time=14, stockout_multiplier_m=m)
        b = a["CLASSICAL"]["TC"]
        rq = pct_reduction(b, a["ML_EMP_QUANTILE"]["TC"])
        rg = pct_reduction(b, a["ML_GAUSSIAN"]["TC"])
        print(f"{m:>5}{('  ' + interp[m]):<34}{rq:>22.1f}%{rg:>16.1f}%")

    # --- Sensitivity to lead time ---
    print("\nSENSITIVITY TO LEAD TIME L   (m = 1.0, SL = 95%)")
    print("-" * 78)
    print(f"{'L (d)':>6}{'Classical TC $':>18}{'ML-Quantile TC $':>20}{'Reduction':>12}")
    for L in [7, 10, 14, 21]:
        a = run_portfolio(series, lead_time=L, stockout_multiplier_m=1.0)
        b = a["CLASSICAL"]["TC"]
        rq = pct_reduction(b, a["ML_EMP_QUANTILE"]["TC"])
        print(f"{L:>6}{b:>18,.0f}{a['ML_EMP_QUANTILE']['TC']:>20,.0f}{rq:>11.1f}%")

    print("\n" + "=" * 78)
    print("READING THIS: the 'Realized SL' column is MEASURED by counting unmet")
    print("demand day-by-day, not assigned. Compare it against the thesis's")
    print("assumed 83% (Gaussian) / 95% (Quantile). The 'vs Cls' reduction is the")
    print("honest number that replaces the assumed 21%.")
    print("=" * 78)

    # --- Save results to CSV for Chapter 5 rewriting ---
    import csv
    rows = []
    for L in [7, 10, 14, 21]:
        for m in [0.4, 1.0, 2.0]:
            a = run_portfolio(series, lead_time=L, stockout_multiplier_m=m)
            for p in ["CLASSICAL", "ML_GAUSSIAN", "ML_EMP_QUANTILE"]:
                rows.append(dict(
                    lead_time=L, stockout_mult=m, policy=p,
                    realized_sl_pct=round(a[p]["achieved_sl"] * 100, 2),
                    safety_stock=round(a[p]["ss"], 1),
                    HC=round(a[p]["HC"], 2),
                    OC=round(a[p]["OC"], 2),
                    SC=round(a[p]["SC"], 2),
                    TC=round(a[p]["TC"], 2),
                ))
    # Demo runs must never overwrite the committed real-data results.
    out = ("data/processed/simulation_results.csv" if use_real
           else "data/processed/simulation_results_demo.csv")
    with open(out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader(); w.writerows(rows)
    print(f"\nSaved full simulation grid to {out}  ({len(rows)} rows)")


if __name__ == "__main__":
    main()
