"""
Falsification tests for the (R, s, S) simulation engine — Chapter 5 §5.3.

Three stress conditions at the central case (L = 14, target SL = 95%):
  1. Safety stock forced to zero
  2. Realised demand scaled ×3
  3. Replenishment suppressed (very long lead time, no orders arrive in test window)

Each should produce stockouts and a positive SC, demonstrating that the engine
measures stockouts correctly and that the baseline result of 100% realised SL
reflects genuinely-met demand rather than a stockout counter that never fires.
"""
import numpy as np
import inventory_simulation as sim


def _simulate_with_overrides(test_demand, d_bar, safety_stock, price, lead_time,
                              initial_on_hand=None, allow_replenishment=True,
                              R=None):
    """
    Local copy of simulate_policy with two extra knobs for stress testing:
      - initial_on_hand: override the starting inventory level
      - allow_replenishment: if False, no orders are ever placed
    Returns the same dict as sim.simulate_policy.
    """
    if R is None:
        R = sim.REVIEW_PERIOD_R
    L = lead_time
    horizon = len(test_demand)

    unit_cost = (1.0 - sim.GROSS_MARGIN) * price
    h = unit_cost * sim.CARRYING_RATE
    D_annual = max(d_bar, 1e-9) * sim.DAYS_PER_YEAR
    Q = np.sqrt(2.0 * D_annual * sim.ORDERING_COST_K / max(h, 1e-9))

    reorder_point_s = d_bar * (R + L) + safety_stock
    order_up_to_S = reorder_point_s + Q

    on_hand = order_up_to_S if initial_on_hand is None else float(initial_on_hand)
    in_transit, stockout_units, served_total, demand_total, on_hand_accum, num_orders = {}, 0.0, 0.0, 0.0, 0.0, 0

    for t in range(horizon):
        if t in in_transit:
            on_hand += in_transit.pop(t)
        d = float(test_demand[t])
        demand_total += d
        served = min(on_hand, d)
        served_total += served
        stockout_units += (d - served)
        on_hand -= served
        on_hand_accum += on_hand
        if allow_replenishment and t % R == 0:
            inv_position = on_hand + sum(in_transit.values())
            if inv_position <= reorder_point_s:
                order_qty = order_up_to_S - inv_position
                if order_qty > 0:
                    arrival = t + L
                    in_transit[arrival] = in_transit.get(arrival, 0.0) + order_qty
                    num_orders += 1

    avg_on_hand = on_hand_accum / horizon
    achieved_sl = served_total / demand_total if demand_total > 0 else 1.0
    f = sim.DAYS_PER_YEAR / horizon
    return dict(
        achieved_sl=achieved_sl, stockout_units=stockout_units,
        annual_stockout_units=stockout_units * f,
        avg_on_hand=avg_on_hand, num_orders=num_orders,
        Q=Q, reorder_point=reorder_point_s, order_up_to=order_up_to_S,
        safety_stock=safety_stock, unit_cost=unit_cost,
        annual_HC=avg_on_hand * h, annual_OC=num_orders * f * sim.ORDERING_COST_K,
    )


def run_stress(series_list, mode, lead_time=14, m=1.0):
    """Run all three policies under a stress mode. Returns aggregate dict."""
    agg = {p: dict(HC=0.0, OC=0.0, SC=0.0, ss=0.0, su=0.0,
                   sl_weighted=0.0, dem=0.0) for p in ["CLASSICAL", "ML_GAUSSIAN", "ML_EMP_QUANTILE"]}

    for s in series_list:
        ss_cls = sim.safety_stock_gaussian(s.sigma_classical, lead_time)
        ss_ml  = sim.safety_stock_gaussian(s.sigma_ml, lead_time)
        ss_emp = sim.safety_stock_empirical(s.q50, s.q95, lead_time)

        test_demand = s.test_demand
        kwargs_extra = {}

        if mode == "zero_ss":
            # Force a true zero-buffer condition: no SS *and* start empty
            # so the (R,s,S) policy has nothing to draw on while orders arrive.
            ss_cls = ss_ml = ss_emp = 0.0
            kwargs_extra = dict(initial_on_hand=0.0)
        elif mode == "demand_x3":
            test_demand = s.test_demand * 3.0
        elif mode == "long_lead_time":
            # No orders arrive within the test window and start empty too,
            # exposing the policy to demand with no resupply available.
            kwargs_extra = dict(initial_on_hand=0.0, allow_replenishment=False)
        else:
            raise ValueError(mode)

        runs = {
            "CLASSICAL":       _simulate_with_overrides(test_demand, s.level_forecast, ss_cls, s.price, lead_time, **kwargs_extra),
            "ML_GAUSSIAN":     _simulate_with_overrides(test_demand, s.level_forecast, ss_ml,  s.price, lead_time, **kwargs_extra),
            "ML_EMP_QUANTILE": _simulate_with_overrides(test_demand, s.level_forecast, ss_emp, s.price, lead_time, **kwargs_extra),
        }
        dem = float(test_demand.sum())
        for p, r in runs.items():
            sc = r["annual_stockout_units"] * r["unit_cost"] * m
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


def fmt_row(label, agg, base_tc=None):
    p = "CLASSICAL"  # all three are similar; report Classical for the table
    a = agg[p]
    sl = a["achieved_sl"] * 100
    sc = a["SC"]
    su = a["su"]
    tc = a["TC"]
    delta = "" if base_tc is None else f"  Δ vs baseline TC: {(tc - base_tc):+,.0f}"
    return f"{label:<42}  realised SL = {sl:5.1f}%   stockout cost SC = ${sc:>11,.0f}   unmet units = {su:>9,.0f}   total ${tc:>11,.0f}{delta}"


def main():
    series = sim.load_real_dataset()
    total_demand = sum(s.test_demand.sum() for s in series)

    print("=" * 100)
    print("FALSIFICATION TESTS — Engine validation for Chapter 5 §5.3")
    print(f"502 series | L = 14 d | target SL = 95% | m = 1.0 | total test demand = {total_demand:,.0f} units")
    print("=" * 100)

    # Baseline (no stress) — for reference
    baseline = sim.run_portfolio(series, lead_time=14, stockout_multiplier_m=1.0)
    base_tc = baseline["CLASSICAL"]["TC"]
    print(fmt_row("Real buffers (baseline)", baseline))

    # Stress 1: SS = 0
    s1 = run_stress(series, mode="zero_ss")
    print(fmt_row("Safety stock set to zero", s1, base_tc))

    # Stress 2: realised demand x3
    s2 = run_stress(series, mode="demand_x3")
    print(fmt_row("Realised demand scaled x3", s2, base_tc))

    # Stress 3: lead time so long no orders arrive
    s3 = run_stress(series, mode="long_lead_time")
    print(fmt_row("Replenishment suppressed (L extended)", s3, base_tc))

    print()
    print("=" * 100)
    print("PER-POLICY VIEW under demand x3 stress")
    print("=" * 100)
    for p in ["CLASSICAL", "ML_GAUSSIAN", "ML_EMP_QUANTILE"]:
        a = s2[p]
        sl = a["achieved_sl"] * 100
        print(f"  {p:<22}  realised SL = {sl:5.1f}%   SC = ${a['SC']:>10,.0f}   unmet = {a['su']:>8,.0f}   TC = ${a['TC']:>11,.0f}")
    qc = s2["CLASSICAL"]["TC"]
    qq = s2["ML_EMP_QUANTILE"]["TC"]
    qg = s2["ML_GAUSSIAN"]["TC"]
    print(f"\n  ML-Quantile vs Classical: {100*(qc - qq)/qc:+.1f}%")
    print(f"  ML-Gaussian vs Classical: {100*(qc - qg)/qc:+.1f}%")
    print()
    print("If the x3 stress flips the comparison toward the quantile policy, that is the regime")
    print("the §5.6 boundary-condition argument predicts. If it does not, the boundary is even")
    print("further out than x3, which itself is a defensible finding.")
    print("=" * 100)


if __name__ == "__main__":
    main()
