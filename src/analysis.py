"""Airline passenger experience study: what makes a passenger recommend an airline?

Data: 41,396 Skytrax user reviews (airlinequality.com, 2008 - Aug 2015, CC0), 362 airlines.
Outputs results/*.json and results/*.csv used by the Excel workbook and the PPT deck.
"""
import itertools, json, math, re
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RES = ROOT / "results"
DRIVERS = {"seat_comfort_rating": "Seat comfort", "cabin_staff_rating": "Cabin staff",
           "food_beverages_rating": "Food & beverages", "inflight_entertainment_rating": "Inflight entertainment",
           "value_money_rating": "Value for money"}
INDIA = {"jet-airways": "Jet Airways", "air-india": "Air India", "air-india-express": "Air India",
         "indigo-airlines": "IndiGo", "spicejet": "SpiceJet", "goair": "GoAir"}
FULL_SERVICE = {"Jet Airways", "Air India"}
THEMES = {  # complaint themes searched in review text (lower-case regex)
    "Delays & cancellations": r"\bdelay|\bdelayed|cancel+ed|cancel+ation|\blate\b|missed (?:my )?connection",
    "Staff attitude": r"\brude\b|unfriendly|unhelpful|impolite|arrogant|attitude",
    "Customer service & refunds": r"customer service|\brefund|compensation|call cent|no response|complain",
    "Baggage": r"baggage|luggage|\bbag\b|\bbags\b",
    "Seat & legroom": r"legroom|leg room|cramped|uncomfortable seat|narrow seat|seat pitch",
    "Food": r"\bfood\b|\bmeal|\bsnack",
    "Fees & charges": r"\bfee\b|\bfees\b|extra charge|hidden charge|charged",
}


def load():
    df = pd.read_csv(ROOT / "data" / "airline.csv")
    for c in DRIVERS:
        df.loc[df[c] == 0, c] = np.nan          # 0 = not rated
    df["year"] = pd.to_datetime(df["date"], errors="coerce").dt.year
    return df


def shapley_r2(df, cols, target="recommended"):
    """Shapley (LMG) split of R-squared among drivers: the standard key-driver method."""
    C = np.corrcoef(df[cols + [target]].values.T)
    rxx, rxy, p = C[:-1, :-1], C[:-1, -1], len(cols)
    r2 = {(): 0.0}
    for k in range(1, p + 1):
        for s in itertools.combinations(range(p), k):
            i = list(s)
            r2[s] = float(rxy[i] @ np.linalg.solve(rxx[np.ix_(i, i)], rxy[i]))
    sh = np.zeros(p)
    for j in range(p):
        for s, v in r2.items():
            if j in s:
                continue
            w = math.factorial(len(s)) * math.factorial(p - len(s) - 1) / math.factorial(p)
            sh[j] += w * (r2[tuple(sorted(s + (j,)))] - v)
    return pd.Series(sh, index=cols), r2[tuple(range(p))]


def main():
    df = load()
    cols = list(DRIVERS)
    cc = df.dropna(subset=cols).copy()            # complete cases for driver analysis
    sh, r2 = shapley_r2(cc, cols)
    imp = (sh / sh.sum()).rename(index=DRIVERS)
    # stability: split-half by review year parity and by random halves
    rng = np.random.default_rng(7)
    half = rng.random(len(cc)) < 0.5
    s1, _ = shapley_r2(cc[half], cols); s2, _ = shapley_r2(cc[~half], cols)
    early, _ = shapley_r2(cc[cc.year <= 2013], cols); late, _ = shapley_r2(cc[cc.year >= 2014], cols)
    perf = df[cols].mean().rename(index=DRIVERS)                       # every review that rated it
    t2b = ((df[cols] >= 4).sum() / df[cols].notna().sum()).rename(index=DRIVERS)
    # recommend rate when a driver is rated 4-5 vs 1-3
    gap = {DRIVERS[c]: [round(float(cc.loc[cc[c] >= 4, "recommended"].mean()), 4),
                        round(float(cc.loc[cc[c] <= 3, "recommended"].mean()), 4)] for c in cols}
    drivers = pd.DataFrame({"importance": imp.round(4), "mean_score": perf.round(3),
                            "top2box": t2b.round(4)})
    drivers["quadrant"] = np.where(drivers.importance >= drivers.importance.mean(),
                                   np.where(drivers.mean_score < drivers.mean_score.mean(), "Fix first", "Protect"),
                                   np.where(drivers.mean_score < drivers.mean_score.mean(), "Monitor", "Maintain"))
    drivers.sort_values("importance", ascending=False).to_csv(RES / "drivers.csv")

    # cabin cut
    cabin = df.dropna(subset=["cabin_flown"]).groupby("cabin_flown").agg(
        reviews=("recommended", "size"), recommend=("recommended", "mean"),
        value=("value_money_rating", "mean"), staff=("cabin_staff_rating", "mean")).round(3)
    cabin.to_csv(RES / "cabin.csv")

    # India benchmark
    df["carrier"] = df.airline_name.map(INDIA)
    ind = df[df.carrier.notna()]
    bench = ind.groupby("carrier").agg(reviews=("recommended", "size"), recommend=("recommended", "mean"),
                                       **{DRIVERS[c]: (c, "mean") for c in cols})
    glob = df.agg({"recommended": "mean", **{c: "mean" for c in cols}})
    bench.loc["All 362 airlines"] = [len(df), glob["recommended"]] + [glob[c] for c in cols]
    bench = bench.round(3).sort_values("recommend", ascending=False)
    bench.to_csv(RES / "india_benchmark.csv")
    ind_fs = ind[ind.carrier.isin(FULL_SERVICE)]
    ind_lcc = ind[~ind.carrier.isin(FULL_SERVICE)]

    # complaint themes in text: share of NOT-recommended reviews vs recommended ones
    txt = df.content.str.lower()
    th = {}
    for name, pat in THEMES.items():
        hit = txt.str.contains(pat, regex=True)
        th[name] = {"not_recommended": round(float(hit[df.recommended == 0].mean()), 4),
                    "recommended": round(float(hit[df.recommended == 1].mean()), 4),
                    "india_not_recommended": round(float(hit[(df.recommended == 0) & df.carrier.notna()].mean()), 4)}
    th = dict(sorted(th.items(), key=lambda kv: -kv[1]["not_recommended"]))
    pd.DataFrame(th).T.to_csv(RES / "themes.csv")

    # sample-client deep dive: Jet Airways vs the 362-airline benchmark
    client = "Jet Airways"
    cl = df[df.carrier == client]
    gaps = (cl[cols].mean() - df[cols].mean()).rename(index=DRIVERS)
    priority = (imp * (-gaps).clip(lower=0)).sort_values(ascending=False)
    # linear probability model on complete cases: recommend-rate points per rating point
    X = np.column_stack([np.ones(len(cc)), cc[cols].values])
    beta = np.linalg.lstsq(X, cc["recommended"].values, rcond=None)[0][1:]
    beta = pd.Series(beta, index=cols).rename(index=DRIVERS)
    whatif = (beta * (-gaps).clip(lower=0))
    n_cl = len(cl); r_cl = cl.recommended.mean()
    client_out = {"client": client, "reviews": n_cl, "recommend": round(float(r_cl), 4),
                  "recommend_ci95": round(float(1.96 * np.sqrt(r_cl * (1 - r_cl) / n_cl)), 4),
                  "gap_vs_benchmark": gaps.round(3).to_dict(),
                  "priority_score": priority.round(4).to_dict(),
                  "pts_per_rating_point": beta.round(4).to_dict(),
                  "whatif_close_gaps_pts": whatif.round(4).to_dict(),
                  "whatif_close_all_gaps_pts": round(float(whatif.sum()), 6),
                  "not_recommended_reviews": int((cl.recommended == 0).sum()),
                  "delay_mention_in_not_recommended": round(float(txt[(df.carrier == client) & (df.recommended == 0)]
                                                                 .str.contains(THEMES["Delays & cancellations"]).mean()), 4)}
    india_ci = {c: [int(n), round(float(m), 4), round(float(1.96 * np.sqrt(m * (1 - m) / n)), 4)]
                for c, (m, n) in ind.groupby("carrier").recommended.agg(["mean", "size"]).iterrows()}
    lift = {k: round(v["not_recommended"] / v["recommended"], 2) for k, v in th.items()}

    out = {
        "reviews": len(df), "airlines": int(df.airline_name.nunique()),
        "years": [int(df.year[df.year > 2000].min()), int(df.year.max())],
        "recommend_rate": round(float(df.recommended.mean()), 4),
        "complete_case_reviews": len(cc), "model_r2": round(r2, 4),
        "importance": imp.round(4).sort_values(ascending=False).to_dict(),
        "stability_spearman_random_halves": round(float(s1.rank().corr(s2.rank(), method="spearman")), 3),
        "stability_spearman_2008_13_vs_2014_15": round(float(early.rank().corr(late.rank(), method="spearman")), 3),
        "top2_driver_share": round(float(imp.sort_values(ascending=False).iloc[:2].sum()), 4),
        "recommend_if_rated_4_5_vs_1_3": gap,
        "india_reviews": len(ind),
        "india_recommend_full_service": round(float(ind_fs.recommended.mean()), 4),
        "india_recommend_low_cost": round(float(ind_lcc.recommended.mean()), 4),
        "india_reviews_full_service": len(ind_fs), "india_reviews_low_cost": len(ind_lcc),
        "themes": th, "theme_lift_not_vs_recommended": lift,
        "india_recommend_ci95": india_ci, "client": client_out,
    }
    json.dump(out, open(RES / "summary.json", "w"), indent=2)
    print(json.dumps({k: v for k, v in out.items() if k not in ("themes",)}, indent=2))
    print(drivers.sort_values("importance", ascending=False)); print(bench); print(cabin); print(pd.DataFrame(th).T)


if __name__ == "__main__":
    main()
