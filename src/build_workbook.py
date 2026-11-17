"""Build deliverables/airline_cx_study.xlsx: review-level data plus formula-driven analysis tabs."""
import json
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, ScatterChart, Reference, Series
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from analysis import load, DRIVERS, INDIA

ROOT = Path(__file__).resolve().parents[1]
S = json.load(open(ROOT / "results" / "summary.json"))
F = "Arial"
H = Font(name=F, bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="0F3D3E")
B = Font(name=F); BB = Font(name=F, bold=True); BLUE = Font(name=F, color="0000FF"); GREEN = Font(name=F, color="008000")
T = Font(name=F, bold=True, size=14); NOTE = Font(name=F, italic=True, size=9, color="555555")
thin = Border(bottom=Side(style="thin", color="BBBBBB"))
N = None  # last data row


def header(ws, row, vals, col=1):
    for i, v in enumerate(vals):
        c = ws.cell(row=row, column=col + i, value=v); c.font = H; c.fill = HF
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x


def main():
    df = load()
    df["carrier"] = df.airline_name.map(INDIA).fillna("")
    cols = list(DRIVERS)
    wb = Workbook()

    # ---------------- README
    ws = wb.active; ws.title = "README"
    rows = [("Airline Passenger Experience Study: what makes a passenger recommend an airline?", T),
            ("", B),
            ("Source: 41,396 Skytrax user reviews (airlinequality.com), 362 airlines, up to Aug 2015. Public CC0 dataset:", B),
            ("https://github.com/quankiquanki/skytrax-reviews-dataset (commit bd4997c).", B),
            ("", B),
            ("Tabs", BB),
            ("Data: one row per review. Ratings are 1-5; blank = not rated (0 in the raw file).", B),
            ("Touchpoints: mean score, top-2-box and recommend rate by rating band, all live formulas on Data.", B),
            ("Drivers: Shapley key-driver importance (from src/analysis.py) and the importance-performance quadrant.", B),
            ("India_Benchmark: recommend rate, 95% CI and driver scores for 5 Indian carriers vs all airlines.", B),
            ("Client_JetAirways: gap vs benchmark, priority score (importance x shortfall) and a model-based what-if.", B),
            ("Themes: share of reviews mentioning each complaint theme (keyword search in review text).", B),
            ("", B),
            ("Colour key: blue = value pasted from the Python analysis (documented in a comment); black = formula; green = link to another tab.", B),
            ("To re-run for another airline: change the client name in Client_JetAirways!B3 (must match a Data!B value).", B)]
    for i, (v, f) in enumerate(rows, 1):
        ws.cell(row=i, column=1, value=v).font = f
    widths(ws, [120])

    # ---------------- Data
    d = wb.create_sheet("Data")
    hdr = ["Review ID", "Airline", "Indian carrier", "Cabin", "Year", "Recommended (1=yes)"] + [DRIVERS[c] for c in cols] + ["Overall (1-10)"]
    header(d, 1, hdr)
    for r, row in enumerate(df.itertuples(index=False), 2):
        vals = [r - 1, row.airline_name, row.carrier, row.cabin_flown if isinstance(row.cabin_flown, str) else None,
                int(row.year) if pd.notna(row.year) else None, int(row.recommended)] + \
               [None if pd.isna(getattr(row, c)) else int(getattr(row, c)) for c in cols] + \
               [None if pd.isna(row.overall_rating) else int(row.overall_rating)]
        for k, v in enumerate(vals, 1):
            d.cell(row=r, column=k, value=v)
    global N
    N = len(df) + 1
    d.freeze_panes = "A2"; d.auto_filter.ref = f"A1:L{N}"
    widths(d, [10, 24, 14, 16, 7, 12, 11, 11, 13, 13, 12, 10])
    rng = lambda col: f"Data!${col}$2:${col}${N}"
    dcol = {DRIVERS[c]: get_column_letter(7 + i) for i, c in enumerate(cols)}
    REC = rng("F")

    # ---------------- Touchpoints
    t = wb.create_sheet("Touchpoints")
    t["A1"] = "Touchpoint performance and the recommend gap"; t["A1"].font = T
    t["A2"] = f"Base: all {N-1:,} reviews; each touchpoint uses reviews that rated it."; t["A2"].font = NOTE
    header(t, 4, ["Touchpoint", "Reviews rating it", "Mean score (1-5)", "Top-2-box (4-5)", "Recommend if rated 4-5", "Recommend if rated 1-3", "Gap (pts)"])
    for i, name in enumerate(DRIVERS.values(), 5):
        c = rng(dcol[name])
        t.cell(row=i, column=1, value=name).font = B
        t.cell(row=i, column=2, value=f"=COUNT({c})")
        t.cell(row=i, column=3, value=f"=AVERAGE({c})").number_format = "0.00"
        t.cell(row=i, column=4, value=f'=COUNTIF({c},">=4")/B{i}').number_format = "0.0%"
        t.cell(row=i, column=5, value=f'=AVERAGEIFS({REC},{c},">=4")').number_format = "0.0%"
        t.cell(row=i, column=6, value=f'=AVERAGEIFS({REC},{c},">=1",{c},"<=3")').number_format = "0.0%"
        t.cell(row=i, column=7, value=f"=(E{i}-F{i})*100").number_format = "0.0"
    t["A11"] = "All reviews"; t["A11"].font = BB
    t["B11"] = f"=COUNT({REC})"; t["E11"] = f"=AVERAGE({REC})"; t["E11"].number_format = "0.0%"
    t["A12"] = "Note: the gap shows how far recommendation swings with each touchpoint; it is descriptive, not causal."; t["A12"].font = NOTE
    widths(t, [26, 16, 15, 15, 20, 20, 10])

    # ---------------- Drivers
    k = wb.create_sheet("Drivers")
    k["A1"] = "Key-driver analysis: importance vs performance"; k["A1"].font = T
    k["A2"] = (f"Importance = Shapley (LMG) share of R-squared, regressing 'recommended' on the 5 ratings, "
               f"{S['complete_case_reviews']:,} reviews that rated all five; R-squared {S['model_r2']:.2f}.")
    k["A2"].font = NOTE
    header(k, 4, ["Driver", "Importance (share)", "Performance (mean, 1-5)", "Quadrant"])
    imp = S["importance"]
    for i, name in enumerate(DRIVERS.values(), 5):
        k.cell(row=i, column=1, value=name)
        c = k.cell(row=i, column=2, value=imp[name]); c.font = BLUE; c.number_format = "0.0%"
        c.comment = Comment("Pasted from src/analysis.py -> results/summary.json (Shapley R-squared share).", "analysis")
        p = k.cell(row=i, column=3, value=f"=Touchpoints!C{i}"); p.font = GREEN; p.number_format = "0.00"
        k.cell(row=i, column=4, value=(f'=IF(B{i}>=AVERAGE($B$5:$B$9),IF(C{i}<AVERAGE($C$5:$C$9),"Fix first","Protect"),'
                                       f'IF(C{i}<AVERAGE($C$5:$C$9),"Monitor","Maintain"))'))
    k["A10"] = "Average"; k["A10"].font = BB
    k["B10"] = "=AVERAGE(B5:B9)"; k["B10"].number_format = "0.0%"; k["C10"] = "=AVERAGE(C5:C9)"; k["C10"].number_format = "0.00"
    k["A11"] = "Quadrant rule: importance above average and score below average = Fix first; above/above = Protect; below/below = Monitor; below/above = Maintain."
    k["A11"].font = NOTE
    k["A12"] = (f"Stability: driver ranking identical on random halves and on 2008-13 vs 2014-15 reviews "
                f"(Spearman {S['stability_spearman_random_halves']:.2f} and {S['stability_spearman_2008_13_vs_2014_15']:.2f}).")
    k["A12"].font = NOTE
    ch = ScatterChart(); ch.title = "Importance vs performance"; ch.style = 13
    ch.x_axis.title = "Performance (mean score)"; ch.y_axis.title = "Importance (share of R-squared)"
    for i in range(5, 10):
        s = Series(Reference(k, min_col=2, min_row=i), Reference(k, min_col=3, min_row=i), title_from_data=False, title=k.cell(row=i, column=1).value)
        s.marker.symbol = "circle"; s.marker.size = 10; s.graphicalProperties.line.noFill = True
        ch.series.append(s)
    ch.roundedCorners = False; ch.legend.overlay = False; ch.title.overlay = False
    ch.x_axis.delete = False; ch.y_axis.delete = False; ch.x_axis.title.overlay = False; ch.y_axis.title.overlay = False
    ch.width, ch.height = 16, 9
    k.add_chart(ch, "F4")
    widths(k, [24, 18, 22, 14])

    # ---------------- India benchmark
    b = wb.create_sheet("India_Benchmark")
    b["A1"] = "Indian carriers vs all 362 airlines"; b["A1"].font = T
    b["A2"] = "95% CI = 1.96 x sqrt(p(1-p)/n). GoAir and SpiceJet have small samples: read their rates as directional."; b["A2"].font = NOTE
    header(b, 4, ["Carrier", "Reviews", "Recommend rate", "95% CI (+/- pts)"] + list(DRIVERS.values()))
    carriers = ["IndiGo", "Air India", "GoAir", "SpiceJet", "Jet Airways"]
    for i, c in enumerate(carriers, 5):
        b.cell(row=i, column=1, value=c)
        b.cell(row=i, column=2, value=f'=COUNTIF({rng("C")},A{i})')
        b.cell(row=i, column=3, value=f'=AVERAGEIFS({REC},{rng("C")},A{i})').number_format = "0.0%"
        b.cell(row=i, column=4, value=f"=1.96*SQRT(C{i}*(1-C{i})/B{i})*100").number_format = "0.0"
        for j, name in enumerate(DRIVERS.values()):
            b.cell(row=i, column=5 + j, value=f'=AVERAGEIFS({rng(dcol[name])},{rng("C")},$A{i})').number_format = "0.00"
    r = 10
    b.cell(row=r, column=1, value="All 362 airlines").font = BB
    b.cell(row=r, column=2, value=f"=COUNT({REC})")
    b.cell(row=r, column=3, value=f"=AVERAGE({REC})").number_format = "0.0%"
    b.cell(row=r, column=4, value=f"=1.96*SQRT(C{r}*(1-C{r})/B{r})*100").number_format = "0.0"
    for j, name in enumerate(DRIVERS.values()):
        b.cell(row=r, column=5 + j, value=f"=AVERAGE({rng(dcol[name])})").number_format = "0.00"
    bc = BarChart(); bc.type = "bar"; bc.title = "Recommend rate"; bc.style = 10
    bc.add_data(Reference(b, min_col=3, min_row=4, max_row=10), titles_from_data=True)
    bc.set_categories(Reference(b, min_col=1, min_row=5, max_row=10))
    bc.roundedCorners = False; bc.legend = None; bc.title.overlay = False; bc.x_axis.delete = False; bc.y_axis.delete = False
    bc.y_axis.numFmt = "0%"; bc.width, bc.height = 16, 8
    b.add_chart(bc, "A13")
    widths(b, [18, 10, 15, 15, 13, 12, 16, 22, 15])

    # ---------------- Client
    c = wb.create_sheet("Client_JetAirways")
    c["A1"] = "Sample client deep dive"; c["A1"].font = T
    c["A3"] = "Client"; c["B3"] = "jet-airways"; c["B3"].font = BLUE; c["B3"].fill = PatternFill("solid", fgColor="FFFF00")
    c["B3"].comment = Comment("Input: any airline name from Data!B (e.g. air-india, emirates).", "analysis")
    c["A4"] = "Reviews"; c["B4"] = f'=COUNTIF({rng("B")},B3)'
    c["A5"] = "Recommend rate"; c["B5"] = f'=AVERAGEIFS({REC},{rng("B")},B3)'; c["B5"].number_format = "0.0%"
    c["A6"] = "Benchmark (all airlines)"; c["B6"] = "=India_Benchmark!C10"; c["B6"].font = GREEN; c["B6"].number_format = "0.0%"
    c["A7"] = "Gap (pts)"; c["B7"] = "=(B5-B6)*100"; c["B7"].number_format = "0.0"
    header(c, 9, ["Driver", "Client score", "Benchmark score", "Gap", "Importance", "Priority score", "Recommend pts per rating point", "What-if: close gap (pts)"])
    beta = S["client"]["pts_per_rating_point"]
    for i, name in enumerate(DRIVERS.values(), 10):
        c.cell(row=i, column=1, value=name)
        c.cell(row=i, column=2, value=f'=AVERAGEIFS({rng(dcol[name])},{rng("B")},$B$3)').number_format = "0.00"
        g = c.cell(row=i, column=3, value=f"=India_Benchmark!{get_column_letter(5 + i - 10)}10"); g.font = GREEN; g.number_format = "0.00"
        c.cell(row=i, column=4, value=f"=B{i}-C{i}").number_format = "+0.00;-0.00;0.00"
        e = c.cell(row=i, column=5, value=f"=Drivers!B{i - 5}"); e.font = GREEN; e.number_format = "0.0%"
        c.cell(row=i, column=6, value=f"=E{i}*MAX(0,-D{i})").number_format = "0.0000"
        bt = c.cell(row=i, column=7, value=beta[name]); bt.font = BLUE; bt.number_format = "0.000"
        bt.comment = Comment("Linear probability model coefficient on 25,569 complete reviews (src/analysis.py).", "analysis")
        c.cell(row=i, column=8, value=f"=G{i}*MAX(0,-D{i})*100").number_format = "0.0"
    c["A15"] = "Total"; c["A15"].font = BB; c["H15"] = "=SUM(H10:H14)"; c["H15"].number_format = "0.0"
    c["A16"] = "Top priority"; c["A16"].font = BB
    c["B16"] = "=INDEX(A10:A14,MATCH(MAX(F10:F14),F10:F14,0))"
    c["A18"] = ("What-if is model-based and correlational: it estimates how much of the recommend gap the rating gaps account for; "
                "the remainder points to factors outside these five ratings, such as delays.")
    c["A18"].font = NOTE
    widths(c, [24, 14, 16, 9, 12, 14, 20, 20])

    # ---------------- Themes
    th = wb.create_sheet("Themes")
    th["A1"] = "What detractors write about"; th["A1"].font = T
    th["A2"] = "Share of reviews mentioning each theme (keyword search on review text, spot-checked). Lift = not recommended / recommended."; th["A2"].font = NOTE
    header(th, 4, ["Theme", "Not recommended", "Recommended", "Lift", "Indian carriers, not recommended"])
    for i, (name, v) in enumerate(sorted(S["themes"].items(), key=lambda kv: -kv[1]["not_recommended"] / kv[1]["recommended"]), 5):
        th.cell(row=i, column=1, value=name)
        for j, key in enumerate(["not_recommended", "recommended"], 2):
            x = th.cell(row=i, column=j, value=v[key]); x.font = BLUE; x.number_format = "0.0%"
        th.cell(row=i, column=4, value=f"=B{i}/C{i}").number_format = "0.0x"
        x = th.cell(row=i, column=5, value=v["india_not_recommended"]); x.font = BLUE; x.number_format = "0.0%"
    th["B5"].comment = Comment("Pasted from results/summary.json (src/analysis.py THEMES regexes).", "analysis")
    widths(th, [28, 16, 14, 8, 30])

    for w in wb.worksheets:
        for row in w.iter_rows():
            for cell in row:
                if cell.font.name != F:
                    cell.font = cell.font.copy(name=F)
    out = ROOT / "deliverables" / "airline_cx_study.xlsx"
    wb.save(out); print("saved", out)


if __name__ == "__main__":
    main()
