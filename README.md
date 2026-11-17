# Airline Passenger Experience Study: Key Drivers, India Benchmark & Client Deck

**The research question:** what makes a passenger recommend an airline, and what should one airline fix first?

**The data:** 41,396 real passenger reviews of 362 airlines from Skytrax (airlinequality.com), up to Aug 2015, public under CC0 ([source repo](https://github.com/quankiquanki/skytrax-reviews-dataset)). Each review has a yes/no "recommend", 1–5 ratings for 5 touchpoints, the cabin flown and the review text.

**The deliverables (`deliverables/`):**
- `airline_cx_study.pptx`: an 11-slide client deck (PDF copy alongside).
- `airline_cx_study.xlsx`: the review-level data plus formula-driven tabs for touchpoint scores, the importance-performance quadrants, the India benchmark with 95% CIs, and a client deep dive. You can switch the client by changing one cell.

## Findings

**1. Value for money and cabin staff drive recommendation.**
- These two carry **55.4%** of what the ratings explain about recommendation, based on a Shapley (LMG) split of R² on 25,569 reviews that rated all five touchpoints (R² 0.72).

  | Driver | Share of explained recommendation | Mean score (1–5) | Quadrant |
  |---|---|---|---|
  | Value for money | 31.7% | 3.21 | Protect |
  | Cabin staff | 23.8% | 3.37 | Protect |
  | Seat comfort | 17.5% | 3.16 | Maintain |
  | Food & beverages | 16.7% | 3.05 | Monitor |
  | Inflight entertainment | 10.4% | 2.84 | Monitor |

- The driver ranking is identical on random halves of the data and on 2008–13 vs 2014–15 reviews (Spearman 1.0).
- Passengers who rate value 4–5 recommend **93.6%** of the time, against **14.2%** for those who rate it 1–3.

**2. Detractors write about staff attitude, fees and delays.** This comes from a keyword search of the review text, spot-checked by reading the matches. Compared with reviews that recommend the airline, reviews that don't mention:
- rude or unhelpful staff **5.4x** as often;
- fees **2.4x** as often;
- delays and cancellations **2.1x** as often. Delays appear in **41.4%** of these reviews.

**3. In India, low-cost IndiGo leads.**
- IndiGo is recommended **82.6% ± 7.1 pts** of the time (109 reviews).
- Air India is at 57.4% ± 5.2 and Jet Airways at 47.3% ± 3.6, against 53.4% for all airlines.
- Across Indian carriers, low-cost airlines are recommended 67.5% of the time vs 50.6% for full-service ones.
- GoAir (30 reviews) and SpiceJet (61) are directional only.

**4. Sample client, Jet Airways: fix cabin staff first.**
- Jet scores below the benchmark on all five touchpoints.
- Priority = importance × shortfall, which ranks **cabin staff** first.
- A linear what-if suggests that closing all five rating gaps would add about **3.6 pts** of its 6.1-pt recommend gap. The rest lies outside the ratings: 39% of Jet's detractor reviews mention delays.

## Run it
```bash
pip install -r requirements.txt
bash get_data.sh                      # downloads data/airline.csv (34 MB), checksum verified
python src/analysis.py                # results/summary.json, drivers.csv, india_benchmark.csv, themes.csv
python src/build_workbook.py          # deliverables/airline_cx_study.xlsx
node src/build_deck.js                # deliverables/airline_cx_study.pptx (needs: npm install pptxgenjs)
```

## Honest limits
- **Who reviews:** reviewers chose to post online; this isn't a random survey, so opinions skew more extreme.
- **Timing:** the reviews run only to 2015. Jet Airways stopped flying in 2019, so it's a historical sample client.
- **Causality:** importance, the rating swing and the what-if are associations. Proving cause needs a designed survey or a field test.
- **Themes:** they come from keyword search, not a trained classifier, so treat them as signals.
- **Data quality:** I first tried the popular Kaggle "airline passenger satisfaction" survey and dropped it. Its labels behave like rules: personal travellers rating wifi 1–3 are 0% satisfied, and anyone rating wifi 5 is about 100% satisfied.
