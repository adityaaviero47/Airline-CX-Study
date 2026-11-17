#!/usr/bin/env bash
# Download the Skytrax user reviews dataset (41,396 airline reviews, CC0) from its public GitHub repo.
set -e
mkdir -p data
curl -L --fail -o data/airline.csv \
  "https://raw.githubusercontent.com/quankiquanki/skytrax-reviews-dataset/bd4997cbfce42aeef9281818560e845adc5382f0/data/airline.csv"
echo "e420602d40048a42d938143f45fdeb23a70b57160301c7ded81be91f80c3105a  data/airline.csv" | sha256sum -c -
