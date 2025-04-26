
#  Buy-outs forecast adjuster

#  1.  Reads a wide Excel sheet (first two columns = SKU, Description;
#      rest are “YYYY-Mon” buckets).
#  2.  Splits each SKU into:
#        – continuous  (≥80 % non-zero in last-12)  ETS level model
#        – intermittent                             TSB model
#        – true buy-out  (<4 hits  or  ADI>9)        0
#  3.  Every future bucket is capped at the tighter of
#        • last-year same month × CAP_MULT
#        • winsorised median of the last-12 months × CAP_MULT
#  4.  Writes two sheets – “Overrides” and “Adjusted” – no duplicated columns.
# ---------------------------------------------------------------

import os, numpy as np, pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from statsforecast import StatsForecast
from statsforecast.models import AutoETS as ETS, TSB   

   

# User Settings
RAW_FILE   = r"C:\path\to\wide_input.xlsx"          # wide sheet
OUT_FILE   = r"C:\path\to\Buyouts_fix_stats.xlsx"   # workbook to write
RUN_DATE   = datetime.today()                       # e.g. 2025-04-25

HORIZON    = 24            # forecast buckets (Apr-25 … Mar-27)
CONT_HIT   = 0.80          # continuous if ≥80 % hits in last-12
ADI_LIMIT  = 9
MIN_HITS   = 4             # <4 hits ⇒ treat as buy-out
CAP_MULT   = 1.00          # 1.00 = hard cap, raise to 1.25 if needed
YEAR_CUTOFF = 2024         # first year to show in output
ID_COLS   = ["SKU", "Description"]
# ---------------------------------------------------------------

df = pd.read_excel(RAW_FILE)
month_cols = df.columns[len(ID_COLS):]
df[month_cols] = df[month_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

col2date  = {c: pd.to_datetime(c, format="%Y-%b") for c in month_cols}
first_run = RUN_DATE.replace(day=1)
trail_start = first_run - relativedelta(months=12)
out_cols = [c for c in month_cols if col2date[c].year >= YEAR_CUTOFF]

long = (df.melt(ID_COLS, month_cols, var_name="Period", value_name="Qty")
          .assign(Date=lambda d: d["Period"].map(col2date),
                  Mo  =lambda d: d["Date"].dt.month,
                  Yr  =lambda d: d["Date"].dt.year))

hist = long[long["Date"] < first_run].copy()
fcst = long[long["Date"] >= first_run].copy()

# Trailing-12 Segmentation
trail = hist[hist["Date"] >= trail_start]
seg   = (trail.assign(hit=lambda d: d["Qty"] > 0)
               .groupby("SKU")
               .agg(hits=('hit','sum'), periods=('hit','size'))
               .reset_index())
seg["hit_ratio"] = seg["hits"] / seg["periods"]
seg["adi"]       = seg["periods"] / seg["hits"].replace(0,1)
seg["buyout"]    = (seg["hits"] < MIN_HITS) | (seg["adi"] > ADI_LIMIT)
seg["continuous"]= (~seg["buyout"]) & (seg["hit_ratio"] >= CONT_HIT)

buyout     = set(seg.loc[seg["buyout"], "SKU"])
continuous = set(seg.loc[seg["continuous"], "SKU"])

#StatsForecast
series = (hist[["SKU","Date","Qty"]]
          .rename(columns={"SKU":"unique_id", "Date":"ds", "Qty":"y"})
          .sort_values(["unique_id","ds"]))

models = [ETS(season_length=1, model="AAN"),
          TSB(alpha_d=0.3, alpha_p=0.3)]

sf = StatsForecast(models=models, freq='MS', n_jobs=1)
fc_df = sf.forecast(df=series, h=HORIZON)

future_cols = fc_df["ds"].dt.strftime("%Y-%b").unique().tolist()
fc_df["Period"] = fc_df["ds"].dt.strftime("%Y-%b")

ets_col = next(c for c in fc_df if c.lower().startswith(("autoets","ets")))
tsb_col = next(c for c in fc_df if c.lower().startswith("tsb"))

def pick(row):
    sku = row["unique_id"]
    if sku in continuous: return row[ets_col]
    if sku in buyout:     return 0.0
    return row[tsb_col]

fc_df["yhat"] = fc_df.apply(pick, axis=1)
fc_df.loc[fc_df["ds"] == first_run, "yhat"] = 0.0   # April-25 = 0

fc_use = fc_df[["unique_id","Period","yhat"]].rename(columns={"unique_id":"SKU"})
fcst   = fcst.merge(fc_use, how="left", on=["SKU","Period"])
fcst["Override"] = fcst["yhat"].fillna(0.0)

# Combine hist + fcst, then cap
long = pd.concat([hist.assign(Override=np.nan), fcst], ignore_index=True)

last_year = hist[hist["Yr"] == first_run.year-1]
ly_tbl = last_year.pivot_table(index="SKU", columns="Mo",
                               values="Qty", aggfunc="first")

def wins_median(s):
    nz = s[s>0].tail(12)
    if nz.empty: return 0
    q1, med, q3 = nz.quantile([.25,.5,.75]); iqr = q3 - q1
    return nz.clip(med-1.5*iqr, med+1.5*iqr).median()

med12 = hist.groupby("SKU")["Qty"].apply(wins_median).rename("Med12")
long  = long.merge(med12, on="SKU", how="left")

long["AdjQty"] = np.where(long["Override"].notna(), long["Override"], long["Qty"])

future = long["Date"] >= first_run
def cap_row(r):
    sku, mo = r["SKU"], r["Mo"]
    ly_cap  = ly_tbl.loc[sku, mo]*CAP_MULT if mo in ly_tbl.columns else np.inf
    med_cap = r["Med12"] * CAP_MULT
    return min(r["AdjQty"], ly_cap, med_cap)

long.loc[future, "AdjQty"] = long.loc[future].apply(cap_row, axis=1)
long.loc[future & (long["AdjQty"] < 1), "AdjQty"] = 0.0
long.drop(columns="Med12", inplace=True)

# Pivot back to wide 
wide = long.pivot(index=ID_COLS, columns="Period", values="AdjQty").reset_index()
keep = list(dict.fromkeys(out_cols + future_cols))   # ensure no duplicates
adj_wide = wide[ID_COLS + [c for c in keep if c in wide.columns]]
ovr_wide = adj_wide.copy()                           # same values for Overrides

# Save 
os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
with pd.ExcelWriter(OUT_FILE, engine="xlsxwriter") as xl:
    ovr_wide.to_excel(xl, sheet_name="Overrides", index=False)
    adj_wide.to_excel(xl, sheet_name="Adjusted", index=False)

print("Workbook written to", OUT_FILE)
