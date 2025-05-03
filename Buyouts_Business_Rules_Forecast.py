"""
Buyouts Forecast & Safety Stock Automation Script
─────────────────────────────────────────
This script automates demand classification, forecasting, and safety stock calculations
for SKUs using historical sales data and lead time inputs.

Overview:
1. Reads historical demand and lead time data from Excel files.
2. Segments SKUs based on demand patterns (e.g., High Run-Rate, True Buy-Out).
3. Calculates Safety Stock using 12-month median usage and item-specific lead times.
4. Applies statistical forecasting models:
   - ETS for continuous demand SKUs
   - TSB for intermittent demand SKUs
5. Outputs a formatted Excel file with three sheets (for executive use):
   - 'Individual': Historical demand with safety stock
   - 'Overrides' : Model-based forecast overrides by month
   - 'Adjusted'  : Final adjusted forecast including safety stock

Libraries used: pandas, numpy, statsforecast, datetime, xlsxwriter
Forecast horizon: 20 months (dynamic from run date)
"""


import os
import numpy as np
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from statsforecast import StatsForecast

pd.set_option('future.no_silent_downcasting', True)
try:
    from statsforecast.models import AutoETS as ETS, TSB
except ImportError:
    from statsforecast.models import ETS, TSB

# User Configuration
RAW_FILE       = #link to file path
LEADTIME_FILE  = #link to file path
OUT_FILE       = #link to file path
RUN_DATE       = datetime.today()
ID_COLS        = ["SKU","Description"]
HORIZON        = 20
CONT_RATIO     = 0.80
MIN_HITS       = 4
CAP_MULT       = 1.00

# Header colors
COL_2024 = '#C5D9F1'
COL_2025 = '#D8E4BC'
COL_2026 = '#92CDDC'
COL_FREQ = '#F2DCDB'
# ────────────────────────────────────

# 1 - READ RAW & CALCULATE ANNUAL FREQUENCIES
raw = pd.read_excel(RAW_FILE)
hits = raw[raw['Hist_Value'] > 0]
freq = (
    hits
    .groupby(ID_COLS + ["Hist_Year"])
    .size()
    .unstack(fill_value=0)
)
for yr in (2024,2025,2026):
    if yr not in freq.columns:
        freq[yr] = 0
freq = freq[[2024,2025,2026]]
freq.columns = ['2024 Frequency','2025 Frequency','2026 Frequency']
df_freq = freq.reset_index()

# 2- LOAD & PREP HISTORICAL DATA
df = raw.drop(columns=['Total'], errors='ignore')
df = df.rename(columns={'Hist_Year':'Year','Hist_Period':'Month','Hist_Value':'Qty'})
df['Date'] = pd.to_datetime(dict(year=df.Year, month=df.Month, day=1))
df = df.sort_values(['SKU','Date'])

# define run‐month and dynamic horizon
first_run = RUN_DATE.replace(day=1)
last_run  = first_run + relativedelta(months=HORIZON-1)

hist      = df[df.Date <= first_run].copy()
hist_long = hist.assign(Period=hist.Date.dt.strftime('%Y-%b'))

# 3 - SEGMENTATION & BUSINESS RULES (2024 calendar year)
trail = hist[hist.Date.dt.year == 2024].copy()
import itertools
stats = (
    trail.assign(hit=lambda d: d.Qty>0)
         .groupby('SKU')
         .agg(
             hits=('hit','sum'),
             periods=('hit','size'),
             median_qty=('Qty', lambda s: s[s>0].median() if (s>0).any() else 0),
             max_qty=('Qty','max'),
             total_qty=('Qty','sum')
         )
         .reset_index()
)
stats['hit_ratio'] = stats.hits / stats.periods
stats['adi']       = stats.periods / stats.hits.replace(0,1)
stats['max_run']   = (
    trail.groupby('SKU')['Qty']
         .apply(lambda arr: max((len(list(g)) for v,g in itertools.groupby(arr>0) if v), default=0))
)
tail = hist[(hist.Date >= first_run - relativedelta(months=3)) & (hist.Date <= first_run)]
stats = stats.merge(
    tail.groupby('SKU')['Qty'].apply(lambda s: (s>0).all()).rename('last3_tail'),
    on='SKU', how='left'
).fillna({'last3_tail': False})

def classify(r):
    if (r.hits <= 2) or (r.total_qty < 10):
        return 'True Buy-Out'
    if (r.hits < MIN_HITS) and r.last3_tail:
        return 'Emerging Buy-Out'
    if r.hit_ratio >= CONT_RATIO:
        return 'High Run-Rate'
    if 4 <= r.hits <= 9 and r.max_qty/r.median_qty <= 3 and r.max_run <= 3:
        return 'Seasonal Buy-Out'
    if 1 <= r.hits <= 3 and r.max_qty/r.median_qty > 3:
        return 'Spike-Driven Buy-Out'
    return 'Intermittent Buy-Out'

stats['Business Rule'] = stats.apply(classify, axis=1)

#  Read lead‐time and compute Safety Stock
lead = pd.read_excel(LEADTIME_FILE)
lead = lead[['SKU','Leadtime Level']].copy()
lead['Leadtime_Days']   = lead['Leadtime Level'] 
lead['Leadtime_Weeks']  = lead['Leadtime_Days']   / 7
lead['Leadtime_Months'] = lead['Leadtime_Days']   / 30

wins_med = (
    trail.groupby('SKU')['Qty']
         .apply(lambda s: s[s>0].tail(12).median() if (s>0).any() else 0)
         .rename('Median12')
)
stats = (
    stats
    .merge(wins_med, on='SKU', how='left')
    .merge(lead,    on='SKU', how='left')
    .fillna({'Median12':0,'Leadtime_Months':0})
)
stats['Safety Stock'] = (stats['Median12'] * stats['Leadtime_Months']).round(0).astype(int)

# ##To zero everything out
# mask = stats['Business Rule']=='True Buy-Out'
# stats.loc[mask, 'Safety Stock'] = 0

# ── ZERO‐OUT for one‐off SKUs ──────────────────────────────
one_off_skus = df_freq.loc[
    (df_freq['2024 Frequency']<=2)
  & (df_freq['2025 Frequency']==0)
  & (df_freq['2026 Frequency']==0),
  'SKU'
].tolist()

# ── cap SS at the largest-ever month for True Buy-Outs ──────
mask = stats['Business Rule'] == 'True Buy-Out'
# stats already has a 'max_qty' column from the initial aggregation
stats.loc[mask, 'Safety Stock'] = stats.loc[mask, ['Safety Stock','max_qty']].min(axis=1).astype(int)

stats.loc[
    stats['SKU'].isin(one_off_skus),
    'Safety Stock'
] = 0
# ───────────────────────────────────────────────────────────

all_skus = pd.DataFrame({'SKU': hist['SKU'].unique()})
stats = (
    all_skus
    .merge(stats[['SKU','Business Rule','Safety Stock']], on='SKU', how='left')
    .assign(
        **{
            'Business Rule': lambda d: d['Business Rule'].fillna('True Buy-Out'),
            'Safety Stock' : lambda d: d['Safety Stock'].fillna(0).astype(int)
        }
    )
)
cat_map = stats.set_index('SKU')['Business Rule'].to_dict()

# 4 - FORECAST & OVERRIDES
series = (
    hist_long.rename(columns={'SKU':'unique_id','Qty':'y'})[['unique_id','Date','y']]
    .rename(columns={'Date':'ds'})
    .sort_values(['unique_id','ds'])
)
def safe_forecast(model, skus, col):
    recs = []
    if not skus:
        return pd.DataFrame(columns=['unique_id','ds',col])
    try:
        sf = StatsForecast(models=[model], freq='MS', n_jobs=1)
        fc = sf.forecast(df=series[series.unique_id.isin(skus)], h=HORIZON)
        raw = next(c for c in fc if c.lower().startswith(col))
        for r in fc.itertuples():
            recs.append({'unique_id':r.unique_id,'ds':r.ds,col:getattr(r,raw)})
    except NotImplementedError:
        # fallback over the same dynamic horizon
        periods = pd.date_range(first_run, last_run, freq='MS')
        for sku in skus:
            for dt in periods:
                recs.append({'unique_id':sku,'ds':dt,col:wins_med.get(sku,0)})
    return pd.DataFrame(recs, columns=['unique_id','ds',col])

ets_skus = [s for s,c in cat_map.items() if c=='High Run-Rate']
tsb_skus = [s for s,c in cat_map.items() if c in ('Spike-Driven Buy-Out','Intermittent Buy-Out')]
fc_ets   = safe_forecast(ETS(season_length=12), ets_skus, 'ets')
fc_tsb   = safe_forecast(TSB(alpha_d=0.3,alpha_p=0.3), tsb_skus, 'tsb')

# dynamic periods for both sheets
periods = pd.date_range(first_run, last_run, freq='MS')
period_cols = [dt.strftime('%Y-%b') for dt in periods]

grid = pd.DataFrame([{'SKU':sku,'ds':dt} for sku in cat_map for dt in periods])
grid['Period'] = grid.ds.dt.strftime('%Y-%b')
grid = (
    grid
    .merge(fc_ets.rename(columns={'unique_id':'SKU'}), on=['SKU','ds'], how='left')
    .merge(fc_tsb.rename(columns={'unique_id':'SKU'}), on=['SKU','ds'], how='left')
    .merge(wins_med.rename('Median12'), left_on='SKU', right_index=True, how='left')
)

def apply_override(r):
    rule = cat_map[r.SKU]
    if   rule in ('True Buy-Out', 'Emerging Buy-Out'):      
        return 0.0
    last = wins_med.get(r.SKU,0)
    med12 = r.Median12 or 0
    # elif rule=='Emerging Buy-Out':
    #     tail3 = hist[(hist.SKU==r.SKU) &
    #                  (hist.Date >= first_run-relativedelta(months=3)) &
    #                  (hist.Date <= first_run)]
    #     base = tail3.Qty.mean() if not tail3.empty else 0.0
    if rule=='High Run-Rate':     base = r.ets
    elif rule=='Seasonal Buy-Out':  base = last
    else:                           base = r.tsb

    # ← no longer zero‐out the run month

    if rule not in ('True Buy-Out','Seasonal Buy-Out'):
        capv = (2*med12) if rule=='Spike-Driven Buy-Out' else med12
        base = min(base, capv*CAP_MULT, last*CAP_MULT)
    return base

grid['Override'] = grid.apply(apply_override, axis=1).fillna(0)

# bring back Description
grid = grid.merge(df[ID_COLS].drop_duplicates(), on="SKU", how="left")

# 5 - BUILD INDIVIDUAL SHEET **with Safety Stock**
df_long = df.assign(Period=df.Date.dt.strftime('%Y-%b'))

# ← original static 2024–2026 range
months = pd.date_range('2024-01-01','2026-12-01', freq='MS')
period_cols_ind = [dt.strftime('%Y-%b') for dt in months]

hw = (
    df_long
      .pivot_table(
          index=ID_COLS,
          columns='Period',
          values='Qty',
          aggfunc='first',
          fill_value=0
      )
      .reindex(columns=period_cols_ind, fill_value=0)
      .reset_index()
      .merge(stats[['SKU','Business Rule','Safety Stock']], on='SKU', how='left')
      .merge(df_freq, on=ID_COLS, how='left')
)

hw = hw[
    ID_COLS
    + ['Business Rule','Safety Stock','2024 Frequency','2025 Frequency','2026 Frequency']
    + period_cols_ind
].fillna(0)

# 6 - WRITE WITH FORMATTING 
with pd.ExcelWriter(OUT_FILE, engine='xlsxwriter') as writer:
    wb   = writer.book
    fmt24 = wb.add_format({'bg_color':COL_2024,'bold':True})
    fmt25 = wb.add_format({'bg_color':COL_2025,'bold':True})
    fmt26 = wb.add_format({'bg_color':COL_2026,'bold':True})
    fmtf  = wb.add_format({'bg_color':COL_FREQ,'bold':True})
    fmtd  = wb.add_format({'bold':True})

    # Individual — exactly as before
    hw.to_excel(writer, sheet_name='Individual', startrow=1, header=False, index=False)
    ws = writer.sheets['Individual']
    for idx, col in enumerate(hw.columns):
        if col.endswith('Frequency'):
            fmt = fmtf
        elif col.startswith('2024-'):
            fmt = fmt24
        elif col.startswith('2025-'):
            fmt = fmt25
        elif col.startswith('2026-'):
            fmt = fmt26
        else:
            fmt = fmtd
        ws.write(0, idx, col, fmt)

    # Overrides (dynamic)
    ov = (
    grid
    .merge(stats[['SKU','Business Rule','Safety Stock']], on='SKU', how='left')
    [['SKU','Description','Business Rule','Safety Stock','Period','Override']]
    )   
    ow = (
    ov
    .pivot_table(
        index=['SKU','Description','Business Rule','Safety Stock'],
        columns='Period',
        values='Override',
        aggfunc='first',
        fill_value=0
    )
    .reindex(columns=period_cols, fill_value=0)
    .reset_index()
)

    # write Overrides
    ow.to_excel(writer, sheet_name='Overrides', startrow=1, header=False, index=False)
    ws2 = writer.sheets['Overrides']
    for idx, col in enumerate(ow.columns):
        if col.startswith('2024-'):
            fmt = fmt24
        elif col.startswith('2025-'):
            fmt = fmt25
        else:
            fmt = fmt26 if col.startswith('2026-') else fmtd
        ws2.write(0, idx, col, fmt)

    # Adjusted (dynamic)
    adjusted = ow[['SKU','Description', 'Business Rule', 'Safety Stock'] + period_cols]
    adjusted.to_excel(writer,
                      sheet_name='Adjusted',
                      startrow=1,
                      header=False,
                      index=False)
    ws3 = writer.sheets['Adjusted']
    for idx, col in enumerate(adjusted.columns):
        if col in period_cols:
            if col.startswith(f"{first_run.year}-"):
                fmt = fmt24
            elif col.startswith(f"{first_run.year+1}-"):
                fmt = fmt25
            else:
                fmt = fmt26
        else:
            fmt = fmtd
        ws3.write(0, idx, col, fmt)

print("Written:", OUT_FILE)

