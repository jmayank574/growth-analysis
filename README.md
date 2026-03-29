# Growth Analytics — Olist Marketing Funnel Analysis

End-to-end growth analytics project covering multi-channel attribution, funnel analysis, CAC/LTV modeling, and seller segmentation using real e-commerce data.

## Dataset
- [Marketing Funnel by Olist](https://www.kaggle.com/datasets/olistbr/marketing-funnel-olist) — 8,000 MQL leads with acquisition channel and conversion data
- [Brazilian E-Commerce by Olist](https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce) — 100,000 orders across sellers, customers, payments, and reviews

## What was built

| Stage | Notebook | Description |
|-------|----------|-------------|
| 1 | `01_data_foundation.ipynb` | Load, clean, and join all 11 datasets into a master table |
| 2 | `02_channel_attribution.ipynb` | Last-touch, linear, and Markov chain attribution models |
| 3 | `03_funnel_analysis.ipynb` | Funnel drop-off, cohort analysis, landing page A/B test |
| 4 | `04_cac_ltv.ipynb` | CAC/LTV by channel and segment, Pareto analysis |
| 5 | `05_segmentation.ipynb` | RFM scoring and K-means clustering with lookalike profile |
| 6 | `06_dashboard.py` | 4-tab Streamlit dashboard with live filters |
| 7 | `07_sql_analysis.ipynb` | 10 DuckDB SQL queries — CTEs, window functions, NTILE |

## Automation
```bash
python run_pipeline.py
```
Executes all 7 stages sequentially in ~51 seconds. Logs every run with timestamps to `logs/`.

```bash
streamlit run 06_dashboard.py
```
Launches the interactive dashboard locally.

## Key findings
- 10.5% overall conversion rate — 842 of 8,000 leads became sellers
- 54.9% activation gap — converted sellers who never placed an order
- Top 22.4% of sellers drive 80% of platform revenue
- Email undervalued 5x by last-touch vs linear attribution
- Paid search and organic search produce 62% of high-value sellers

## Stack
Python · pandas · scikit-learn · DuckDB · Streamlit · plotly
