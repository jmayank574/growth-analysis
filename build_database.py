import duckdb
import pandas as pd
import os
import logging
import sys
from datetime import datetime

PROJECT_ROOT = r"C:\Users\Mayank Joshi\Downloads\Marketing_Channel_Project"
DATA_PATH    = os.path.join(PROJECT_ROOT, "data")
DB_PATH      = os.path.join(PROJECT_ROOT, "olist_analytics.db")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger("build_database")

# ── TABLE DEFINITIONS ─────────────────────────────────────────────────────────
TABLES = [
    {
        "name":    "mql",
        "file":    "olist_marketing_qualified_leads_dataset.csv",
        "desc":    "Marketing qualified leads — top of funnel"
    },
    {
        "name":    "closed_deals",
        "file":    "olist_closed_deals_dataset.csv",
        "desc":    "Converted sellers — closed deals"
    },
    {
        "name":    "orders",
        "file":    "olist_orders_dataset.csv",
        "desc":    "All platform orders with status and timestamps"
    },
    {
        "name":    "order_items",
        "file":    "olist_order_items_dataset.csv",
        "desc":    "Line items per order — price and seller"
    },
    {
        "name":    "order_payments",
        "file":    "olist_order_payments_dataset.csv",
        "desc":    "Payment method and value per order"
    },
    {
        "name":    "order_reviews",
        "file":    "olist_order_reviews_dataset.csv",
        "desc":    "Customer review scores and comments"
    },
    {
        "name":    "customers",
        "file":    "olist_customers_dataset.csv",
        "desc":    "Customer profiles and locations"
    },
    {
        "name":    "sellers",
        "file":    "olist_sellers_dataset.csv",
        "desc":    "Seller profiles and locations"
    },
    {
        "name":    "products",
        "file":    "olist_products_dataset.csv",
        "desc":    "Product catalog with categories and dimensions"
    },
    {
        "name":    "category_translation",
        "file":    "product_category_name_translation.csv",
        "desc":    "Portuguese to English category name translation"
    },
    {
        "name":    "geolocation",
        "file":    "olist_geolocation_dataset.csv",
        "desc":    "Zip code to lat/lng mapping"
    },
    {
        "name":    "master",
        "file":    "master_table.csv",
        "desc":    "Unified master table — all stages joined"
    },
]

ANALYTICAL_TABLES = [
    {
        "name": "attribution_comparison",
        "file": os.path.join("channel_attribution", "attribution_comparison.csv"),
        "desc": "Last touch vs linear vs Markov attribution"
    },
    {
        "name": "channel_funnel",
        "file": os.path.join("funnel_analysis", "channel_funnel.csv"),
        "desc": "End-to-end funnel metrics by channel"
    },
    {
        "name": "cohort_monthly",
        "file": os.path.join("funnel_analysis", "cohort_monthly.csv"),
        "desc": "Monthly cohort conversion rates"
    },
    {
        "name": "ltv_by_channel",
        "file": os.path.join("cac_ltv", "ltv_by_channel.csv"),
        "desc": "LTV and CAC metrics by acquisition channel"
    },
    {
        "name": "ltv_by_segment",
        "file": os.path.join("cac_ltv", "ltv_by_segment.csv"),
        "desc": "LTV metrics by business segment"
    },
    {
        "name": "activated_sellers",
        "file": os.path.join("cac_ltv", "activated_sellers_ltv.csv"),
        "desc": "Sellers who placed at least one order with LTV data"
    },
    {
        "name": "rfm_scores",
        "file": os.path.join("segmentation", "rfm_scores.csv"),
        "desc": "RFM scores and segments for all active sellers"
    },
    {
        "name": "sellers_segmented",
        "file": os.path.join("segmentation", "sellers_segmented_final.csv"),
        "desc": "K-means cluster assignments for all sellers"
    },
    {
        "name": "high_value_cluster",
        "file": os.path.join("segmentation", "high_value_cluster_final.csv"),
        "desc": "High-value seller cluster — lookalike profile"
    },
]

def build_database():
    log.info("=" * 60)
    log.info("BUILDING OLIST ANALYTICS DATABASE")
    log.info(f"Database path : {DB_PATH}")
    log.info(f"Start time    : {datetime.now().strftime('%d %B %Y, %H:%M:%S')}")
    log.info("=" * 60)

    # Remove existing database to rebuild clean
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        log.info("Removed existing database — rebuilding clean")

    con = duckdb.connect(DB_PATH)

    # ── LOAD RAW SOURCE TABLES ─────────────────────────────────────────────
    log.info("")
    log.info("Loading raw source tables:")
    source_success = 0
    for table in TABLES:
        file_path = os.path.join(DATA_PATH, table["file"])
        try:
            df = pd.read_csv(file_path, low_memory=False)
            con.execute(f"DROP TABLE IF EXISTS {table['name']}")
            con.execute(
                f"CREATE TABLE {table['name']} AS SELECT * FROM df"
            )
            row_count = con.execute(
                f"SELECT COUNT(*) FROM {table['name']}"
            ).fetchone()[0]
            log.info(
                f"  {table['name']:<30} {row_count:>10,} rows  —  {table['desc']}"
            )
            source_success += 1
        except Exception as e:
            log.error(f"  FAILED: {table['name']} — {e}")

    # ── LOAD ANALYTICAL TABLES ─────────────────────────────────────────────
    log.info("")
    log.info("Loading analytical output tables:")
    analytical_success = 0
    for table in ANALYTICAL_TABLES:
        file_path = os.path.join(DATA_PATH, table["file"])
        try:
            df = pd.read_csv(file_path, low_memory=False)
            con.execute(f"DROP TABLE IF EXISTS {table['name']}")
            con.execute(
                f"CREATE TABLE {table['name']} AS SELECT * FROM df"
            )
            row_count = con.execute(
                f"SELECT COUNT(*) FROM {table['name']}"
            ).fetchone()[0]
            log.info(
                f"  {table['name']:<30} {row_count:>10,} rows  —  {table['desc']}"
            )
            analytical_success += 1
        except Exception as e:
            log.error(f"  FAILED: {table['name']} — {e}")

    # ── DATABASE SUMMARY ───────────────────────────────────────────────────
    all_tables = con.execute(
        "SELECT table_name FROM information_schema.tables "
        "WHERE table_schema = 'main' ORDER BY table_name"
    ).fetchall()

    total_size_mb = round(os.path.getsize(DB_PATH) / 1024 / 1024, 1)

    log.info("")
    log.info("=" * 60)
    log.info("DATABASE BUILD COMPLETE")
    log.info("=" * 60)
    log.info(f"  Total tables loaded  : {len(all_tables)}")
    log.info(f"  Source tables        : {source_success} / {len(TABLES)}")
    log.info(f"  Analytical tables    : {analytical_success} / {len(ANALYTICAL_TABLES)}")
    log.info(f"  Database size        : {total_size_mb} MB")
    log.info(f"  Database file        : {DB_PATH}")
    log.info("")
    log.info("  All tables:")
    for t in all_tables:
        log.info(f"    - {t[0]}")
    log.info("=" * 60)

    con.close()
    return True

if __name__ == "__main__":
    build_database()