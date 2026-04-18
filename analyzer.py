"""
analyzer.py
Core analysis engine for Stock Performance Analyzer.
"""
import pandas as pd
import numpy as np
from datetime import datetime

REQUIRED_COLS = {"ITEM NAME", "QTY AT HAND", "QTY SOLD", "LAST SALES DATE"}

BUCKET_CONFIG = [
    ("0-7 days",    "Active",    0,   7,   "Fast mover - high reorder priority",       "#E2EFDA", "#375623"),
    ("8-14 days",   "Active",    8,  14,   "Healthy mover - monitor stock",             "#EBF5E0", "#375623"),
    ("15-30 days",  "Active",   15,  30,   "Recent sale - watch velocity",              "#F2F8EC", "#375623"),
    ("31-60 days",  "Slow",     31,  60,   "Slowing - consider promotion",              "#FFF2CC", "#7D5A00"),
    ("61-90 days",  "Slow",     61,  90,   "At risk - promote or bundle now",           "#FFEB9C", "#7D5A00"),
    ("91-120 days", "Dormant",  91, 120,   "Dormant - discount to clear",               "#FCE4D6", "#843C0C"),
    ("121-180 days","Dormant", 121, 180,   "Deeply dormant - urgent clearance",         "#F8CBAD", "#843C0C"),
    ("181-270 days","Near Dead",181, 270,  "Near dead - liquidate immediately",         "#F4CCCC", "#7F0000"),
    ("271-365 days","Near Dead",271, 365,  "Critical - last chance to recover cost",    "#EAAAA0", "#7F0000"),
    ("366+ days",   "Near Dead",366,9999,  "Write-off risk - supplier return/discard",  "#D9534F", "#FFFFFF"),
]


def detect_columns(df):
    mapping = {}
    canon = {c.upper().strip(): c for c in df.columns}
    targets = {
        "ITEM NAME":       ["ITEM NAME", "PRODUCT NAME", "PRODUCT", "NAME", "DESCRIPTION", "ITEM"],
        "QTY AT HAND":     ["QTY AT HAND", "QTY ON HAND", "QUANTITY AT HAND", "STOCK QTY", "ON HAND", "STOCK"],
        "QTY SOLD":        ["QTY SOLD", "QUANTITY SOLD", "SOLD QTY", "UNITS SOLD", "SALES QTY"],
        "LAST SALES DATE": ["LAST SALES DATE", "LAST SALE DATE", "LAST SOLD", "LAST DATE", "DATE"],
        "UNIT COST PRICE": ["UNIT COST PRICE", "UNIT COST", "COST PRICE", "COST", "PRICE", "UNIT PRICE"],
        "TOTAL COST":      ["TOTAL COST", "TOTAL", "COST TOTAL"],
        "BARCODE":         ["BARCODE", "BAR CODE", "SKU", "CODE", "PRODUCT CODE"],
        "WAREHOUSE":       ["WAREHOUSE", "STORE", "LOCATION", "BRANCH"],
        "DATE RANGE":      ["DATE RANGE", "PERIOD", "RANGE"],
    }
    for key, aliases in targets.items():
        for alias in aliases:
            if alias.upper() in canon:
                mapping[key] = canon[alias.upper()]
                break
    return mapping


def load_and_prepare(df, snapshot_date=None):
    col_map = detect_columns(df)
    missing = REQUIRED_COLS - set(col_map.keys())
    if missing:
        raise ValueError(
            f"Could not find required columns: {missing}. "
            f"Please ensure your file has: ITEM NAME, QTY AT HAND, QTY SOLD, LAST SALES DATE"
        )
    rename = {v: k for k, v in col_map.items()}
    df = df.rename(columns=rename).copy()
    df = df.dropna(subset=["ITEM NAME"])

    for col in ["QTY AT HAND", "QTY SOLD"]:
        df[col] = pd.to_numeric(
            df[col].astype(str).str.replace(",", "").str.strip(), errors="coerce"
        ).fillna(0)

    if "UNIT COST PRICE" in df.columns:
        df["UNIT_COST"] = pd.to_numeric(
            df["UNIT COST PRICE"].astype(str).str.replace(",", "").str.strip(), errors="coerce"
        ).fillna(0)
    else:
        df["UNIT_COST"] = 0

    df["LAST SALES DATE"] = pd.to_datetime(df["LAST SALES DATE"], errors="coerce")

    snap = pd.Timestamp(snapshot_date) if snapshot_date else pd.Timestamp.today().normalize()
    df["SNAPSHOT_DATE"] = snap
    df["DAYS_SINCE_SALE"] = (snap - df["LAST SALES DATE"]).dt.days

    df["REVENUE"] = df["QTY SOLD"] * df["UNIT_COST"]
    df["CAPITAL_TIED"] = df["QTY AT HAND"].clip(lower=0) * df["UNIT_COST"]
    df["LAST_SALE_STR"] = df["LAST SALES DATE"].dt.strftime("%d/%m/%Y").fillna("Never Sold")
    df["BARCODE_STR"] = df.get(
        "BARCODE", pd.Series([""] * len(df), index=df.index)
    ).astype(str).str.replace("nan", "")

    df["STATUS"] = df.apply(_assign_status, axis=1)
    df["BUCKET_LABEL"] = df.apply(_assign_bucket_label, axis=1)
    df["RISK_LEVEL"] = df["BUCKET_LABEL"].map(_bucket_risk())
    return df


def _assign_status(row):
    if pd.isna(row["LAST SALES DATE"]):
        return "Never Sold"
    d = row["DAYS_SINCE_SALE"]
    if d <= 30:  return "Active"
    if d <= 90:  return "Slow Moving"
    if d <= 180: return "Dormant"
    return "Near Dead"


def _assign_bucket_label(row):
    if pd.isna(row["LAST SALES DATE"]):
        return "Never Sold"
    d = row["DAYS_SINCE_SALE"]
    for label, _, lo, hi, *_ in BUCKET_CONFIG:
        if lo <= d <= hi:
            return label
    return "Never Sold"


def _bucket_risk():
    m = {}
    for label, cat, *_ in BUCKET_CONFIG:
        if cat == "Active":    m[label] = "Low"
        elif cat == "Slow":    m[label] = "Medium"
        elif cat == "Dormant": m[label] = "High"
        else:                  m[label] = "Critical"
    m["Never Sold"] = "Critical"
    return m


def get_summary_metrics(df):
    total     = len(df)
    active    = (df["STATUS"] == "Active").sum()
    slow      = (df["STATUS"] == "Slow Moving").sum()
    dormant   = (df["STATUS"] == "Dormant").sum()
    near_dead = (df["STATUS"] == "Near Dead").sum()
    never_sold= (df["STATUS"] == "Never Sold").sum()
    with_stock= df[df["QTY AT HAND"] > 0]
    neg_stock = (df["QTY AT HAND"] < 0).sum()
    total_rev = df["REVENUE"].sum()
    total_qty = df["QTY SOLD"].sum()
    idle_cap  = df[df["STATUS"].isin(["Slow Moving","Dormant","Near Dead","Never Sold"])]["CAPITAL_TIED"].sum()
    dead_cap  = df[df["STATUS"] == "Never Sold"]["CAPITAL_TIED"].sum()
    top10_rev_share = (df.nlargest(10,"REVENUE")["REVENUE"].sum() / total_rev * 100) if total_rev > 0 else 0

    return {
        "total_skus":      total,
        "active":          int(active),
        "slow_moving":     int(slow),
        "dormant":         int(dormant),
        "near_dead":       int(near_dead),
        "never_sold":      int(never_sold),
        "with_stock":      len(with_stock),
        "negative_stock":  int(neg_stock),
        "total_revenue":   total_rev,
        "total_qty_sold":  total_qty,
        "idle_capital":    idle_cap,
        "dead_capital":    dead_cap,
        "top10_rev_share": top10_rev_share,
        "snapshot_date":   df["SNAPSHOT_DATE"].iloc[0],
    }


def get_monthly_trend(df):
    has_date = df[df["LAST SALES DATE"].notna()].copy()
    has_date["MONTH"] = has_date["LAST SALES DATE"].dt.to_period("M")
    grp = has_date.groupby("MONTH").agg(
        skus=("ITEM NAME","count"),
        qty_sold=("QTY SOLD","sum"),
        revenue=("REVENUE","sum"),
    ).reset_index()
    grp["MONTH_STR"] = grp["MONTH"].astype(str)
    return grp.sort_values("MONTH")


def get_top_products(df, n=20, by="REVENUE"):
    cols = ["ITEM NAME","QTY SOLD","QTY AT HAND","UNIT_COST","REVENUE","LAST_SALE_STR","STATUS","DAYS_SINCE_SALE"]
    cols = [c for c in cols if c in df.columns]
    return df.nlargest(n, by)[cols].reset_index(drop=True)


def get_bucket_summary(df):
    rows = []
    for label, cat, lo, hi, action, bg, tc in BUCKET_CONFIG:
        sub = df[df["DAYS_SINCE_SALE"].between(lo, hi, inclusive="both")]
        ws  = sub[sub["QTY AT HAND"] > 0]
        rows.append({
            "Age Bucket":       label,
            "Category":         cat,
            "Total SKUs":       len(sub),
            "With Stock":       len(ws),
            "Units in Stock":   int(ws["QTY AT HAND"].sum()),
            "Capital Tied (N)": ws["CAPITAL TIED"].sum(),
            "Avg Days Idle":    round(sub["DAYS_SINCE_SALE"].mean(), 1) if len(sub) else 0,
            "Risk Level":       "Low" if cat=="Active" else "Medium" if cat=="Slow" else "High" if cat=="Dormant" else "Critical",
            "Action":           action,
            "_bg":              bg,
            "_tc":              tc,
        })
    ns    = df[df["LAST SALES DATE"].isna()]
    ns_ws = ns[ns["QTY AT HAND"] > 0]
    rows.append({
        "Age Bucket":       "Never Sold",
        "Category":         "Dead Stock",
        "Total SKUs":       len(ns),
        "With Stock":       len(ns_ws),
        "Units in Stock":   int(ns_ws["QTY AT HAND"].sum()),
        "Capital Tied (N)": ns_ws["CAPITAL_TIED"].sum(),
        "Avg Days Idle":    None,
        "Risk Level":       "Critical",
        "Action":           "Liquidate stock / De-list catalogue items",
        "_bg":              "#F2F2F2",
        "_tc":              "#595959",
    })
    return pd.DataFrame(rows)


def get_slow_moving(df):
    sub = df[df["DAYS_SINCE_SALE"] >= 31].copy()
    return sub.sort_values(["DAYS_SINCE_SALE","CAPITAL_TIED"], ascending=[False,False]).reset_index(drop=True)


def get_dead_stock(df):
    sub = df[(df["LAST SALES DATE"].isna()) & (df["QTY AT HAND"] > 0)].copy()
    sub["CAPITAL_TIED"] = sub["QTY AT HAND"].clip(lower=0) * sub["UNIT_COST"]
    sub["RECOMMENDATION"] = sub["CAPITAL_TIED"].apply(
        lambda x: "Liquidate urgently" if x > 50000 else "Promote or bundle" if x > 10000 else "Clear or de-list"
    )
    return sub.sort_values("CAPITAL_TIED", ascending=False).reset_index(drop=True)


def get_near_stockout(df, max_stock=10, min_sold=50):
    mask = (df["QTY AT HAND"] > 0) & (df["QTY AT HAND"] <= max_stock) & (df["QTY SOLD"] >= min_sold)
    sub  = df[mask].copy()
    sub["TURNOVER_RATIO"] = sub["QTY SOLD"] / (sub["QTY SOLD"] + sub["QTY AT HAND"].clip(lower=1))
    return sub.sort_values("TURNOVER_RATIO", ascending=False).reset_index(drop=True)


def get_negative_stock(df):
    return df[df["QTY AT HAND"] < 0].copy().sort_values("QTY AT HAND").reset_index(drop=True)


def get_all_stock_by_idle(df):
    sub = df[df["QTY AT HAND"] > 0].copy()
    sub["_sort"] = sub["DAYS_SINCE_SALE"].fillna(99999)
    return sub.sort_values(["_sort","CAPITAL_TIED"], ascending=[False,False]).drop(columns=["_sort"]).reset_index(drop=True)
