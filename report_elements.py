# -*- coding: utf-8 -*-
"""
Shared HTML report components and badge logic.
"""

from __future__ import annotations
import pandas as pd
import numpy as np
import re
import datetime as dt

# --- UI Constants ---
TEXT_DARK  = "#0F172A"
TEXT_MUTED = "#6B7280"
DIVIDER    = "#E5E7EB"
PILL_BG    = "#EDE9FE"
PILL_TEXT  = "#3730A3"
BR_SMALL   = "font:600 11px Segoe UI,Arial;color:#6B7280;"

# --- Basic UI Helpers ---
def _pill(txt: str, bg: str=PILL_BG, fg: str=PILL_TEXT) -> str:
    return f"""<span style="display:inline-block;padding:5px 10px;border-radius:999px;background:{bg};color:{fg};font:600 12px Segoe UI,Arial;">{txt}</span>"""

def small(txt: str) -> str:
    return f"""<span style="{BR_SMALL}">{txt}</span>"""

def wrap_card(inner_html: str, background: str="#FFFFFF") -> str:
    return f"""<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:{background};border:1px solid {DIVIDER};border-radius:12px;"><tr><td style="padding:14px;">{inner_html}</td></tr></table>"""

def hlabel(text: str) -> str:
    return f"""<div style="font:800 14px/1.2 Segoe UI,Arial;color:{TEXT_DARK};margin-bottom:8px;">{text}</div>"""

# --- Badge Logic ---
_BADGE_KEYS = {
    "Top Submission": {"subm", "submission"},
    "Top Total Comments": {"comment", "cmt"},
    "Top Total Offer Follow Ups": {"offer", "follow"},
    "Top Total Pending from Partner F/U": {"pending", "partner", "f/u", "fu"},
    "Top Total Status Changed": {"status", "change"},
}

def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+"," ",str(s or "").lower()).strip()

def abbrev_label(lbl: str) -> str:
    _ABBR_MAP = {
        "Submission": "Subm", "Assessed": "Assess", "Total Comments": "Cmt(T)", "Unique Comments": "Cmt(U)",
        "Total Offer Follow Ups": "OFU(T)", "Unique Offer Follow Up": "OFU(U)",
        "Total Pending from Partner F/U": "PFU(T)", "Unique Pending from Partner F/U": "PFU(U)",
        "Total Status Changed": "Stat(T)", "Unique Status Changed": "Stat(U)", "Month-to-Date": "MTD",
    }
    s = (lbl or "").strip()
    for k, v in _ABBR_MAP.items(): s = re.sub(k, v, s, flags=re.I)
    return s

def resolve_measure_for_badge(labels_all: dict, columns: list[str], title: str) -> str | None:
    keys = _BADGE_KEYS.get(title, set())
    if not keys: return None
    best = None; score = -1
    for col in columns:
        if col.startswith("MTD_"): continue
        base = col
        lab  = labels_all.get(base, base)
        tokens = _norm(base).split() + _norm(lab).split()
        points = sum(1 for k in keys if any(k in t for t in tokens))
        if points > score:
            score = points; best = col
    return best if score>0 else None

def build_org_team_badges_today(df_day_all: pd.DataFrame, labels_all: dict) -> str:
    from data_loader import region_key
    temp = df_day_all.copy()
    temp["Team"] = temp.apply(region_key, axis=1)
    measure_today_cols = [c for c in temp.columns if (not str(c).startswith("MTD_")) and c not in ["KRA_Score","ReportDate","Role","Region","SubRegion","EmployeeEmail","EmployeeName","Manager","SendTo"]]
    grp_kra = temp[pd.to_numeric(temp.get("KRA_Score", 0), errors="coerce").fillna(0.0)>0] \
                .groupby("Team", dropna=True)["KRA_Score"].mean(numeric_only=True).reset_index() \
                .sort_values("KRA_Score", ascending=False)
    items=[]
    if not grp_kra.empty:
        items.append(f"<div style='margin:6px 0;'>{_pill('Top KRA', bg='#DBEAFE', fg='#1E3A8A')} {small(f'— Avg KRA: {round(float(grp_kra.iloc[0].KRA_Score),1)}')}<br><b>Team:</b> {grp_kra.iloc[0].Team}</div>")
    for title in _BADGE_KEYS.keys():
        col = resolve_measure_for_badge(labels_all, measure_today_cols, title)
        if not col or col not in temp.columns: continue
        agg = temp.groupby("Team", dropna=True)[col].sum(numeric_only=True).reset_index().sort_values(col, ascending=False)
        if agg.empty or int(round(float(agg.iloc[0][col])))<=0: continue
        items.append(f"<div style='margin:6px 0;'>{_pill(title, bg='#FEF3C7', fg='#92400E')} {small(f'— {abbrev_label(labels_all.get(col, col))}: {int(round(float(agg.iloc[0][col])))}')}<br><b>Team:</b> {agg.iloc[0].Team}</div>")
    body = "".join(items) if items else f"<div style='color:{TEXT_MUTED}'>No team badges today.</div>"
    return wrap_card(hlabel("TEAM BADGES (TODAY) — ORG (Team-wise)") + body)

def build_org_team_badges_mtd(df_day_snapshot: pd.DataFrame, df_all_days: pd.DataFrame, labels_all: dict) -> str:
    from data_loader import region_key
    temp_all = df_all_days.copy(); temp_all["Team"] = temp_all.apply(region_key, axis=1)
    temp_all["KRA_Score"] = pd.to_numeric(temp_all.get("KRA_Score", 0), errors="coerce").fillna(0.0)
    items = []
    day = temp_all[temp_all["KRA_Score"]>0].groupby(["Team", "ReportDate"], dropna=True)["KRA_Score"].agg(avg="mean", cnt="count").reset_index()
    if not day.empty:
        day["w_sum"] = day["avg"] * day["cnt"]
        agg = day.groupby("Team", dropna=True)[["w_sum","cnt"]].sum().reset_index()
        agg["MTD_KRA_AW"] = np.where(agg["cnt"]>0, agg["w_sum"]/agg["cnt"], 0.0)
        agg = agg.sort_values("MTD_KRA_AW", ascending=False)
        if not agg.empty and float(agg.iloc[0]["MTD_KRA_AW"]) > 0:
            items.append(f"<div style='margin:6px 0;'>{_pill('Top KRA (MTD, AW)', bg='#DBEAFE', fg='#1E3A8A')} {small(f'— Avg KRA: {round(float(agg.iloc[0].MTD_KRA_AW),1)}')}<br><b>Team:</b> {agg.iloc[0].Team}</div>")
    snap = df_day_snapshot.copy(); snap["Team"] = snap.apply(region_key, axis=1); mtd_cols = [c for c in snap.columns if str(c).startswith("MTD_")]
    for title, keys in _BADGE_KEYS.items():
        best = None; score = -1
        for col in mtd_cols:
            base = col[4:]; tokens = _norm(base).split() + _norm(labels_all.get(base, base)).split()
            pts = sum(1 for k in keys if any(k in t for t in tokens))
            if pts > score: best, score = col, pts
        if not best or score <= 0: continue
        agg2 = snap.groupby("Team", dropna=True)[best].sum(numeric_only=True).reset_index().sort_values(best, ascending=False)
        if agg2.empty or int(round(float(agg2.iloc[0][best]))) <= 0: continue
        items.append(f"<div style='margin:6px 0;'>{_pill(title+' (MTD)', bg='#E0F2FE', fg='#075985')} {small(f'— {abbrev_label(labels_all.get(best[4:], best[4:]))}: {int(round(float(agg2.iloc[0][best])))}')}<br><b>Team:</b> {agg2.iloc[0].Team}</div>")
    body = "".join(items) if items else f"<div style='color:{TEXT_MUTED}'>No team badges (MTD) yet.</div>"
    return wrap_card(hlabel("TEAM BADGES (MTD) — ORG (Team-wise)") + body)

def winners_daily_badges(df_day: pd.DataFrame, labels_all: dict):
    out = []
    top_kra = df_day[pd.to_numeric(df_day["KRA_Score"], errors='coerce').fillna(0.0)>0].sort_values("KRA_Score", ascending=False)
    if not top_kra.empty:
        best = top_kra.iloc[0]
        out.append(("Top KRA", best["EmployeeName"] or best["EmployeeEmail"], round(float(best["KRA_Score"]),1)))
    for col in [c for c in df_day.columns if c.startswith("Rank_")]:
        code = col.replace("Rank_","")
        if code not in df_day.columns: continue
        title = f"Top {labels_all.get(code, code)}"
        sub = df_day[(df_day[col]==1) & (pd.to_numeric(df_day[code], errors='coerce').fillna(0)>0)]
        if not sub.empty: out.append((title, sub.iloc[0]["EmployeeName"] or sub.iloc[0]["EmployeeEmail"], int(round(float(sub.iloc[0][code])))))
    return out

def build_badges_today_block(df_day: pd.DataFrame, labels_all: dict) -> str:
    winners = winners_daily_badges(df_day, labels_all); items = []
    for title, who, val in winners:
        items.append(f"<div style='padding:4px 0;'>{_pill(f'{title} — {val}', bg='#DCFCE7', fg='#065F46')}<div style='color:{TEXT_MUTED};font:600 12px/1.2 Inter;'>Winner: <b style='color:{TEXT_DARK}'>{who}</b></div></div>")
    body = ''.join(items) or "<div style='color:#6B7280'>No badges were earned today.</div>"
    return f"<div style='background:#FFFFFF;border:1px solid {DIVIDER};border-radius:12px;padding:16px;'><div style='font:800 15px Inter;color:{TEXT_DARK};margin-bottom:6px;'>Badges — Today</div>{body}</div>"

def build_monthly_badges_block(summary: pd.DataFrame, work_day: dt.date, labels_all: dict) -> str:
    on_day = summary[summary["ReportDate"] == work_day].copy()
    if on_day.empty: return f"<div style='background:#FFFFFF;border:1px solid {DIVIDER};border-radius:12px;padding:16px;'><div style='font:800 15px Inter;color:{TEXT_DARK};margin-bottom:6px;'>Badges — Monthly Leaders</div><div style='color:{TEXT_MUTED}'>No data.</div></div>"
    items_html, seen_titles = [], set()
    mtd_cols = [c for c in on_day.columns if c.startswith("MTD_")]
    for t, keys in _BADGE_KEYS.items():
        best, score = None, -1
        for col in mtd_cols:
            base = col[4:]; tokens = _norm(base).split() + _norm(labels_all.get(base, base)).split()
            pts = sum(1 for k in keys if any(k in t for t in tokens))
            if pts > score: best, score = col, pts
        if not best or score <= 0: continue
        sub = on_day[[best, "EmployeeName", "EmployeeEmail"]].copy()
        sub[best] = pd.to_numeric(sub[best], errors="coerce").fillna(0)
        sub = sub.sort_values(best, ascending=False)
        if sub.empty or sub.iloc[0][best] <= 0: continue
        who = sub.iloc[0]["EmployeeName"] or sub.iloc[0]["EmployeeEmail"]
        items_html.append(f"<div style='padding:4px 0;'>{_pill(f'{t} — {int(round(sub.iloc[0][best]))}', bg='#DBEAFE', fg='#1E40AF')}<div style='color:{TEXT_MUTED};font:600 12px/1.2 Inter;'>Leader: <b style='color:{TEXT_DARK}'>{who}</b></div></div>")
    body = ''.join(items_html) or f"<div style='color:{TEXT_MUTED}'>No leaders.</div>"
    return f"<div style='background:#FFFFFF;border:1px solid {DIVIDER};border-radius:12px;padding:16px;'><div style='font:800 15px Inter;color:{TEXT_DARK};margin-bottom:6px;'>Badges — Monthly Leaders</div>{body}</div>"
