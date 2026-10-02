import hmac
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import altair as alt
import pandas as pd
import streamlit as st

import kobo
from species import WARDS

st.set_page_config(page_title="FFS Tree Seedlings – Vihiga", page_icon="vihiga_logo.png", layout="wide")

NAVY, BLUE, GREEN = "#0A2A5E", "#2390CE", "#2E8B57"


def secret(name, default=""):
    try:
        return st.secrets.get(name, os.environ.get(name, default))
    except Exception:
        return os.environ.get(name, default)


LIVE = bool(secret("KOBO_TOKEN") and secret("KOBO_ASSET_UID"))

# ------------------------------------------------------------------ header
c1, c2 = st.columns([1, 9])
c1.image("vihiga_logo.png", width=90)
c2.title("FFS Tree Seedling Selection")
c2.caption("County Government of Vihiga · Department of Agriculture, Livestock & Fisheries · with KEFRI")

# ------------------------------------------------------------------ access
if "role" not in st.session_state:
    st.session_state.role = None

if st.session_state.role is None:
    if not LIVE:
        st.warning("Demo mode: sample data, no Kobo connection. Set the Kobo secrets to go live.")
        role = st.radio("Preview as", ["KEFRI (totals only)", "DoALF (full access)"], horizontal=True)
        if st.button("Open"):
            st.session_state.role = "doalf" if role.startswith("DoALF") else "kefri"
            st.rerun()
    else:
        pw = st.text_input("Password", type="password")
        if st.button("Open") and pw:
            if hmac.compare_digest(pw, str(secret("DOALF_PASSWORD", "\x00"))):
                st.session_state.role = "doalf"
                st.rerun()
            elif hmac.compare_digest(pw, str(secret("KEFRI_PASSWORD", "\x00"))):
                st.session_state.role = "kefri"
                st.rerun()
            else:
                st.error("Wrong password.")
    st.stop()

ROLE = st.session_state.role
IS_DOALF = ROLE == "doalf"


# ------------------------------------------------------------------ data
@st.cache_data(ttl=60, show_spinner="Loading submissions…")
def load(live: bool):
    rows = (kobo.fetch_submissions(secret("KOBO_SERVER", "https://kf.kobotoolbox.org"),
                                   secret("KOBO_TOKEN"), secret("KOBO_ASSET_UID"))
            if live else kobo.demo_rows())
    return kobo.flatten(rows), datetime.now(ZoneInfo("Africa/Nairobi"))


try:
    raw, loaded_at = load(LIVE)
except Exception as e:  # network, token, wrong asset id
    st.error(f"Could not read from Kobo: {e}")
    st.stop()

top = st.columns([6, 2, 2])
top[0].caption(f"{'Live' if LIVE else 'DEMO'} data · updated {loaded_at:%a %d %b %H:%M} · refreshes every minute · "
               f"signed in as {'DoALF' if IS_DOALF else 'KEFRI (totals only)'}")
if top[1].button("Refresh now"):
    load.clear()
    st.rerun()
if top[2].button("Sign out"):
    st.session_state.role = None
    st.rerun()

dedupe = st.sidebar.checkbox("Remove duplicate farmers (keep latest)", value=True,
                             help="Duplicate = same ward, FFS group and farmer number.")
df_all, info = kobo.prepare(raw, dedupe)

if df_all.empty:
    st.info("No submissions yet. Numbers will appear here as facilitators submit.")
    st.stop()

# ------------------------------------------------------------------ filters
st.sidebar.header("Filters")
subs = sorted(df_all["subcounty_name"].unique())
sel_sub = st.sidebar.multiselect("Sub-county", subs)
df = df_all[df_all["subcounty_name"].isin(sel_sub)] if sel_sub else df_all
wards = sorted(df["ward_name"].unique())
sel_ward = st.sidebar.multiselect("Ward", wards)
df = df[df["ward_name"].isin(sel_ward)] if sel_ward else df

# ------------------------------------------------------------------ KPIs
k = st.columns(5)
k[0].metric("Farmers", f"{len(df):,}")
k[1].metric("FFS groups", f"{df[['ward_name', 'ffs_group_clean']].drop_duplicates().shape[0]:,}")
k[2].metric("Wards reporting", f"{df['ward_name'].nunique()} of {len(WARDS)}")
k[3].metric("Indigenous seedlings", f"{int(df['total_indigenous'].sum()):,}")
k[4].metric("Fruit seedlings", f"{int(df['total_fruit'].sum()):,}")

tabs = ["Overview", "FFS groups", "Wards & sub-counties", "County by species"]
if IS_DOALF:
    tabs += ["Farmers", "Data checks"]
T = st.tabs(tabs)


def hbar(data, color, title):
    d = data[data["Seedlings"] > 0]
    if d.empty:
        st.caption(f"{title}: nothing yet")
        return
    ch = (alt.Chart(d).mark_bar(cornerRadiusEnd=4, size=16, color=color)
          .encode(y=alt.Y("Species:N", sort="-x", title=None, axis=alt.Axis(labelLimit=380, labelOverlap=False)),
                  x=alt.X("Seedlings:Q", title="Seedlings"),
                  tooltip=["Species", "Seedlings", "Farmers choosing"])
          .properties(title=title, height=alt.Step(28)))
    st.altair_chart(ch, width="stretch")


def download(label, data, name):
    st.download_button(label, kobo.to_excel(data, IS_DOALF), file_name=name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


stamp = datetime.now(ZoneInfo("Africa/Nairobi")).strftime("%Y%m%d_%H%M")

# ------------------------------------------------------------------ Overview
with T[0]:
    cs = kobo.county_by_species(df)
    a, b = st.columns(2)
    with a:
        hbar(cs[cs["Category"] == "Indigenous"], BLUE, "Indigenous species – seedlings requested")
    with b:
        hbar(cs[cs["Category"] != "Indigenous"], GREEN, "Fruit / agroforestry – seedlings requested")
    a, b = st.columns(2)
    with a:
        s = df.groupby("subcounty_name")["total_seedlings"].sum().reset_index()
        s.columns = ["Sub-county", "Seedlings"]
        st.altair_chart(alt.Chart(s).mark_bar(cornerRadiusEnd=4, size=28, color=NAVY)
                        .encode(x=alt.X("Sub-county:N", sort="-y", title=None), y=alt.Y("Seedlings:Q"),
                                tooltip=["Sub-county", "Seedlings"])
                        .properties(title="Seedlings by sub-county", height=260), width="stretch")
    with b:
        d = df.dropna(subset=["submitted"]).copy()
        d["Day"] = d["submitted"].dt.tz_convert("Africa/Nairobi").dt.date
        d = d.groupby("Day").size().reset_index(name="Farmers")
        st.altair_chart(alt.Chart(d).mark_bar(cornerRadiusEnd=4, size=28, color=BLUE)
                        .encode(x=alt.X("Day:T", title=None), y=alt.Y("Farmers:Q", title="Farmers recorded"),
                                tooltip=["Day:T", "Farmers"])
                        .properties(title="Farmers recorded per day", height=260), width="stretch")
    download("Download everything (Excel)", df, f"FFS_seedlings_{stamp}.xlsx")

# ------------------------------------------------------------------ FFS groups
with T[1]:
    st.caption("One row per FFS group, one column per species – the single request per group for KEFRI.")
    g = kobo.group_table(df)
    st.dataframe(g, width="stretch", hide_index=True)
    ob = kobo.others_by_group(df)
    if not ob.empty:
        st.subheader("Species added by farmers (not on the list)")
        st.caption("Written in by farmers under 'Other'. Spelling differences in capitals and spaces are merged.")
        st.dataframe(ob, width="stretch", hide_index=True)
    download("Download group consolidation (Excel)", df, f"FFS_seedlings_{stamp}.xlsx")

# ------------------------------------------------------------------ Wards & sub-counties
with T[2]:
    st.subheader("By ward")
    st.dataframe(kobo.ward_table(df), width="stretch", hide_index=True)
    st.subheader("By sub-county")
    st.dataframe(kobo.subcounty_table(df), width="stretch", hide_index=True)
    exp = "expected_groups.csv"
    if os.path.exists(exp):
        st.subheader("Reporting progress")
        e = pd.read_csv(exp)  # columns: ward, expected_groups
        rep = df.groupby("ward_name")["ffs_group_clean"].nunique().reset_index()
        rep.columns = ["ward", "groups_reporting"]
        p = e.merge(rep, on="ward", how="left").fillna({"groups_reporting": 0})
        p["remaining"] = (p["expected_groups"] - p["groups_reporting"]).clip(lower=0).astype(int)
        st.dataframe(p, width="stretch", hide_index=True)

# ------------------------------------------------------------------ County by species
with T[3]:
    st.caption("Countywide total per species – the distribution list for KEFRI.")
    st.dataframe(kobo.county_by_species(df), width="stretch", hide_index=True)
    download("Download county totals (Excel)", df, f"FFS_seedlings_{stamp}.xlsx")

# ------------------------------------------------------------------ DoALF only
if IS_DOALF:
    with T[4]:
        st.caption("Personal data – DoALF only. Not visible on the KEFRI login.")
        st.dataframe(kobo.farmer_table(df), width="stretch", hide_index=True)
    with T[5]:
        st.write(f"Submissions received: **{info['raw']}**")
        st.write(f"Duplicate farmers removed: **{info['duplicates_removed']}**"
                 + ("" if dedupe else " (removal is switched off)"))
        st.write(f"Rows without a ward (ignored): **{info['no_ward']}**")
        st.write(f"Farmers with fewer than 3 species in a category: **{info['under_3']}**")
        if (df_all["ffs_group"] == "").any():
            st.warning("Some submissions have no FFS group name.")
        per = df_all.groupby(["ward_name", "ffs_group_clean"]).size().reset_index(name="Farmers")
        per.columns = ["Ward", "FFS group", "Farmers"]
        st.write("Farmers per group – spot groups with unusually few or many:")
        st.dataframe(per, width="stretch", hide_index=True)
