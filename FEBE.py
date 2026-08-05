"""
FEBE Academics Dashboard
Run with:  streamlit run febe_dashboard.py
Expects FEBE_Academics_Clean.xlsx (or febe_academics_clean.csv) in the same folder.
"""

import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

CURRENT_YEAR = 2026

st.set_page_config(page_title="FEBE Academics Dashboard", layout="wide")


# ---------------------------------------------------------------------------
# DATA LOADING
# ---------------------------------------------------------------------------
@st.cache_data
def load_data():
    try:
        df = pd.read_excel("FEBE_Academics.xlsx")
    except FileNotFoundError:
        df = pd.read_csv("FEBE_Academics.csv")

    # Derived fields used across pages
    reg_year = pd.to_numeric(df["Year of first Registration"].astype(str).str.strip(), errors="coerce")
    exp_year = pd.to_numeric(df["Expected Completion"].astype(str).str.strip(), errors="coerce")
    is_registered = reg_year.notna()
    overdue = is_registered & exp_year.notna() & (exp_year < CURRENT_YEAR)
    df["Registration Status"] = np.where(
        overdue, "Overdue", np.where(is_registered, "On track / no exp. date", "Not registered")
    )
    df["_reg_year"] = reg_year
    df["_exp_year"] = exp_year

    ret_year = pd.to_datetime(df["Retirement year"], errors="coerce").dt.year
    df["_retire_year"] = ret_year.where(ret_year >= CURRENT_YEAR)

    return df


df = load_data()

# ---------------------------------------------------------------------------
# SIDEBAR FILTERS
# ---------------------------------------------------------------------------
st.sidebar.header("Filters")

dept_options = sorted(df["Department Name"].dropna().unique())
depts = st.sidebar.multiselect("Department", dept_options, default=[])

campus_options = sorted(df["CAMPUS NAME"].dropna().unique())
campuses = st.sidebar.multiselect("Campus", campus_options, default=[])

qual_options = sorted(df["Qualification Level"].dropna().unique())
quals = st.sidebar.multiselect("Highest Qualification Level", qual_options, default=[])

fdf = df.copy()
if depts:
    fdf = fdf[fdf["Department Name"].isin(depts)]
if campuses:
    fdf = fdf[fdf["CAMPUS NAME"].isin(campuses)]
if quals:
    fdf = fdf[fdf["Qualification Level"].isin(quals)]

st.sidebar.markdown(f"**{len(fdf)}** of {len(df)} staff shown")

# ---------------------------------------------------------------------------
# HEADER + KPI ROW
# ---------------------------------------------------------------------------
st.title("FEBE Academics Dashboard")

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Staff", len(fdf))
k2.metric("PhD Holders", int((fdf["Qualification Level"] == "Doctorate").sum()))
k3.metric("Currently Registered", int(fdf["_reg_year"].notna().sum()))
k4.metric("Overdue Qualifications", int((fdf["Registration Status"] == "Overdue").sum()))
k5.metric("NRF Rated", int(fdf["NRF Rating"].isin(["C1", "C2", "C3", "Y1", "Y2", "B1", "B2", "B3", "A1", "A2"]).sum()))

tab_overview, tab_quals, tab_research, tab_workforce = st.tabs(
    ["Overview", "Qualifications", "Research & Funding", "Workforce Planning"]
)

# ---------------------------------------------------------------------------
# TAB 1 — OVERVIEW
# ---------------------------------------------------------------------------
with tab_overview:
    c1, c2 = st.columns(2)

    with c1:
        qual_order = ["Doctorate", "Masters", "Bachelors/Honours", "Diploma/Other", "Other/Unclassified"]
        counts = fdf["Qualification Level"].value_counts().reindex(qual_order).dropna().reset_index()
        counts.columns = ["Qualification Level", "Count"]
        fig = px.bar(counts, x="Count", y="Qualification Level", orientation="h",
                     title="Highest Qualification Level", color="Qualification Level",
                     color_discrete_sequence=px.colors.sequential.Blues_r)
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        nrf = fdf["NRF Status"].fillna("No rating").replace("", "No rating")
        nrf_counts = nrf.value_counts().reset_index()
        nrf_counts.columns = ["Status", "Count"]
        fig = px.pie(nrf_counts, names="Status", values="Count", title="NRF Rating Status", hole=0.35)
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)

    with c3:
        ct = pd.crosstab(fdf["Ethnic Group Name"], fdf["Gender"]).reset_index()
        ct_melt = ct.melt(id_vars="Ethnic Group Name", var_name="Gender", value_name="Count")
        fig = px.bar(ct_melt, x="Ethnic Group Name", y="Count", color="Gender", barmode="stack",
                     title="Staff by Ethnic Group & Gender")
        st.plotly_chart(fig, use_container_width=True)

    with c4:
        camp = fdf["CAMPUS NAME"].value_counts().reset_index()
        camp.columns = ["Campus", "Count"]
        fig = px.bar(camp, x="Count", y="Campus", orientation="h", title="Staff by Campus")
        st.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------------------------
# TAB 2 — QUALIFICATIONS
# ---------------------------------------------------------------------------
with tab_quals:
    c1, c2 = st.columns(2)

    with c1:
        reg = fdf[fdf["_reg_year"].notna()]
        fig = px.histogram(reg, x="_reg_year", nbins=15, title="Qualification Registration Start Year",
                            labels={"_reg_year": "Year"})
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        counts = fdf["Registration Status"].value_counts().reset_index()
        counts.columns = ["Status", "Count"]
        fig = px.bar(counts, x="Status", y="Count", color="Status",
                     color_discrete_map={"Overdue": "#a33", "On track / no exp. date": "#2c5f8a",
                                         "Not registered": "#d9d9d9"},
                     title="Registered Qualification Status")
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Overdue staff (Expected Completion has passed)")
    overdue_cols = ["First name", "Surname", "Department Name", "Higher qualification registered for",
                    "Year of first Registration", "Expected Completion", "Name of Insitution"]
    overdue_cols = [c for c in overdue_cols if c in fdf.columns]
    st.dataframe(fdf.loc[fdf["Registration Status"] == "Overdue", overdue_cols], use_container_width=True)

# ---------------------------------------------------------------------------
# TAB 3 — RESEARCH & FUNDING
# ---------------------------------------------------------------------------
with tab_research:
    out_cols = {"2022": "2022  Research output", "2023": "2023  Research output",
                "2024": "2024 Research Output", "2025": "2025 Research Output"}
    out_cols = {yr: col for yr, col in out_cols.items() if col in fdf.columns}
    required = pd.to_numeric(fdf["Required  Research Units"].astype(str).str.strip().replace("-", np.nan),
                              errors="coerce")

    c1, c2 = st.columns(2)

    with c1:
        avgs = {yr: pd.to_numeric(fdf[col], errors="coerce").mean() for yr, col in out_cols.items()}
        avg_df = pd.DataFrame({"Year": list(avgs.keys()), "Average Output": list(avgs.values())})
        fig = px.bar(avg_df, x="Year", y="Average Output", title="Average Research Output by Year")
        st.plotly_chart(fig, use_container_width=True)
        st.caption("2025 is the current year and likely under-reported so far.")

    with c2:
        # Under/on/over target counts, one grouped bar per year
        rows = []
        for yr, col in out_cols.items():
            actual = pd.to_numeric(fdf[col], errors="coerce")
            gap = (actual - required).dropna()
            status = pd.cut(gap, bins=[-np.inf, -0.01, 0.01, np.inf],
                             labels=["Under target", "On target", "Over target"])
            vc = status.value_counts().reindex(["Under target", "On target", "Over target"])
            for s, v in vc.items():
                rows.append({"Year": yr, "Status": s, "Count": v})
        gap_df = pd.DataFrame(rows)
        fig = px.bar(gap_df, x="Year", y="Count", color="Status", barmode="group",
                     color_discrete_map={"Under target": "#a33", "On target": "#d9d9d9", "Over target": "#2c5f8a"},
                     title="Output vs Required Units by Year")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Staff under target, by year")
    under_rows = []
    for yr, col in out_cols.items():
        actual = pd.to_numeric(fdf[col], errors="coerce")
        gap = (actual - required)
        under_rows.append({"Year": yr, "Staff Under Target": int((gap < -0.01).sum())})
    st.dataframe(pd.DataFrame(under_rows), use_container_width=True, hide_index=True)

    st.subheader("Supervision load")
    sup_cols = {
        "Masters Co-Sup": "Masters Co-Supervision",
        "Doctoral Co-Sup": "Doctoral Co-Supervision",
        "Principal Sup. (Masters)": "Principal Supervisor-Masters",
        "Principal Sup. (Doctorate)": "Principal Supervisor- Doctorate",
    }
    sup_totals = {label: pd.to_numeric(fdf[col], errors="coerce").sum() for label, col in sup_cols.items()
                  if col in fdf.columns}
    sup_df = pd.DataFrame({"Role": list(sup_totals.keys()), "Students": list(sup_totals.values())})
    fig = px.bar(sup_df, x="Students", y="Role", orientation="h", title="Total Supervision Load by Role")
    st.plotly_chart(fig, use_container_width=True)

    st.info(
        "Funding-by-source (NRF/TIA/ESKOM/Erasmus/Industry) and NRF/CPUT workshop-attendance "
        "columns are almost entirely blank in the source data, so they aren't charted here — "
        "revisit once those fields are actually populated."
    )

# ---------------------------------------------------------------------------
# TAB 4 — WORKFORCE PLANNING
# ---------------------------------------------------------------------------
with tab_workforce:
    c1, c2 = st.columns(2)

    with c1:
      senior = fdf["Direct Senior Surname"].astype(str).str.strip().str.upper()
      senior = senior[(senior != "") & (senior != "NAN")]
    # Exclude Balkaran and Ramsuroop from THIS chart only — their reports
    # still count everywhere else in the dashboard.
      senior = senior[~senior.isin(["BALKARAN", "RAMSUROOP"])]

      span = senior.value_counts().reset_index()
      span.columns = ["Direct Senior", "Direct Reports"]
      fig = px.bar(span.sort_values("Direct Reports"), x="Direct Reports", y="Direct Senior",
                 orientation="h", title="Span of Control by Direct Senior")
      st.plotly_chart(fig, use_container_width=True)

    with c2:
        window = fdf["_retire_year"].between(CURRENT_YEAR, CURRENT_YEAR + 9)
        sub = fdf.loc[window].copy()
        sub["5yr bucket"] = pd.cut(
            sub["_retire_year"], bins=[CURRENT_YEAR - 1, CURRENT_YEAR + 4, CURRENT_YEAR + 9],
            labels=[f"{CURRENT_YEAR}-{CURRENT_YEAR+4}", f"{CURRENT_YEAR+5}-{CURRENT_YEAR+9}"]
        )
        sub["Qual Group"] = sub["Qualification Level"].apply(lambda x: "PhD holder" if x == "Doctorate" else "Other")
        ret_counts = sub.groupby(["5yr bucket", "Qual Group"], observed=True).size().reset_index(name="Count")
        fig = px.bar(ret_counts, x="5yr bucket", y="Count", color="Qual Group", barmode="stack",
                     title="Retirement Window — Next 10 Years",
                     color_discrete_map={"PhD holder": "#a33", "Other": "#d9d9d9"})
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Staff retiring within 5 years")
    near_term = fdf[fdf["_retire_year"].between(CURRENT_YEAR, CURRENT_YEAR + 4)]
    ret_cols = ["First name", "Surname", "Department Name", "Post Name", "Qualification Level", "Retirement year"]
    ret_cols = [c for c in ret_cols if c in fdf.columns]
    st.dataframe(near_term[ret_cols].sort_values("Retirement year"), use_container_width=True)
