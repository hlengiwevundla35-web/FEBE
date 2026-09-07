"""
FEBE Academics Dashboard with Partnerships
Run with: streamlit run febe_dashboard_with_partnerships.py
Expects FEBE_Academics.xlsx (or febe_academics_clean.csv) and Partneships.xlsx in the same folder.
"""

import re
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from datetime import datetime

CURRENT_YEAR = 2026

st.set_page_config(page_title="FEBE Dashboard", layout="wide")

# ---------------------------------------------------------------------------
# DATA LOADING
# ---------------------------------------------------------------------------
@st.cache_data
def load_academics_data():
    try:
        df = pd.read_excel("FEBE_Academics.xlsx")
    except FileNotFoundError:
        df = pd.read_csv("febe_academics_clean.csv")

    # Normalize column headers
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]

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

@st.cache_data
def load_partnerships_data():
    df = pd.read_excel("Partneships.xlsx", sheet_name="Sheet1")
    
    # Clean column names
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    
    # Clean up the data
    # Clean the "Active /Not Active" column
    if "Active /Not Active" in df.columns:
        df["Active Status"] = df["Active /Not Active"].astype(str).str.strip()
        df["Active Status"] = df["Active Status"].replace({
            "nan": "Unknown",
            "": "Unknown",
            "Not active": "Not Active",
            "Not Avtive": "Not Active",
            "Elapsed": "Not Active",
            "Not Active": "Not Active",
        })
    else:
        df["Active Status"] = "Unknown"
    
    # Clean funding
    if "Funding (if applicable)" in df.columns:
        df["Has Funding"] = df["Funding (if applicable)"].astype(str).str.strip().str.upper()
        df["Has Funding"] = df["Has Funding"].apply(
            lambda x: "Yes" if x in ["YES", "Y", "TRUE", "1"] 
            else "No" if x in ["NO", "N", "FALSE", "0", "N/A", "NOT APPLICABLE", "NAN", ""] 
            else "Unknown"
        )
    else:
        df["Has Funding"] = "Unknown"
    
    # Clean Formal Agreement
    if "Formal Agreement in place (Y/N)" in df.columns:
        df["Has Formal Agreement"] = df["Formal Agreement in place (Y/N)"].astype(str).str.strip().str.upper()
        df["Has Formal Agreement"] = df["Has Formal Agreement"].apply(
            lambda x: "Yes" if x in ["YES", "Y", "TRUE"] 
            else "No" if x in ["NO", "N", "FALSE", "", "NAN"] 
            else "Unknown"
        )
    else:
        df["Has Formal Agreement"] = "Unknown"
    
    # Clean continent
    if "Continent" in df.columns:
        df["Continent"] = df["Continent"].astype(str).str.strip()
        df["Continent"] = df["Continent"].replace({
            "": "Unknown",
            "nan": "Unknown",
            "North Africa": "Africa",
            "West Africa": "Africa",
            "East Asia": "Asia",
        })
    else:
        df["Continent"] = "Unknown"
    
    # Clean agreement type
    if "Agreement type" in df.columns:
        df["Agreement Type"] = df["Agreement type"].astype(str).str.strip()
        df["Agreement Type"] = df["Agreement Type"].replace({
            "": "Unknown",
            "nan": "Unknown",
        })
    else:
        df["Agreement Type"] = "Unknown"
    
    # Clean partner type
    if "Type University or Company or organisation" in df.columns:
        df["Partner Type"] = df["Type University or Company or organisation"].astype(str).str.strip()
        df["Partner Type"] = df["Partner Type"].replace({
            "": "Unknown",
            "nan": "Unknown",
            "International Partnership": "University",  # Some rows have this misclassified
            "University or Company or organisation": "Unknown",
        })
    else:
        df["Partner Type"] = "Unknown"
    
    # Clean partnership type
    if "Type of Partnerships" in df.columns:
        df["Partnership Type"] = df["Type of Partnerships"].astype(str).str.strip()
        df["Partnership Type"] = df["Partnership Type"].replace({
            "": "Unknown",
            "nan": "Unknown",
        })
    else:
        df["Partnership Type"] = "Unknown"
    
    # Extract country from location if needed
    if "Partner's Country)" in df.columns:
        df["Country"] = df["Partner's Country)"].astype(str).str.strip()
        df["Country"] = df["Country"].replace({
            "": "Unknown",
            "nan": "Unknown",
        })
    else:
        df["Country"] = "Unknown"
    
    # Clean up partner names
    if "Name of Partner" in df.columns:
        df["Partner Name"] = df["Name of Partner"].astype(str).str.strip()
        df["Partner Name"] = df["Partner Name"].replace({
            "": "Unknown",
            "nan": "Unknown",
        })
    else:
        df["Partner Name"] = "Unknown"
    
    return df

# Load data
academics_df = load_academics_data()
partnerships_df = load_partnerships_data()

# ---------------------------------------------------------------------------
# SIDEBAR FILTERS - Academics
# ---------------------------------------------------------------------------
st.sidebar.header("Filters")

dept_options = sorted(academics_df["Department Name"].dropna().unique())
depts = st.sidebar.multiselect("Department", dept_options, default=[])

campus_options = sorted(academics_df["CAMPUS NAME"].dropna().unique())
campuses = st.sidebar.multiselect("Campus", campus_options, default=[])

qual_options = sorted(academics_df["Qualification Level"].dropna().unique())
quals = st.sidebar.multiselect("Highest Qualification Level", qual_options, default=[])

# Apply filters to academics data
fdf = academics_df.copy()
if depts:
    fdf = fdf[fdf["Department Name"].isin(depts)]
if campuses:
    fdf = fdf[fdf["CAMPUS NAME"].isin(campuses)]
if quals:
    fdf = fdf[fdf["Qualification Level"].isin(quals)]

st.sidebar.markdown(f"**{len(fdf)}** of {len(academics_df)} staff shown")

# ---------------------------------------------------------------------------
# TABS
# ---------------------------------------------------------------------------
tab_overview, tab_quals, tab_research, tab_workforce, tab_partnerships = st.tabs(
    ["Overview", "Qualifications", "Research & Funding", "Workforce Planning", "Partnerships"]
)

# ---------------------------------------------------------------------------
# TAB 1 — OVERVIEW (unchanged from original)
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
# TAB 2 — QUALIFICATIONS (unchanged)
# ---------------------------------------------------------------------------
with tab_quals:
    c1, c2 = st.columns(2)

    with c1:
        reg = fdf[fdf["_reg_year"].notna()]
        fig = px.histogram(reg, x="_reg_year", nbins=15,
                            title="Qualification Registration Start Year", labels={"_reg_year": "Year"})
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
# TAB 3 — RESEARCH & FUNDING (unchanged)
# ---------------------------------------------------------------------------
with tab_research:
    out_cols = {"2022": "2022 Research output", "2023": "2023 Research output",
                "2024": "2024 Research Output", "2025": "2025 Research Output"}
    out_cols = {yr: col for yr, col in out_cols.items() if col in fdf.columns}
    required = pd.to_numeric(fdf["Required Research Units"].astype(str).str.strip().replace("-", np.nan),
                              errors="coerce")

    c1, c2 = st.columns(2)

    with c1:
        avgs = {yr: pd.to_numeric(fdf[col], errors="coerce").mean() for yr, col in out_cols.items()}
        avg_df = pd.DataFrame({"Year": list(avgs.keys()), "Average Output": list(avgs.values())})
        fig = px.bar(avg_df, x="Year", y="Average Output", title="Average Research Output by Year")
        st.plotly_chart(fig, use_container_width=True)
        st.caption("2025 is the current year and likely under-reported so far.")

    with c2:
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
                     color_discrete_map={"Under target": "#a33", "On target": "#d9d9d9",
                                         "Over target": "#2c5f8a"},
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
    sup_totals = {label: pd.to_numeric(fdf[col], errors="coerce").sum()
                  for label, col in sup_cols.items() if col in fdf.columns}
    sup_df = pd.DataFrame({"Role": list(sup_totals.keys()), "Students": list(sup_totals.values())})
    fig = px.bar(sup_df, x="Students", y="Role", orientation="h", title="Total Supervision Load by Role")
    st.plotly_chart(fig, use_container_width=True)

    st.info(
        "Funding-by-source (NRF/TIA/ESKOM/Erasmus/Industry) and NRF/CPUT workshop-attendance "
        "columns are almost entirely blank in the source data, so they aren't charted here — "
        "revisit once those fields are actually populated."
    )

# ---------------------------------------------------------------------------
# TAB 4 — WORKFORCE PLANNING (unchanged)
# ---------------------------------------------------------------------------
with tab_workforce:
    c1, c2 = st.columns(2)

    with c1:
        senior = fdf["HoD Surname"].astype(str).str.strip().str.upper()
        senior = senior[(senior != "") & (senior != "NAN")]
        senior = senior[~senior.isin(["BALKARAN", "RAMSUROOP"])]

        span = senior.value_counts().reset_index()
        span.columns = ["Direct Senior", "Direct Reports"]
        fig = px.bar(
            span.sort_values("Direct Reports"), x="Direct Reports", y="Direct Senior",
            orientation="h",
            title=f"Span of Control by Direct Senior (n={span['Direct Reports'].sum()} staff, "
                  f"{span.shape[0]} seniors)"
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        window = fdf["_retire_year"].between(CURRENT_YEAR, CURRENT_YEAR + 9)
        sub = fdf.loc[window].copy()
        sub["5yr bucket"] = pd.cut(
            sub["_retire_year"],
            bins=[CURRENT_YEAR - 1, CURRENT_YEAR + 4, CURRENT_YEAR + 9],
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


# ---------------------------------------------------------------------------
# TAB 5 — PARTNERSHIPS
# ---------------------------------------------------------------------------
with tab_partnerships:
    st.header("FEBE Partnerships Overview")
    
    # KPIs for partnerships
    p_k1, p_k2, p_k3, p_k4, p_k5 = st.columns(5)
    total_partnerships = len(partnerships_df)
    active_partnerships = len(partnerships_df[partnerships_df["Active Status"] == "Active"])
    has_agreement = len(partnerships_df[partnerships_df["Has Formal Agreement"] == "Yes"])
    has_funding = len(partnerships_df[partnerships_df["Has Funding"] == "Yes"])
    
    # Get unique partners
    unique_partners = partnerships_df["Partner Name"].nunique()
    
    p_k1.metric("Total Partnerships", total_partnerships)
    p_k2.metric("Active Partnerships", active_partnerships)
    p_k3.metric("Formal Agreements", has_agreement)
    p_k4.metric("With Funding", has_funding)
    p_k5.metric("Unique Partners", unique_partners)
    
    st.markdown("---")
    
    # Filter section for partnerships
    st.subheader("Filter Partnerships")
    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
    
    with f_col1:
        status_filter = st.multiselect(
            "Active Status",
            options=sorted(partnerships_df["Active Status"].unique()),
            default=[]
        )
    
    with f_col2:
        partnership_type_filter = st.multiselect(
            "Partnership Type",
            options=sorted(partnerships_df["Partnership Type"].unique()),
            default=[]
        )
    
    with f_col3:
        continent_filter = st.multiselect(
            "Continent",
            options=sorted(partnerships_df["Continent"].unique()),
            default=[]
        )
    
    with f_col4:
        funding_filter = st.multiselect(
            "Has Funding",
            options=sorted(partnerships_df["Has Funding"].unique()),
            default=[]
        )
    
    # Apply filters
    p_fdf = partnerships_df.copy()
    if status_filter:
        p_fdf = p_fdf[p_fdf["Active Status"].isin(status_filter)]
    if partnership_type_filter:
        p_fdf = p_fdf[p_fdf["Partnership Type"].isin(partnership_type_filter)]
    if continent_filter:
        p_fdf = p_fdf[p_fdf["Continent"].isin(continent_filter)]
    if funding_filter:
        p_fdf = p_fdf[p_fdf["Has Funding"].isin(funding_filter)]
    
    st.caption(f"Showing {len(p_fdf)} of {len(partnerships_df)} partnerships")
    
    # Partnership Charts
    chart_col1, chart_col2 = st.columns(2)
    
    with chart_col1:
        # Active Status Distribution
        status_counts = p_fdf["Active Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]
        fig = px.pie(
            status_counts, 
            names="Status", 
            values="Count", 
            title="Partnership Status",
            color_discrete_sequence=px.colors.qualitative.Set2,
            hole=0.3
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with chart_col2:
        # Partnership Type Distribution
        type_counts = p_fdf["Partnership Type"].value_counts().head(10).reset_index()
        type_counts.columns = ["Type", "Count"]
        fig = px.bar(
            type_counts, 
            x="Count", 
            y="Type", 
            orientation="h",
            title="Top Partnership Types",
            color="Type",
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    # Second row of charts
    chart_col3, chart_col4 = st.columns(2)
    
    with chart_col3:
        # Geographic Distribution by Continent
        continent_counts = p_fdf["Continent"].value_counts().reset_index()
        continent_counts.columns = ["Continent", "Count"]
        fig = px.bar(
            continent_counts,
            x="Continent",
            y="Count",
            title="Partnerships by Continent",
            color="Continent",
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    with chart_col4:
        # Partner Type Distribution
        partner_type_counts = p_fdf["Partner Type"].value_counts().head(8).reset_index()
        partner_type_counts.columns = ["Partner Type", "Count"]
        fig = px.bar(
            partner_type_counts,
            x="Count",
            y="Partner Type",
            orientation="h",
            title="Partner Types",
            color="Partner Type",
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    # Third row - Funding and Agreement
    chart_col5, chart_col6 = st.columns(2)
    
    with chart_col5:
        # Funding Distribution
        funding_counts = p_fdf["Has Funding"].value_counts().reset_index()
        funding_counts.columns = ["Funding", "Count"]
        fig = px.pie(
            funding_counts,
            names="Funding",
            values="Count",
            title="Partnerships with Funding",
            color_discrete_sequence=px.colors.qualitative.Set2,
            hole=0.3
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with chart_col6:
        # Formal Agreement Distribution
        agreement_counts = p_fdf["Has Formal Agreement"].value_counts().reset_index()
        agreement_counts.columns = ["Formal Agreement", "Count"]
        fig = px.pie(
            agreement_counts,
            names="Formal Agreement",
            values="Count",
            title="Formal Agreements in Place",
            color_discrete_sequence=px.colors.qualitative.Set2,
            hole=0.3
        )
        st.plotly_chart(fig, use_container_width=True)
    
    # Top Partners by Activity
    st.subheader("Top Partners by Activity")
    partner_activity = p_fdf.groupby("Partner Name").agg({
        "Active Status": lambda x: (x == "Active").sum(),
        "Has Funding": lambda x: (x == "Yes").sum(),
        "Has Formal Agreement": lambda x: (x == "Yes").sum(),
    }).reset_index()
    partner_activity.columns = ["Partner Name", "Active Partnerships", "With Funding", "With Formal Agreement"]
    partner_activity = partner_activity.sort_values("Active Partnerships", ascending=False).head(15)
    
    if not partner_activity.empty:
        st.dataframe(partner_activity, use_container_width=True, hide_index=True)
    
    # Map of partnerships by country
    st.subheader("Partnerships by Country")
    country_counts = p_fdf["Country"].value_counts().head(15).reset_index()
    country_counts.columns = ["Country", "Number of Partnerships"]
    
    if not country_counts.empty:
        fig = px.bar(
            country_counts,
            x="Country",
            y="Number of Partnerships",
            title="Top 15 Countries by Partnership Count",
            color="Number of Partnerships",
            color_continuous_scale="Blues"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    # Detailed partnership table
    st.subheader("Partnership Details")
    
    # Select columns to display
    display_cols = [
        "Partner Name", 
        "Partnership Type", 
        "Partner Type", 
        "Active Status", 
        "Continent", 
        "Country",
        "Has Funding", 
        "Has Formal Agreement", 
        "Agreement Type"
    ]
    # Filter to only columns that exist
    display_cols = [c for c in display_cols if c in p_fdf.columns]
    
    # Add Project Leader if available
    if "CPUT Principal Partner" in p_fdf.columns and "Name of project leader; Faculty" in p_fdf.columns:
        p_fdf["Project Leader"] = p_fdf["CPUT Principal Partner"].astype(str) + " / " + p_fdf["Name of project leader; Faculty"].astype(str)
        display_cols.insert(1, "Project Leader")
    
    if display_cols:
        st.dataframe(
            p_fdf[display_cols].sort_values("Partner Name"),
            use_container_width=True,
            hide_index=True
        )
    
    # Export option
    st.download_button(
        label="📥 Download Partnership Data (CSV)",
        data=p_fdf.to_csv(index=False),
        file_name="febe_partnerships_export.csv",
        mime="text/csv"
    )