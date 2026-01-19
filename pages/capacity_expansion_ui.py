import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from datetime import datetime, timedelta

def fiscal_day_to_label(day_num: int, base_year: int = 2027) -> str:
    """
    Converts fiscal day index (1..365) to date label assuming:
    Day 1 = Apr 1, Day 365 = Mar 31.
    base_year is only used for date calculation (year label),
    it doesn't affect month-day mapping logic.
    """
    # Fiscal year starts Apr 1 of base_year
    start_date = datetime(base_year, 4, 1)
    dt = start_date + timedelta(days=int(day_num) - 1)
    return dt.strftime("%d-%b")  # e.g. "01-Apr"

# ---------------------------
# Page configuration & style
# ---------------------------
st.set_page_config(page_title="Annual Power Procurement Planning Outcomes", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .reportview-container .main .block-container{padding-top:1rem;padding-left:2rem;padding-right:2rem}
    .stButton>button {height:2.6rem}
    .stDownloadButton>button {height:2.6rem}
    table {text-align: center !important;}
    th, td {text-align: center !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------
# Helper functions
# ---------------------------
def load_summary_sheet(file_path: str) -> pd.DataFrame:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Excel file not found at {file_path}")
    df = pd.read_excel(file_path, sheet_name="Summary")
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    df = df.rename(columns={df.columns[0]: 'Year'})
    df['Year'] = pd.to_numeric(df['Year'], errors='coerce').astype('Int64')
    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    df = df.sort_values('Year').reset_index(drop=True)
    return df

def wide_to_long(df_wide: pd.DataFrame) -> pd.DataFrame:
    long = df_wide.melt(id_vars=['Year'], var_name='Technology', value_name='MW')
    long['MW'] = pd.to_numeric(long['MW'], errors='coerce').fillna(0)
    return long

def make_stacked_bar(df_long: pd.DataFrame, techs: list, years: list, percent: bool = False) -> px.bar:
    df = df_long[df_long['Technology'].isin(techs) & df_long['Year'].isin(years)]
    if percent:
        total_by_year = df.groupby('Year', as_index=False)['MW'].sum().rename(columns={'MW': 'TotalMW'})
        df = df.merge(total_by_year, on='Year')
        df['Share'] = df['MW'] / df['TotalMW'] * 100
        y_col, y_title = 'Share', 'Share (%)'
    else:
        y_col, y_title = 'MW', 'MW'
    fig = px.bar(df, x='Year', y=y_col, color='Technology', labels={y_col: y_title, 'Year': 'Year'})
    fig.update_layout(
        barmode='stack',
        legend_title_text='Technology',
        hovermode='x unified',
        margin={'t': 30, 'b': 40, 'l': 40, 'r': 20},
    )
    return fig

def make_tech_line_chart(df_long: pd.DataFrame, techs: list, years: list) -> px.line:
    df = df_long[df_long['Technology'].isin(techs) & df_long['Year'].isin(years)]
    fig = px.line(df, x='Year', y='MW', color='Technology', markers=True)
    fig.update_layout(
        legend_title_text='Technology',
        hovermode='x unified',
        margin={'t': 30, 'b': 40, 'l': 40, 'r': 20},
    )
    return fig

# ---------------------------
# Load data
# ---------------------------
df_wide = load_summary_sheet("Total Capacities 2026 - 2036.xlsx")
df_long = wide_to_long(df_wide)

# Daily supply-demand
demand_file = "Demand balance gen.xlsx"
xls = pd.ExcelFile(demand_file)
demand_data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

# RAR Data
rar_summary_fmt = None
rar_additions_fmt = None
rar_thermal_plf = None
rar_prm_fmt = None

try:
    rar = pd.ExcelFile("RAR1.xlsx")
    # Sheet1
    rar_summary = pd.read_excel(rar, "Sheet1")
    rar_summary_fmt = rar_summary.copy()

    # Sheet2
    rar_additions = pd.read_excel(rar, "Sheet2")
    rar_additions_fmt = rar_additions.copy()

    # Sheet3
    rar_thermal_plf = pd.read_excel(rar, "Sheet3")
    rar_thermal_plf = rar_thermal_plf.rename(columns={rar_thermal_plf.columns[0]: "Plant"})
    for c in rar_thermal_plf.columns[1:]:
        rar_thermal_plf[c] = pd.to_numeric(rar_thermal_plf[c], errors="coerce")

    # Sheet4 - PRM vs LoLP & NENS
    rar_prm = pd.read_excel(rar, "Sheet4")
    rar_prm = rar_prm.rename(columns={rar_prm.columns[0]:"PRM", rar_prm.columns[1]:"LoLP", rar_prm.columns[2]:"NENS"})
    rar_prm = rar_prm[1:]
    rar_prm['PRM'] = (rar_prm['PRM'] * 100).round(1)
    for c in ["PRM","LoLP","NENS"]:
        rar_prm[c] = pd.to_numeric(rar_prm[c], errors="coerce")
    rar_prm_fmt = rar_prm.copy()

except Exception as e:
    st.warning(f"RAR1.xlsx not found or could not be read: {e}")

# ---------------------------
# Tabs
# ---------------------------
st.title("Annual Power Procurement Planning Outcomes")
tab6, tab7, tab4, tab5, tab2, tab3= st.tabs([
    "Demand Forecasts",
    "Price Forecasts",
    # "RAR Requirements & PRM",
    "Optimal Capacity Mix",
    "Regions-wise Capacity Procurement",
    "Daily Supply-Demand",
    "Thermal PLF Heatmap"
])

# # ---------------------------
# # Tab 1: RAR Requirements + PRM chart
# # ---------------------------
# with tab1:
#     if rar_summary_fmt is not None:
#         st.subheader("Forecasted Demand and Resource Adequacy Requirement")
#         st.dataframe(rar_summary_fmt, use_container_width=True, height=140)
#         st.markdown("---")

#     # PRM vs LoLP/NENS
#     if rar_prm_fmt is not None:
#         st.subheader("PRM vs LoLP / NENS")

#         # Keep PRM as string (like "3%", "5%")
#         df = rar_prm_fmt.copy()
#         df["PRM_str"] = df["PRM"].astype(str) if df["PRM"].dtype == object else df["PRM"].astype(str) + "%"

#         # Convert LoLP & NENS to percentages
#         df["LoLP_pct"] = df["LoLP"] * 100
#         df["NENS_pct"] = df["NENS"] * 100

#         # Highlight least PRM where LoLP <=0.2% & NENS <=0.05%
#         cond = (df["LoLP_pct"] <= 0.2) & (df["NENS_pct"] <= 0.05)
#         highlight_prm = df.loc[cond, "PRM_str"].values[0] if cond.any() else None

#         fig_prm = go.Figure()

#         # LoLP bar
#         fig_prm.add_trace(go.Bar(
#             x=df["PRM_str"],
#             y=df["LoLP_pct"],
#             name="LoLP (%)",
#             marker_color="#a8e6a3",  # light green
#             hovertemplate="PRM: %{x}<br>LoLP: %{y:.2f}%",
#         ))

#         # Highlight bar without legend
#         if highlight_prm:
#             fig_prm.add_trace(go.Bar(
#                 x=[highlight_prm],
#                 y=df.loc[df["PRM_str"]==highlight_prm, "LoLP_pct"],
#                 marker_color="#a8e6a3",
#                 marker_pattern_shape="/",
#                 showlegend=False,
#                 hovertemplate="PRM: %{x}<br>LoLP: %{y:.2f}% (Target)"
#             ))

#         # NENS line
#         fig_prm.add_trace(go.Scatter(
#             x=df["PRM_str"],
#             y=df["NENS_pct"],
#             name="NENS (%)",
#             mode="lines+markers",
#             line=dict(color="#7fa8d5", width=3),  # light navy blue
#             marker=dict(symbol="circle", size=8),
#             yaxis="y2",
#             hovertemplate="PRM: %{x}<br>NENS: %{y:.3f}%",
#         ))

#         # Highlight NENS point without legend
#         if highlight_prm:
#             highlight_nens = df.loc[df["PRM_str"]==highlight_prm, "NENS_pct"].values[0]
#             fig_prm.add_trace(go.Scatter(
#                 x=[highlight_prm],
#                 y=[highlight_nens],
#                 mode="markers",
#                 marker=dict(symbol="circle", size=12, color="#7fa8d5", line=dict(color="black", width=2)),
#                 yaxis="y2",
#                 showlegend=False,
#                 hovertemplate="PRM: %{x}<br>NENS: %{y:.3f}% (Target)"
#             ))

#         fig_prm.update_layout(
#             xaxis_title="PRM (%)",
#             yaxis=dict(title="LoLP (%)", color="#a8e6a3", showgrid=False),
#             yaxis2=dict(title="NENS (%)", overlaying="y", side="right", color="#7fa8d5", showgrid=False),
#             barmode="overlay",
#             legend=dict(x=0.9, y=0.98),
#             margin=dict(l=60, r=60, t=40, b=40),
#             hovermode="x unified",
#             template="plotly_white"
#         )
#         st.plotly_chart(fig_prm, use_container_width=True)

# ---------------------------
# Tab 2: Daily Supply-Demand
# ---------------------------
with tab2:
    col1, col2 = st.columns([1,3])
    with col1:
        year_sel = st.selectbox("Select Year", sorted([int(y) for y in demand_data.keys()]))
        df_year = demand_data[str(year_sel)]
        day_sel = st.selectbox("Select Day", sorted(df_year["Day"].unique()))
    with col2:
        df_day = df_year[df_year["Day"] == day_sel].copy()
        slot = df_day["Slot"]
        tech_cols_day = [c for c in df_day.columns if c not in ["Day","Slot","Demand",
                                                                "BESS-D","BESS-C","PSP-D","PSP-C","Unmet","Excess"]]
        pos_cols = tech_cols_day + ["BESS-D","PSP-D","Unmet"]
        neg_cols = ["BESS-C","PSP-C","Excess"]
        for c in pos_cols + neg_cols:
            if c not in df_day.columns:
                df_day[c] = 0
        df_contrib = pd.DataFrame({c:(df_day[c] if c in pos_cols else -df_day[c]) for c in pos_cols+neg_cols})
        color_seq = px.colors.qualitative.Set3 + px.colors.qualitative.Plotly + px.colors.qualitative.Safe
        colors = {c: color_seq[i%len(color_seq)] for i,c in enumerate(df_contrib.columns)}
        fig_daily = go.Figure()
        for c in df_contrib.columns:
            fig_daily.add_trace(go.Bar(x=slot, y=df_contrib[c], name=c, marker_color=colors[c],
                                       hovertemplate=f"{c}: %{{y}} MW"))
        fig_daily.add_trace(go.Scatter(x=slot, y=df_day["Demand"], name="Demand",
                                       mode="lines+markers", line=dict(color="black", width=3),
                                       hovertemplate="Demand: %{y} MW"))
        fig_daily.update_layout(title=f"Daily Supply-Demand Balance — Year {year_sel}, Day {day_sel}",
                                xaxis_title="Hour", yaxis_title="MW", barmode='relative',
                                legend_title="Technology", template="plotly_white",
                                hovermode="x unified")
        st.plotly_chart(fig_daily, use_container_width=True)

# ---------------------------
# Tab 3: Thermal PLF Heatmap
# ---------------------------
with tab3:
    st.subheader("Thermal Plants PLF Heatmap")
    if rar_thermal_plf is None or rar_thermal_plf.empty:
        st.warning("Sheet3 (Thermal PLF) not found in RAR1.xlsx.")
    else:
        rar_thermal_plf["Plant"] = pd.Categorical(rar_thermal_plf["Plant"], categories=rar_thermal_plf["Plant"], ordered=True)
        df_heat = rar_thermal_plf.melt(id_vars="Plant", var_name="Year", value_name="PLF")
        heatmap_data = df_heat.pivot(index="Plant", columns="Year", values="PLF")
        pleasant_r_y_g = [[0.0,"#ff9999"], [0.5,"#ffff99"], [1.0,"#99ff99"]]
        fig_heat = px.imshow(heatmap_data, aspect="auto", color_continuous_scale=pleasant_r_y_g,
                             labels=dict(color="PLF (%)"), text_auto=True)
        fig_heat.update_traces(hovertemplate="Plant: %{y}<br>Year: %{x}<br>PLF: %{z:.2f}%")
        fig_heat.update_layout(xaxis_title="Year", yaxis_title="Thermal Plant",
                               coloraxis_colorbar=dict(title="PLF (%)"),
                               xaxis=dict(tickmode="linear"),
                               yaxis=dict(tickfont=dict(color="black", size=12, family="Arial")),
                               margin=dict(l=60,r=30,t=40,b=40))
        st.plotly_chart(fig_heat, use_container_width=True)
        st.dataframe(rar_thermal_plf, use_container_width=True)

# ---------------------------
# Tab 4: Cumulative + YoY Capacities
# ---------------------------
with tab4:
    techs_available = sorted(df_long['Technology'].unique())
    years_available = sorted(df_long['Year'].dropna().unique())
    col1, col2 = st.columns([1,3])
    with col1:
        view_mode = st.radio("Select View", ["Year-wise (Stacked Bar)", "Technology Trends (Line Chart)"])
        percent_toggle = st.checkbox("Show Percent Shares", value=False)
        selected_techs = st.multiselect("Select Technologies", options=techs_available, default=techs_available)
        selected_years = st.slider("Select Year Range", 
                                   min_value=int(min(years_available)),
                                   max_value=int(max(years_available)),
                                   value=(int(min(years_available)), int(max(years_available))),
                                   step=1)
    with col2:
        placeholder = st.empty()
    years_to_plot = list(range(selected_years[0], selected_years[1]+1))
    with placeholder.container():
        if len(selected_techs) == 0:
            st.warning("Please select at least one technology to display.")
        else:
            if view_mode.startswith('Year-wise'):
                fig = make_stacked_bar(df_long, techs=selected_techs, years=years_to_plot, percent=percent_toggle)
            else:
                fig = make_tech_line_chart(df_long, techs=selected_techs, years=years_to_plot)
            st.plotly_chart(fig, use_container_width=True)
    filtered_df = df_long[df_long['Technology'].isin(selected_techs) & df_long['Year'].isin(years_to_plot)]
    pivot = filtered_df.pivot_table(
        index='Year',
        columns='Technology',
        values='MW',
        aggfunc='sum',
        fill_value=0
    )
    
    # keep selected tech order
    pivot = pivot.reindex(columns=selected_techs)
    
    # ---- FORMAT: whole numbers + replace 0 with "-" ----
    pivot_fmt = pivot.round(0).astype(int)          # remove decimals
    pivot_fmt = pivot_fmt.replace(0, "-")           # replace zero with dash
    
    st.markdown("---")
    st.subheader("Cumulative Capacities (MW)")
    st.dataframe(pivot_fmt.reset_index(), height=400)
    
    csv_bytes = pivot_fmt.reset_index().to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download Data as CSV",
        data=csv_bytes,
        file_name="capacity_outcomes.csv",
        mime="text/csv"
    )

    if rar_additions_fmt is not None:
        st.markdown("---")
        st.subheader("Year-on-Year Capacity Additions (MW)")
        st.dataframe(rar_additions_fmt, use_container_width=True, height=320)
        csv_add = rar_additions_fmt.to_csv(index=False).encode('utf-8')
        st.download_button("Download Additions Data", data=csv_add, file_name="Capacity_Additions.csv", mime="text/csv")
        
with tab5:
    st.subheader("Regional Capacity Procurement Plan")

    # ---- Dummy Data Generation ----
    years = list(range(2026, 2036))
    regions = ["Rajasthan", "Tamil Nadu", "Gujarat", "Target State"]
    technologies = ["Solar", "Wind", "Hydro", "BESS", "PSP", "Thermal", "STOA", "Market"]

    # Create dummy capacity additions (MW)
    import numpy as np
    np.random.seed(42)
    data = []

    for year in years:
        # Solar - Rajasthan + Target State
        data.append([year, "Rajasthan", "Solar", np.random.randint(300, 800)])
        data.append([year, "Target State", "Solar", np.random.randint(400, 900)])

        # Wind - Tamil Nadu + Gujarat + Target State
        data.append([year, "Tamil Nadu", "Wind", np.random.randint(300, 700)])
        data.append([year, "Gujarat", "Wind", np.random.randint(300, 700)])
        data.append([year, "Target State", "Wind", np.random.randint(200, 600)])

        # All other sources from Target State
        for tech in ["Hydro", "BESS", "PSP", "Thermal", "STOA", "Market"]:
            data.append([year, "Target State", tech, np.random.randint(200, 800)])

    df_sankey = pd.DataFrame(data, columns=["Year", "Region", "Technology", "Capacity_MW"])

    # ---- Filters ----
    year_sel = st.selectbox("Select Year", options=years, index=0)
    df_year = df_sankey[df_sankey["Year"] == year_sel]

    # ---- Sankey Node Preparation ----
    all_nodes = list(df_year["Region"].unique()) + list(df_year["Technology"].unique())
    node_dict = {node: i for i, node in enumerate(all_nodes)}

    sources = df_year["Region"].map(node_dict)
    targets = df_year["Technology"].map(node_dict)
    values = df_year["Capacity_MW"]

    # ---- Node Colors ----
    node_colors = []
    for node in all_nodes:
        if node in ["Rajasthan", "Tamil Nadu", "Gujarat"]:
            node_colors.append("#90caf9")  # light blue for external states
        elif node == "Target State":
            node_colors.append("#81c784")  # green for target
        else:
            node_colors.append("#ffd54f")  # yellow for technologies

    # ---- Link Colors ----
    link_colors = []
    for s, t in zip(df_year["Region"], df_year["Technology"]):
        if s == "Rajasthan":
            link_colors.append("rgba(33,150,243,0.4)")  # blueish
        elif s == "Tamil Nadu":
            link_colors.append("rgba(255,87,34,0.4)")  # orange
        elif s == "Gujarat":
            link_colors.append("rgba(156,39,176,0.4)")  # purple
        else:
            link_colors.append("rgba(102,187,106,0.4)")  # green

    # ---- Create Sankey Figure ----
    fig_sankey = go.Figure(
        go.Sankey(
            node=dict(
                pad=18,
                thickness=18,
                line=dict(color="gray", width=0.5),
                label=all_nodes,
                color=node_colors,
            ),
            link=dict(
                source=sources,
                target=targets,
                value=values,
                color=link_colors,
                hovertemplate="Region: %{source.label}<br>Technology: %{target.label}<br>Capacity: %{value} MW<extra></extra>"
            )
        )
    )

    fig_sankey.update_layout(
        title=dict(
            text=f"Capacity Procurement from Different Regions – {year_sel}",
            x=0.02,
            xanchor="left",
            font=dict(size=18, color="black", family="Arial"),
        ),
        font=dict(size=13, color="black", family="Arial"),
        margin=dict(l=40, r=40, t=60, b=30),
        height=600,
        template="plotly_white"
    )
    
    # Make node border cleaner (removes the “boxed” feel)
    fig_sankey.update_traces(
        node=dict(
            pad=22,
            thickness=16,
            line=dict(color="rgba(0,0,0,0.2)", width=0.5),  # lighter border
            label=all_nodes,
            color=node_colors,
        ),
        selector=dict(type="sankey")
    )


    # ---- Display Sankey ----
    st.plotly_chart(fig_sankey, use_container_width=True)

    # ---- Show Underlying Data Table ----
    st.markdown("#### Yearly Regional-Technology Capacity Additions (MW)")
    st.dataframe(df_year, use_container_width=True, height=350)
    
# ---------------------------
# Tab 6: Demand Forecasts
# ---------------------------
with tab6:
    # st.subheader("Demand Forecasts (Assam)")

    demand_xlsx_path = "Demand.xlsx"   # adjust if needed (or use OUTPUT_DIR)
    if not os.path.exists(demand_xlsx_path):
        st.warning(f"Demand.xlsx not found at: {demand_xlsx_path}")
    else:
        xls_demand = pd.ExcelFile(demand_xlsx_path)

        # ---- Yearly MU Trend ----
        if "Demand MU Assam" in xls_demand.sheet_names:
            df_mu = pd.read_excel(xls_demand, sheet_name="Demand MU Assam")
            df_mu = df_mu.rename(columns={df_mu.columns[0]: "Year", df_mu.columns[1]: "MU"})
            df_mu["Year"] = pd.to_numeric(df_mu["Year"], errors="coerce")
            df_mu["MU"] = pd.to_numeric(df_mu["MU"], errors="coerce")
            df_mu = df_mu.dropna().sort_values("Year")

            st.markdown("### Energy Requirement (MU)")
            
            fig_mu = px.bar(df_mu, x="Year", y="MU", text="MU")
            fig_mu.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            fig_mu.update_layout(
                xaxis_title="Year",
                yaxis_title="Energy (MU)",
                hovermode="x unified",
                template="plotly_white",
                margin=dict(t=40, b=40, l=40, r=20)
            )
            st.plotly_chart(fig_mu, use_container_width=True)
        else:
            st.warning("Sheet 'Demand MU Assam' not found in Demand.xlsx")

        mu_csv = df_mu.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download Energy Requirement (MU) Data",
            data=mu_csv,
            file_name="Yearly_Demand_MU_Assam.csv",
            mime="text/csv"
        )

        # ---- Yearly Peak Demand Trend ----
        if "Peak Demand Assam" in xls_demand.sheet_names:
            df_peak = pd.read_excel(xls_demand, sheet_name="Peak Demand Assam")
            df_peak = df_peak.rename(columns={df_peak.columns[0]: "Year", df_peak.columns[1]: "PeakDemand_MW"})
            df_peak["Year"] = pd.to_numeric(df_peak["Year"], errors="coerce")
            df_peak["PeakDemand_MW"] = pd.to_numeric(df_peak["PeakDemand_MW"], errors="coerce")
            df_peak = df_peak.dropna().sort_values("Year")

            st.markdown("### Peak Demand (MW)")
            
            fig_peak = px.bar(df_peak, x="Year", y="PeakDemand_MW", text="PeakDemand_MW")
            fig_peak.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            fig_peak.update_layout(
                xaxis_title="Year",
                yaxis_title="Peak Demand (MW)",
                hovermode="x unified",
                template="plotly_white",
                margin=dict(t=40, b=40, l=40, r=20)
            )
            st.plotly_chart(fig_peak, use_container_width=True)
        else:
            st.warning("Sheet 'Peak Demand Assam' not found in Demand.xlsx")

        peak_csv = df_peak.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download Peak Demand (MW) Data",
            data=peak_csv,
            file_name="Yearly_Peak_Demand_Assam.csv",
            mime="text/csv"
        )

        st.markdown("---")
        st.markdown("### Hourly Demand Pattern")
        
        # ---- Year selection (show only year, internally map to "<year> Assam") ----
        year_map = {}
        for y in range(2027, 2037):   # 2027 to 2036 inclusive
            sheet_name = f"{y} Assam"
            if sheet_name in xls_demand.sheet_names:
                year_map[y] = sheet_name
        
        available_years = sorted(year_map.keys())
        
        if len(available_years) == 0:
            st.warning("No sheets found like '2027 Assam' ... '2036 Assam' in Demand.xlsx")
        else:
            col1, col2 = st.columns([1, 3])
        
            with col1:
                year_sel = st.selectbox("Select Year", options=available_years, index=0, key="demand_year_sel")
        
            sheet_sel = year_map[year_sel]
            df_year = pd.read_excel(xls_demand, sheet_name=sheet_sel)
            df_year = df_year.dropna(axis=0, how="all").dropna(axis=1, how="all")
        
            # first column = Slot
            slot_col = df_year.columns[0]
            df_year = df_year.rename(columns={slot_col: "Slot"})
        
            # day columns are like 324, 12, 250 etc
            day_cols = [c for c in df_year.columns if c != "Slot"]
            
            # Convert to int safely for sorting (day identifiers)
            day_nums = []
            day_other = []  # if any non-numeric day columns exist (rare)
            for c in day_cols:
                try:
                    day_nums.append(int(str(c)))
                except:
                    day_other.append(str(c))
            
            day_nums = sorted(day_nums)              # numeric sort low -> high
            day_cols_str = [str(d) for d in day_nums] + sorted(day_other)
        
            with col1:
                day_sel_str = st.selectbox("Select Representative Day", options=day_cols_str, index=0, key="demand_day_sel")
        
            # map selected day back to column name type
            if day_sel_str.isdigit() and int(day_sel_str) in df_year.columns:
                day_sel = int(day_sel_str)
            else:
                day_sel = day_sel_str
        
            df_plot = df_year[["Slot", day_sel]].copy()
            df_plot = df_plot.rename(columns={day_sel: "Demand_MW"})
            df_plot["Slot"] = pd.to_numeric(df_plot["Slot"], errors="coerce")
            df_plot["Demand_MW"] = pd.to_numeric(df_plot["Demand_MW"], errors="coerce")
            df_plot = df_plot.dropna()
        
            with col2:
                fig_day = px.area(df_plot, x="Slot", y="Demand_MW")
                fig_day.update_layout(
                    title=f"Hourly Demand Profile — {year_sel} | Day {day_sel_str}",
                    xaxis_title="Slot (Hour)",
                    yaxis_title="Demand (MW)",
                    hovermode="x unified",
                    template="plotly_white",
                    margin=dict(t=50, b=40, l=40, r=20)
                )
                st.plotly_chart(fig_day, use_container_width=True)
                
            day_profile_download = df_plot.copy()
            day_profile_download.insert(0, "Year", year_sel)
            day_profile_download.insert(1, "Day", day_sel_str)
            
            day_csv = day_profile_download.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Download Selected Day Demand Profile Data",
                data=day_csv,
                file_name=f"DemandProfile_{year_sel}_Day{day_sel_str}_Assam.csv",
                mime="text/csv"
            )

        # # Optional: Show raw tables
        # with st.expander("Show Raw Tables"):
        #     if "Demand MU Assam" in xls_demand.sheet_names:
        #         st.write("Demand MU Assam")
        #         st.dataframe(df_mu, use_container_width=True)
        #     if "Peak Demand Assam" in xls_demand.sheet_names:
        #         st.write("Peak Demand Assam")
        #         st.dataframe(df_peak, use_container_width=True)
        
# ---------------------------
# Tab 7: Price Forecasts
# ---------------------------
with tab7:
    # st.subheader("Price Forecasts (Market)")

    market_xlsx_path = "Market.xlsx"   # or os.path.join(OUTPUT_DIR, "Market.xlsx")
    if not os.path.exists(market_xlsx_path):
        st.warning(f"Market.xlsx not found at: {market_xlsx_path}")
    else:
        xls_price = pd.ExcelFile(market_xlsx_path)

        # Sheets expected: 2027 ... 2036
        year_map = {}
        for y in range(2027, 2037):
            if str(y) in xls_price.sheet_names:
                year_map[y] = str(y)

        available_years = sorted(year_map.keys())

        if len(available_years) == 0:
            st.warning("No sheets found in Market.xlsx like '2027' ... '2036'")
        else:
            st.markdown("### Day-wise Market Price Profile (₹/kWh)")

            col1, col2 = st.columns([1, 3])

            with col1:
                year_sel = st.selectbox("Select Year", options=available_years, index=0, key="price_year_sel")

            df_year = pd.read_excel(xls_price, sheet_name=year_map[year_sel])
            df_year = df_year.dropna(axis=0, how="all").dropna(axis=1, how="all")

            # Rename first column as Slot
            slot_col = df_year.columns[0]
            df_year = df_year.rename(columns={slot_col: "Slot"})

            # Day columns
            day_cols = [c for c in df_year.columns if c != "Slot"]

            # Sort day list numerically low -> high
            day_nums = []
            day_other = []
            for c in day_cols:
                try:
                    day_nums.append(int(str(c)))
                except:
                    day_other.append(str(c))
            day_nums = sorted(day_nums)
            day_cols_str = [str(d) for d in day_nums] + sorted(day_other)

            with col1:
                day_sel_str = st.selectbox("Select Representative Day", options=day_cols_str, index=0, key="price_day_sel")

            # Map selection back to correct dataframe column type
            if day_sel_str.isdigit() and int(day_sel_str) in df_year.columns:
                day_sel = int(day_sel_str)
            else:
                day_sel = day_sel_str

            # Build plot DF
            df_plot = df_year[["Slot", day_sel]].copy()
            df_plot = df_plot.rename(columns={day_sel: "Price (Rs./kWh)"})
            df_plot["Slot"] = pd.to_numeric(df_plot["Slot"], errors="coerce")
            df_plot["Price (Rs./kWh)"] = pd.to_numeric(df_plot["Price (Rs./kWh)"], errors="coerce")
            df_plot = df_plot.dropna()

            with col2:
                fig_price = px.area(df_plot, x="Slot", y="Price (Rs./kWh)")
                fig_price.update_layout(
                    title=f"Market Price Profile — Year {year_sel} | Day {day_sel_str}",
                    xaxis_title="Slot (Hour)",
                    yaxis_title="Price (₹/kWh)",
                    hovermode="x unified",
                    template="plotly_white",
                    margin=dict(t=50, b=40, l=40, r=20)
                )
                st.plotly_chart(fig_price, use_container_width=True)

                # Download selected day CSV
                dl_df = df_plot.copy()
                dl_df.insert(0, "Year", year_sel)
                dl_df.insert(1, "Day", day_sel_str)

                price_day_csv = dl_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="Download Selected Day Market Price Profile",
                    data=price_day_csv,
                    file_name=f"MarketPrice_{year_sel}_Day{day_sel_str}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

            # # -------------------
            # # Additional Analytics
            # # -------------------
            # st.markdown("---")
            # st.markdown(f"### Market Price Analytics — Year {year_sel}")

            # Convert full year sheet to long format for analytics
            df_long_price = df_year.melt(id_vars="Slot", var_name="Day", value_name="Price (Rs./kWh)")
            df_long_price["Slot"] = pd.to_numeric(df_long_price["Slot"], errors="coerce")
            df_long_price["Price (Rs./kWh)"] = pd.to_numeric(df_long_price["Price (Rs./kWh)"], errors="coerce")
            
            # Day to numeric
            df_long_price["Day_num"] = pd.to_numeric(df_long_price["Day"], errors="coerce")
            df_long_price = df_long_price.dropna(subset=["Slot", "Price (Rs./kWh)", "Day_num"])
            df_long_price["Day_num"] = df_long_price["Day_num"].astype(int)
            
            df_long_price["DateLabel"] = df_long_price["Day_num"].apply(
                lambda d: fiscal_day_to_label(d, base_year=year_sel)
            )
            
            sorted_days = (
                df_long_price[["Day_num", "DateLabel"]]
                .drop_duplicates()
                .sort_values("Day_num")
            )
            
            ordered_labels = sorted_days["DateLabel"].tolist()
            
            df_long_price["DateLabel"] = pd.Categorical(
                df_long_price["DateLabel"],
                categories=ordered_labels,
                ordered=True
            )

            # avg_price = df_long_price["Price (Rs./kWh)"].mean()
            # min_price = df_long_price["Price (Rs./kWh)"].min()
            # max_price = df_long_price["Price (Rs./kWh)"].max()
            # p90_price = df_long_price["Price (Rs./kWh)"].quantile(0.90)

            # left_pad, mid, right_pad = st.columns([1, 3, 1])
            
            # with mid:
            #     k1, k2, k3 = st.columns(3)
            #     k1.metric("Avg Price (₹/kWh)", f"{avg_price:.2f}")
            #     k2.metric("Min Price (₹/kWh)", f"{min_price:.2f}")
            #     k3.metric("Max Price (₹/kWh)", f"{max_price:.2f}")
            # k4.metric("P90 Price (₹/kWh)", f"{p90_price:.2f}")

            # Avg price by hour slot across all representative days
            st.markdown(f"### Average Hourly Price Pattern — Year {year_sel}")
            hourly_avg = df_long_price.groupby("Slot", as_index=False)["Price (Rs./kWh)"].mean()

            fig_hourly_avg = px.line(hourly_avg, x="Slot", y="Price (Rs./kWh)", markers=True)
            fig_hourly_avg.update_layout(
                xaxis_title="Slot (Hour)",
                yaxis_title="Average Price (₹/kWh)",
                hovermode="x unified",
                template="plotly_white",
                margin=dict(t=40, b=40, l=40, r=20)
            )
            st.plotly_chart(fig_hourly_avg, use_container_width=True)

            # Heatmap: Slot vs Day
            st.markdown(f"### Market Price Heatmap — Year {year_sel}")
            
            heat_df = df_long_price.pivot(index="Slot", columns="DateLabel", values="Price (Rs./kWh)")
            
            pleasant_r_y_g = [
                [0.0, "#26A69A"],   # low = teal
                [0.5, "#FFF59D"],   # mid = pale yellow
                [1.0, "#D32F2F"]    # high = deeper red
            ]
            
            fig_heat = px.imshow(
                heat_df.T,
                aspect="auto",
                labels=dict(x="Slot (Hour)", y="Representative Day", color="₹/kWh"),
                text_auto=False,
                color_continuous_scale=pleasant_r_y_g
            )
            
            fig_heat.update_layout(
                margin=dict(l=120, r=30, t=40, b=40),
                template="plotly_white"
            )
            
            st.plotly_chart(fig_heat, use_container_width=True)