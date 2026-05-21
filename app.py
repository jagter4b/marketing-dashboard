import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# Set Streamlit Page Configuration
st.set_page_config(
    page_title="Campaign Performance Analysis - Summer Connect 2026",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom premium Egypt-inspired midnight-gold theme CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Playfair+Display:ital,wght@0,600;1,400&display=swap');
    
    /* Base theme override */
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    .main {
        background-color: #0b0e14;
    }
    
    /* Custom Header Styling */
    .header-container {
        background: linear-gradient(135deg, #0f1624 0%, #152035 100%);
        border-left: 5px solid #d4af37;
        padding: 24px;
        border-radius: 10px;
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
    }
    
    .header-title {
        color: #f3e5ab;
        font-family: 'Playfair Display', serif;
        font-size: 32px;
        font-weight: 700;
        margin: 0 0 5px 0;
        letter-spacing: 1px;
    }
    
    .header-subtitle {
        color: #a0aec0;
        font-size: 16px;
        margin: 0;
    }
    
    /* Metric Card Styling */
    .metric-card {
        background: #111a2e;
        border: 1px solid #1e2d4a;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.25);
        transition: all 0.3s ease;
    }
    .metric-card:hover {
        transform: translateY(-4px);
        border-color: #d4af37;
        box-shadow: 0 8px 25px rgba(212, 175, 55, 0.15);
    }
    .metric-label {
        font-size: 13px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #718096;
        margin-bottom: 5px;
    }
    .metric-value {
        font-size: 28px;
        font-weight: 700;
        color: #ffffff;
        margin-bottom: 5px;
    }
    .metric-delta {
        font-size: 13px;
        font-weight: 600;
    }
    .delta-up { color: #00ff87; }
    .delta-down { color: #ff416c; }
    .delta-neutral { color: #a0aec0; }
    
    /* Segment Card Styling */
    .segment-card {
        background: #121f35;
        border-top: 4px solid #d4af37;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 15px;
        height: 100%;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }
    
    /* Tabs Custom CSS */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #0f1624;
        padding: 10px 10px 0px 10px;
        border-radius: 8px 8px 0px 0px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 4px 4px 0px 0px;
        color: #a0aec0;
        font-weight: 600;
        border: none;
        padding: 0px 15px;
        transition: all 0.2s;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #d4af37;
        background-color: rgba(212, 175, 55, 0.05);
    }
    .stTabs [aria-selected="true"] {
        background-color: #1a263f !important;
        color: #d4af37 !important;
        border-bottom: 3px solid #d4af37 !important;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- DATA PIPELINE -----------------
excel_path = "Egypt_Marketing_Campaign_Dashboard (1).xlsx"

@st.cache_data
def load_excel_data(path):
    wb_data = {}
    try:
        # Load sheets with standard naming
        wb_data["campaign_overview"] = pd.read_excel(path, sheet_name="Campaign Overview")
        wb_data["target_audience"] = pd.read_excel(path, sheet_name="Target Audience")
        wb_data["channels_budget"] = pd.read_excel(path, sheet_name="Channels & Budget")
        wb_data["timeline"] = pd.read_excel(path, sheet_name="Timeline")
        wb_data["kpis_metrics"] = pd.read_excel(path, sheet_name="KPIs & Metrics")
        wb_data["content_calendar"] = pd.read_excel(path, sheet_name="Content Calendar")
        wb_data["competitor_analysis"] = pd.read_excel(path, sheet_name="Competitor Analysis")
        wb_data["risk_register"] = pd.read_excel(path, sheet_name="Risk Register")
        wb_data["weekly_performance"] = pd.read_excel(path, sheet_name="Weekly Performance")
        wb_data["financial_summary"] = pd.read_excel(path, sheet_name="Financial Summary")
        wb_data["regional_breakdown"] = pd.read_excel(path, sheet_name="Regional Breakdown")
    except Exception as e:
        st.error(f"Error loading Excel workbook: {e}")
    return wb_data

data = load_excel_data(excel_path)

# Custom color palette mapping
BRAND_COLORS = {
    "gold": "#D4AF37",
    "deep_blue": "#0F2C59",
    "dark_navy": "#0C1A30",
    "light_gold": "#F3E5AB",
    "slate": "#1E293B",
    "teal": "#00F2FE",
    "green": "#00FF87",
    "red": "#FF416C",
    "purple": "#8A2BE2"
}

# ----------------- SIDEBAR FILTERS & OVERVIEW -----------------
st.sidebar.markdown(f"""
<div style="text-align: center; padding: 15px 0;">
    <h2 style="color: #d4af37; font-family: 'Playfair Display', serif; margin: 0;">[Confidential Client]</h2>
    <p style="color: #a0aec0; font-size: 12px; margin: 0;">SUMMER CONNECT 2026</p>
</div>
<hr style="border-top: 1px solid #1e2d4a; margin: 10px 0;"/>
""", unsafe_allow_html=True)

st.sidebar.subheader("🎛️ Dashboard Controls")

# Global filter: Weeks
all_weeks = ["All Campaign (13 Weeks)"] + [f"Week {i}" for i in range(1, 14)]
selected_week_filter = st.sidebar.selectbox("Filter by Timeline Range", all_weeks)

# Global filter: Status Selector
if "channels_budget" in data:
    statuses = ["All Channels"] + list(data["channels_budget"]["Status"].dropna().unique())
    selected_status_filter = st.sidebar.selectbox("Filter Channels by Status", statuses)
else:
    selected_status_filter = "All Channels"

st.sidebar.markdown("<br>", unsafe_allow_html=True)

# Small interactive checklist on sidebar for executives
st.sidebar.subheader("🎯 Campaign Tasks Status")
st.sidebar.checkbox("Phase 1: Planning & Setup", value=True, disabled=True)
st.sidebar.checkbox("Phase 2: Pre-Launch (Teaser)", value=True, disabled=True)
st.sidebar.checkbox("Phase 3: Launch & Awareness", value=True, disabled=False)
st.sidebar.checkbox("Phase 4: Engagement & Conversion", value=True, disabled=False)
st.sidebar.checkbox("Phase 5: Optimization", value=False, disabled=False)
st.sidebar.checkbox("Phase 6: Wrap-up & Evaluation", value=False, disabled=False)

st.sidebar.markdown("""
<div style="background-color: #111a2e; border: 1px solid #1e2d4a; border-radius: 8px; padding: 15px; margin-top: 20px;">
    <h4 style="color: #d4af37; margin: 0 0 5px 0; font-size: 14px;">📤 Export Report</h4>
    <p style="color: #a0aec0; font-size: 11px; margin-bottom: 10px;">Download custom filtered Excel tables.</p>
</div>
""", unsafe_allow_html=True)

if st.sidebar.button("Generate CSV of Active Metrics"):
    if "weekly_performance" in data:
        csv = data["weekly_performance"].to_csv(index=False)
        st.sidebar.download_button(
            label="💾 Download Weekly CSV",
            data=csv,
            file_name="Egypt_Campaign_Weekly_Performance.csv",
            mime="text/csv",
        )

# ----------------- MAIN APP HEADER -----------------
st.markdown("""
<div class="header-container">
    <div class="header-title">Campaign Performance Analysis</div>
    <div class="header-subtitle">Performance Tracker — [Confidential Client] "Summer Connect" 2026</div>
</div>
""", unsafe_allow_html=True)

# Filter Data based on selections
weekly_df = data.get("weekly_performance", pd.DataFrame())
channels_df = data.get("channels_budget", pd.DataFrame())

if selected_week_filter != "All Campaign (13 Weeks)":
    week_num = int(selected_week_filter.split(" ")[1])
    filtered_weekly = weekly_df[weekly_df["Week"] == f"Week {week_num}"]
else:
    filtered_weekly = weekly_df

if selected_status_filter != "All Channels" and not channels_df.empty:
    filtered_channels = channels_df[channels_df["Status"] == selected_status_filter]
else:
    filtered_channels = channels_df

# Helper function to render a custom metric card
def render_metric_card(label, value, delta, delta_type="up"):
    delta_class = "delta-up" if delta_type == "up" else ("delta-down" if delta_type == "down" else "delta-neutral")
    delta_symbol = "+" if delta_type == "up" and not str(delta).startswith("+") else ""
    
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-delta {delta_class}">{delta_symbol}{delta}</div>
    </div>
    """, unsafe_allow_html=True)

# ----------------- TABS CREATION -----------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🏆 Executive Overview",
    "📊 Channels & Budget",
    "🎯 Audience & Geography",
    "📅 Timeline & Calendar",
    "⚔️ Competitors & SWOT",
    "⚠️ Risk & Briefing"
])

# ==================== TAB 1: EXECUTIVE OVERVIEW ====================
with tab1:
    # 1. Metric Cards Grid
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    # Calculate values dynamically based on filtered data
    if not filtered_weekly.empty:
        total_reach = filtered_weekly["Reach"].sum()
        total_impressions = filtered_weekly["Impressions"].sum()
        total_clicks = filtered_weekly["Clicks"].sum()
        total_conversions = filtered_weekly["Conversions"].sum()
        total_spend = filtered_weekly["Spend (EGP)"].sum()
        total_revenue = filtered_weekly["Revenue (EGP)"].sum()
        
        avg_roas = total_revenue / total_spend if total_spend > 0 else 0
        avg_conv_rate = (total_conversions / total_clicks * 100) if total_clicks > 0 else 0
        avg_ctr = (total_clicks / total_impressions * 100) if total_impressions > 0 else 0
    else:
        total_reach, total_spend, total_revenue, avg_roas, avg_conv_rate = 3247891, 4500000, 17110000, 3.8, 2.9
    
    with col1:
        render_metric_card("Total Reach (YTD)", f"{total_reach:,}", "+12.4% vs Target", "up")
    with col2:
        render_metric_card("Campaign Spend", f"EGP {total_spend/1e6:.2f}M", "100% Allocated", "neutral")
    with col3:
        render_metric_card("Gross Revenue", f"EGP {total_revenue/1e6:.2f}M", "+5.2% vs Forecast", "up")
    with col4:
        render_metric_card("Average ROAS", f"{avg_roas:.2f}x", "-0.2x vs Target", "down")
    with col5:
        render_metric_card("Conversions", f"{total_conversions:,}", "Target: 12.5K", "neutral")
    with col6:
        render_metric_card("Conversion Rate", f"{avg_conv_rate:.2%}" if avg_conv_rate < 1 else f"{avg_conv_rate:.2f}%", "-0.6pp vs Target", "down")
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 2. Main Row: Weekly Trend Chart and Regional Geo-Bubble Map
    chart_col, map_col = st.columns([3, 2])
    
    with chart_col:
        st.subheader("📈 Weekly Spend, Revenue, and ROAS")
        
        fig_trend = go.Figure()
        # Bar chart for Spend
        fig_trend.add_trace(go.Bar(
            x=weekly_df["Week"],
            y=weekly_df["Spend (EGP)"],
            name="Spend (EGP)",
            marker_color="#1E293B",
            yaxis="y"
        ))
        # Bar chart for Revenue
        fig_trend.add_trace(go.Bar(
            x=weekly_df["Week"],
            y=weekly_df["Revenue (EGP)"],
            name="Revenue (EGP)",
            marker_color=BRAND_COLORS["gold"],
            yaxis="y"
        ))
        # Line chart for ROAS (secondary Y-axis)
        fig_trend.add_trace(go.Scatter(
            x=weekly_df["Week"],
            y=weekly_df["ROAS"],
            name="ROAS (Right Axis)",
            line=dict(color=BRAND_COLORS["teal"], width=3),
            yaxis="y2"
        ))
        
        fig_trend.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#a0aec0', family="Outfit"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=0, r=0, t=30, b=0),
            yaxis=dict(
                title=dict(text="EGP (Egyptian Pound)", font=dict(color='#a0aec0')),
                tickfont=dict(color='#a0aec0'),
                gridcolor='#1e2d4a'
            ),
            yaxis2=dict(
                title=dict(text="ROAS (x-multiple)", font=dict(color=BRAND_COLORS["teal"])),
                tickfont=dict(color=BRAND_COLORS["teal"]),
                overlaying="y",
                side="right",
                gridcolor='rgba(0,0,0,0)'
            ),
            xaxis=dict(gridcolor='#1e2d4a'),
            barmode='group'
        )
        st.plotly_chart(fig_trend, use_container_width=True)

    with map_col:
        st.subheader("🗺️ Regional Coverage")
        
        # Regional Coordinates for Stylized Egypt Bubble plot
        regional_df = data.get("regional_breakdown", pd.DataFrame()).copy()
        
        # Coordinates mapping of Egypt major governorates
        coords = {
            "Greater Cairo": {"lat": 30.0444, "lon": 31.2357},
            "Alexandria": {"lat": 31.2001, "lon": 29.9187},
            "Giza": {"lat": 30.0131, "lon": 31.2089},
            "Delta (Mansoura/Tanta/Zagazig)": {"lat": 30.7885, "lon": 31.0003},
            "Upper Egypt (Sohag/Assiut)": {"lat": 27.1810, "lon": 31.1837},
            "Canal (Suez/Ismailia/Port Said)": {"lat": 30.5901, "lon": 32.2647},
            "Red Sea (Hurghada)": {"lat": 27.2579, "lon": 33.8116},
            "Sinai (North/South)": {"lat": 29.5000, "lon": 33.7500}
        }
        
        regional_df["lat"] = regional_df["Region"].map(lambda r: coords.get(r, {}).get("lat", 30))
        regional_df["lon"] = regional_df["Region"].map(lambda r: coords.get(r, {}).get("lon", 30))
        
        fig_map = px.scatter(
            regional_df,
            x="lon",
            y="lat",
            size="Media Spend (EGP)",
            color="ROAS",
            hover_name="Region",
            hover_data=["Population (M)", "Reach", "Revenue (EGP)", "Conversion Rate"],
            color_continuous_scale=[[0, BRAND_COLORS["red"]], [0.5, BRAND_COLORS["gold"]], [1, BRAND_COLORS["green"]]],
            size_max=35,
            title="Regional Revenue and ROAS"
        )
        
        # High styling to make it resemble a stylized geographic locator
        fig_map.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='#0f1624',
            font=dict(color='#a0aec0', family="Outfit"),
            xaxis=dict(showgrid=False, zeroline=False, showticklabels=False, title=""),
            yaxis=dict(showgrid=False, zeroline=False, showticklabels=False, title=""),
            margin=dict(l=0, r=0, t=30, b=0),
            coloraxis_colorbar=dict(title="ROAS", thickness=15, len=0.8)
        )
        
        # Add labels on scatter
        for i, row in regional_df.iterrows():
            fig_map.add_annotation(
                x=row["lon"],
                y=row["lat"] + 0.15,
                text=row["Region"].split(" ")[0],
                showarrow=False,
                font=dict(size=10, color="#ffffff", family="Outfit")
            )
            
        st.plotly_chart(fig_map, use_container_width=True)

    # 3. Quick SWOT alignment & Summary of the first sheet
    st.markdown("<hr style='border-color: #1e2d4a;'/>", unsafe_allow_html=True)
    o1, o2 = st.columns([1, 1])
    with o1:
        st.markdown(f"""
        <div style="background-color: #121f35; border-radius: 8px; padding: 20px; border-left: 4px solid {BRAND_COLORS["gold"]};">
            <h4 style="margin: 0 0 10px 0; color: #f3e5ab;">📝 Summer Connect 2026 Overview</h4>
            <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                <tr><td style="padding: 5px 0; font-weight: bold; color:#718096">Client/Brand:</td><td style="color:#e2e8f0">[Confidential Client] (Consumer Electronics)</td></tr>
                <tr><td style="padding: 5px 0; font-weight: bold; color:#718096">Objective:</td><td style="color:#e2e8f0">Increase Brand Awareness & Drive Q3 Sales by 25%</td></tr>
                <tr><td style="padding: 5px 0; font-weight: bold; color:#718096">Key Target Market:</td><td style="color:#e2e8f0">Cairo, Alexandria, Giza, Delta Region, Upper Egypt</td></tr>
                <tr><td style="padding: 5px 0; font-weight: bold; color:#718096">Core Value Proposition:</td><td style="color:#e2e8f0">Premium mobile accessories at local prices with 2-year warranty</td></tr>
            </table>
        </div>
        """, unsafe_allow_html=True)
        
    with o2:
        st.markdown(f"""
        <div style="background-color: #121f35; border-radius: 8px; padding: 20px; border-left: 4px solid {BRAND_COLORS["teal"]};">
            <h4 style="margin: 0 0 10px 0; color: #00F2FE;">🔍 SWOT Analysis</h4>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 11px; line-height: 1.4;">
                <div><b style="color: #00FF87;">🟢 STRENGTHS</b><br/>Competitive pricing, local warranty, strong distributor network.</div>
                <div><b style="color: #FF416C;">🔴 WEAKNESSES</b><br/>Low brand recognition vs. international giants (Anker, Baseus).</div>
                <div><b style="color: #00F2FE;">🔵 OPPORTUNITIES</b><br/>60% of population under 30, high mobile penetration (95%).</div>
                <div><b style="color: #FF8C00;">🟡 THREATS</b><br/>Economic inflation, import restrictions, currency fluctuation.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)


# ==================== TAB 2: CHANNELS & BUDGET ====================
with tab2:
    st.subheader("📊 Channels & Budget")
    
    col_chan1, col_chan2 = st.columns([3, 2])
    
    with col_chan1:
        if not filtered_channels.empty:
            # Prepare data for plotting
            plot_channels = filtered_channels[filtered_channels["Channel"] != "TOTAL"]
            
            fig_channels = px.bar(
                plot_channels,
                y="Channel",
                x="Budget (EGP)",
                color="Status",
                orientation="h",
                text="Allocation %",
                color_discrete_map={
                    "Active": BRAND_COLORS["gold"],
                    "Planned": BRAND_COLORS["teal"],
                    "Reserved": BRAND_COLORS["slate"]
                },
                title="Budget Allocation & Status"
            )
            fig_channels.update_traces(texttemplate='%{text}%', textposition='outside')
            fig_channels.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#a0aec0', family="Outfit"),
                xaxis=dict(gridcolor='#1e2d4a', title="Budget (EGP)"),
                yaxis=dict(gridcolor='rgba(0,0,0,0)', title=""),
                margin=dict(l=0, r=0, t=40, b=0)
            )
            st.plotly_chart(fig_channels, use_container_width=True)
            
    with col_chan2:
        # Donut Chart for Budget Allocation
        if "channels_budget" in data:
            donut_data = data["channels_budget"][data["channels_budget"]["Channel"] != "TOTAL"].dropna(subset=["Allocation %"])
            fig_donut = px.pie(
                donut_data,
                values="Allocation %",
                names="Channel",
                hole=.4,
                title="Budget Share % by Marketing Channel",
                color_discrete_sequence=px.colors.sequential.YlOrBr_r
            )
            fig_donut.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#a0aec0', family="Outfit"),
                legend=dict(orientation="h", yanchor="bottom", y=-0.5, xanchor="center", x=0.5),
                margin=dict(l=0, r=0, t=40, b=120)
            )
            st.plotly_chart(fig_donut, use_container_width=True)
            
    st.markdown("### 📋 Media Plan & KPIs")
    if not filtered_channels.empty:
        # Format columns for display
        df_display = filtered_channels.copy()
        
        # Format large numbers for better visual appearance in table
        for c in ["Budget (EGP)", "Weekly Spend"]:
            if c in df_display.columns:
                df_display[c] = df_display[c].map(lambda val: f"EGP {val:,.2f}" if not pd.isna(val) else "")
        if "Allocation %" in df_display.columns:
            df_display["Allocation %"] = df_display["Allocation %"].map(lambda val: f"{val:.1f}%" if not pd.isna(val) else "")
            
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
    # 3. Interactive ROI Simulator Card
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"""
    <div style="background-color: #111a2e; border: 1px solid #1e2d4a; border-radius: 12px; padding: 25px;">
        <h3 style="color: #d4af37; margin: 0 0 10px 0; font-family: 'Playfair Display', serif;">🎛️ ROI Simulator</h3>
        <p style="color: #a0aec0; font-size: 13px; margin-bottom: 20px;">Use this sandbox tool to simulate budget re-allocations and project profit outcomes.</p>
    </div>
    """, unsafe_allow_html=True)
    
    sim_col1, sim_col2, sim_col3 = st.columns(3)
    
    with sim_col1:
        sim_spend = st.slider("Target Campaign Ad Spend (EGP)", min_value=1000000, max_value=15000000, value=4500000, step=100000)
    with sim_col2:
        sim_roas = st.slider("Expected ROAS Multiple (x)", min_value=1.0, max_value=8.0, value=3.8, step=0.1)
    with sim_col3:
        sim_cogs = st.slider("Cost of Goods Sold (COGS) % of Revenue", min_value=10, max_value=70, value=30, step=5)
        
    # Run simulation calculations
    sim_gross_revenue = sim_spend * sim_roas
    sim_cogs_cost = sim_gross_revenue * (sim_cogs / 100)
    sim_net_revenue = sim_gross_revenue - sim_cogs_cost
    sim_profit = sim_net_revenue - sim_spend
    sim_roi = (sim_profit / sim_spend) * 100 if sim_spend > 0 else 0
    
    r_col1, r_col2, r_col3, r_col4 = st.columns(4)
    with r_col1:
        st.metric("Projected Gross Revenue", f"EGP {sim_gross_revenue:,.2f}", help="Total sales generated")
    with r_col2:
        st.metric("Net Margin (after COGS)", f"EGP {sim_net_revenue:,.2f}", help="Revenue after manufacturing & product costs")
    with r_col3:
        st.metric("Net Campaign Profit", f"EGP {sim_profit:,.2f}", delta=f"ROI: {sim_roi:.1f}%", delta_color="normal")
    with r_col4:
        # Display recommendation based on ROI
        if sim_roi > 100:
            rec_text = "🟢 HIGH PROFITABILITY: Scale campaign aggressively!"
        elif sim_roi > 20:
            rec_text = "🟡 MODERATE ROI: Target specific channel optimization."
        else:
            rec_text = "🔴 LOW ROI: Re-negotiate COGS or pivot messaging."
        st.markdown(f"""
        <div style="background-color: #0f1624; border-radius: 8px; padding: 12px; border: 1px solid #1e2d4a; height: 100%; display: flex; align-items: center;">
            <span style="font-size: 12px; font-weight: bold; color: #f3e5ab;">{rec_text}</span>
        </div>
        """, unsafe_allow_html=True)


# ==================== TAB 3: TARGET AUDIENCE & GEOGRAPHY ====================
with tab3:
    st.subheader("🎯 Target Audience & Reach")
    
    aud_col1, aud_col2 = st.columns([3, 2])
    
    with aud_col1:
        # Scatter bubble plot of audience segments
        if "target_audience" in data:
            aud_df = data["target_audience"].copy()
            
            # Scatter Plot: Conversion Potential vs Estimated Reach
            fig_aud = px.scatter(
                aud_df,
                x="Conversion Potential",
                y="Estimated Reach",
                size="Estimated Reach",
                color="Priority",
                text="Segment",
                color_discrete_map={
                    "High": BRAND_COLORS["gold"],
                    "Medium": BRAND_COLORS["teal"],
                    "Low": BRAND_COLORS["slate"]
                },
                size_max=50,
                title="Audience Size & Potential"
            )
            fig_aud.update_traces(textposition='top center')
            fig_aud.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#a0aec0', family="Outfit"),
                xaxis=dict(gridcolor='#1e2d4a', title="Conversion Potential Index (Scale 1-15)"),
                yaxis=dict(gridcolor='#1e2d4a', title="Estimated Reach"),
                margin=dict(l=0, r=0, t=40, b=0)
            )
            st.plotly_chart(fig_aud, use_container_width=True)
            
    with aud_col2:
        st.markdown("### 👥 Target Market Segments")
        if "target_audience" in data:
            for idx, row in data["target_audience"].iterrows():
                priority_color = BRAND_COLORS["gold"] if row["Priority"] == "High" else (BRAND_COLORS["teal"] if row["Priority"] == "Medium" else "#a0aec0")
                st.markdown(f"""
                <div class="segment-card">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <h4 style="color: #ffffff; margin: 0; font-size: 15px;">{row['Segment']} ({row['Age Range']})</h4>
                        <span style="background-color: {priority_color}; color: #0b0e14; padding: 2px 8px; font-weight: bold; border-radius: 4px; font-size: 10px;">{row['Priority']} Priority</span>
                    </div>
                    <p style="color: #a0aec0; font-size: 11px; margin: 5px 0;">📍 <b>Location:</b> {row['Location']} | <b>Gender:</b> {row['Gender Split']}</p>
                    <p style="color: #e2e8f0; font-size: 12px; margin: 5px 0;">⚡ <b>Pain Point:</b> {row['Primary Pain Point']}</p>
                    <p style="color: #f3e5ab; font-size: 11px; margin: 0;">📱 <b>Preferred Channels:</b> {row['Preferred Channels']}</p>
                </div>
                """, unsafe_allow_html=True)
                
    st.markdown("<hr style='border-color: #1e2d4a;'/>", unsafe_allow_html=True)
    st.subheader("🗺️ Regional Breakdowns")
    
    if "regional_breakdown" in data:
        reg_df = data["regional_breakdown"].copy()
        
        reg_tab1, reg_tab2 = st.columns(2)
        with reg_tab1:
            fig_reg_spend = px.bar(
                reg_df.sort_values(by="Media Spend (EGP)", ascending=False),
                x="Region",
                y="Media Spend (EGP)",
                color="Priority",
                color_discrete_map={
                    "High": BRAND_COLORS["gold"],
                    "Medium": BRAND_COLORS["teal"],
                    "Low": BRAND_COLORS["slate"]
                },
                title="Regional Media Spend Distribution"
            )
            fig_reg_spend.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#a0aec0', family="Outfit"),
                xaxis=dict(gridcolor='rgba(0,0,0,0)', title=""),
                yaxis=dict(gridcolor='#1e2d4a'),
                margin=dict(l=0, r=0, t=40, b=0)
            )
            st.plotly_chart(fig_reg_spend, use_container_width=True)
            
        with reg_tab2:
            fig_reg_rev = px.bar(
                reg_df.sort_values(by="Revenue (EGP)", ascending=False),
                x="Region",
                y="Revenue (EGP)",
                color="ROAS",
                color_continuous_scale=[[0, BRAND_COLORS["red"]], [0.5, BRAND_COLORS["gold"]], [1, BRAND_COLORS["green"]]],
                title="Regional Revenue Attributed vs ROAS Efficiency"
            )
            fig_reg_rev.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#a0aec0', family="Outfit"),
                xaxis=dict(gridcolor='rgba(0,0,0,0)', title=""),
                yaxis=dict(gridcolor='#1e2d4a'),
                margin=dict(l=0, r=0, t=40, b=0)
            )
            st.plotly_chart(fig_reg_rev, use_container_width=True)


# ==================== TAB 4: TIMELINE & CALENDAR ====================
with tab4:
    st.subheader("📅 Campaign Timeline & Calendar")
    
    # 1. Gantt Timeline View
    if "timeline" in data:
        st.markdown("### 📍 Operational Phases & Milestones")
        t_df = data["timeline"].copy()
        
        # Create a visually striking vertical process flow representation
        for idx, row in t_df.iterrows():
            status_color = "#00FF87" if row["Status"] == "Completed" else ("#00F2FE" if row["Status"] == "In Progress" else "#718096")
            st.markdown(f"""
            <div style="background-color: #0f1624; border: 1px solid #1e2d4a; border-radius: 8px; padding: 15px; margin-bottom: 12px; border-left: 6px solid {status_color}">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px;">
                    <span style="font-weight: bold; font-size: 16px; color: #f3e5ab;">{row['Phase']}</span>
                    <span style="background-color: {status_color}22; color: {status_color}; padding: 3px 10px; font-weight: bold; border-radius: 4px; font-size: 11px;">{row['Status']} ({row['Progress %']}% Done)</span>
                </div>
                <div style="color: #a0aec0; font-size: 12px; margin-bottom: 8px;">⏱️ <b>Schedule:</b> {row['Start Date']} to {row['End Date']} ({row['Duration (Days)']} Days)</div>
                <p style="color: #e2e8f0; font-size: 13px; margin: 0 0 5px 0;">🔧 <b>Activities:</b> {row['Key Activities']}</p>
                <p style="color: #a0aec0; font-size: 12px; margin: 0;">📦 <b>Deliverables:</b> {row['Deliverables']}</p>
            </div>
            """, unsafe_allow_html=True)

    # 2. Content Calendar Spreadsheet Search Layout
    st.markdown("<hr style='border-color: #1e2d4a;'/>", unsafe_allow_html=True)
    st.subheader("📰 Content Calendar")
    
    if "content_calendar" in data:
        cc_df = data["content_calendar"].copy()
        
        # Multi-filters specifically for Content Calendar
        cc_cols = st.columns(4)
        with cc_cols[0]:
            cc_channels = st.multiselect("Filter Content Channel", list(cc_df["Primary Channel"].unique()))
        with cc_cols[1]:
            cc_types = st.multiselect("Filter Content Type", list(cc_df["Content Type"].unique()))
        with cc_cols[2]:
            cc_owners = st.multiselect("Filter Team Owner", list(cc_df["Owner"].unique()))
        with cc_cols[3]:
            cc_states = st.multiselect("Filter Content Status", list(cc_df["Status"].unique()))
            
        # Apply filters
        if cc_channels:
            cc_df = cc_df[cc_df["Primary Channel"].isin(cc_channels)]
        if cc_types:
            cc_df = cc_df[cc_df["Content Type"].isin(cc_types)]
        if cc_owners:
            cc_df = cc_df[cc_df["Owner"].isin(cc_owners)]
        if cc_states:
            cc_df = cc_df[cc_df["Status"].isin(cc_states)]
            
        st.dataframe(cc_df, use_container_width=True, hide_index=True)


# ==================== TAB 5: COMPETITORS & SWOT ====================
with tab5:
    st.subheader("⚔️ Competitor Analysis")
    
    comp_col1, comp_col2 = st.columns([3, 2])
    
    with comp_col1:
        if "competitor_analysis" in data:
            c_df = data["competitor_analysis"].copy()
            
            # Interactive Scatter Chart: Est Spend vs Market Share
            fig_comp = px.scatter(
                c_df,
                x="Digital Spend (Est. EGP)",
                y="Market Share",
                size="Social Followers",
                color="Competitor",
                hover_data=["Price Positioning", "Threat Level"],
                size_max=45,
                title="Market Share vs Ad Spend"
            )
            fig_comp.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#a0aec0', family="Outfit"),
                xaxis=dict(gridcolor='#1e2d4a', title="Estimated Digital Spend (EGP)"),
                yaxis=dict(gridcolor='#1e2d4a', title="Market Share %"),
                margin=dict(l=0, r=0, t=40, b=0)
            )
            st.plotly_chart(fig_comp, use_container_width=True)
            
    with comp_col2:
        st.markdown("### 🏆 Competitor Profiles")
        if "competitor_analysis" in data:
            for idx, row in data["competitor_analysis"].iterrows():
                threat_color = BRAND_COLORS["red"] if row["Threat Level"] == "High" else (BRAND_COLORS["gold"] if row["Threat Level"] == "Medium" else BRAND_COLORS["green"])
                st.markdown(f"""
                <div style="background-color: #0f1624; border: 1px solid #1e2d4a; border-radius: 8px; padding: 15px; margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <h4 style="color: #ffffff; margin: 0; font-size: 15px;">{row['Competitor']} ({row['Price Positioning']})</h4>
                        <span style="background-color: {threat_color}22; color: {threat_color}; padding: 2px 8px; font-weight: bold; border-radius: 4px; font-size: 10px;">{row['Threat Level']} Threat</span>
                    </div>
                    <p style="color: #a0aec0; font-size: 11px; margin: 5px 0;">📊 <b>Market Share:</b> {row['Market Share']}% | 👥 <b>Social Followers:</b> {row['Social Followers']:,}</p>
                    <p style="color: #e2e8f0; font-size: 12px; margin: 5px 0;">💪 <b>Strengths:</b> {row['Strengths']}</p>
                    <p style="color: #ff416c; font-size: 12px; margin: 5px 0;">⚠️ <b>Weaknesses:</b> {row['Weaknesses']}</p>
                    <p style="color: #00ff87; font-size: 12px; margin: 0;">⚡ <b>Our Advantage:</b> {row['Our Advantage']}</p>
                </div>
                """, unsafe_allow_html=True)


# ==================== TAB 6: RISK & BRIEFING ====================
with tab6:
    st.subheader("⚠️ Risk Register")
    
    risk_col1, risk_col2 = st.columns([3, 2])
    
    with risk_col1:
        # Render a custom 3x3 Risk Matrix scatter map
        if "risk_register" in data:
            r_df = data["risk_register"].copy().dropna(subset=["Probability", "Impact"])
            
            # Map strings to coordinates for clean matrix display
            val_map = {"Low": 1, "Medium": 2, "High": 3}
            r_df["x"] = r_df["Probability"].map(val_map)
            r_df["y"] = r_df["Impact"].map(val_map)
            
            # Add jitter to prevent exact overlaps on scatter
            np.random.seed(42)
            r_df["x_jitter"] = r_df["x"] + np.random.uniform(-0.15, 0.15, size=len(r_df))
            r_df["y_jitter"] = r_df["y"] + np.random.uniform(-0.15, 0.15, size=len(r_df))
            
            fig_risk = px.scatter(
                r_df,
                x="x_jitter",
                y="y_jitter",
                color="Risk Score",
                text="Risk ID",
                hover_data=["Risk Description", "Probability", "Impact", "Mitigation Strategy"],
                color_continuous_scale=[[0, BRAND_COLORS["green"]], [0.5, BRAND_COLORS["gold"]], [1, BRAND_COLORS["red"]]],
                size_max=35,
                title="Risk Matrix"
            )
            fig_risk.update_traces(textposition='top center', marker=dict(size=20))
            
            fig_risk.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='#0f1624',
                font=dict(color='#a0aec0', family="Outfit"),
                xaxis=dict(
                    title="Probability",
                    tickmode='array',
                    tickvals=[1, 2, 3],
                    ticktext=['Low', 'Medium', 'High'],
                    range=[0.5, 3.5],
                    gridcolor='#1e2d4a',
                    zeroline=False
                ),
                yaxis=dict(
                    title="Impact",
                    tickmode='array',
                    tickvals=[1, 2, 3],
                    ticktext=['Low', 'Medium', 'High'],
                    range=[0.5, 3.5],
                    gridcolor='#1e2d4a',
                    zeroline=False
                ),
                margin=dict(l=0, r=0, t=40, b=0)
            )
            st.plotly_chart(fig_risk, use_container_width=True)
            
    with risk_col2:
        st.markdown("### 🛡️ Risk Mitigations")
        if "risk_register" in data:
            for idx, row in data["risk_register"].copy().dropna(subset=["Risk ID"]).iterrows():
                score_color = BRAND_COLORS["red"] if row["Risk Score"] >= 9 else (BRAND_COLORS["gold"] if row["Risk Score"] >= 4 else BRAND_COLORS["green"])
                st.markdown(f"""
                <div style="background-color: #0f1624; border-radius: 8px; padding: 12px; margin-bottom: 8px; border: 1px solid #1e2d4a;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-weight: bold; color: {BRAND_COLORS['gold']}; font-size: 13px;">{row['Risk ID']}: {row['Risk Description']}</span>
                        <span style="background-color: {score_color}22; color: {score_color}; border: 1px solid {score_color}; padding: 1px 6px; font-weight: bold; border-radius: 4px; font-size: 10px;">Score: {row['Risk Score']}</span>
                    </div>
                    <p style="color: #a0aec0; font-size: 11px; margin: 4px 0 0 0;">🛠️ <b>Mitigation:</b> {row['Mitigation Strategy']}</p>
                </div>
                """, unsafe_allow_html=True)
                
    st.markdown("<hr style='border-color: #1e2d4a;'/>", unsafe_allow_html=True)
    st.subheader("📊 Financial Summary")
    
    # Financial Statement View
    if "financial_summary" in data:
        fin_df = data["financial_summary"].copy()
        
        # Separate cost columns and metrics
        st.markdown("#### 📂 Operations Expenses Ledger")
        exp_df = fin_df.iloc[0:7].copy()
        # Clean totals display
        for col in ["BUDGETED (EGP)", "ACTUAL (EGP)", "VARIANCE (EGP)"]:
            exp_df[col] = exp_df[col].map(lambda x: f"EGP {float(x):,.2f}" if not pd.isna(x) and str(x) != "-" else "")
        st.dataframe(exp_df, use_container_width=True, hide_index=True)
        
        # Separate revenue metrics
        st.markdown("#### 💳 Attributed Revenue & Capital Multiples")
        rev_metrics_df = fin_df.iloc[10:15][["CATEGORY", "ACTUAL (EGP)", "NOTES"]].copy()
        rev_metrics_df = rev_metrics_df.rename(columns={"CATEGORY": "Metric", "ACTUAL (EGP)": "Value"})
        
        # Format metric values properly
        def format_metrics(row):
            val = row["Value"]
            m_name = row["Metric"]
            if pd.isna(val): return ""
            if m_name in ["Gross Revenue", "Net Revenue (after COGS)", "Campaign Profit"]:
                return f"EGP {float(val):,.2f}"
            elif m_name == "ROI":
                return f"{float(val):.2f}%"
            elif m_name == "ROAS":
                return f"{float(val):.2f}x"
            return str(val)
            
        rev_metrics_df["Value"] = rev_metrics_df.apply(format_metrics, axis=1)
        st.dataframe(rev_metrics_df, use_container_width=True, hide_index=True)
