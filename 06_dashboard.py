import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
 
st.set_page_config(
    page_title="Growth Analytics Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)
 
DATA_PATH = r"C:\Users\Mayank Joshi\Downloads\Marketing_Channel_Project\data"
 
@st.cache_data
def load_data():
    master         = pd.read_csv(os.path.join(DATA_PATH, 'master_table.csv'))
    attribution    = pd.read_csv(os.path.join(DATA_PATH, 'channel_attribution', 'attribution_comparison.csv'))
    ltv_channel    = pd.read_csv(os.path.join(DATA_PATH, 'cac_ltv', 'ltv_by_channel.csv'))
    ltv_segment    = pd.read_csv(os.path.join(DATA_PATH, 'cac_ltv', 'ltv_by_segment.csv'))
    activated      = pd.read_csv(os.path.join(DATA_PATH, 'cac_ltv', 'activated_sellers_ltv.csv'))
    rfm            = pd.read_csv(os.path.join(DATA_PATH, 'segmentation', 'rfm_scores.csv'))
    segments       = pd.read_csv(os.path.join(DATA_PATH, 'segmentation', 'sellers_segmented_final.csv'))
    lookalike      = pd.read_csv(os.path.join(DATA_PATH, 'segmentation', 'high_value_cluster_final.csv'))
    channel_funnel = pd.read_csv(os.path.join(DATA_PATH, 'funnel_analysis', 'channel_funnel.csv'))
    cohort         = pd.read_csv(os.path.join(DATA_PATH, 'funnel_analysis', 'cohort_monthly.csv'))
 
    master['first_contact_date'] = pd.to_datetime(master['first_contact_date'], errors='coerce')
    master['won_date']           = pd.to_datetime(master['won_date'], errors='coerce')
 
    return master, attribution, ltv_channel, ltv_segment, activated, rfm, segments, lookalike, channel_funnel, cohort
 
master, attribution, ltv_channel, ltv_segment, activated, rfm, segments, lookalike, channel_funnel, cohort = load_data()
 
# ── SIDEBAR ───────────────────────────────────────────────────────────────────
st.sidebar.title("Growth Analytics")
st.sidebar.markdown("**Olist Marketing Funnel Analysis**")
st.sidebar.markdown("---")
 
selected_channels = st.sidebar.multiselect(
    "Filter by channel",
    options=sorted(master['origin'].unique().tolist()),
    default=master['origin'].unique().tolist()
)
 
# ── PAGE TITLE ────────────────────────────────────────────────────────────────
st.title("Growth Analytics Dashboard")
 
tab1, tab2, tab3, tab4 = st.tabs([
    "Funnel overview",
    "Attribution analysis",
    "CAC / LTV",
    "Seller segments"
])
 
# ── TAB 1 — FUNNEL OVERVIEW ──────────────────────────────────────────────────
with tab1:
    st.header("Seller acquisition funnel")
 
    filtered        = master[master['origin'].isin(selected_channels)]
    total_leads     = len(filtered)
    converted       = int(filtered['converted'].sum())
    activated_n     = int((filtered['total_orders'] > 0).sum())
    total_rev       = filtered['total_revenue'].sum()
    cvr             = round(converted / total_leads * 100, 2) if total_leads > 0 else 0
    activation_rate = round(activated_n / converted * 100, 1) if converted > 0 else 0
 
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total leads",       f"{total_leads:,}")
    c2.metric("Converted sellers", f"{converted:,}")
    c3.metric("Conversion rate",   f"{cvr}%")
    c4.metric("Activated sellers", f"{activated_n:,}")
    c5.metric("Total revenue",     f"${total_rev:,.0f}")
 
    st.markdown("---")
    col1, col2 = st.columns(2)
 
    with col1:
        st.subheader("Funnel volume")
        funnel_data = pd.DataFrame({
            'Stage': [
                'Leads generated',
                'Leads with contact date',
                'Closed deals',
                'Made first order',
                'Repeat seller (3+ orders)'
            ],
            'Count': [
                len(filtered),
                int(filtered['first_contact_date'].notna().sum()),
                converted,
                activated_n,
                int((filtered['total_orders'] >= 3).sum())
            ]
        })
        fig_funnel = go.Figure(go.Funnel(
            y=funnel_data['Stage'],
            x=funnel_data['Count'],
            textinfo="value+percent initial",
            marker=dict(color=['#1D9E75', '#5DCAA5', '#7F77DD', '#EF9F27', '#D85A30'])
        ))
        fig_funnel.update_layout(height=400, margin=dict(l=0, r=0, t=20, b=0))
        st.plotly_chart(fig_funnel, use_container_width=True)
 
    with col2:
        st.subheader("Conversion rate by channel")
        ch_filtered = channel_funnel[
            channel_funnel['origin'].isin(selected_channels)
        ].sort_values('cvr', ascending=True)
 
        fig_cvr = px.bar(
            ch_filtered, x='cvr', y='origin', orientation='h',
            color='cvr', color_continuous_scale='Teal',
            labels={'cvr': 'Conversion rate (%)', 'origin': 'Channel'},
            text='cvr'
        )
        fig_cvr.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fig_cvr.update_layout(height=400, margin=dict(l=0, r=0, t=20, b=0),
                              coloraxis_showscale=False)
        st.plotly_chart(fig_cvr, use_container_width=True)
 
    st.subheader("Monthly cohort — conversion rate trend")
    fig_cohort = make_subplots(specs=[[{"secondary_y": True}]])
    fig_cohort.add_trace(
        go.Bar(x=cohort['contact_month'], y=cohort['leads'],
               name='Total leads', marker_color='#B5D4F4', opacity=0.7),
        secondary_y=False
    )
    fig_cohort.add_trace(
        go.Bar(x=cohort['contact_month'], y=cohort['converted'],
               name='Converted', marker_color='#1D9E75', opacity=0.9),
        secondary_y=False
    )
    fig_cohort.add_trace(
        go.Scatter(x=cohort['contact_month'], y=cohort['cvr'],
                   name='CVR %', mode='lines+markers',
                   line=dict(color='#D85A30', width=2.5),
                   marker=dict(size=6)),
        secondary_y=True
    )
    fig_cohort.update_yaxes(title_text="Count", secondary_y=False)
    fig_cohort.update_yaxes(title_text="Conversion rate (%)", secondary_y=True)
    fig_cohort.update_layout(
        height=380, margin=dict(l=0, r=0, t=20, b=0),
        legend=dict(orientation='h', y=1.1)
    )
    st.plotly_chart(fig_cohort, use_container_width=True)
 
    with st.expander("Channel funnel detail — end to end"):
        ch_display = channel_funnel[channel_funnel['origin'].isin(selected_channels)].copy()
        ch_display = ch_display[[
            'origin', 'total_leads', 'converted', 'cvr',
            'activated', 'activation_rate', 'repeat_sellers',
            'repeat_rate', 'revenue_per_lead'
        ]].sort_values('cvr', ascending=False)
        ch_display.columns = [
            'Channel', 'Leads', 'Converted', 'CVR%',
            'Activated', 'Act%', 'Repeat', 'Rep%', 'Rev/Lead'
        ]
        st.dataframe(ch_display, use_container_width=True, hide_index=True)
 
# ── TAB 2 — ATTRIBUTION ANALYSIS ─────────────────────────────────────────────
with tab2:
    st.header("Channel attribution — model comparison")
 
    model_choice = st.radio(
        "Select attribution model",
        ["Last touch", "Linear", "Markov chain", "Compare all"],
        horizontal=True
    )
 
    col1, col2 = st.columns([2, 1])
 
    with col1:
        if model_choice == "Last touch":
            fig_attr = px.bar(
                attribution.sort_values('last_touch_rev', ascending=True),
                x='last_touch_rev', y='origin', orientation='h',
                color_discrete_sequence=['#1D9E75'],
                labels={'last_touch_rev': 'Attributed revenue ($)', 'origin': 'Channel'},
                title='Last touch attribution',
                text='last_touch_rev'
            )
            fig_attr.update_traces(texttemplate='$%{text:,.0f}', textposition='outside')
 
        elif model_choice == "Linear":
            fig_attr = px.bar(
                attribution.sort_values('linear_attributed_revenue', ascending=True),
                x='linear_attributed_revenue', y='origin', orientation='h',
                color_discrete_sequence=['#7F77DD'],
                labels={'linear_attributed_revenue': 'Attributed revenue ($)', 'origin': 'Channel'},
                title='Linear attribution',
                text='linear_attributed_revenue'
            )
            fig_attr.update_traces(texttemplate='$%{text:,.0f}', textposition='outside')
 
        elif model_choice == "Markov chain":
            fig_attr = px.bar(
                attribution.sort_values('markov_attributed_revenue', ascending=True),
                x='markov_attributed_revenue', y='origin', orientation='h',
                color_discrete_sequence=['#EF9F27'],
                labels={'markov_attributed_revenue': 'Attributed revenue ($)', 'origin': 'Channel'},
                title='Markov chain attribution',
                text='markov_attributed_revenue'
            )
            fig_attr.update_traces(texttemplate='$%{text:,.0f}', textposition='outside')
 
        else:
            fig_attr = go.Figure()
            fig_attr.add_trace(go.Bar(
                name='Last touch', x=attribution['origin'],
                y=attribution['last_touch_rev'], marker_color='#1D9E75'
            ))
            fig_attr.add_trace(go.Bar(
                name='Linear', x=attribution['origin'],
                y=attribution['linear_attributed_revenue'], marker_color='#7F77DD'
            ))
            fig_attr.add_trace(go.Bar(
                name='Markov chain', x=attribution['origin'],
                y=attribution['markov_attributed_revenue'], marker_color='#EF9F27'
            ))
            fig_attr.update_layout(barmode='group', title='All models — side by side')
 
        fig_attr.update_layout(height=420, margin=dict(l=0, r=0, t=40, b=0))
        st.plotly_chart(fig_attr, use_container_width=True)
 
    with col2:
        st.subheader("Model guide")
        st.markdown("""
**Last touch** gives 100% credit to the origin channel.
Simple but misleading for multi-touch journeys.
 
**Linear** splits credit equally across all landing pages used.
Better represents assisted conversions.
 
**Markov chain** uses removal effect — how many conversions
would be lost if this channel didn't exist. Most accurate.
        """)
 
        st.warning(
            "Email and direct traffic are undervalued by "
            "2-5x under last touch vs linear and Markov models."
        )
 
        st.subheader("Comparison table")
        comp = attribution[['origin', 'last_touch_rev',
                             'linear_attributed_revenue',
                             'markov_attributed_revenue']].copy()
        comp.columns = ['Channel', 'Last touch', 'Linear', 'Markov']
        comp = comp.sort_values('Last touch', ascending=False)
        for col in ['Last touch', 'Linear', 'Markov']:
            comp[col] = comp[col].apply(lambda x: f"${x:,.0f}")
        st.dataframe(comp, use_container_width=True, hide_index=True)
 
# ── TAB 3 — CAC / LTV ────────────────────────────────────────────────────────
with tab3:
    st.header("CAC / LTV — seller unit economics")
 
    total_converted = int(master['converted'].sum())
    activation_pct  = round(len(activated) / total_converted * 100, 1)
 
    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Total platform revenue", f"${activated['total_revenue'].sum():,.0f}")
    a2.metric("Active sellers",         f"{len(activated):,}")
    a3.metric("Avg LTV per seller",     f"${activated['total_revenue'].mean():,.0f}")
    a4.metric("Activation rate",        f"{activation_pct}%")
 
    st.markdown("---")
    col1, col2 = st.columns(2)
 
    with col1:
        st.subheader("Avg LTV by channel")
        ltv_ch = ltv_channel.sort_values('avg_ltv', ascending=True)
        fig_ltv_ch = px.bar(
            ltv_ch, x='avg_ltv', y='origin', orientation='h',
            color='avg_ltv', color_continuous_scale='Teal',
            labels={'avg_ltv': 'Avg LTV ($)', 'origin': 'Channel'},
            text='avg_ltv'
        )
        fig_ltv_ch.update_traces(texttemplate='$%{text:,.0f}', textposition='outside')
        fig_ltv_ch.update_layout(height=380, margin=dict(l=0, r=0, t=20, b=0),
                                 coloraxis_showscale=False)
        st.plotly_chart(fig_ltv_ch, use_container_width=True)
 
    with col2:
        st.subheader("Avg LTV by business segment (top 12)")
        top_seg = ltv_segment.head(12).sort_values('avg_ltv', ascending=True)
        fig_ltv_seg = px.bar(
            top_seg, x='avg_ltv', y='business_segment', orientation='h',
            color='avg_ltv', color_continuous_scale='Purples',
            labels={'avg_ltv': 'Avg LTV ($)', 'business_segment': 'Segment'},
            text='avg_ltv'
        )
        fig_ltv_seg.update_traces(texttemplate='$%{text:,.0f}', textposition='outside')
        fig_ltv_seg.update_layout(height=380, margin=dict(l=0, r=0, t=20, b=0),
                                  coloraxis_showscale=False)
        st.plotly_chart(fig_ltv_seg, use_container_width=True)
 
    st.subheader("Pareto curve — revenue concentration")
    act_sorted = activated.sort_values('total_revenue', ascending=False).copy()
    act_sorted['cumulative_revenue'] = act_sorted['total_revenue'].cumsum()
    act_sorted['cumulative_pct']     = (
        act_sorted['cumulative_revenue'] / act_sorted['total_revenue'].sum() * 100
    )
    act_sorted['seller_pct'] = (
        np.arange(1, len(act_sorted) + 1) / len(act_sorted) * 100
    )
 
    threshold_row = act_sorted[act_sorted['cumulative_pct'] >= 80].iloc[0]
    pct_sellers   = round(threshold_row['seller_pct'], 1)
 
    fig_pareto = go.Figure()
    fig_pareto.add_trace(go.Scatter(
        x=act_sorted['seller_pct'],
        y=act_sorted['cumulative_pct'],
        fill='tozeroy',
        fillcolor='rgba(29,158,117,0.15)',
        line=dict(color='#1D9E75', width=2.5),
        name='Revenue concentration'
    ))
    fig_pareto.add_hline(
        y=80, line_dash='dash', line_color='#D85A30',
        annotation_text='80% revenue threshold'
    )
    fig_pareto.add_vline(
        x=pct_sellers, line_dash='dash', line_color='#7F77DD',
        annotation_text=f'{pct_sellers}% of sellers'
    )
    fig_pareto.update_layout(
        xaxis_title='Cumulative % of sellers (ranked by revenue)',
        yaxis_title='Cumulative % of total revenue',
        height=350, margin=dict(l=0, r=0, t=20, b=0)
    )
    st.plotly_chart(fig_pareto, use_container_width=True)
    st.info(
        f"Top **{pct_sellers}%** of sellers drive **80%** of total platform revenue."
    )
 
    with st.expander("LTV by channel — full detail"):
        ltv_display = ltv_channel[[
            'origin', 'activated_sellers', 'avg_ltv',
            'median_ltv', 'leads_per_activation', 'ltv_to_cac_index'
        ]].copy()
        ltv_display.columns = [
            'Channel', 'Active sellers', 'Avg LTV',
            'Median LTV', 'Leads per activation', 'LTV/CAC index'
        ]
        ltv_display['Avg LTV']    = ltv_display['Avg LTV'].apply(lambda x: f"${x:,.0f}")
        ltv_display['Median LTV'] = ltv_display['Median LTV'].apply(lambda x: f"${x:,.0f}")
        st.dataframe(ltv_display, use_container_width=True, hide_index=True)
 
# ── TAB 4 — SELLER SEGMENTS ──────────────────────────────────────────────────
with tab4:
    st.header("Seller segmentation")
 
    col1, col2 = st.columns(2)
 
    with col1:
        st.subheader("RFM segments")
        rfm_summary = (
            rfm.groupby('rfm_segment')
            .agg(
                sellers     = ('seller_id', 'count'),
                avg_revenue = ('monetary', 'mean'),
                avg_orders  = ('frequency', 'mean')
            )
            .reset_index()
            .sort_values('avg_revenue', ascending=False)
        )
        colors_rfm = {
            'champions': '#1D9E75', 'loyal': '#7F77DD',
            'promising': '#EF9F27', 'at_risk': '#D85A30',
            'dormant':   '#B4B2A9'
        }
        fig_rfm = px.bar(
            rfm_summary, x='rfm_segment', y='sellers',
            color='rfm_segment',
            color_discrete_map=colors_rfm,
            text='sellers',
            labels={'rfm_segment': 'Segment', 'sellers': 'Number of sellers'}
        )
        fig_rfm.update_traces(textposition='outside')
        fig_rfm.update_layout(
            height=320, margin=dict(l=0, r=0, t=20, b=0), showlegend=False
        )
        st.plotly_chart(fig_rfm, use_container_width=True)
 
        rfm_display = rfm_summary.copy()
        rfm_display['avg_revenue'] = rfm_display['avg_revenue'].apply(lambda x: f"${x:,.0f}")
        rfm_display['avg_orders']  = rfm_display['avg_orders'].apply(lambda x: f"{x:.1f}")
        rfm_display.columns = ['Segment', 'Sellers', 'Avg revenue', 'Avg orders']
        st.dataframe(rfm_display, use_container_width=True, hide_index=True)
 
    with col2:
        st.subheader("K-means clusters")
        if 'cluster_label' in segments.columns:
            cluster_summary = (
                segments.groupby('cluster_label')
                .agg(
                    sellers       = ('seller_id', 'count'),
                    avg_revenue   = ('monetary', 'mean'),
                    total_revenue = ('monetary', 'sum')
                )
                .reset_index()
            )
            cluster_summary['revenue_share'] = (
                cluster_summary['total_revenue'] /
                cluster_summary['total_revenue'].sum() * 100
            ).round(1)
            cluster_summary = cluster_summary.sort_values('avg_revenue', ascending=False)
 
            fig_cluster = px.pie(
                cluster_summary,
                values='sellers',
                names='cluster_label',
                hole=0.4,
                color_discrete_sequence=['#1D9E75', '#7F77DD', '#EF9F27', '#D85A30', '#B4B2A9']
            )
            fig_cluster.update_layout(height=300, margin=dict(l=0, r=0, t=20, b=0))
            st.plotly_chart(fig_cluster, use_container_width=True)
 
            cl_display = cluster_summary[['cluster_label', 'sellers', 'avg_revenue', 'revenue_share']].copy()
            cl_display['avg_revenue']   = cl_display['avg_revenue'].apply(lambda x: f"${x:,.0f}")
            cl_display['revenue_share'] = cl_display['revenue_share'].apply(lambda x: f"{x:.1f}%")
            cl_display.columns = ['Cluster', 'Sellers', 'Avg revenue', 'Revenue share']
            st.dataframe(cl_display, use_container_width=True, hide_index=True)
 
    st.markdown("---")
    st.subheader("High value seller — lookalike profile")
 
    l1, l2, l3, l4 = st.columns(4)
    l1.metric("Cluster size", f"{len(lookalike)} sellers")
    l2.metric("Avg revenue",  f"${lookalike['monetary'].mean():,.0f}")
    l3.metric("Avg orders",   f"{lookalike['frequency'].mean():.1f}")
    l4.metric("Avg review",   f"{lookalike['avg_review_score'].mean():.2f} / 5.0")
 
    col1, col2, col3 = st.columns(3)
 
    with col1:
        st.markdown("**Top acquisition channels**")
        ch_data = lookalike['origin'].value_counts().reset_index()
        ch_data.columns = ['Channel', 'Sellers']
        ch_data['%'] = (ch_data['Sellers'] / len(lookalike) * 100).round(1).astype(str) + '%'
        st.dataframe(ch_data, use_container_width=True, hide_index=True)
 
    with col2:
        st.markdown("**Top business segments**")
        seg_data = lookalike['business_segment'].value_counts().head(5).reset_index()
        seg_data.columns = ['Segment', 'Sellers']
        seg_data['%'] = (seg_data['Sellers'] / len(lookalike) * 100).round(1).astype(str) + '%'
        st.dataframe(seg_data, use_container_width=True, hide_index=True)
 
    with col3:
        st.markdown("**Lead type breakdown**")
        lt_data = lookalike['lead_type'].value_counts().reset_index()
        lt_data.columns = ['Lead type', 'Sellers']
        lt_data['%'] = (lt_data['Sellers'] / len(lookalike) * 100).round(1).astype(str) + '%'
        st.dataframe(lt_data, use_container_width=True, hide_index=True)
 
    st.success(
        "Ideal acquisition target: **paid_search or organic_search** channel · "
        "**household_utilities or health_beauty** segment · "
        "**reseller** business type · **online_big or online_medium** lead type"
    )