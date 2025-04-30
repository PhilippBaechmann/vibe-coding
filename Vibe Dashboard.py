import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime
from PIL import Image
from io import BytesIO

# Set page config
st.set_page_config(
    page_title="Ireland High-Growth Firms Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-title {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1e3a8a;
        text-align: center;
        margin-bottom: 1rem;
    }
    .subtitle {
        font-size: 1.2rem;
        color: #475569;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-container {
        background-color: #f1f5f9;
        border-radius: 8px;
        padding: 1.2rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        margin-bottom: 1rem;
    }
    .metric-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: #0f172a;
        text-align: center;
    }
    .metric-label {
        font-size: 0.9rem;
        color: #64748b;
        text-align: center;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: 600;
        color: #334155;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #e2e8f0;
    }
    .chart-container {
        background-color: white;
        border-radius: 8px;
        padding: 1rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        margin-bottom: 2rem;
    }
    .filters-container {
        background-color: #f8fafc;
        border-radius: 8px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        border: 1px solid #e2e8f0;
    }
    .filter-title {
        font-size: 1.1rem;
        font-weight: 600;
        color: #334155;
        margin-bottom: 1rem;
    }
    .stButton>button {
        background-color: #2563eb;
        color: white;
        font-weight: 500;
    }
    .info-box {
        background-color: #eff6ff;
        border-radius: 8px;
        padding: 1rem;
        border-left: 5px solid #3b82f6;
        margin-bottom: 1rem;
    }
    .footer {
        text-align: center;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #e2e8f0;
        color: #64748b;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# Load and process data
@st.cache_data(show_spinner="Loading high-growth firms data...")
def load_data(file_path='ireland_cleaned_CHGF.xlsx'):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # Filter for consistent high-growth firms
        df['ConsistentHighGrowthFirm 2023'] = pd.to_numeric(df['ConsistentHighGrowthFirm 2023'], errors='coerce')
        high_growth_firms = df[df['ConsistentHighGrowthFirm 2023'] == 1].copy()
        
        if high_growth_firms.empty:
            st.error("No consistent high-growth firms found in the dataset")
            return None
        
        # Process text columns
        text_columns = ['Company Name', 'City', 'NACE_Industry', 'Topic', 'Business_Model', 'Region in country']
        for col in text_columns:
            if col in high_growth_firms.columns:
                high_growth_firms[col] = high_growth_firms[col].fillna('Unknown').astype(str)
        
        # Process numeric columns
        numeric_columns = ['Estimated_Revenue_mn', 'Number of employees 2023']
        for col in numeric_columns:
            if col in high_growth_firms.columns:
                high_growth_firms[col] = pd.to_numeric(high_growth_firms[col], errors='coerce').fillna(0)
        
        # Calculate company age
        current_year = datetime.datetime.now().year
        
        # Handle 'Founded Year' column
        if 'Founded Year' in high_growth_firms.columns:
            high_growth_firms['Founded Year'] = pd.to_numeric(high_growth_firms['Founded Year'], errors='coerce')
            # Filter out unreasonable founded years
            high_growth_firms.loc[high_growth_firms['Founded Year'] < 1900, 'Founded Year'] = np.nan
            high_growth_firms.loc[high_growth_firms['Founded Year'] > current_year, 'Founded Year'] = np.nan
            # Calculate company age
            high_growth_firms['Company Age'] = current_year - high_growth_firms['Founded Year']
        # Alternative: Use 'Date of incorporation' if available
        elif 'Date of incorporation' in high_growth_firms.columns:
            try:
                high_growth_firms['Date of incorporation'] = pd.to_datetime(high_growth_firms['Date of incorporation'], errors='coerce')
                # Calculate company age
                high_growth_firms['Company Age'] = current_year - high_growth_firms['Date of incorporation'].dt.year
            except:
                # If date conversion fails, create a placeholder age column
                high_growth_firms['Company Age'] = np.nan
        else:
            # If no founding date is available, create a placeholder
            high_growth_firms['Company Age'] = np.nan
        
        # Create age categories
        high_growth_firms['Age Category'] = pd.cut(
            high_growth_firms['Company Age'],
            bins=[0, 3, 5, 10, 20, float('inf')],
            labels=['0-3 years', '3-5 years', '5-10 years', '10-20 years', '20+ years'],
            right=False
        )
        
        # Replace NaN age categories
        high_growth_firms['Age Category'] = high_growth_firms['Age Category'].fillna('Unknown')
        
        return high_growth_firms
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

# Create geographical distribution chart
def create_geo_chart(df, region_col, value_col=None, title=None):
    if df is None or len(df) == 0:
        return None
    
    if region_col not in df.columns:
        return None
    
    # Get counts by region
    if value_col:
        region_data = df.groupby(region_col)[value_col].sum().reset_index()
        region_data.columns = [region_col, 'Value']
    else:
        region_data = df[region_col].value_counts().reset_index()
        region_data.columns = [region_col, 'Count']
    
    # Sort by count/value
    value_col_name = 'Value' if value_col else 'Count'
    region_data = region_data.sort_values(by=value_col_name, ascending=False)
    
    # Create chart
    fig = px.bar(
        region_data, 
        x=region_col, 
        y=value_col_name,
        title=title or f'Distribution by {region_col}',
        color=value_col_name,
        color_continuous_scale='Viridis',
        labels={region_col: region_col.replace('_', ' '), value_col_name: value_col_name}
    )
    
    fig.update_layout(
        height=500,
        xaxis_title=region_col.replace('_', ' '),
        yaxis_title=value_col_name,
        xaxis_tickangle=-45,
        title_x=0.5,
        title_font_size=16,
        margin=dict(l=40, r=40, t=60, b=80),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)'
    )
    
    return fig

# Create topic distribution chart
def create_topic_chart(df, topic_col, title=None):
    if df is None or len(df) == 0 or topic_col not in df.columns:
        return None
    
    # Get topic counts
    topic_counts = df[topic_col].value_counts().reset_index()
    topic_counts.columns = [topic_col, 'Count']
    
    # Limit to top 15 topics for readability
    if len(topic_counts) > 15:
        other_count = topic_counts.iloc[15:]['Count'].sum()
        topic_counts = topic_counts.iloc[:15]
        other_row = pd.DataFrame({topic_col: ['Other'], 'Count': [other_count]})
        topic_counts = pd.concat([topic_counts, other_row], ignore_index=True)
    
    # Create pie chart
    fig = px.pie(
        topic_counts, 
        names=topic_col, 
        values='Count',
        title=title or f'Distribution by {topic_col}',
        hole=0.4,
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    
    fig.update_layout(
        height=600,
        title_x=0.5,
        title_font_size=16,
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5)
    )
    
    fig.update_traces(textposition='inside', textinfo='percent+label')
    
    return fig

# Create age distribution chart
def create_age_chart(df, age_col, age_cat_col):
    if df is None or len(df) == 0:
        return None
    
    if age_col not in df.columns or age_cat_col not in df.columns:
        return None
    
    # Create histogram for company age
    fig = make_subplots(
        rows=1, cols=2,
        specs=[[{"type": "histogram"}, {"type": "pie"}]],
        subplot_titles=("Company Age Distribution", "Age Categories")
    )
    
    # Histogram of company age
    fig.add_trace(
        go.Histogram(
            x=df[age_col].dropna(),
            nbinsx=20,
            marker_color='rgba(55, 83, 109, 0.7)',
            marker_line_color='rgba(55, 83, 109, 1)',
            marker_line_width=1
        ),
        row=1, col=1
    )
    
    # Pie chart of age categories
    age_cats = df[age_cat_col].value_counts().reset_index()
    age_cats.columns = [age_cat_col, 'Count']
    
    fig.add_trace(
        go.Pie(
            labels=age_cats[age_cat_col],
            values=age_cats['Count'],
            hole=0.4,
            marker_colors=px.colors.sequential.Viridis
        ),
        row=1, col=2
    )
    
    # Update layout
    fig.update_layout(
        height=500,
        title='Company Age Analysis',
        title_x=0.5,
        title_font_size=16,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        margin=dict(l=20, r=20, t=80, b=80)
    )
    
    # Update x and y axis labels
    fig.update_xaxes(title_text="Age (Years)", row=1, col=1)
    fig.update_yaxes(title_text="Number of Companies", row=1, col=1)
    
    return fig

# Create industry breakdown chart
def create_industry_chart(df, industry_col, title=None):
    if df is None or len(df) == 0 or industry_col not in df.columns:
        return None
    
    # Get industry counts
    industry_counts = df[industry_col].value_counts().reset_index()
    industry_counts.columns = [industry_col, 'Count']
    
    # Limit to top 10 industries for readability
    if len(industry_counts) > 10:
        industry_counts = industry_counts.head(10)
    
    # Create horizontal bar chart
    fig = px.bar(
        industry_counts, 
        x='Count', 
        y=industry_col,
        title=title or f'Top Industries',
        orientation='h',
        color='Count',
        color_continuous_scale='Viridis',
        labels={industry_col: industry_col.replace('_', ' '), 'Count': 'Number of Companies'}
    )
    
    fig.update_layout(
        height=500,
        yaxis_title="",
        xaxis_title="Number of Companies",
        title_x=0.5,
        title_font_size=16,
        margin=dict(l=20, r=40, t=60, b=40),
        yaxis={'categoryorder':'total ascending'}
    )
    
    return fig

# Create correlation analysis
def create_correlation_chart(df, columns, title=None):
    if df is None or len(df) == 0:
        return None
    
    # Ensure all columns exist and contain numeric data
    valid_columns = []
    for col in columns:
        if col in df.columns:
            try:
                # Convert to numeric and drop NaN values
                df[col] = pd.to_numeric(df[col], errors='coerce')
                if not df[col].isna().all():  # If not all values are NaN
                    valid_columns.append(col)
            except:
                pass
    
    if len(valid_columns) < 2:
        return None
    
    # Calculate correlation matrix
    corr_matrix = df[valid_columns].corr()
    
    # Create heatmap
    fig = px.imshow(
        corr_matrix,
        text_auto=True,
        color_continuous_scale='RdBu_r',
        zmin=-1, zmax=1,
        title=title or "Correlation Analysis"
    )
    
    fig.update_layout(
        height=600,
        title_x=0.5,
        title_font_size=16
    )
    
    return fig

# Main dashboard function
def main():
    # Title
    st.markdown("<h1 class='main-title'>Ireland High-Growth Firms Dashboard</h1>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Insights into the distribution and characteristics of consistent high-growth firms in Ireland</p>", unsafe_allow_html=True)
    
    # Load data
    df = load_data()
    
    if df is not None:
        # Sidebar filters
        st.sidebar.markdown("## Dashboard Filters")
        st.sidebar.markdown("Use these filters to explore specific segments of high-growth firms.")
        
        # Industry filter
        industries = sorted(df['NACE_Industry'].unique())
        selected_industries = st.sidebar.multiselect("Filter by Industry", industries)
        
        # Topic filter
        if 'Topic' in df.columns:
            topics = sorted(df['Topic'].unique())
            selected_topics = st.sidebar.multiselect("Filter by Topic", topics)
        else:
            selected_topics = []
        
        # Age category filter
        age_categories = sorted(df['Age Category'].unique())
        selected_age_cats = st.sidebar.multiselect("Filter by Company Age", age_categories)
        
        # Revenue range filter
        if 'Estimated_Revenue_mn' in df.columns:
            min_revenue = float(df['Estimated_Revenue_mn'].min())
            max_revenue = float(df['Estimated_Revenue_mn'].max())
            revenue_range = st.sidebar.slider(
                "Revenue Range (€ millions)", 
                min_value=min_revenue,
                max_value=max_revenue,
                value=(min_revenue, max_revenue)
            )
        else:
            revenue_range = None
        
        # Apply filters
        filtered_df = df.copy()
        
        if selected_industries:
            filtered_df = filtered_df[filtered_df['NACE_Industry'].isin(selected_industries)]
        
        if selected_topics and 'Topic' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Topic'].isin(selected_topics)]
        
        if selected_age_cats:
            filtered_df = filtered_df[filtered_df['Age Category'].isin(selected_age_cats)]
        
        if revenue_range and 'Estimated_Revenue_mn' in filtered_df.columns:
            filtered_df = filtered_df[
                (filtered_df['Estimated_Revenue_mn'] >= revenue_range[0]) & 
                (filtered_df['Estimated_Revenue_mn'] <= revenue_range[1])
            ]
        
        # Reset filters button
        if st.sidebar.button("Reset All Filters"):
            st.experimental_rerun()
        
        # Show filter status
        if len(filtered_df) < len(df):
            st.markdown(f"<div class='info-box'>Showing {len(filtered_df)} out of {len(df)} high-growth firms based on current filters.</div>", unsafe_allow_html=True)
        
        # Display key metrics
        st.markdown("<h2 class='section-header'>Key Metrics</h2>", unsafe_allow_html=True)
        
        metric_cols = st.columns(4)
        
        with metric_cols[0]:
            st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-value'>{len(filtered_df)}</div>", unsafe_allow_html=True)
            st.markdown("<div class='metric-label'>High-Growth Firms</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        
        with metric_cols[1]:
            unique_regions = filtered_df['City'].nunique() if 'City' in filtered_df.columns else 0
            st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-value'>{unique_regions}</div>", unsafe_allow_html=True)
            st.markdown("<div class='metric-label'>Unique Locations</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        
        with metric_cols[2]:
            topics_count = filtered_df['Topic'].nunique() if 'Topic' in filtered_df.columns else 0
            st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-value'>{topics_count}</div>", unsafe_allow_html=True)
            st.markdown("<div class='metric-label'>Unique Topics</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        
        with metric_cols[3]:
            avg_age = filtered_df['Company Age'].mean() if 'Company Age' in filtered_df.columns else 0
            st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-value'>{avg_age:.1f}</div>", unsafe_allow_html=True)
            st.markdown("<div class='metric-label'>Average Age (Years)</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        
        # Geographical Distribution Section
        st.markdown("<h2 class='section-header'>Geographical Distribution</h2>", unsafe_allow_html=True)
        
        geo_tab1, geo_tab2 = st.tabs(["City Distribution", "Regional Distribution"])
        
        with geo_tab1:
            city_chart = create_geo_chart(filtered_df, 'City', title='Distribution of High-Growth Firms by City')
            if city_chart:
                st.plotly_chart(city_chart, use_container_width=True)
            else:
                st.info("Insufficient city data for visualization.")
        
        with geo_tab2:
            if 'Region in country' in filtered_df.columns:
                region_chart = create_geo_chart(filtered_df, 'Region in country', title='Distribution of High-Growth Firms by Region')
                if region_chart:
                    st.plotly_chart(region_chart, use_container_width=True)
                else:
                    st.info("Insufficient region data for visualization.")
            else:
                st.info("Regional data not available in the dataset.")
        
        # Topic Distribution Section
        st.markdown("<h2 class='section-header'>Topic Distribution</h2>", unsafe_allow_html=True)
        
        topic_tab1, topic_tab2 = st.tabs(["Topic Overview", "Industry Breakdown"])
        
        with topic_tab1:
            if 'Topic' in filtered_df.columns:
                topic_chart = create_topic_chart(filtered_df, 'Topic', title='Distribution of High-Growth Firms by Topic')
                if topic_chart:
                    st.plotly_chart(topic_chart, use_container_width=True)
                else:
                    st.info("Insufficient topic data for visualization.")
            else:
                st.info("Topic data not available in the dataset.")
        
        with topic_tab2:
            industry_chart = create_industry_chart(filtered_df, 'NACE_Industry', title='Top Industries among High-Growth Firms')
            if industry_chart:
                st.plotly_chart(industry_chart, use_container_width=True)
            else:
                st.info("Insufficient industry data for visualization.")
        
        # Company Age Section
        st.markdown("<h2 class='section-header'>Company Age Analysis</h2>", unsafe_allow_html=True)
        
        age_chart = create_age_chart(filtered_df, 'Company Age', 'Age Category')
        if age_chart:
            st.plotly_chart(age_chart, use_container_width=True)
        else:
            st.info("Insufficient company age data for visualization.")
        
        # Advanced Analytics Section
        st.markdown("<h2 class='section-header'>Advanced Analytics</h2>", unsafe_allow_html=True)
        
        analytics_tab1, analytics_tab2 = st.tabs(["Correlation Analysis", "Data Explorer"])
        
        with analytics_tab1:
            corr_cols = ['Company Age', 'Estimated_Revenue_mn', 'Number of employees 2023', 'Growth 2023']
            corr_chart = create_correlation_chart(filtered_df, corr_cols, title='Correlation Between Key Metrics')
            if corr_chart:
                st.plotly_chart(corr_chart, use_container_width=True)
                st.markdown("""
                **Understanding the correlation matrix:**
                * Values close to 1 indicate a strong positive correlation
                * Values close to -1 indicate a strong negative correlation
                * Values close to 0 indicate little to no correlation
                """)
            else:
                st.info("Insufficient numeric data for correlation analysis.")
        
        with analytics_tab2:
            st.markdown("#### Explore the Raw Data")
            st.markdown("View and interact with the filtered dataset:")
            
            # Display columns selector
            available_cols = filtered_df.columns.tolist()
            selected_cols = st.multiselect("Select columns to display", available_cols, default=available_cols[:8])
            
            if selected_cols:
                st.dataframe(filtered_df[selected_cols], use_container_width=True)
            else:
                st.dataframe(filtered_df, use_container_width=True)
            
            # Add download button
            csv_buffer = BytesIO()
            filtered_df.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)
            st.download_button(
                label="Download Filtered Data as CSV",
                data=csv_buffer,
                file_name=f"high_growth_firms_data_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )
    else:
        st.error("Failed to load data. Please ensure the Excel file is available and properly formatted.")
    
    # Footer
    st.markdown("<div class='footer'>© 2025 Ireland High-Growth Firms Dashboard</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()