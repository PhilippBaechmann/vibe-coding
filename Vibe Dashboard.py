import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

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
        current_year = datetime.now().year
        
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
        
        # Create age categories - using string type instead of categorical to avoid issues
        bins = [0, 3, 5, 10, 20, float('inf')]
        labels = ['0-3 years', '3-5 years', '5-10 years', '10-20 years', '20+ years']
        
        # Handle categorical data safely
        high_growth_firms['Age Category'] = pd.cut(high_growth_firms['Company Age'], bins=bins, labels=labels)
        # Convert to string type and handle NaN values
        high_growth_firms['Age Category'] = high_growth_firms['Age Category'].astype(str)
        high_growth_firms.loc[high_growth_firms['Age Category'] == 'nan', 'Age Category'] = 'Unknown'
        
        return high_growth_firms
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

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
        
        # Apply filters
        filtered_df = df.copy()
        
        if selected_industries:
            filtered_df = filtered_df[filtered_df['NACE_Industry'].isin(selected_industries)]
        
        if selected_topics and 'Topic' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Topic'].isin(selected_topics)]
        
        if selected_age_cats:
            filtered_df = filtered_df[filtered_df['Age Category'].isin(selected_age_cats)]
        
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
        
        # City distribution
        city_counts = filtered_df['City'].value_counts().reset_index()
        city_counts.columns = ['City', 'Count']
        city_counts = city_counts.sort_values('Count', ascending=False).head(10)
        
        st.subheader("Top Cities with High-Growth Firms")
        st.bar_chart(city_counts.set_index('City'))
        
        # Regional distribution if available
        if 'Region in country' in filtered_df.columns:
            region_counts = filtered_df['Region in country'].value_counts().reset_index()
            region_counts.columns = ['Region', 'Count']
            region_counts = region_counts.sort_values('Count', ascending=False)
            
            st.subheader("Regional Distribution")
            st.bar_chart(region_counts.set_index('Region'))
        
        # Topic Distribution Section
        st.markdown("<h2 class='section-header'>Topic Distribution</h2>", unsafe_allow_html=True)
        
        # Topic distribution if available
        if 'Topic' in filtered_df.columns:
            topic_counts = filtered_df['Topic'].value_counts().reset_index()
            topic_counts.columns = ['Topic', 'Count']
            topic_counts = topic_counts.sort_values('Count', ascending=False).head(10)
            
            st.subheader("Top Business Topics")
            st.bar_chart(topic_counts.set_index('Topic'))
        
        # Industry distribution
        industry_counts = filtered_df['NACE_Industry'].value_counts().reset_index()
        industry_counts.columns = ['Industry', 'Count']
        industry_counts = industry_counts.sort_values('Count', ascending=False).head(10)
        
        st.subheader("Top Industries")
        st.bar_chart(industry_counts.set_index('Industry'))
        
        # Company Age Section
        st.markdown("<h2 class='section-header'>Company Age Analysis</h2>", unsafe_allow_html=True)
        
        # Age category distribution
        age_counts = filtered_df['Age Category'].value_counts().reset_index()
        age_counts.columns = ['Age Category', 'Count']
        
        st.subheader("Distribution by Age Category")
        st.bar_chart(age_counts.set_index('Age Category'))
        
        # Data explorer tab
        st.markdown("<h2 class='section-header'>Data Explorer</h2>", unsafe_allow_html=True)
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
        csv = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download Filtered Data as CSV",
            data=csv,
            file_name=f"high_growth_firms_data_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )
    else:
        st.error("Failed to load data. Please ensure the Excel file is available and properly formatted.")
    
    # Footer
    st.markdown("<div class='footer'>© 2025 Ireland High-Growth Firms Dashboard</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()