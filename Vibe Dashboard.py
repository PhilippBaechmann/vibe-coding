import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

# Set page config with light theme
st.set_page_config(
    page_title="Ireland High-Growth Firms Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': 'Ireland High-Growth Firms Dashboard'
    }
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
    .warning-box {
        background-color: #fffbeb;
        border-radius: 8px;
        padding: 1rem;
        border-left: 5px solid #f59e0b;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #f0fdf4;
        border-radius: 8px;
        padding: 1rem;
        border-left: 5px solid #22c55e;
        margin-bottom: 1rem;
    }
    .company-card {
        background-color: white;
        border-radius: 8px;
        padding: 1.5rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        margin-bottom: 1.5rem;
    }
    .company-name {
        font-size: 1.2rem;
        font-weight: 600;
        color: #1e3a8a;
        margin-bottom: 0.5rem;
    }
    .company-detail {
        margin-bottom: 0.3rem;
        color: #334155;
    }
    .data-table {
        border-radius: 8px;
        overflow: hidden;
        margin-bottom: 1.5rem;
    }
    .footer {
        text-align: center;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #e2e8f0;
        color: #64748b;
        font-size: 0.9rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f8fafc;
        border-radius: 4px 4px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
        font-size: 16px;
        font-weight: 500;
    }
    
    /* Improving metric display */
    .metric-row {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
    }
    .metric-item {
        flex: 1;
        min-width: 120px;
        padding: 10px;
        background-color: #f1f5f9;
        border-radius: 8px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.1);
    }
    .metric-item-value {
        font-size: 24px;
        font-weight: 600;
        color: #1e40af;
    }
    .metric-item-label {
        font-size: 12px;
        color: #64748b;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state for filters
if 'selected_industries' not in st.session_state:
    st.session_state['selected_industries'] = []
if 'selected_topics' not in st.session_state:
    st.session_state['selected_topics'] = []
if 'selected_age_cats' not in st.session_state:
    st.session_state['selected_age_cats'] = []
if 'search_term' not in st.session_state:
    st.session_state['search_term'] = ""
if 'year_range' not in st.session_state:
    st.session_state['year_range'] = None
if 'revenue_range' not in st.session_state:
    st.session_state['revenue_range'] = None

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
        numeric_columns = [
            'Estimated_Revenue_mn', 
            'Number of employees 2023', 
            'Number of employees 2022',
            'Number of employees 2021',
            'Growth 2023',
            'Growth 2022',
            'Growth 2021',
            'aagr 2023'
        ]
        for col in numeric_columns:
            if col in high_growth_firms.columns:
                # Handle percentage strings if present
                if high_growth_firms[col].dtype == object:
                    high_growth_firms[col] = high_growth_firms[col].astype(str).str.replace('%', '').str.strip()
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
        
        # Create company size categories
        if 'Number of employees 2023' in high_growth_firms.columns:
            emp_bins = [0, 10, 50, 250, float('inf')]
            emp_labels = ['Micro (1-9)', 'Small (10-49)', 'Medium (50-249)', 'Large (250+)']
            high_growth_firms['Size Category'] = pd.cut(
                high_growth_firms['Number of employees 2023'],
                bins=emp_bins,
                labels=emp_labels
            )
            high_growth_firms['Size Category'] = high_growth_firms['Size Category'].astype(str)
            high_growth_firms.loc[high_growth_firms['Size Category'] == 'nan', 'Size Category'] = 'Unknown'
        
        # Add simple geocoding for demonstration (normally would use real geo data)
        # This is a placeholder to enable the map feature
        if 'City' in high_growth_firms.columns:
            # Create a simple mapping of cities to coordinates for Ireland
            # In a real app, you would use proper geocoding or have real coordinates
            city_coords = {
                'DUBLIN': (53.3498, -6.2603),
                'CORK': (51.8969, -8.4863),
                'GALWAY': (53.2707, -9.0568),
                'LIMERICK': (52.6638, -8.6267),
                'WATERFORD': (52.2593, -7.1101),
                'KILKENNY': (52.6541, -7.2448),
                'DROGHEDA': (53.7144, -6.3501),
                'SWORDS': (53.4599, -6.2205),
                'DUNDALK': (54.0034, -6.4063),
                'BRAY': (53.2009, -6.1110),
                'Unknown': (53.1424, -7.6921)  # Center of Ireland for unknown cities
            }
            
            # Create latitude and longitude columns
            high_growth_firms['latitude'] = high_growth_firms['City'].apply(
                lambda x: city_coords.get(x.upper(), city_coords['Unknown'])[0]
            )
            high_growth_firms['longitude'] = high_growth_firms['City'].apply(
                lambda x: city_coords.get(x.upper(), city_coords['Unknown'])[1]
            )
        
        return high_growth_firms
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

# Create a simple correlation matrix visualization
def show_correlation_matrix(df, columns):
    """Display a correlation matrix for selected numeric columns"""
    # Filter out columns that don't exist or aren't numeric
    valid_columns = []
    for col in columns:
        if col in df.columns:
            try:
                # Verify column is numeric
                pd.to_numeric(df[col], errors='raise')
                valid_columns.append(col)
            except:
                pass
    
    if len(valid_columns) < 2:
        st.info("Not enough numeric data available for correlation analysis.")
        return
    
    # Calculate correlation matrix
    corr_df = df[valid_columns].corr()
    
    # Display the correlation matrix as a heatmap-styled dataframe
    st.dataframe(
        corr_df.style.background_gradient(cmap='coolwarm', axis=None, vmin=-1, vmax=1),
        use_container_width=True
    )
    
    # Explain correlation values
    st.markdown("""
    **Understanding the correlation matrix:**
    * Values close to 1 indicate a strong positive correlation
    * Values close to -1 indicate a strong negative correlation
    * Values close to 0 indicate little to no correlation
    """)

# Main dashboard function
def main():
    # Title
    st.markdown("<h1 class='main-title'>Ireland High-Growth Firms Dashboard</h1>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Insights into the distribution and characteristics of consistent high-growth firms in Ireland</p>", unsafe_allow_html=True)
    
    # Load data
    df = load_data()
    
    if df is not None:
        # Help information in the sidebar
        with st.sidebar.expander("ℹ️ How to Use This Dashboard"):
            st.markdown("""
            **Filtering Data**: 
            Use the sidebar filters to focus on specific segments of firms. You can filter by industry, topic, company age, and more.
            
            **Search**: 
            Use the search box to find specific companies by name.
            
            **Visualizations**: 
            Navigate through the different tabs to explore various aspects of the data - geographical distribution, topic breakdown, company age analysis, etc.
            
            **Company Profiles**: 
            Explore detailed information about individual companies in the Company Profiles section.
            
            **Data Explorer**: 
            Use the data table to examine raw data, with options to sort, filter, and download for further analysis.
            
            **Industry Benchmarking**:
            Compare metrics across different industries to understand competitive landscapes.
            """)
        
        # Sidebar filters
        st.sidebar.markdown("## Dashboard Filters")
        
        # Text search filter
        search_term = st.sidebar.text_input(
            "Search Companies", 
            value=st.session_state['search_term'],
            placeholder="Enter company name..."
        )
        st.session_state['search_term'] = search_term
        
        # Industry filter
        industries = sorted(df['NACE_Industry'].unique())
        selected_industries = st.sidebar.multiselect(
            "Filter by Industry", 
            industries, 
            default=st.session_state['selected_industries']
        )
        st.session_state['selected_industries'] = selected_industries
        
        # Topic filter
        if 'Topic' in df.columns:
            topics = sorted(df['Topic'].unique())
            selected_topics = st.sidebar.multiselect(
                "Filter by Topic", 
                topics,
                default=st.session_state['selected_topics']
            )
            st.session_state['selected_topics'] = selected_topics
        else:
            selected_topics = []
        
        # Age category filter
        age_categories = sorted(df['Age Category'].unique())
        selected_age_cats = st.sidebar.multiselect(
            "Filter by Company Age", 
            age_categories,
            default=st.session_state['selected_age_cats']
        )
        st.session_state['selected_age_cats'] = selected_age_cats
        
        # Company size filter
        if 'Size Category' in df.columns:
            size_categories = sorted(df['Size Category'].unique())
            selected_size_cats = st.sidebar.multiselect(
                "Filter by Company Size", 
                size_categories
            )
        else:
            selected_size_cats = []
        
        # Year founded range filter
        if 'Founded Year' in df.columns:
            min_year = int(df['Founded Year'].min())
            max_year = int(df['Founded Year'].max())
            year_range = st.sidebar.slider(
                "Founded Year Range", 
                min_value=min_year,
                max_value=max_year,
                value=(min_year, max_year)
            )
        else:
            year_range = None
        
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
        
        # Apply search filter
        if search_term:
            filtered_df = filtered_df[filtered_df['Company Name'].str.contains(search_term, case=False)]
        
        # Apply industry filter
        if selected_industries:
            filtered_df = filtered_df[filtered_df['NACE_Industry'].isin(selected_industries)]
        
        # Apply topic filter
        if selected_topics and 'Topic' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Topic'].isin(selected_topics)]
        
        # Apply age category filter
        if selected_age_cats:
            filtered_df = filtered_df[filtered_df['Age Category'].isin(selected_age_cats)]
        
        # Apply size category filter
        if selected_size_cats and 'Size Category' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Size Category'].isin(selected_size_cats)]
        
        # Apply year founded filter
        if year_range and 'Founded Year' in filtered_df.columns:
            filtered_df = filtered_df[
                (filtered_df['Founded Year'] >= year_range[0]) & 
                (filtered_df['Founded Year'] <= year_range[1])
            ]
        
        # Apply revenue filter
        if revenue_range and 'Estimated_Revenue_mn' in filtered_df.columns:
            filtered_df = filtered_df[
                (filtered_df['Estimated_Revenue_mn'] >= revenue_range[0]) & 
                (filtered_df['Estimated_Revenue_mn'] <= revenue_range[1])
            ]
        
        # Reset filters button
        if st.sidebar.button("Reset All Filters"):
            # Clear session state for filters
            st.session_state['selected_industries'] = []
            st.session_state['selected_topics'] = []
            st.session_state['selected_age_cats'] = []
            st.session_state['search_term'] = ""
            st.session_state['year_range'] = None
            st.session_state['revenue_range'] = None
            st.experimental_rerun()
        
        # Show filter status
        if len(filtered_df) < len(df):
            st.markdown(f"<div class='info-box'>Showing {len(filtered_df)} out of {len(df)} high-growth firms based on current filters.</div>", unsafe_allow_html=True)
        
        # Main tabs for different sections
        main_tabs = st.tabs([
            "📊 Overview",
            "🗺️ Geographic Analysis",
            "🏷️ Topic Analysis",
            "⏳ Age Analysis",
            "🔍 Company Profiles",
            "📈 Industry Benchmarking",
            "📋 Data Explorer"
        ])
        
        # OVERVIEW TAB
        with main_tabs[0]:
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
            
            # High-level breakdown charts
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Top Industries")
                industry_counts = filtered_df['NACE_Industry'].value_counts().reset_index()
                industry_counts.columns = ['Industry', 'Count']
                industry_counts = industry_counts.sort_values('Count', ascending=False).head(5)
                st.bar_chart(industry_counts.set_index('Industry'), use_container_width=True)
            
            with col2:
                st.subheader("Company Size Distribution")
                if 'Size Category' in filtered_df.columns:
                    size_counts = filtered_df['Size Category'].value_counts().reset_index()
                    size_counts.columns = ['Size', 'Count']
                    st.bar_chart(size_counts.set_index('Size'), use_container_width=True)
                else:
                    st.info("Company size data not available.")
            
            # Company age and growth rate summary
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Age Distribution")
                age_counts = filtered_df['Age Category'].value_counts().reset_index()
                age_counts.columns = ['Age Category', 'Count']
                st.bar_chart(age_counts.set_index('Age Category'), use_container_width=True)
            
            with col2:
                st.subheader("Growth Rate Analysis")
                if 'Growth 2023' in filtered_df.columns:
                    # Group by industry and calculate mean growth rate
                    growth_by_industry = filtered_df.groupby('NACE_Industry')['Growth 2023'].mean().sort_values(ascending=False).head(5)
                    st.bar_chart(growth_by_industry, use_container_width=True)
                else:
                    st.info("Growth rate data not available.")
            
            # Trend analysis
            st.markdown("<h2 class='section-header'>Trend Analysis</h2>", unsafe_allow_html=True)
            st.markdown("Analyze how the landscape of high-growth firms has evolved over time:")
            
            # Count firms by founding year
            if 'Founded Year' in filtered_df.columns:
                # Create a clean series of founded year counts
                founded_years = filtered_df['Founded Year'].dropna().astype(int)
                year_counts = founded_years.value_counts().sort_index()
                
                # Only include years with reasonable data
                valid_years = year_counts[year_counts.index >= 1980]
                if not valid_years.empty:
                    st.subheader("Companies Founded by Year")
                    st.bar_chart(valid_years, use_container_width=True)
                    
                    # Provide a download button for this data
                    year_data = pd.DataFrame({
                        'Year': valid_years.index,
                        'Number of Companies': valid_years.values
                    })
                    csv = year_data.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Download Year Trend Data",
                        data=csv,
                        file_name="founding_year_trend.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("Insufficient data to show founding year trends.")
            else:
                st.info("Founded year data not available.")
        
        # GEOGRAPHIC ANALYSIS TAB
        with main_tabs[1]:
            st.markdown("<h2 class='section-header'>Geographical Distribution</h2>", unsafe_allow_html=True)
            
            geo_tab1, geo_tab2, geo_tab3 = st.tabs(["City Distribution", "Regional Distribution", "Map View"])
            
            with geo_tab1:
                # City distribution
                city_counts = filtered_df['City'].value_counts().reset_index()
                city_counts.columns = ['City', 'Count']
                city_counts = city_counts.sort_values('Count', ascending=False).head(10)
                
                st.subheader("Top Cities with High-Growth Firms")
                st.bar_chart(city_counts.set_index('City'), use_container_width=True, height=400)
                
                # Download option for city data
                city_csv = city_counts.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download City Distribution Data",
                    data=city_csv,
                    file_name="city_distribution.csv",
                    mime="text/csv"
                )
            
            with geo_tab2:
                # Regional distribution if available
                if 'Region in country' in filtered_df.columns:
                    region_counts = filtered_df['Region in country'].value_counts().reset_index()
                    region_counts.columns = ['Region', 'Count']
                    region_counts = region_counts.sort_values('Count', ascending=False)
                    
                    st.subheader("Regional Distribution")
                    st.bar_chart(region_counts.set_index('Region'), use_container_width=True, height=400)
                    
                    # Download option for region data
                    region_csv = region_counts.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Download Regional Distribution Data",
                        data=region_csv,
                        file_name="regional_distribution.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("Regional data not available in the dataset.")
            
            with geo_tab3:
                # Map view of company locations
                if 'latitude' in filtered_df.columns and 'longitude' in filtered_df.columns:
                    st.subheader("Geographic Distribution Map")
                    
                    # Prepare data for map - ensure we have valid coordinates
                    map_data = filtered_df[['latitude', 'longitude']].copy()
                    # Remove any rows with invalid coordinates
                    map_data = map_data.dropna()
                    
                    if not map_data.empty:
                        st.map(map_data, use_container_width=True)
                    else:
                        st.info("No valid geographic coordinates available for mapping.")
                    
                    st.caption("Note: This map shows approximate locations based on city centers.")
                else:
                    st.info("Geographic coordinate data not available for mapping.")
            
            # Additional analysis - Growth by region
            if 'City' in filtered_df.columns and 'Growth 2023' in filtered_df.columns:
                st.markdown("<h3 class='section-header'>Growth Analysis by Location</h3>", unsafe_allow_html=True)
                
                # Calculate average growth by city
                growth_by_city = filtered_df.groupby('City')['Growth 2023'].mean().sort_values(ascending=False)
                # Filter for cities with at least 2 companies to ensure meaningful averages
                city_counts = filtered_df['City'].value_counts()
                valid_cities = city_counts[city_counts >= 2].index
                growth_by_city = growth_by_city[growth_by_city.index.isin(valid_cities)].head(10)
                
                if not growth_by_city.empty:
                    st.subheader("Top Cities by Average Growth Rate")
                    st.bar_chart(growth_by_city, use_container_width=True)
                else:
                    st.info("Insufficient data to analyze growth rates by city.")
        
        # TOPIC ANALYSIS TAB
        with main_tabs[2]:
            st.markdown("<h2 class='section-header'>Topic Distribution</h2>", unsafe_allow_html=True)
            
            topic_tab1, topic_tab2 = st.tabs(["Topic Overview", "Industry Breakdown"])
            
            with topic_tab1:
                # Topic distribution if available
                if 'Topic' in filtered_df.columns:
                    topic_counts = filtered_df['Topic'].value_counts().reset_index()
                    topic_counts.columns = ['Topic', 'Count']
                    topic_counts = topic_counts.sort_values('Count', ascending=False).head(10)
                    
                    st.subheader("Top Business Topics")
                    st.bar_chart(topic_counts.set_index('Topic'), use_container_width=True, height=400)
                    
                    # Download option for topic data
                    topic_csv = topic_counts.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Download Topic Distribution Data",
                        data=topic_csv,
                        file_name="topic_distribution.csv",
                        mime="text/csv"
                    )
                    
                    # Topic growth analysis if growth data available
                    if 'Growth 2023' in filtered_df.columns:
                        st.subheader("Average Growth Rate by Topic")
                        # Calculate mean growth by topic
                        topic_growth = filtered_df.groupby('Topic')['Growth 2023'].mean().sort_values(ascending=False)
                        # Filter for topics with at least 2 companies
                        topic_counts = filtered_df['Topic'].value_counts()
                        valid_topics = topic_counts[topic_counts >= 2].index
                        topic_growth = topic_growth[topic_growth.index.isin(valid_topics)].head(10)
                        
                        if not topic_growth.empty:
                            st.bar_chart(topic_growth, use_container_width=True)
                        else:
                            st.info("Insufficient data to analyze growth rates by topic.")
                else:
                    st.info("Topic data not available in the dataset.")
            
            with topic_tab2:
                # Industry distribution
                industry_counts = filtered_df['NACE_Industry'].value_counts().reset_index()
                industry_counts.columns = ['Industry', 'Count']
                industry_counts = industry_counts.sort_values('Count', ascending=False).head(10)
                
                st.subheader("Top Industries")
                st.bar_chart(industry_counts.set_index('Industry'), use_container_width=True, height=400)
                
                # Download option for industry data
                industry_csv = industry_counts.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Industry Distribution Data",
                    data=industry_csv,
                    file_name="industry_distribution.csv",
                    mime="text/csv"
                )
                
                # Business model distribution if available
                if 'Business_Model' in filtered_df.columns:
                    st.subheader("Business Model Distribution")
                    model_counts = filtered_df['Business_Model'].value_counts().reset_index()
                    model_counts.columns = ['Business Model', 'Count']
                    model_counts = model_counts.sort_values('Count', ascending=False).head(10)
                    
                    st.bar_chart(model_counts.set_index('Business Model'), use_container_width=True)
            
            # Topic correlation analysis
            if 'Topic' in filtered_df.columns and len(filtered_df) > 5:
                st.markdown("<h3 class='section-header'>Cross-Topic Analysis</h3>", unsafe_allow_html=True)
                st.markdown("Explore relationships between topics and other attributes:")
                
                # Topic vs Age analysis
                if 'Age Category' in filtered_df.columns:
                    # Create a crosstab of Topics vs Age Categories
                    topic_age_pivot = pd.crosstab(
                        filtered_df['Topic'], 
                        filtered_df['Age Category'], 
                        normalize='index'
                    ) * 100  # Convert to percentages
                    
                    # Select top topics for better visualization
                    top_topics = filtered_df['Topic'].value_counts().head(8).index
                    topic_age_pivot = topic_age_pivot.loc[topic_age_pivot.index.isin(top_topics)]
                    
                    if not topic_age_pivot.empty:
                        st.subheader("Topic Distribution by Company Age (%)")
                        st.dataframe(
                            topic_age_pivot.style.background_gradient(cmap='Blues'),
                            use_container_width=True
                        )
                        st.caption("Values show the percentage of companies within each topic that fall into each age category.")
        
        # AGE ANALYSIS TAB
        with main_tabs[3]:
            st.markdown("<h2 class='section-header'>Company Age Analysis</h2>", unsafe_allow_html=True)
            
            # Age distribution visualization
            age_tab1, age_tab2 = st.tabs(["Age Distribution", "Age vs Performance"])
            
            with age_tab1:
                # Age category distribution
                st.subheader("Distribution by Age Category")
                age_counts = filtered_df['Age Category'].value_counts().reset_index()
                age_counts.columns = ['Age Category', 'Count']
                
                # Sort age categories in logical order
                age_order = ['0-3 years', '3-5 years', '5-10 years', '10-20 years', '20+ years', 'Unknown']
                age_counts['Age Category'] = pd.Categorical(
                    age_counts['Age Category'], 
                    categories=age_order, 
                    ordered=True
                )
                age_counts = age_counts.sort_values('Age Category')
                
                st.bar_chart(age_counts.set_index('Age Category'), use_container_width=True, height=400)
                
                # Company age histogram if available
                if 'Company Age' in filtered_df.columns:
                    valid_ages = filtered_df['Company Age'].dropna()
                    if len(valid_ages) > 0:
                        st.subheader("Company Age Distribution (Years)")
                        
                        # Create age bins for histogram
                        age_bins = list(range(0, int(valid_ages.max()) + 5, 5))
                        age_hist = pd.cut(valid_ages, bins=age_bins).value_counts().sort_index()
                        age_hist.index = [f"{i.left}-{i.right}" for i in age_hist.index]
                        
                        st.bar_chart(age_hist, use_container_width=True)
            
            with age_tab2:
                # Age vs performance metrics
                if 'Company Age' in filtered_df.columns:
                    st.subheader("Age vs Performance Metrics")
                    
                    performance_metrics = []
                    if 'Growth 2023' in filtered_df.columns:
                        performance_metrics.append('Growth 2023')
                    if 'Number of employees 2023' in filtered_df.columns:
                        performance_metrics.append('Number of employees 2023')
                    if 'Estimated_Revenue_mn' in filtered_df.columns:
                        performance_metrics.append('Estimated_Revenue_mn')
                    
                    if performance_metrics:
                        # Group by age category and calculate mean of performance metrics
                        age_performance = filtered_df.groupby('Age Category')[performance_metrics].mean()
                        
                        # Sort age categories in logical order
                        age_performance = age_performance.reindex(['0-3 years', '3-5 years', '5-10 years', '10-20 years', '20+ years'])
                        
                        for metric in performance_metrics:
                            metric_name = metric.replace('_', ' ')
                            st.subheader(f"Average {metric_name} by Company Age")
                            st.bar_chart(age_performance[metric], use_container_width=True)
                    else:
                        st.info("No performance metrics available for age analysis.")
                else:
                    st.info("Company age data not available for performance analysis.")
            
            # Additional Age Analysis - Founded Year Trends
            if 'Founded Year' in filtered_df.columns:
                st.markdown("<h3 class='section-header'>Historical Founding Trends</h3>", unsafe_allow_html=True)
                
                # Group companies by decade founded
                filtered_df['Decade Founded'] = (filtered_df['Founded Year'] // 10) * 10
                decade_counts = filtered_df['Decade Founded'].value_counts().sort_index()
                
                # Filter out unreasonable years
                valid_decades = decade_counts[decade_counts.index >= 1950]
                
                if not valid_decades.empty:
                    # Create a dataframe with decade labels
                    decade_df = pd.DataFrame({
                        'Decade': [f"{int(decade)}s" for decade in valid_decades.index],
                        'Count': valid_decades.values
                    })
                    
                    st.subheader("Companies Founded by Decade")
                    st.bar_chart(decade_df.set_index('Decade'), use_container_width=True)
                else:
                    st.info("Insufficient data to show founding decade trends.")
        
        # COMPANY PROFILES TAB
        with main_tabs[4]:
            st.markdown("<h2 class='section-header'>Company Profiles</h2>", unsafe_allow_html=True)
            st.markdown("Explore detailed information about individual high-growth firms:")
            
            # Company selector
            if not filtered_df.empty:
                # Sort companies alphabetically for the selector
                companies = sorted(filtered_df['Company Name'].unique())
                selected_company = st.selectbox("Select a company to view details", companies)
                
                # Get the selected company's data
                company_data = filtered_df[filtered_df['Company Name'] == selected_company].iloc[0]
                
                # Display company profile card
                st.markdown("<div class='company-card'>", unsafe_allow_html=True)
                
                # Company name and basic info
                st.markdown(f"<h2>{company_data['Company Name']}</h2>", unsafe_allow_html=True)
                
                # Organize company details in columns
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("#### Company Information")
                    
                    if 'NACE_Industry' in company_data:
                        st.markdown(f"**Industry:** {company_data['NACE_Industry']}")
                    
                    if 'Business_Model' in company_data and str(company_data['Business_Model']) != 'Unknown':
                        st.markdown(f"**Business Model:** {company_data['Business_Model']}")
                    
                    if 'Topic' in company_data and str(company_data['Topic']) != 'Unknown':
                        st.markdown(f"**Topic:** {company_data['Topic']}")
                    
                    if 'City' in company_data and str(company_data['City']) != 'Unknown':
                        st.markdown(f"**Location:** {company_data['City']}")
                        
                        if 'Region in country' in company_data and str(company_data['Region in country']) != 'Unknown':
                            st.markdown(f"**Region:** {company_data['Region in country']}")
                    
                    if 'Founded Year' in company_data and not pd.isna(company_data['Founded Year']):
                        founded_year = int(company_data['Founded Year'])
                        if founded_year > 1900:
                            st.markdown(f"**Founded:** {founded_year}")
                            
                            if 'Company Age' in company_data:
                                st.markdown(f"**Company Age:** {int(company_data['Company Age'])} years")
                
                with col2:
                    st.markdown("#### Performance Metrics")
                    
                    if 'Number of employees 2023' in company_data:
                        emp_count = int(company_data['Number of employees 2023'])
                        st.markdown(f"**Employees (2023):** {emp_count}")
                    
                    if 'Estimated_Revenue_mn' in company_data and company_data['Estimated_Revenue_mn'] > 0:
                        revenue = company_data['Estimated_Revenue_mn']
                        st.markdown(f"**Estimated Revenue:** €{revenue:.2f} million")
                    
                    if 'Growth 2023' in company_data:
                        growth = company_data['Growth 2023']
                        st.markdown(f"**Growth Rate (2023):** {growth:.1f}%")
                    
                    if 'aagr 2023' in company_data:
                        aagr = company_data['aagr 2023']
                        st.markdown(f"**Annual Average Growth Rate:** {aagr:.1f}%")
                    
                    if 'Size Category' in company_data:
                        st.markdown(f"**Size Category:** {company_data['Size Category']}")
                
                # Company description if available
                if 'Description' in company_data and str(company_data['Description']) != 'Unknown' and str(company_data['Description']) != 'nan':
                    st.markdown("#### Description")
                    st.markdown(f"{company_data['Description']}")
                
                st.markdown("</div>", unsafe_allow_html=True)
                
                # Year-over-year performance if available
                employee_cols = [col for col in filtered_df.columns if 'employees' in col.lower()]
                growth_cols = [col for col in filtered_df.columns if 'growth' in col.lower() and col != 'ConsistentHighGrowthFirm 2023']
                
                if employee_cols or growth_cols:
                    st.markdown("#### Performance Trends")
                    
                    perf_col1, perf_col2 = st.columns(2)
                    
                    # Employee trend
                    if len(employee_cols) >= 2:
                        emp_data = {}
                        for col in sorted(employee_cols, reverse=True):
                            year = col.split()[-1]
                            if not pd.isna(company_data[col]):
                                emp_data[year] = int(company_data[col])
                        
                        if emp_data:
                            with perf_col1:
                                st.subheader("Employee Count")
                                emp_df = pd.DataFrame({'Year': emp_data.keys(), 'Employees': emp_data.values()})
                                st.bar_chart(emp_df.set_index('Year'), use_container_width=True)
                    
                    # Growth trend
                    if len(growth_cols) >= 2:
                        growth_data = {}
                        for col in sorted(growth_cols, reverse=True):
                            year = col.split()[-1]
                            if not pd.isna(company_data[col]):
                                growth_data[year] = float(company_data[col])
                        
                        if growth_data:
                            with perf_col2:
                                st.subheader("Growth Rate (%)")
                                growth_df = pd.DataFrame({'Year': growth_data.keys(), 'Growth': growth_data.values()})
                                st.bar_chart(growth_df.set_index('Year'), use_container_width=True)
            else:
                st.warning("No companies match the current filter criteria.")
        
        # INDUSTRY BENCHMARKING TAB
        with main_tabs[5]:
            st.markdown("<h2 class='section-header'>Industry Benchmarking</h2>", unsafe_allow_html=True)
            st.markdown("Compare key metrics across different industries to understand competitive landscapes:")
            
            # Industry selector
            if not filtered_df.empty:
                industries = sorted(filtered_df['NACE_Industry'].unique())
                selected_industry = st.selectbox("Select an industry to benchmark", industries)
                
                # Filter data for selected industry
                industry_df = filtered_df[filtered_df['NACE_Industry'] == selected_industry]
                
                # Display industry overview
                st.markdown(f"### Overview: {selected_industry}")
                st.markdown(f"**Number of High-Growth Firms:** {len(industry_df)}")
                
                # Key metrics comparison
                metrics_to_compare = []
                if 'Number of employees 2023' in industry_df.columns:
                    metrics_to_compare.append('Number of employees 2023')
                if 'Estimated_Revenue_mn' in industry_df.columns:
                    metrics_to_compare.append('Estimated_Revenue_mn')
                if 'Growth 2023' in industry_df.columns:
                    metrics_to_compare.append('Growth 2023')
                if 'Company Age' in industry_df.columns:
                    metrics_to_compare.append('Company Age')
                
                if metrics_to_compare:
                    # Calculate industry averages and overall averages
                    industry_avgs = industry_df[metrics_to_compare].mean()
                    overall_avgs = filtered_df[metrics_to_compare].mean()
                    
                    # Display comparison
                    st.markdown("### Key Metrics Comparison")
                    
                    metric_cols = st.columns(len(metrics_to_compare))
                    
                    for i, metric in enumerate(metrics_to_compare):
                        with metric_cols[i]:
                            industry_val = industry_avgs[metric]
                            overall_val = overall_avgs[metric]
                            
                            # Calculate percent difference
                            if overall_val != 0:
                                diff_pct = ((industry_val - overall_val) / overall_val) * 100
                            else:
                                diff_pct = 0
                            
                            # Format the display name
                            display_name = metric.replace('_', ' ').replace('Number of ', '')
                            
                            # Display the metric with comparison to average
                            st.metric(
                                label=display_name,
                                value=f"{industry_val:.1f}",
                                delta=f"{diff_pct:.1f}% vs Average"
                            )
                    
                    # Industry distribution charts
                    st.markdown("### Distribution Within Industry")
                    
                    # Age distribution only (removed size distribution)
                    if 'Age Category' in industry_df.columns:
                        st.subheader("Company Age Distribution")
                        age_counts = industry_df['Age Category'].value_counts()
                        st.bar_chart(age_counts, use_container_width=True)
                    
                    # Top companies in the industry
                    st.markdown("### Top Companies in This Industry")
                    
                    # Sort by revenue or employees if available
                    if 'Estimated_Revenue_mn' in industry_df.columns:
                        top_companies = industry_df.sort_values('Estimated_Revenue_mn', ascending=False).head(10)
                        sort_by = 'Estimated_Revenue_mn'
                        sort_label = 'Revenue (€M)'
                    elif 'Number of employees 2023' in industry_df.columns:
                        top_companies = industry_df.sort_values('Number of employees 2023', ascending=False).head(10)
                        sort_by = 'Number of employees 2023'
                        sort_label = 'Employees'
                    else:
                        top_companies = industry_df.head(10)
                        sort_by = None
                    
                    # Display top companies
                    if sort_by:
                        # Create dataframe for chart
                        top_df = pd.DataFrame({
                            'Company': top_companies['Company Name'],
                            sort_label: top_companies[sort_by]
                        })
                        st.bar_chart(top_df.set_index('Company'), use_container_width=True)
                    
                    # Display table with more details
                    display_cols = ['Company Name']
                    if 'City' in industry_df.columns:
                        display_cols.append('City')
                    display_cols.extend([col for col in ['Estimated_Revenue_mn', 'Number of employees 2023', 'Growth 2023', 'Company Age'] 
                                        if col in industry_df.columns])
                    
                    st.dataframe(
                        top_companies[display_cols],
                        use_container_width=True
                    )
                else:
                    st.warning("Not enough metric data available for benchmarking.")
            else:
                st.warning("No data available for industry benchmarking.")
        
        # DATA EXPLORER TAB
        with main_tabs[6]:
            st.markdown("<h2 class='section-header'>Data Explorer</h2>", unsafe_allow_html=True)
            st.markdown("#### Explore and Download the Raw Data")
            st.markdown("View, sort and filter the detailed data for all high-growth firms matching your criteria:")
            
            # Sort options
            sort_options = ['Company Name']
            for col in ['Estimated_Revenue_mn', 'Number of employees 2023', 'Growth 2023', 'Founded Year', 'Company Age']:
                if col in filtered_df.columns:
                    sort_options.append(col)
            
            col1, col2, col3 = st.columns([2, 1, 1])
            
            with col1:
                sort_by = st.selectbox("Sort By", sort_options)
            
            with col2:
                sort_order = st.radio("Order", ["Ascending", "Descending"])
            
            with col3:
                # Column selector
                columns_expander = st.expander("Select Columns")
                with columns_expander:
                    # Display columns selector
                    available_cols = filtered_df.columns.tolist()
                    selected_cols = st.multiselect("Select columns to display", available_cols, default=available_cols[:8])
            
            # Sort the dataframe
            if sort_order == "Ascending":
                sorted_df = filtered_df.sort_values(by=sort_by)
            else:
                sorted_df = filtered_df.sort_values(by=sort_by, ascending=False)
            
            # Display the dataframe
            if selected_cols:
                st.dataframe(sorted_df[selected_cols], use_container_width=True)
            else:
                st.dataframe(sorted_df, use_container_width=True)
            
            # Download options
            col1, col2 = st.columns(2)
            
            with col1:
                # Download full data
                csv = filtered_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download All Data (CSV)",
                    data=csv,
                    file_name=f"high_growth_firms_data_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv"
                )
            
            with col2:
                # Download selected columns only
                if selected_cols:
                    selected_csv = filtered_df[selected_cols].to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Download Selected Columns (CSV)",
                        data=selected_csv,
                        file_name=f"high_growth_firms_selected_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv"
                    )
            
            # Correlation analysis
            st.markdown("<h3 class='section-header'>Correlation Analysis</h3>", unsafe_allow_html=True)
            st.markdown("Explore relationships between numerical variables:")
            
            # Select columns for correlation analysis
            corr_cols = [
                'Company Age', 'Estimated_Revenue_mn', 'Number of employees 2023', 
                'Growth 2023', 'aagr 2023'
            ]
            
            # Show correlation matrix
            show_correlation_matrix(filtered_df, corr_cols)
    else:
        st.error("Failed to load data. Please ensure the Excel file is available and properly formatted.")
    
    # Footer
    st.markdown("<div class='footer'>Ireland High-Growth Firms Dashboard | © 2025</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()