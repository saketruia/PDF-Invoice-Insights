import streamlit as st
import json
import pandas as pd
import datetime
import time
import google.generativeai as genai
import os
from dotenv import load_dotenv
import tempfile
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re

# ------------------------------------------------------------------------------
# 1. Loading API key from .env
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")
genai.configure(api_key=API_KEY)

def extract_field(pdf_file):
    """Save PDF temporarily and send it to Gemini to extract pincode, delivery/shipment charges, and main date in pure JSON format."""
    prompt = """Analyze this PDF and extract pincode, delivery/shipment charges, and a main date (such as delivery, billing, or invoice date) in pure JSON format with keys pincode, delivery_charge, main_date. If not present, put "NA".

    Format the main date in DD-MM-YYYY format.

    Provide ONLY the raw JSON and nothing else."""    
    
    model = genai.GenerativeModel('gemini-1.5-flash')  # Free and faster
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(pdf_file.read())  # write the uploaded file's content
        tmp_file_path = tmp_file.name

    for attempt in range(5):
        try:
            response = model.generate_content([prompt, genai.upload_file(tmp_file_path)])

            raw = response.text
            start = raw.find('{')
            end = raw.rfind('}')
            if start == -1 or end == -1:
                raise ValueError("Invalid format.")
            raw_json = raw[start:end+1]
            return json.loads(raw_json)

        except Exception as e:
            st.error(f"Error retrieving from Gemini (attempt {attempt + 1}): {e}")
            time.sleep(30)  # Wait and retry
    
    return {"pincode": "NA", "delivery_charge": "NA", "main_date": "NA"}

def append_to_file(filename, new_df):
    """Append new data to an existing CSV or create it if it doesn't exist."""
    if os.path.exists(filename):
        existing_df = pd.read_excel(filename)
        df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        df = new_df
    
    df.to_excel(filename, index=False)
    return df

def extract_numeric_value(delivery_charge):
    """Extract numeric value from delivery charge string."""
    if pd.isna(delivery_charge) or delivery_charge == "NA":
        return 0
    
    # Convert to string if not already
    delivery_str = str(delivery_charge)
    
    # Remove currency symbols and extract numbers
    numbers = re.findall(r'\d+\.?\d*', delivery_str)
    if numbers:
        return float(numbers[0])
    return 0

def load_existing_data(filename):
    """Load existing data from Excel file."""
    if os.path.exists(filename):
        return pd.read_excel(filename)
    else:
        return pd.DataFrame(columns=['File Name', 'Delivery/Shipment Charges', 'Main Date', 'Pincode'])

def create_dashboard(df):
    """Create dashboard with analytics and visualizations."""
    st.header("📊 Dashboard Analytics")
    
    if df.empty:
        st.warning("No data available. Please upload some PDFs first!")
        return
    
    # Date range filter
    st.subheader("📅 Filter by Date Range")
    
    # Convert dates and get valid date range
    df_with_valid_dates = df[df['Main Date'] != 'NA'].copy()
    
    if not df_with_valid_dates.empty:
        try:
            df_with_valid_dates['Date'] = pd.to_datetime(df_with_valid_dates['Main Date'], format='%d-%m-%Y', errors='coerce')
            df_with_valid_dates = df_with_valid_dates.dropna(subset=['Date'])
            
            if not df_with_valid_dates.empty:
                min_date = df_with_valid_dates['Date'].min().date()
                max_date = df_with_valid_dates['Date'].max().date()
                
                col1, col2 = st.columns(2)
                with col1:
                    start_date = st.date_input("Start Date", value=min_date, min_value=min_date, max_value=max_date)
                with col2:
                    end_date = st.date_input("End Date", value=max_date, min_value=min_date, max_value=max_date)
                
                # Filter dataframe based on date range
                if start_date and end_date:
                    mask = (df_with_valid_dates['Date'].dt.date >= start_date) & (df_with_valid_dates['Date'].dt.date <= end_date)
                    filtered_df = df_with_valid_dates[mask]
                    
                    # Update main dataframe with filtered data
                    if not filtered_df.empty:
                        df = filtered_df
                        st.success(f"📊 Showing data from {start_date} to {end_date} ({len(df)} records)")
                    else:
                        st.warning("No data found in the selected date range.")
                        return
            else:
                st.info("📅 No valid dates found. Showing all data.")
        except:
            st.info("📅 No valid dates found. Showing all data.")
    else:
        st.info("📅 No dates available in data. Showing all data.")
    
    st.divider()
    
    # Process delivery charges
    df['Numeric_Delivery_Charge'] = df['Delivery/Shipment Charges'].apply(extract_numeric_value)
    
    # Calculate metrics - properly detect NA values
    total_delivery_spent = df['Numeric_Delivery_Charge'].sum()
    na_count = ((df['Delivery/Shipment Charges'] == 'NA') | 
                (df['Delivery/Shipment Charges'].isna()) | 
                (df['Delivery/Shipment Charges'] == '') |
                (df['Delivery/Shipment Charges'].astype(str).str.strip() == 'NA')).sum()
    total_orders = len(df)
    
    # Create metrics row
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="💰 Total Delivery Spent",
            value=f"₹{total_delivery_spent:,.2f}"
        )
    
    with col2:
        st.metric(
            label="📋 Total Orders",
            value=total_orders
        )
    
    with col3:
        st.metric(
            label="❌ NA Delivery Charges",
            value=na_count
        )
    
    with col4:
        na_percentage = (na_count / total_orders * 100) if total_orders > 0 else 0
        st.metric(
            label="📊 NA Percentage",
            value=f"{na_percentage:.1f}%"
        )
    
    st.divider()
    
    # Pincode Analysis
    st.subheader("📍 Pincode Analysis")
    
    # Filter out NA pincodes for analysis - improved NA detection
    valid_pincodes_mask = ((df['Pincode'] != 'NA') & 
                          (~df['Pincode'].isna()) & 
                          (df['Pincode'] != '') &
                          (df['Pincode'].astype(str).str.strip() != 'NA'))
    valid_pincodes = df[valid_pincodes_mask]['Pincode']
    
    if not valid_pincodes.empty:
        # Convert pincodes to string and clean them
        valid_pincodes = valid_pincodes.astype(str).str.strip()
        pincode_counts = valid_pincodes.value_counts()
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            # Create bar chart for pincode distribution with proper formatting
            fig_bar = px.bar(
                x=pincode_counts.index[:10],  # Top 10 pincodes
                y=pincode_counts.values[:10],
                title="Top 10 Pincodes by Order Count",
                labels={'x': 'Pincode', 'y': 'Number of Orders'},
                color=pincode_counts.values[:10],
                color_continuous_scale='viridis',
                text=pincode_counts.values[:10]  # Show values on bars
            )
            
            # Format the x-axis to show pincodes as strings without scientific notation
            fig_bar.update_layout(
                xaxis_title="Pincode",
                yaxis_title="Number of Orders",
                showlegend=False,
                xaxis={'type': 'category', 'tickmode': 'array', 'tickvals': pincode_counts.index[:10], 'ticktext': pincode_counts.index[:10]}
            )
            
            # Add text on bars
            fig_bar.update_traces(texttemplate='%{text}', textposition='outside')
            
            st.plotly_chart(fig_bar, use_container_width=True)
        
        with col2:
            st.write("**Top 5 Pincodes:**")
            for i, (pincode, count) in enumerate(pincode_counts.head().items(), 1):
                st.write(f"{i}. **{pincode}**: {count} orders")
            
            # Most frequent pincode
            most_frequent_pincode = pincode_counts.index[0]
            most_frequent_count = pincode_counts.iloc[0]
            st.success(f"🏆 Most frequent pincode: **{most_frequent_pincode}** ({most_frequent_count} orders)")
    
    else:
        st.warning("No valid pincodes found in the data.")
    
    st.divider()
    
    # Delivery Charges Analysis
    st.subheader("💳 Delivery Charges Analysis")
    
    # Filter out zero and NA delivery charges
    valid_charges = df[df['Numeric_Delivery_Charge'] > 0]['Numeric_Delivery_Charge']
    
    if not valid_charges.empty:
        col1, col2 = st.columns(2)
        
        with col1:
            # Histogram of delivery charges
            fig_hist = px.histogram(
                valid_charges,
                nbins=20,
                title="Distribution of Delivery Charges",
                labels={'value': 'Delivery Charge (₹)', 'count': 'Frequency'},
                color_discrete_sequence=['#1f77b4']
            )
            st.plotly_chart(fig_hist, use_container_width=True)
        
        with col2:
            # Box plot for delivery charges
            fig_box = px.box(
                y=valid_charges,
                title="Delivery Charges Box Plot",
                labels={'y': 'Delivery Charge (₹)'}
            )
            st.plotly_chart(fig_box, use_container_width=True)
        
        # Statistics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Average Charge", f"₹{valid_charges.mean():.2f}")
        with col2:
            st.metric("Median Charge", f"₹{valid_charges.median():.2f}")
        with col3:
            st.metric("Min Charge", f"₹{valid_charges.min():.2f}")
        with col4:
            st.metric("Max Charge", f"₹{valid_charges.max():.2f}")
    
    st.divider()
    
    # Time Series Analysis (if dates are available)
    st.subheader("📅 Time Series Analysis")
    
    # Convert dates and filter valid ones
    df_with_dates = df[df['Main Date'] != 'NA'].copy()
    
    if not df_with_dates.empty:
        try:
            df_with_dates['Date'] = pd.to_datetime(df_with_dates['Main Date'], format='%d-%m-%Y', errors='coerce')
            df_with_dates = df_with_dates.dropna(subset=['Date'])
            
            if not df_with_dates.empty:
                # Group by date and count orders
                daily_orders = df_with_dates.groupby('Date').size().reset_index(name='Orders')
                
                fig_line = px.line(
                    daily_orders,
                    x='Date',
                    y='Orders',
                    title="Orders Over Time",
                    labels={'Date': 'Date', 'Orders': 'Number of Orders'}
                )
                st.plotly_chart(fig_line, use_container_width=True)
                
                # Monthly summary
                df_with_dates['Month'] = df_with_dates['Date'].dt.to_period('M')
                monthly_summary = df_with_dates.groupby('Month').agg({
                    'File Name': 'count',
                    'Numeric_Delivery_Charge': 'sum'
                }).rename(columns={'File Name': 'Orders', 'Numeric_Delivery_Charge': 'Total_Delivery'})
                
                st.write("**Monthly Summary:**")
                st.dataframe(monthly_summary, use_container_width=True)
        
        except Exception as e:
            st.warning(f"Could not process dates for time series analysis: {e}")
    
    st.divider()
    
    # Data Table
    st.subheader("📋 Raw Data")
    st.dataframe(df, use_container_width=True)

# ------------------------------------------------------------------------------

# Streamlit UI
st.set_page_config(page_title="Invoice Analyzer Dashboard", page_icon="📊", layout="wide")

st.title("📊 Invoice Information Extractor & Dashboard")

# Create tabs for different functionalities
tab1, tab2 = st.tabs(["📤 Upload & Extract", "📊 Dashboard"])

with tab1:
    st.write("**Upload PDF files to extract pincode, delivery/shipment charges, and main date.**")
    uploaded_files = st.file_uploader("Choose PDF files", accept_multiple_files=True, type=['pdf'])

    if uploaded_files:
        data = []

        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Processing {uploaded_file.name} ({i+1}/{len(uploaded_files)})")
            fields = extract_field(uploaded_file)

            pincode = fields.get("pincode", "NA")
            delivery_charge = fields.get("delivery_charge", "NA")
            main_date = fields.get("main_date", "NA")

            data.append([uploaded_file.name, delivery_charge, main_date, pincode])
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        status_text.text("Processing complete!")
        
        df = pd.DataFrame(data, columns=['File Name', 'Delivery/Shipment Charges', 'Main Date', 'Pincode'])

        st.success("✅ Data extracted successfully!")
        st.write("**Extracted Data:**")
        st.dataframe(df, use_container_width=True)

        output_file = "extracted_data.xlsx"
        final_df = append_to_file(output_file, df)

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Download as Excel",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        st.info("💡 Data has been added to the existing Excel file. Check the Dashboard tab to see updated analytics!")

with tab2:
    # Load existing data for dashboard
    output_file = "extracted_data.xlsx"
    existing_df = load_existing_data(output_file)
    
    # Add refresh button
    if st.button("🔄 Refresh Dashboard"):
        st.rerun()
    
    create_dashboard(existing_df)