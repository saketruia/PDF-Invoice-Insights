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
import zipfile
import io
from pdf_report_generator import generate_pdf_report, create_download_link

# ------------------------------------------------------------------------------
# 1. Loading API key from .env
API_KEY = st.secrets["GEMINI_API_KEY"]["API_KEY"]

genai.configure(api_key=API_KEY)

def extract_field(pdf_file):
    """Save PDF temporarily and send it to Gemini to extract invoice number, sender pincode, receiver pincode, delivery/shipment charges, and main date in pure JSON format."""
    prompt = """Analyze this PDF and extract invoice number, sender pincode, receiver pincode, delivery/shipment charges, and a main date in pure JSON format with keys invoice_number, sender_pincode, receiver_pincode, delivery_charge, main_date.

    IMPORTANT INSTRUCTIONS:
    
    1. INVOICE NUMBER EXTRACTION:
    - Look for invoice number throughout the document
    - Common terms: "Invoice No", "Invoice Number", "Bill No", "Bill Number", "Receipt No", "Document No"
    - For Porter invoices: Check top right corner specifically
    - For other invoices: Search anywhere in the document
    - Extract the complete alphanumeric invoice number
    - If not found, put "NA"
    
    2. PORTER INVOICE DETECTION:
    - If you see "PORTER" written in the top left corner of the document, this is a Porter invoice
    - For Porter invoices:
      * Use "Total Amount" or "Grand Total" as the delivery_charge (NOT delivery/shipping charges)
      * Look for pickup and drop locations (usually in bottom right section)
      * Extract pincodes from pickup location as sender_pincode and drop location as receiver_pincode
      * If pickup/drop locations or their pincodes are not found, put "NA"
    
    3. REGULAR INVOICES (Non-Porter):
    - Look for sender and receiver addresses throughout the document
    - Common terms: "Bill To", "Ship To", "From", "To", "Sender", "Recipient", "Billing Address", "Shipping Address"
    - Extract pincodes from sender address as sender_pincode and receiver address as receiver_pincode
    - For delivery_charge: Look for delivery/shipping charges including GST/tax (same as before)
    
    4. DELIVERY CHARGE CALCULATION:
    - For Porter: Use Total Amount/Grand Total
    - For Regular: If there's a total delivery amount (including GST/tax), use that total amount
    - If only base delivery charge is available without tax, use that amount
    - Look for terms like: delivery charge, shipping charge, freight charge, courier charge, GST, tax, CGST, SGST, IGST
    - Calculate total = base delivery charge + any applicable taxes/GST
    - If not present, put "NA"

    5. DATE FORMAT:
    - Format the main date in DD-MM-YYYY format
    - Look for delivery date, billing date, or invoice date

    6. PINCODES:
    - Extract 6-digit pincodes from addresses
    - If sender or receiver pincode not found, put "NA" for that field

    Provide ONLY the raw JSON and nothing else with keys: invoice_number, sender_pincode, receiver_pincode, delivery_charge, main_date"""    
    
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
    
    return {"invoice_number": "NA", "sender_pincode": "NA", "receiver_pincode": "NA", "delivery_charge": "NA", "main_date": "NA"}

def extract_pdfs_from_zip(zip_file):
    """Extract all PDF files from a ZIP file and return them as a list of file-like objects."""
    pdf_files = []
    
    try:
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            for file_info in zip_ref.infolist():
                if file_info.filename.lower().endswith('.pdf') and not file_info.is_dir():
                    # Extract the PDF file content
                    pdf_content = zip_ref.read(file_info.filename)
                    
                    # Create a file-like object
                    pdf_file_obj = io.BytesIO(pdf_content)
                    pdf_file_obj.name = file_info.filename
                    
                    pdf_files.append(pdf_file_obj)
        
        return pdf_files
    
    except Exception as e:
        st.error(f"Error extracting ZIP file: {e}")
        return []

def process_uploaded_files(uploaded_files):
    """Process uploaded files (PDFs and ZIP files containing PDFs)."""
    all_pdf_files = []
    
    for uploaded_file in uploaded_files:
        if uploaded_file.name.lower().endswith('.pdf'):
            # Direct PDF file
            all_pdf_files.append(uploaded_file)
        elif uploaded_file.name.lower().endswith('.zip'):
            # ZIP file containing PDFs
            st.info(f"📦 Extracting PDFs from {uploaded_file.name}...")
            pdf_files_from_zip = extract_pdfs_from_zip(uploaded_file)
            if pdf_files_from_zip:
                all_pdf_files.extend(pdf_files_from_zip)
                st.success(f"✅ Extracted {len(pdf_files_from_zip)} PDF files from {uploaded_file.name}")
            else:
                st.warning(f"⚠️ No PDF files found in {uploaded_file.name}")
        else:
            st.warning(f"⚠️ Unsupported file type: {uploaded_file.name}")
    
    return all_pdf_files

def check_duplicate_invoice(filename, invoice_number):
    """Check if invoice number already exists in the Excel file."""
    try:
        if os.path.exists(filename) and invoice_number != "NA" and invoice_number != "":
            existing_df = pd.read_excel(filename)
            if 'Invoice Number' in existing_df.columns:
                # Check for duplicate invoice numbers (case-insensitive)
                existing_invoices = existing_df['Invoice Number'].astype(str).str.upper()
                return invoice_number.upper() in existing_invoices.values
        return False
    except Exception as e:
        st.error(f"Error checking duplicates: {str(e)}")
        return False

def append_to_file(filename, new_df):
    """Append new data to an existing Excel file or create it if it doesn't exist."""
    max_attempts = 5
    
    for attempt in range(max_attempts):
        try:
            if os.path.exists(filename):
                existing_df = pd.read_excel(filename)
                df = pd.concat([existing_df, new_df], ignore_index=True)
            else:
                df = new_df
            
            # Try to save the file
            df.to_excel(filename, index=False)
            return df
            
        except PermissionError:
            if attempt < max_attempts - 1:
                st.warning(f"⚠️ Cannot access '{filename}'. Please close the file if it's open in Excel or another program. Retrying in 3 seconds... (Attempt {attempt + 1}/{max_attempts})")
                time.sleep(3)
            else:
                # Final attempt - try with a different filename
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_filename = f"invoice_backup_{timestamp}.xlsx"
                st.error(f"❌ Could not save to '{filename}'. Saving as '{backup_filename}' instead.")
                
                try:
                    if os.path.exists(filename):
                        existing_df = pd.read_excel(filename)
                        df = pd.concat([existing_df, new_df], ignore_index=True)
                    else:
                        df = new_df
                    
                    df.to_excel(backup_filename, index=False)
                    st.success(f"✅ Data saved successfully as '{backup_filename}'")
                    return df
                    
                except Exception as e:
                    st.error(f"❌ Failed to save file: {str(e)}")
                    return new_df
                    
        except Exception as e:
            st.error(f"❌ Unexpected error while saving file: {str(e)}")
            return new_df
    
    return new_df

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
    """Load existing data from Excel file with error handling."""
    try:
        if os.path.exists(filename):
            return pd.read_excel(filename)
        else:
            return pd.DataFrame(columns=['File Name', 'Invoice Number', 'Delivery/Shipment Charges', 'Main Date', 'Sender Pincode', 'Receiver Pincode'])
    except PermissionError:
        st.error(f"❌ Cannot read '{filename}'. Please close the file if it's open in Excel.")
        return pd.DataFrame(columns=['File Name', 'Invoice Number', 'Delivery/Shipment Charges', 'Main Date', 'Sender Pincode', 'Receiver Pincode'])
    except Exception as e:
        st.error(f"❌ Error reading file: {str(e)}")
        return pd.DataFrame(columns=['File Name', 'Invoice Number', 'Delivery/Shipment Charges', 'Main Date', 'Sender Pincode', 'Receiver Pincode'])

def create_pincode_analysis(df, pincode_column, title_prefix):
    """Create pincode analysis for either sender or receiver pincodes."""
    # Filter out NA pincodes for analysis - improved NA detection
    valid_pincodes_mask = ((df[pincode_column] != 'NA') & 
                          (~df[pincode_column].isna()) & 
                          (df[pincode_column] != '') &
                          (df[pincode_column].astype(str).str.strip() != 'NA'))
    valid_pincodes = df[valid_pincodes_mask][pincode_column]
    
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
                title=f"Top 10 {title_prefix} Pincodes by Order Count",
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
            st.write(f"**Top 5 {title_prefix} Pincodes:**")
            for i, (pincode, count) in enumerate(pincode_counts.head().items(), 1):
                st.write(f"{i}. **{pincode}**: {count} orders")
            
            # Most frequent pincode
            most_frequent_pincode = pincode_counts.index[0]
            most_frequent_count = pincode_counts.iloc[0]
            st.success(f"🏆 Most frequent {title_prefix.lower()} pincode: **{most_frequent_pincode}** ({most_frequent_count} orders)")
    
    else:
        st.warning(f"No valid {title_prefix.lower()} pincodes found in the data.")

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
    
    # Sender Pincode Analysis
    st.subheader("📍 Sender Pincode Analysis")
    create_pincode_analysis(df, 'Sender Pincode', 'Sender')
    
    st.divider()
    
    # Receiver Pincode Analysis
    st.subheader("📍 Receiver Pincode Analysis")
    create_pincode_analysis(df, 'Receiver Pincode', 'Receiver')
    
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


# Streamlit UI
st.set_page_config(page_title="Invoice Analyzer Dashboard", page_icon="📊", layout="wide")

st.title("📊 Invoice Information Extractor & Dashboard")

# Create tabs for different functionalities
tab1, tab2 = st.tabs(["📤 Upload & Extract", "📊 Dashboard"])

with tab1:
    st.write("**Upload PDF files or ZIP files containing PDFs to extract invoice number, sender pincode, receiver pincode, delivery/shipment charges, and main date.**")
    uploaded_files = st.file_uploader("Choose PDF or ZIP files", accept_multiple_files=True, type=['pdf', 'zip'])

    if uploaded_files:
        # Process all uploaded files (PDFs and ZIPs)
        all_pdf_files = process_uploaded_files(uploaded_files)
        
        if not all_pdf_files:
            st.error("❌ No PDF files found to process!")
        else:
            st.info(f"📋 Total PDF files to process: {len(all_pdf_files)}")
            
            data = []
            skipped_files = []
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, pdf_file in enumerate(all_pdf_files):
                file_name = getattr(pdf_file, 'name', f'file_{i+1}.pdf')
                status_text.text(f"Processing {file_name} ({i+1}/{len(all_pdf_files)})")
                
                # Reset file pointer for processing
                if hasattr(pdf_file, 'seek'):
                    pdf_file.seek(0)
                
                fields = extract_field(pdf_file)

                invoice_number = fields.get("invoice_number", "NA")
                sender_pincode = fields.get("sender_pincode", "NA")
                receiver_pincode = fields.get("receiver_pincode", "NA")
                delivery_charge = fields.get("delivery_charge", "NA")
                main_date = fields.get("main_date", "NA")

                # Check for duplicate invoice number
                output_file = "invoice.xlsx"
                if check_duplicate_invoice(output_file, invoice_number):
                    skipped_files.append(f"{file_name} (Invoice: {invoice_number})")
                    st.warning(f"⚠️ Skipped {file_name}: Invoice number '{invoice_number}' already exists!")
                else:
                    data.append([file_name, invoice_number, delivery_charge, main_date, sender_pincode, receiver_pincode])
                
                progress_bar.progress((i + 1) / len(all_pdf_files))

            status_text.text("Processing complete!")
            
            # Show summary of processing
            if data:
                st.success(f"✅ Successfully processed {len(data)} files!")
            
            if skipped_files:
                st.warning(f"⚠️ Skipped {len(skipped_files)} duplicate files:")
                for skipped in skipped_files:
                    st.write(f"- {skipped}")
            
            if data:
                df = pd.DataFrame(data, columns=['File Name', 'Invoice Number', 'Delivery/Shipment Charges', 'Main Date', 'Sender Pincode', 'Receiver Pincode'])

                st.write("**Extracted Data (New Records Only):**")
                st.dataframe(df, use_container_width=True)

                final_df = append_to_file(output_file, df)

                # Check which file was actually created/updated
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_filename = f"invoice_backup_{timestamp}.xlsx"
                
                download_filename = output_file if os.path.exists(output_file) else backup_filename
                
                if os.path.exists(download_filename):
                    with open(download_filename, "rb") as f:
                        st.download_button(
                            label="📥 Download as Excel",
                            data=f,
                            file_name=download_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                st.info("💡 Data has been added to the Excel file. Check the Dashboard tab to see updated analytics!")
                
                # Show file status
                if os.path.exists(output_file):
                    st.success(f"✅ Data successfully saved to: {output_file}")
                elif os.path.exists(backup_filename):
                    st.warning(f"⚠️ Original file was locked, data saved to: {backup_filename}")
                    st.info("💡 To merge with your main file, close Excel and run the upload again.")
            else:
                st.info("ℹ️ No new records to add. All files were either duplicates or had processing errors.")

with tab2:
    # Load existing data for dashboard
    output_file = "invoice.xlsx"
    existing_df = load_existing_data(output_file)
    
    # Add refresh button and PDF download
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("🔄 Refresh Dashboard"):
            st.rerun()
    
    with col2:
        if st.button("📄 Generate PDF Report") and not existing_df.empty:
            with st.spinner("Generating PDF report..."):
                try:
                    # Get date range if available
                    start_date = None
                    end_date = None
                    
                    # Check if date filtering was applied
                    df_with_valid_dates = existing_df[existing_df['Main Date'] != 'NA'].copy()
                    if not df_with_valid_dates.empty:
                        try:
                            df_with_valid_dates['Date'] = pd.to_datetime(df_with_valid_dates['Main Date'], format='%d-%m-%Y', errors='coerce')
                            df_with_valid_dates = df_with_valid_dates.dropna(subset=['Date'])
                            if not df_with_valid_dates.empty:
                                start_date = df_with_valid_dates['Date'].min().strftime('%d-%m-%Y')
                                end_date = df_with_valid_dates['Date'].max().strftime('%d-%m-%Y')
                        except:
                            pass
                    
                    # Generate PDF
                    pdf = generate_pdf_report(existing_df, start_date, end_date)
                    
                    if pdf:
                        # Create filename with timestamp
                        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"invoice_analysis_report_{timestamp}.pdf"
                        
                        # Convert bytearray to bytes for Streamlit download_button
                        pdf_bytearray = pdf.output()
                        pdf_bytes = bytes(pdf_bytearray)
                        
                        # Provide download button
                        st.download_button(
                            label="📥 Download PDF Report",
                            data=pdf_bytes,
                            file_name=filename,
                            mime="application/pdf"
                        )
                        st.success("✅ PDF report generated successfully!")
                    else:
                        st.error("❌ Failed to generate PDF report")
                        
                except Exception as e:
                    st.error(f"❌ Error generating PDF: {str(e)}")
    
    create_dashboard(existing_df)