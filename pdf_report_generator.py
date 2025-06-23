import pandas as pd
from fpdf import FPDF
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import tempfile
import os
import base64
import streamlit as st

class InvoiceReportPDF(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
        
        # Add Unicode font support
        try:
            # Try to use DejaVu Sans (commonly available)
            self.add_font('DejaVu', '', 'DejaVuSans.ttf')
            self.add_font('DejaVu', 'B', 'DejaVuSans-Bold.ttf')
            self.font_family = 'DejaVu'
        except:
            try:
                # Fallback to system fonts
                import platform
                if platform.system() == "Windows":
                    self.add_font('Arial', '', 'arial.ttf')
                    self.add_font('Arial', 'B', 'arialbd.ttf')
                    self.font_family = 'Arial'
                else:
                    # For Linux/Mac, try common Unicode fonts
                    self.add_font('Liberation', '', 'LiberationSans-Regular.ttf')
                    self.add_font('Liberation', 'B', 'LiberationSans-Bold.ttf')
                    self.font_family = 'Liberation'
            except:
                # Final fallback - use built-in fonts but replace Unicode chars
                self.font_family = 'Arial'
                self.unicode_fallback = True
    
    def safe_text(self, text):
        """Convert Unicode characters to safe alternatives for non-Unicode fonts."""
        if hasattr(self, 'unicode_fallback') and self.unicode_fallback:
            # Replace common Unicode characters with ASCII alternatives
            replacements = {
                '₹': 'Rs.',
                '–': '-',
                '—': '--',
                ''': "'",
                ''': "'",
                '"': '"',
                '"': '"',
                '…': '...',
                '•': '*',
                '€': 'EUR',
                '£': 'GBP',
                '$': 'USD'
            }
            for unicode_char, ascii_char in replacements.items():
                text = text.replace(unicode_char, ascii_char)
        return str(text)
    
    def header(self):
        self.set_font(self.font_family, 'B', 16)
        self.cell(0, 10, self.safe_text('Invoice Analysis Report'), 0, 1, 'C')
        self.set_font(self.font_family, '', 10)
        self.cell(0, 10, self.safe_text(f'Generated on: {datetime.now().strftime("%B %d, %Y at %I:%M %p")}'), 0, 1, 'C')
        self.ln(10)
    
    def footer(self):
        self.set_y(-15)
        self.set_font(self.font_family, 'I', 8)
        self.cell(0, 10, self.safe_text(f'Page {self.page_no()}'), 0, 0, 'C')

def extract_numeric_value(delivery_charge):
    """Extract numeric value from delivery charge string."""
    if pd.isna(delivery_charge) or delivery_charge == "NA":
        return 0
    
    import re
    delivery_str = str(delivery_charge)
    numbers = re.findall(r'\d+\.?\d*', delivery_str)
    if numbers:
        return float(numbers[0])
    return 0

def create_summary_metrics(df):
    """Calculate summary metrics from the dataframe."""
    df['Numeric_Delivery_Charge'] = df['Delivery/Shipment Charges'].apply(extract_numeric_value)
    
    total_delivery_spent = df['Numeric_Delivery_Charge'].sum()
    na_count = ((df['Delivery/Shipment Charges'] == 'NA') | 
                (df['Delivery/Shipment Charges'].isna()) | 
                (df['Delivery/Shipment Charges'] == '') |
                (df['Delivery/Shipment Charges'].astype(str).str.strip() == 'NA')).sum()
    total_orders = len(df)
    na_percentage = (na_count / total_orders * 100) if total_orders > 0 else 0
    
    # Valid charges statistics
    valid_charges = df[df['Numeric_Delivery_Charge'] > 0]['Numeric_Delivery_Charge']
    avg_charge = valid_charges.mean() if not valid_charges.empty else 0
    median_charge = valid_charges.median() if not valid_charges.empty else 0
    min_charge = valid_charges.min() if not valid_charges.empty else 0
    max_charge = valid_charges.max() if not valid_charges.empty else 0
    
    return {
        'total_delivery_spent': total_delivery_spent,
        'total_orders': total_orders,
        'na_count': na_count,
        'na_percentage': na_percentage,
        'avg_charge': avg_charge,
        'median_charge': median_charge,
        'min_charge': min_charge,
        'max_charge': max_charge
    }

def get_pincode_analysis(df, pincode_column):
    """Get pincode analysis for sender or receiver."""
    valid_pincodes_mask = ((df[pincode_column] != 'NA') & 
                          (~df[pincode_column].isna()) & 
                          (df[pincode_column] != '') &
                          (df[pincode_column].astype(str).str.strip() != 'NA'))
    valid_pincodes = df[valid_pincodes_mask][pincode_column]
    
    if not valid_pincodes.empty:
        valid_pincodes = valid_pincodes.astype(str).str.strip()
        pincode_counts = valid_pincodes.value_counts()
        return pincode_counts.head(10)  # Top 10
    return pd.Series()

def generate_pdf_report(df, start_date=None, end_date=None):
    """Generate a comprehensive PDF report from the dataframe."""
    
    if df.empty:
        return None
    
    # Create PDF instance
    pdf = InvoiceReportPDF()
    pdf.add_page()
    
    # Title and date range
    pdf.set_font(pdf.font_family, 'B', 14)
    pdf.cell(0, 10, pdf.safe_text('Executive Summary'), 0, 1, 'L')
    pdf.ln(5)
    
    if start_date and end_date:
        pdf.set_font(pdf.font_family, '', 10)
        pdf.cell(0, 8, pdf.safe_text(f'Report Period: {start_date} to {end_date}'), 0, 1, 'L')
        pdf.ln(5)
    
    # Summary Metrics
    metrics = create_summary_metrics(df)
    
    pdf.set_font(pdf.font_family, 'B', 12)
    pdf.cell(0, 10, pdf.safe_text('Key Metrics'), 0, 1, 'L')
    pdf.set_font(pdf.font_family, '', 10)
    
    # Create metrics table - using Rs. instead of ₹ for compatibility
    metrics_data = [
        ['Total Orders', f"{metrics['total_orders']:,}"],
        ['Total Delivery Spent', f"Rs.{metrics['total_delivery_spent']:,.2f}"],
        ['Average Delivery Charge', f"Rs.{metrics['avg_charge']:.2f}"],
        ['Median Delivery Charge', f"Rs.{metrics['median_charge']:.2f}"],
        ['Minimum Delivery Charge', f"Rs.{metrics['min_charge']:.2f}"],
        ['Maximum Delivery Charge', f"Rs.{metrics['max_charge']:.2f}"],
        ['Orders with NA Charges', f"{metrics['na_count']} ({metrics['na_percentage']:.1f}%)"]
    ]
    
    for metric, value in metrics_data:
        pdf.cell(80, 8, pdf.safe_text(metric + ':'), 0, 0, 'L')
        pdf.cell(0, 8, pdf.safe_text(value), 0, 1, 'L')
    
    pdf.ln(10)
    
    # Sender Pincode Analysis
    pdf.set_font(pdf.font_family, 'B', 12)
    pdf.cell(0, 10, pdf.safe_text('Top Sender Pincodes'), 0, 1, 'L')
    pdf.set_font(pdf.font_family, '', 10)
    
    sender_pincodes = get_pincode_analysis(df, 'Sender Pincode')
    if not sender_pincodes.empty:
        for i, (pincode, count) in enumerate(sender_pincodes.head(5).items(), 1):
            pdf.cell(0, 6, pdf.safe_text(f"{i}. {pincode}: {count} orders"), 0, 1, 'L')
    else:
        pdf.cell(0, 6, pdf.safe_text("No valid sender pincodes found"), 0, 1, 'L')
    
    pdf.ln(5)
    
    # Receiver Pincode Analysis
    pdf.set_font(pdf.font_family, 'B', 12)
    pdf.cell(0, 10, pdf.safe_text('Top Receiver Pincodes'), 0, 1, 'L')
    pdf.set_font(pdf.font_family, '', 10)
    
    receiver_pincodes = get_pincode_analysis(df, 'Receiver Pincode')
    if not receiver_pincodes.empty:
        for i, (pincode, count) in enumerate(receiver_pincodes.head(5).items(), 1):
            pdf.cell(0, 6, pdf.safe_text(f"{i}. {pincode}: {count} orders"), 0, 1, 'L')
    else:
        pdf.cell(0, 6, pdf.safe_text("No valid receiver pincodes found"), 0, 1, 'L')
    
    pdf.ln(10)
    
    # Monthly Analysis (if dates are available)
    df_with_dates = df[df['Main Date'] != 'NA'].copy()
    if not df_with_dates.empty:
        try:
            df_with_dates['Date'] = pd.to_datetime(df_with_dates['Main Date'], format='%d-%m-%Y', errors='coerce')
            df_with_dates = df_with_dates.dropna(subset=['Date'])
            
            if not df_with_dates.empty:
                pdf.set_font(pdf.font_family, 'B', 12)
                pdf.cell(0, 10, pdf.safe_text('Monthly Summary'), 0, 1, 'L')
                pdf.set_font(pdf.font_family, '', 10)
                
                df_with_dates['Month'] = df_with_dates['Date'].dt.to_period('M')
                monthly_summary = df_with_dates.groupby('Month').agg({
                    'File Name': 'count',
                    'Numeric_Delivery_Charge': 'sum'
                }).rename(columns={'File Name': 'Orders', 'Numeric_Delivery_Charge': 'Total_Delivery'})
                
                for month, row in monthly_summary.iterrows():
                    pdf.cell(0, 6, pdf.safe_text(f"{month}: {row['Orders']} orders, Rs.{row['Total_Delivery']:.2f} total delivery"), 0, 1, 'L')
        except:
            pass
    
    # Add new page for detailed data
    pdf.add_page()
    pdf.set_font(pdf.font_family, 'B', 14)
    pdf.cell(0, 10, pdf.safe_text('Detailed Data'), 0, 1, 'L')
    pdf.ln(5)
    
    # Table headers
    pdf.set_font(pdf.font_family, 'B', 8)
    col_widths = [60, 40, 25, 30, 30]
    headers = ['File Name', 'Delivery Charge', 'Date', 'Sender PIN', 'Receiver PIN']
    
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 8, pdf.safe_text(header), 1, 0, 'C')
    pdf.ln()
    
    # Table data
    pdf.set_font(pdf.font_family, '', 7)
    for _, row in df.head(50).iterrows():  # Limit to first 50 rows to avoid too many pages
        # Truncate long file names
        file_name = str(row['File Name'])[:40] + '...' if len(str(row['File Name'])) > 40 else str(row['File Name'])
        
        pdf.cell(col_widths[0], 6, pdf.safe_text(file_name), 1, 0, 'L')
        pdf.cell(col_widths[1], 6, pdf.safe_text(str(row['Delivery/Shipment Charges'])[:20]), 1, 0, 'C')
        pdf.cell(col_widths[2], 6, pdf.safe_text(str(row['Main Date'])[:15]), 1, 0, 'C')
        pdf.cell(col_widths[3], 6, pdf.safe_text(str(row['Sender Pincode'])[:15]), 1, 0, 'C')
        pdf.cell(col_widths[4], 6, pdf.safe_text(str(row['Receiver Pincode'])[:15]), 1, 0, 'C')
        pdf.ln()
    
    if len(df) > 50:
        pdf.ln(5)
        pdf.set_font(pdf.font_family, 'I', 8)
        pdf.cell(0, 6, pdf.safe_text(f"Note: Showing first 50 records out of {len(df)} total records"), 0, 1, 'L')
    
    return pdf

def create_download_link(pdf_content, filename):
    """Create a download link for the PDF."""
    b64 = base64.b64encode(pdf_content).decode()
    return f'<a href="data:application/pdf;base64,{b64}" download="{filename}">Download PDF Report</a>'
