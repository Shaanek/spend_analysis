############################################################
# Supplier and Purchase Order (PO) Spend Analysis
# Purpose: Analyze procurement data to understand spending patterns
# Created by: [Your Name]
# Last Updated: [Date]
############################################################

"""
SECTION 0: IMPORT REQUIRED LIBRARIES
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

"""
SECTION 1: DATA LOADING AND PREPROCESSING
"""
def load_po_data(file_path):
    """
    Load PO data from Excel file and perform initial preprocessing
    Args:
        file_path (str): Path to the Excel file containing PO data
    Returns:
        pandas.DataFrame: Processed PO data
    """
    try:
        # Read the Excel file
        df_po_report = pd.read_excel(file_path)
        print("Successfully loaded the input file.")
        
        # Convert date columns to datetime
        df_po_report['Po Creation Date'] = pd.to_datetime(df_po_report['Po Creation Date'])
        
        return df_po_report
    
    except Exception as e:
        print(f"Error loading the file: {str(e)}")
        return None

"""
SECTION 2: SUPPLIER SPEND ANALYSIS
"""
def create_supplier_spend_charts(df):
    """
    Create visualizations for supplier spend analysis
    Args:
        df (pandas.DataFrame): PO data
    """
    # Calculate total spend per supplier
    spend_by_supplier = df.groupby('Supplier Name')['PO Amount Due '].sum().reset_index()
    
    # 2.1 Create Bar Chart
    fig_bar = px.bar(
        spend_by_supplier,
        x='Supplier Name',
        y='PO Amount Due ',
        title='Spend by Supplier - Bar Chart',
        labels={'PO Amount Due ': 'Total Spend ($)', 
                'Supplier Name': 'Supplier Name'},
        template='plotly_white'
    )
    
    # Customize bar chart for better readability
    fig_bar.update_layout(
        xaxis_tickangle=-45,  # Rotate labels for better visibility
        showlegend=False,
        height=600,
        title_x=0.5,         # Center the title
        yaxis_title="Total Spend ($)",
        xaxis_title="Supplier Name",
        hoverlabel=dict(bgcolor="white"),  # White background for hover text
    )
    
    # Display bar chart
    fig_bar.show()
    
    # 2.2 Create Pie Chart
    fig_pie = px.pie(
        spend_by_supplier,
        values='PO Amount Due ',
        names='Supplier Name',
        title='Supplier Spend Distribution',
        template='plotly_white',
        hole=0.3  # Creates a donut chart
    )
    
    # Customize pie chart
    fig_pie.update_layout(
        height=600,
        title_x=0.5,
        legend=dict(
            orientation="h",     # Horizontal legend
            yanchor="bottom",    # Position at bottom
            y=-0.3,
            xanchor="center",
            x=0.5
        )
    )
    
    # Display pie chart
    fig_pie.show()

"""
SECTION 3: ITEM MODEL ANALYSIS
"""
def create_item_model_analysis(df):
    """
    Create visualizations for item model analysis
    Args:
        df (pandas.DataFrame): PO data
    """
    # Calculate spend by item model
    spend_by_model = df.groupby('Item Model Description')['PO Amount Due '].sum().reset_index()
    spend_by_model = spend_by_model.sort_values('PO Amount Due ', ascending=False)
    
    # 3.1 Create Bar Chart for All Models
    fig_model_bar = px.bar(
        spend_by_model,
        x='Item Model Description',
        y='PO Amount Due ',
        title='Spend by Item Model',
        labels={
            'PO Amount Due ': 'Total Spend ($)', 
            'Item Model Description': 'Item Model'
        },
        template='plotly_white'
    )
    
    # Customize model bar chart
    fig_model_bar.update_layout(
        xaxis_tickangle=-45,
        height=700,
        margin=dict(b=150),
        yaxis_title="Total Spend ($)"
    )
    
    # Display model bar chart
    fig_model_bar.show()
    
    # 3.2 Create Top 10 Models Pie Chart
    top_10_models = spend_by_model.head(10)
    fig_model_pie = px.pie(
        top_10_models,
        values='PO Amount Due ',
        names='Item Model Description',
        title='Top 10 Item Models by Spend',
        template='plotly_white',
        hole=0.3
    )
    
    # Display top 10 pie chart
    fig_model_pie.show()

"""
SECTION 4: BUYER ANALYSIS
"""
def create_buyer_analysis(df):
    """
    Create visualizations for buyer analysis
    Args:
        df (pandas.DataFrame): PO data
    """
    # Calculate buyer metrics
    buyer_metrics = df.groupby('Buyer').agg({
        'PO #': 'nunique',          # Count unique POs
        'PO Amount Due ': 'sum',     # Total spend
        'Supplier Name': 'nunique'   # Number of suppliers
    }).reset_index()
    
    # Rename columns for clarity
    buyer_metrics.columns = ['Buyer', 'Number of POs', 'Total Spend', 'Number of Suppliers']
    
    # 4.1 Create Bubble Chart
    fig_bubble = px.scatter(
        buyer_metrics,
        x='Number of POs',
        y='Total Spend',
        size='Number of Suppliers',
        text='Buyer',
        title='Buyer Analysis Dashboard',
        template='plotly_white',
        labels={
            'Number of POs': 'Number of Purchase Orders',
            'Total Spend': 'Total Spend ($)',
            'Number of Suppliers': 'Number of Suppliers'
        }
    )
    
    # Customize bubble chart
    fig_bubble.update_traces(
        textposition='top center',
        hovertemplate="<b>%{text}</b><br>" +
                      "POs: %{x}<br>" +
                      "Spend: $%{y:,.2f}<br>" +
                      "Suppliers: %{marker.size}<extra></extra>"
    )
    
    # Display bubble chart
    fig_bubble.show()

"""
SECTION 5: MAIN EXECUTION
"""
def main():
    # File path - Replace with your file path
    file_path = "path/to/your/PO_Report.xlsx"
    
    # Load data
    df_po_report = load_po_data(file_path)
    
    if df_po_report is not None:
        # Create all visualizations
        create_supplier_spend_charts(df_po_report)
        create_item_model_analysis(df_po_report)
        create_buyer_analysis(df_po_report)
        
        # Print summary statistics
        print("\nAnalysis Summary:")
        print("================")
        print(f"Total POs: {df_po_report['PO #'].nunique():,}")
        print(f"Total Spend: ${df_po_report['PO Amount Due '].sum():,.2f}")
        print(f"Total Suppliers: {df_po_report['Supplier Name'].nunique():,}")
        print(f"Total Buyers: {df_po_report['Buyer'].nunique():,}")

if __name__ == "__main__":
    main()



