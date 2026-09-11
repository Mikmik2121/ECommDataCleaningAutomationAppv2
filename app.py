# E-COMMERCE DATA PROCESSING TOOL
# Version: 2.0.0
# Date: September 10, 2026

import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import zipfile

st.set_page_config(page_title="Ecommerce Data Processing Tool", layout="wide")

st.title("E-Commerce Data Processing Tool")

@st.dialog("How to Use?")
def show_instructions():
    st.markdown("""
    ### Instructions:

    - Before uploading, rename the files using this format: ```<BRAND>_<PLATFORM> (ex. VANS_SHOPEE)```

    - Upload ```.xlsx``` and ```.csv files``` (csv files for Shopify only). 
    
    - If Auto detector doesn't work, you can manually set the platform of the files you want to format.
    
    - Wait for the app to process your data until a ```'Download ALL as ZIP'``` button appears at the very bottom of the page.
    
    - Download your clean files and you're done! :DD
    
    """)

if st.button("How to Use?"):
    show_instructions()

st.divider()

# =========================
# PLATFORM DETECTION
# =========================
def detect_platform(filename):
    name = filename.upper()
    if "LAZADA" in name:
        return "lazada"
    elif "SHOPEE" in name:
        return "shopee"
    elif "ZALORA" in name:
        return "zalora"
    elif "SHOPIFY" in name or "WEBSITE" in name:
        return "shopify"
    elif "TIKTOK" in name:
        return "tiktok"
    return None

# =================================
# RAW SALES DATA CLEANING FUNCTIONS
# =================================

def clean_lazada_sales(df):
    columns_to_keep = [
        'orderItemId',
        'sellerSku',
        'lazadaSku',
        'createTime', 
        'updateTime', 
        'rtsSla',
        'ttsSla',
        'orderNumber', 
        'deliveredDate',
        'paidPrice',
        'unitPrice',
        'sellerDiscountTotal',
        'platformDiscountTotal',
        'shippingFee',
        'itemName',
        'variation', 
        'shippingProvider',
        'trackingCode',
        'status',
        'buyerFailedDeliveryReturnInitiator', 
        'buyerFailedDeliveryReason'
    ]
    
    df = df[columns_to_keep].copy()

    df['createTime'] = pd.to_datetime(df['createTime'], format='%d %b %Y %H:%M')

    col_index = df.columns.get_loc('createTime')
    df.insert(col_index, "createTime(Date)", df['createTime'].dt.strftime('%B %d, %Y'))
    df.insert(col_index + 1, "Time", df['createTime'].dt.strftime('%H:%M:%S'))

    df = df.drop(columns=["createTime"])

    df['Date_sort'] = pd.to_datetime(df['createTime(Date)'], format='%B %d, %Y')
    df['Time_sort'] = pd.to_datetime(df['Time'], format='%H:%M:%S').dt.time

    df = df.sort_values(by=['Date_sort', 'Time_sort'])

    df = df.drop(columns=['Date_sort', 'Time_sort'])

    df['unitPrice'] = pd.to_numeric(df['unitPrice'], errors='coerce').fillna(0)
    df['sellerDiscountTotal'] = pd.to_numeric(df['sellerDiscountTotal'], errors='coerce').fillna(0).abs()
    df['platformDiscountTotal'] = pd.to_numeric(df['platformDiscountTotal'], errors='coerce').fillna(0).abs()

    col_index = df.columns.get_loc("paidPrice")
    df.insert(col_index, "Amount Paid", df['unitPrice'] - df['sellerDiscountTotal'])
    df['Amount Paid'] = df['Amount Paid'].where(df['Amount Paid'] >= 0, 0) # if result displays a negative number, convert to 0 instead

    for col in ['orderItemId','orderNumber']:
        df[col] = df[col].astype(str)

    return df


def clean_shopee_sales(df):
    columns_to_keep = ['Order ID',
        'Order Status',
        'Cancel reason',
        'Return / Refund Status',
        'Tracking Number*',
        'Shipping Option',
        'Estimated Ship Out Date',
        'Ship Time',
        'Order Creation Date',
        'Order Paid Time',
        'Parent SKU Reference No.',
        'Product Name',
        'SKU Reference No.',
        'Variation Name',
        'Original Price',
        'Deal Price',
        'Quantity',
        'Returned quantity',
        'Product Subtotal',
        'Total Discount(PHP)',
        'Price Discount(from Seller)(PHP)',
        'Shopee Rebate(PHP)',
        'Seller Voucher(PHP)',
        'Coin Cashback Voucher Amount Sponsored by Seller',
        #'Seller Absorbed Coin Cashback',
        'Shopee Voucher(PHP)',
        'Shopee Coins Offset(PHP)',
        'Credit Card Discount Total(PHP)',
        "Products' Price Paid by Buyer (PHP)",
        'Buyer Paid Shipping Fee',
        'Shipping Rebate Estimate',
        'Reverse Shipping Fee',
        'Service Fee',
        'Grand Total',
        'Estimated Shipping Fee',
    ]
    
    df = df[columns_to_keep].copy()

    mask = (
        df['Return / Refund Status'].notna() &
        (df['Return / Refund Status'].astype(str).str.strip() != '')
    )
    df.loc[mask, 'Order Status'] = df.loc[mask, 'Return / Refund Status']

    df['Order Creation Date'] = pd.to_datetime(df['Order Creation Date'], format='%Y-%m-%d %H:%M')

    col_index = df.columns.get_loc("Order Creation Date")
    df.insert(col_index, "Order Creation Date (Date)", df['Order Creation Date'].dt.strftime('%B %d, %Y'))
    df.insert(col_index + 1, "Time", df['Order Creation Date'].dt.strftime('%H:%M:%S'))

    df = df.drop(columns=["Order Creation Date"])

    df['Date_sort'] = pd.to_datetime(df['Order Creation Date (Date)'], format='%B %d, %Y')
    df['Time_sort'] = pd.to_datetime(df['Time'], format='%H:%M:%S').dt.time

    df = df.sort_values(by=['Date_sort', 'Time_sort'])

    df = df.drop(columns=['Date_sort', 'Time_sort'])

    df['Product Subtotal'] = df['Product Subtotal'].astype(float)
    df['Quantity'] = df['Quantity'].astype(int)

    return df


def clean_zalora_sales(df):
    columns_to_keep = [
        'Order Item Id',
        'Zalora Id',
        'Seller SKU',
        'Zalora SKU',
        'Created at',
        'Updated at',
        'Order Number',
        'Paid Price',
        'Unit Price',
        'Tax Amount',
        'Shipping Fee',
        'Wallet Credits',
        'Item Name',
        'Variation',
        'Shipping Provider',
        'Tracking Code',
        'Status',
        'Reason',
        'voucher: discount',
        'Store Credits',
        'Shipping Voucher'
    ]
    df = df[columns_to_keep].copy()

    df['Created at'] = pd.to_datetime(df['Created at'], format='%Y-%m-%d %H:%M:%S', errors='coerce')

    col_index = df.columns.get_loc('Created at')
    df.insert(col_index, "Created at(Date)", df['Created at'].dt.strftime('%B %d, %Y'))
    df.insert(col_index + 1, "Time", df['Created at'].dt.strftime('%H:%M:%S'))

    df = df.drop(columns=["Created at"])

    df['Date_sort'] = pd.to_datetime(df['Created at(Date)'], format='%B %d, %Y')
    df['Time_sort'] = pd.to_datetime(df['Time'], format='%H:%M:%S').dt.time

    df = df.sort_values(by=['Date_sort', 'Time_sort'])

    df = df.drop(columns=['Date_sort', 'Time_sort'])

    df['Paid Price'] = pd.to_numeric(df['Paid Price'], errors='coerce').fillna(0)
    df['Wallet Credits'] = pd.to_numeric(df['Wallet Credits'], errors='coerce').fillna(0)

    col_index = df.columns.get_loc("Paid Price")
    df.insert(col_index, "Amount Paid", df['Paid Price'] + df['Wallet Credits'])

    df['Order Number'] = df['Order Number'].astype(str)

    return df


def clean_shopify_sales(df):
    columns_to_keep = [ 
        "Name",
        "Email",
        "Financial Status",
        "Paid at",
        "Fulfillment Status",
        "Fulfilled at",
        "Accepts Marketing",
        "Subtotal",
        "Shipping",
        "Taxes",
        "Total",
        "Discount Code",
        "Discount Amount",
        "Shipping Method",
        "Created at",
        "Lineitem quantity",
        "Lineitem name",
        "Lineitem price",
        "Lineitem compare at price",
        "Lineitem sku",
        "Payment Method",
        "Refunded Amount",
        "Outstanding Balance",
        "Lineitem discount",
        "Tax 1 Name",
        "Tax 1 Value",
        "Payment Terms Name",
        "Next Payment Due At",
        "Payment References"
    ] 
    for col in columns_to_keep:
        if col not in df.columns:
            df[col] = pd.NA

    df = df[columns_to_keep].copy()

    cols_to_fill = ['Financial Status', 'Fulfillment Status', 'Paid at', 'Fulfilled at', 'Email', 'Accepts Marketing', 'Subtotal', 'Shipping', 
                    'Taxes', 'Total', 'Discount Code', 'Discount Amount', 'Shipping Method', 'Payment Method', 'Refunded Amount',
                   'Outstanding Balance', 'Tax 1 Name', 'Tax 1 Value', 'Payment Terms Name', 'Next Payment Due At', 'Payment References']
    df[cols_to_fill] = df.groupby('Name')[cols_to_fill].transform('ffill')
    
    col_index = df.columns.get_loc("Lineitem price")
    df.insert(col_index, "Amount Paid", df['Lineitem price'] * df['Lineitem quantity'])

    def parse_datetime(dt_str):
        formats = ["%Y-%m-%d %H:%M:%S %z", "%d/%m/%Y %I:%M:%S %p"]
        for fmt in formats:
            try:
                return pd.to_datetime(dt_str, format=fmt)
            except (ValueError, TypeError):
                continue
        return pd.to_datetime(dt_str, errors='coerce')

    df['Created at'] = df['Created at'].apply(parse_datetime)

    col_index = df.columns.get_loc('Created at')
    df.insert(col_index, "Created at(Date)", df['Created at'].dt.strftime('%B %d, %Y'))
    df.insert(col_index + 1, "Time", df['Created at'].dt.strftime('%H:%M:%S'))

    df = df.drop(columns=["Created at"])

    df['Date_sort'] = pd.to_datetime(df['Created at(Date)'], format='%B %d, %Y')
    df['Time_sort'] = pd.to_datetime(df['Time'], format='%H:%M:%S').dt.time

    df = df.sort_values(by=['Date_sort', 'Time_sort'])

    df = df.drop(columns=['Date_sort', 'Time_sort'])

    return df


def clean_tiktok_sales(df):
    columns_to_keep = [
      "Order ID",
      "Order Status",
      "Order Substatus",
      "Cancelation/Return Type",
      "SKU ID",
      "Seller SKU",
      "Product Name",
      "Variation",
      "Quantity",
      "Sku Quantity of return",
      "SKU Unit Original Price",
      "SKU Subtotal Before Discount",
      "SKU Platform Discount",
      "SKU Seller Discount",
      "SKU Subtotal After Discount",
      "Shipping Fee After Discount",
      "Original Shipping Fee",
      "Shipping Fee Seller Discount",
      "Shipping Fee Platform Discount",
      "Payment platform discount",
      "Taxes",
      "Order Amount",
      "Order Refund Amount",
      "Created Time",
      "Paid Time",
      "RTS Time",
      "Shipped Time",
      "Delivered Time",
      "Cancelled Time",
      "Cancel By",
      "Cancel Reason",
      "Tracking ID",
      "Shipping Provider Name",
      "Package ID"
    ] 
    df = df[columns_to_keep].copy()

    df['Created Time'] = pd.to_datetime(df['Created Time'], format="%m/%d/%Y %I:%M:%S %p", errors='coerce')
    
    col_index = df.columns.get_loc('Created Time')
    df.insert(col_index, "Created Time(Date)", df['Created Time'].dt.strftime('%B %d, %Y'))
    df.insert(col_index + 1, "Time", df['Created Time'].dt.strftime('%H:%M:%S'))

    df = df.drop(columns=["Created Time"])

    df['Date_sort'] = pd.to_datetime(df['Created Time(Date)'], format='%B %d, %Y')
    df['Time_sort'] = pd.to_datetime(df['Time'], format='%H:%M:%S').dt.time

    df = df.sort_values(by=['Date_sort', 'Time_sort'])

    df = df.drop(columns=['Date_sort', 'Time_sort'])

    df['SKU Subtotal Before Discount'] = pd.to_numeric(df['SKU Subtotal Before Discount'], errors='coerce').fillna(0)
    df['SKU Seller Discount'] = pd.to_numeric(df['SKU Seller Discount'], errors='coerce').fillna(0)

    col_index = df.columns.get_loc("SKU Subtotal Before Discount")
    df.insert(col_index, "Amount Paid", df['SKU Subtotal Before Discount'] - df['SKU Seller Discount'])

    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0).astype(int)

    for col in ['Order ID','SKU ID']:
        df[col] = df[col].astype(str)

    df['Package ID'] = df['Package ID'].apply(
        lambda x: str(int(x)) if pd.notnull(x) else None)

    return df

# =============================================================================================
# RAW PRODUCT TRAFFIC DATA CLEANING FUNCTIONS
# =============================================================================================

def clean_lazada_traffic(df):
    df = df[df['Seller SKU'] == "-"]

    columns_to_keep = [
        'Product ID',
        'Seller SKU',
        'Product Name',
        'Product Pageviews',
        'Product Clicks',
        'Orders',
        'Add to Cart Units',
        'CTR',
        'Conversion Rate'
    ]
    
    for col in columns_to_keep:
        if col not in df.columns:
            df[col] = pd.NA
    
    df = df[columns_to_keep].copy()

    df['Conversion Rate'] = (df['Conversion Rate'].str.replace('%', '', regex=True).astype(float) / 100)

    df['Product Pageviews'] = df['Product Pageviews'].astype(int)
    df['Orders'] = df['Orders'].astype(int)
    df['Add to Cart Units'] = df['Add to Cart Units'].astype(int)
    df['Conversion Rate'] = df['Conversion Rate'].astype(float)

    id_columns = ['Product ID']
    for col in id_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)

    return df


def clean_shopee_traffic(df):
    df = df[df['SKU'] == "-"]

    # Columns to keep
    columns_to_keep = [
        'Item ID',
        'Parent SKU',
        'Product',
        'Product Impression',
        'Product Clicks',
        'Confirmed Order',
        'Units (Add to Cart)',
        'CTR',
        'Order Conversion Rate (Confirmed Order)'
    ]

    df = df[columns_to_keep].copy()

    df['CTR'] = df['CTR'].str.replace('%', '', regex=True).astype(float) / 100
    df['Order Conversion Rate (Confirmed Order)'] = (
        df['Order Conversion Rate (Confirmed Order)']
        .str.replace('%', '', regex=True)
        .astype(float) / 100
    )

    df['Product Impression'] = df['Product Impression'].astype(int)
    df['Product Clicks'] = df['Product Clicks'].astype(int)
    df['Confirmed Order'] = df['Confirmed Order'].astype(int)
    df['Units (Add to Cart)'] = df['Units (Add to Cart)'].astype(int)
    df['CTR'] = df['CTR'].astype(float)
    df['Order Conversion Rate (Confirmed Order)'] = df['Order Conversion Rate (Confirmed Order)'].astype(float)

    id_columns = ['Item ID', 'Parent SKU']
    for col in id_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)

    return df


def clean_zalora_traffic(df):
    df['Add to Cart Units'] = df['Gross Orders']

    columns_to_keep = [
        'Config sku',
        'Parent SKU',
        'Product Name',
        'Impression views',
        'Product Clicks',
        'Gross Orders',
        'Add to Cart Units',
        'CTR',
        'Conversion Rate'
    ]
    
    for col in columns_to_keep:
        if col not in df.columns:
            df[col] = pd.NA

    df = df[columns_to_keep].copy()

    df['Conversion Rate'] = (df['Conversion Rate'].astype(float) / 100)

    df['Impression views'] = df['Impression views'].astype(int)
    df['Gross Orders'] = df['Gross Orders'].astype(int)
    df['Add to Cart Units'] = df['Add to Cart Units'].astype(int)
    df['Conversion Rate'] = df['Conversion Rate'].astype(float)

    id_columns = ['Config sku']
    for col in id_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)

    return df


def clean_tiktok_traffic(df):
    df['Conversion Rate'] = df['Orders']/df['Product clicks']

    columns_to_keep = [
        'Product ID',
        'Product Name',
        'Product impressions',
        'Product clicks',
        'Orders',
        'Add-to-cart count',
        'CTR',
        'Conversion Rate'
    ]

    df = df[columns_to_keep].copy()

    df['CTR'] = df['CTR'].str.replace('%', '', regex=True).astype(float) / 100

    df['Product impressions'] = df['Product impressions'].astype(int)
    df['Product clicks'] = df['Product clicks'].astype(int)
    df['Orders'] = df['Orders'].astype(int)
    df['Add-to-cart count'] = df['Add-to-cart count'].astype(int)
    df['CTR'] = df['CTR'].astype(float)
    df['Conversion Rate'] = df['Conversion Rate'].astype(float)

    id_columns = ['Product ID']
    for col in id_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)

    return df

# =======================================
# RAW RETURN/REFUND DATA CLEANING FUNCTIONS
# =======================================

def clean_lazada_returns(df):
    columns_to_keep = [
        'Return Order Date',
        'Order ID',
        'Seller SKU ID',
        'Return Reason', 
        'Status', 
        'Logistic Status'
    ]
    
    df = df[columns_to_keep].copy()

    df['Return Order Date'] = pd.to_datetime(df['Return Order Date'], format='%Y-%m-%d %H:%M:%S')
    df['Return Order Date'] = df['Return Order Date'].dt.strftime('%B %d, %Y')

    df = df.sort_values(by=['Return Order Date'])

    df['Order ID'] = df['Order ID'].astype(str)

    return df

def clean_shopee_returns(df):
    columns_to_keep = [
        'Return Creation Time',
        'Order ID',
        'SKU',
        'Return Reason', 
        'Return / Refund Status', 
        'Return Tracking Status'
    ]
    
    df = df[columns_to_keep].copy()

    df['Return Creation Time'] = pd.to_datetime(df['Return Creation Time'], format='%Y-%m-%d %H:%M')
    df['Return Creation Time'] = df['Return Creation Time'].dt.strftime('%B %d, %Y')

    df = df.sort_values(by=['Return Creation Time'])

    df['Order ID'] = df['Order ID'].astype(str)

    return df

def clean_tiktok_returns(df):
    columns_to_keep = [
        "Time Requested",
        "Order ID",
        "Seller SKU",
        "Return Reason", 
        "Return Status", 
        "Return Sub Status"
    ]
    
    df = df[columns_to_keep].copy()

    df['Time Requested'] = pd.to_datetime(df['Time Requested'], format="%m/%d/%Y %I:%M:%S %p", errors='coerce')
    df['Time Requested'] = df['Time Requested'].dt.strftime('%B %d, %Y')

    df = df.sort_values(by=['Time Requested'])

    df['Order ID'] = df['Order ID'].astype(str)

    return df

# =========================
# SALES DATA UI
# =========================

st.subheader("Sales Data")

with st.container(border=True):
    uploaded_sales_files = st.file_uploader(
        label="Drag & drop your Excel/CSV files here",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        help="You can upload multiple files at once.",
        width="stretch",
        key="sales_uploader"
    )

manual_override_sales = st.selectbox(
    "Manual Platform Override (optional)",
    ["Auto Detect","lazada","shopee","zalora","shopify","tiktok"],
    key="sales_platform_override"
)

st.divider()
st.subheader("Progress:")

# Progress UI
progress_bar_sales = st.progress(0)
progress_text_sales = st.empty()

# Processing uploaded files
if uploaded_sales_files:
    total_sales_files = len(uploaded_sales_files)
    
    zip_buffer = io.BytesIO()
    zip_file = zipfile.ZipFile(zip_buffer, "w")

    for i, file in enumerate(uploaded_sales_files, start=1):

        platform = detect_platform(file.name) if manual_override_sales == "Auto Detect" else manual_override_sales

        st.write(f"### 📄 {file.name}")
        st.write(f"Platform: **{platform.upper()}**")

        try:
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)

            if platform == "lazada":
                cleaned = clean_lazada_sales(df)
            elif platform == "shopee":
                cleaned = clean_shopee_sales(df)
            elif platform == "zalora":
                cleaned = clean_zalora_sales(df)
            elif platform == "shopify":
                cleaned = clean_shopify_sales(df)
            elif platform == "tiktok":
                cleaned = clean_tiktok_sales(df)
            else:
                st.error("Unknown platform")
                continue

            st.dataframe(cleaned.head(20))

            output = io.BytesIO()
            cleaned.to_excel(output, index=False)
            output.seek(0)

            base, ext = os.path.splitext(file.name)
            filename = f"{base}_cleaned.xlsx"

            st.download_button(
                f"Download {file.name}",
                data=output,
                file_name=filename,
                key=f"sales_download_{i}_{file.name}"
            )

            zip_file.writestr(filename, output.getvalue())

        except Exception as e:
            st.error(f"Error: {e}")
            
        # Update progress
        progress_sales = i / total_sales_files
        progress_bar_sales.progress(progress_sales)
        progress_text_sales.text(f"Processing files... {i}/{total_sales_files}")

    zip_file.close()

    progress_bar_sales.progress(1.0)
    progress_text_sales.success(f"✅ All files processed successfully! {i}/{total_sales_files}")

    st.divider()

    st.download_button(
        "⬇️ Download ALL as ZIP",
        data=zip_buffer.getvalue(),
        file_name="cleaned_sales_files.zip",
        key="download_all_sales"
    )

# =========================
# PRODUCT TRAFFIC DATA UI
# =========================

st.subheader("Product Traffic Data")

with st.container(border=True):
    uploaded_traffic_files = st.file_uploader(
        label="Drag & drop your Excel/CSV files here",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="You can upload multiple files at once.",
        width="stretch",
        key="traffic_uploader"
    )

manual_override_traffic = st.selectbox(
    "Manual Platform Override (optional)",
    ["Auto Detect","lazada","shopee","zalora","shopify","tiktok"],
    key="traffic_platform_override"
)

st.divider()
st.subheader("Progress:")

# Progress UI
progress_bar_traffic = st.progress(0)
progress_text_traffic = st.empty()

# Processing uploaded files
if uploaded_traffic_files:
    total_traffic_files = len(uploaded_traffic_files)
    
    zip_buffer = io.BytesIO()
    zip_file = zipfile.ZipFile(zip_buffer, "w")

    for i, file in enumerate(uploaded_traffic_files, start=1):

        platform = detect_platform(file.name) if manual_override_traffic == "Auto Detect" else manual_override_traffic

        st.write(f"### 📄 {file.name}")
        st.write(f"Platform: **{platform.upper()}**")

        try:
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)

            if platform == "lazada":
                cleaned = clean_lazada_traffic(df)
            elif platform == "shopee":
                cleaned = clean_shopee_traffic(df)
            elif platform == "zalora":
                cleaned = clean_zalora_traffic(df)
            elif platform == "tiktok":
                cleaned = clean_tiktok_traffic(df)
            else:
                st.error("Unknown platform")
                continue

            st.dataframe(cleaned.head(20))

            output = io.BytesIO()
            cleaned.to_excel(output, index=False)
            output.seek(0)

            base, ext = os.path.splitext(file.name)
            filename = f"{base}_cleaned.xlsx"

            st.download_button(
                f"Download {file.name}",
                data=output,
                file_name=filename,
                key=f"traffic_download_{i}_{file.name}"
            )

            zip_file.writestr(filename, output.getvalue())

        except Exception as e:
            st.error(f"Error: {e}")
            
        # Update progress
        progress_traffic = i / total_traffic_files
        progress_bar_traffic.progress(progress_traffic)
        progress_text_traffic.text(f"Processing files... {i}/{total_traffic_files}")

    zip_file.close()

    progress_bar_traffic.progress(1.0)
    progress_text_traffic.success(f"✅ All files processed successfully! {i}/{total_traffic_files}")

    st.divider()

    st.download_button(
        "⬇️ Download ALL as ZIP",
        data=zip_buffer.getvalue(),
        file_name="cleaned_traffic_files.zip",
        key="download_all_traffic"
    )

# =========================
# RETURN/REFUND DATA UI
# =========================

st.subheader("Return/Refund Data")

with st.container(border=True):
    uploaded_returns_files = st.file_uploader(
        label="Drag & drop your Excel/CSV files here",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="You can upload multiple files at once.",
        width="stretch",
        key="returns_uploader"
    )

manual_override_returns = st.selectbox(
    "Manual Platform Override (optional)",
    ["Auto Detect","lazada","shopee","zalora","shopify","tiktok"],
    key="returns_platform_override"
)

st.divider()
st.subheader("Progress:")

# Progress UI
progress_bar_returns = st.progress(0)
progress_text_returns = st.empty()

# Processing uploaded files
if uploaded_returns_files:
    total_returns_files = len(uploaded_returns_files)
    
    zip_buffer = io.BytesIO()
    zip_file = zipfile.ZipFile(zip_buffer, "w")

    for i, file in enumerate(uploaded_returns_files, start=1):

        platform = detect_platform(file.name) if manual_override_returns == "Auto Detect" else manual_override_returns

        st.write(f"### 📄 {file.name}")
        st.write(f"Platform: **{platform.upper()}**")

        try:
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)

            if platform == "lazada":
                cleaned = clean_lazada_returns(df)
            elif platform == "shopee":
                cleaned = clean_shopee_returns(df)
            elif platform == "tiktok":
                cleaned = clean_tiktok_returns(df)
            else:
                st.error("Unknown platform")
                continue

            st.dataframe(cleaned.head(20))

            output = io.BytesIO()
            cleaned.to_excel(output, index=False)
            output.seek(0)

            base, ext = os.path.splitext(file.name)
            filename = f"{base}_cleaned.xlsx"

            st.download_button(
                f"Download {file.name}",
                data=output,
                file_name=filename,
                key=f"returns_download_{i}_{file.name}"
            )

            zip_file.writestr(filename, output.getvalue())

        except Exception as e:
            st.error(f"Error: {e}")
            
        # Update progress
        progress_returns = i / total_returns_files
        progress_bar_returns.progress(progress_returns)
        progress_text_returns.text(f"Processing files... {i}/{total_returns_files}")

    zip_file.close()

    progress_bar_returns.progress(1.0)
    progress_text_returns.success(f"✅ All files processed successfully! {i}/{total_returns_files}")

    st.divider()

    st.download_button(
        "⬇️ Download ALL as ZIP",
        data=zip_buffer.getvalue(),
        file_name="cleaned_returns_files.zip",
        key="download_all_returns"
    )
