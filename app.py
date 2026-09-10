import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import zipfile

st.set_page_config(page_title="Ecommerce Cleaner", layout="wide")

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
    elif "SHOPIFY" in name:
        return "shopify"
    elif "SW_TIKTOK" in name or "SUNNYWOOD_TIKTOK" in name:
        return "tiktok (Sunnywood)"
    elif "TIKTOK" in name and "SW" not in name and "SUNNYWOOD_TIKTOK" not in name:
        return "tiktok"
    return None


# =========================
# CLEANING FUNCTIONS (FULL)
# =========================

def clean_lazada(df):
    columns_to_keep = [
        'orderItemId',
        'lazadaId',
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
        'shippingFee',
        'itemName',
        'variation',
        'shippingProvider',
        'trackingCode',
        'status',
        'buyerFailedDeliveryReturnInitiator',
        'buyerFailedDeliveryReason',
        'buyerFailedDeliveryDetail',
        'refundAmount',
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

    df['paidPrice'] = df['paidPrice'].astype(float)
    df['unitPrice'] = df['unitPrice'].astype(float)
    df['sellerDiscountTotal'] = df['sellerDiscountTotal'].astype(float).fillna(0)

    df['paidPrice'] = df['unitPrice'] + df['sellerDiscountTotal']
    df['paidPrice'] = df['paidPrice'].where(df['paidPrice'] >= 0, 0)

    for col in ['orderItemId','lazadaId','orderNumber']:
        df[col] = df[col].astype(str)

    return df


def clean_shopee(df):
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
        # 'Username (Buyer)'
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


def clean_zalora(df):
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
    df = df.sort_values(by=['Created at'])

    df['Created at'] = df['Created at'].dt.strftime('%B %d, %Y')

    df['Paid Price'] = pd.to_numeric(df['Paid Price'], errors='coerce').fillna(0)
    df['Wallet Credits'] = pd.to_numeric(df['Wallet Credits'], errors='coerce').fillna(0)

    col_index = df.columns.get_loc("Paid Price")
    df.insert(col_index, "Amount", df['Paid Price'] + df['Wallet Credits'])

    df['Order Number'] = df['Order Number'].astype(str)

    return df


def clean_shopify(df):
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
        "Payment Reference",
        "Refunded Amount",
        "Outstanding Balance",
        "Id",
        "Lineitem discount",
        "Tax 1 Name",
        "Tax 1 Value",
        "Payment ID",
        "Payment Terms Name",
        "Next Payment Due At",
        "Payment References"
    ] 
    for col in columns_to_keep:
        if col not in df.columns:
            df[col] = pd.NA

    df = df[columns_to_keep].copy()

    cols_to_fill = ['Financial Status', 'Fulfillment Status']
    df[cols_to_fill] = df.groupby('Name')[cols_to_fill].transform('ffill')

    def try_parse_datetime(dt_str):
        from datetime import datetime
        for fmt in ["%Y-%m-%d %H:%M:%S %z","%d/%m/%Y %I:%M:%S %p"]:
            try:
                return datetime.strptime(str(dt_str), fmt)
            except:
                continue
        return pd.to_datetime(dt_str, errors='coerce')

    df['Created at'] = df['Created at'].apply(try_parse_datetime)
    df = df.sort_values(by='Created at')

    df['Lineitem price'] = df['Lineitem price'] * df['Lineitem quantity']

    df['Created at'] = pd.to_datetime(df['Created at']).dt.strftime('%B %d, %Y')

    return df


def clean_tiktok(df):
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
      # "Buyer Username",
      "Product Category",
      "Package ID"
    ] 
    df = df[columns_to_keep].copy()

    df['Created Time'] = pd.to_datetime(df['Created Time'],
                                        format="%m/%d/%Y %I:%M:%S %p",
                                        errors='coerce')

    df = df.sort_values(by='Created Time', na_position='last')

    df['Created Time'] = df['Created Time'].dt.strftime('%B %d, %Y')

    df['SKU Subtotal After Discount'] = pd.to_numeric(
        df['SKU Subtotal After Discount'], errors='coerce').fillna(0)

    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0).astype(int)

    for col in ['Order ID','SKU ID']:
        df[col] = df[col].astype(str)

    df['Package ID'] = df['Package ID'].apply(
        lambda x: str(int(x)) if pd.notnull(x) else None)

    return df

def clean_SW_tiktok(df):
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
      # "Buyer Username",
      "Product Category",
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
    
    df['SKU Subtotal After Discount'] = pd.to_numeric(
        df['SKU Subtotal After Discount'], errors='coerce').fillna(0)

    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0).astype(int)

    for col in ['Order ID','SKU ID']:
        df[col] = df[col].astype(str)

    df['Package ID'] = df['Package ID'].apply(
        lambda x: str(int(x)) if pd.notnull(x) else None)

    return df

# =========================
# UI
# =========================

st.subheader("Upload Files")

with st.container(border=True):
    uploaded_files = st.file_uploader(
        label="Drag & drop your Excel/CSV files here",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        help="You can upload multiple files at once.",
        width="stretch"
    )

manual_override = st.selectbox(
    "Manual Platform Override (optional)",
    ["Auto Detect","lazada","shopee","zalora","shopify","tiktok", "tiktok (Sunnywood)"]
)

st.divider()
st.subheader("Progress:")

# Progress UI
progress_bar = st.progress(0)
progress_text = st.empty()

# Processing uploaded files
if uploaded_files:
    total_files = len(uploaded_files)
    
    zip_buffer = io.BytesIO()
    zip_file = zipfile.ZipFile(zip_buffer, "w")

    for i, file in enumerate(uploaded_files, start=1):

        platform = detect_platform(file.name) if manual_override == "Auto Detect" else manual_override

        st.write(f"### 📄 {file.name}")
        st.write(f"Platform: **{platform.upper()}**")

        try:
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)

            if platform == "lazada":
                cleaned = clean_lazada(df)
            elif platform == "shopee":
                cleaned = clean_shopee(df)
            elif platform == "zalora":
                cleaned = clean_zalora(df)
            elif platform == "shopify":
                cleaned = clean_shopify(df)
            elif platform == "tiktok":
                cleaned = clean_tiktok(df)
            elif platform == "tiktok (Sunnywood)":
                cleaned = clean_SW_tiktok(df)
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
                file_name=filename
            )

            zip_file.writestr(filename, output.getvalue())

        except Exception as e:
            st.error(f"Error: {e}")
            
        # Update progress
        progress = i / total_files
        progress_bar.progress(progress)
        progress_text.text(f"Processing files... {i}/{total_files}")

    zip_file.close()

    progress_bar.progress(1.0)
    progress_text.success("✅ All files processed successfully!")

    st.divider()

    st.download_button(
        "⬇️ Download ALL as ZIP",
        data=zip_buffer.getvalue(),
        file_name="cleaned_files.zip"
    )
