from dotenv import load_dotenv
load_dotenv()

import io
import os
import sqlite3
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import fpdf
import joblib
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from sqlalchemy import create_engine, text

# ==========================================
# PAGE CONFIGURATION & CUSTOM STYLING
# ==========================================
st.set_page_config(
    page_title="SmartStock - Wholesale System",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# DATABASE ENGINE (PostgreSQL with SQLite Fallback)
# ==========================================
DB_URI = os.getenv("DATABASE_URL")

import streamlit as st

@st.cache_resource
def get_db_connection():
    """Returns PostgreSQL SQLAlchemy engine connection or SQLite fallback."""
    if DB_URI:
        uri = DB_URI.strip('"\'')
        if uri.startswith("postgresql://"):
            uri = uri.replace("postgresql://", "postgresql+psycopg2://", 1)
        elif uri.startswith("postgres://"):
            uri = uri.replace("postgres://", "postgresql+psycopg2://", 1)
        
        if "&channel_binding=" in uri:
            uri = uri.split("&channel_binding=")[0]
            
        # pool_pre_ping and pool_size for speed
        engine = create_engine(uri, pool_size=5, max_overflow=10, pool_pre_ping=True)
        return engine
    return None
def init_db():
    engine = get_db_connection()
    if engine:
        # PostgreSQL Schema Execution
        with engine.connect() as conn:
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS sales_history (
                    id SERIAL PRIMARY KEY,
                    date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    item_name VARCHAR(255),
                    quantity_sold NUMERIC,
                    unit_price NUMERIC,
                    customer_name VARCHAR(255),
                    recorded_by VARCHAR(100)
                );
            """))
            conn.commit()
    else:
        # SQLite Schema Execution (Fallback)
        conn = sqlite3.connect("inventory.db")
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sales_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                item_name TEXT,
                quantity_sold REAL,
                unit_price REAL,
                customer_name TEXT,
                recorded_by TEXT
            )
        """)
        conn.commit()
        conn.close()

init_db()


# ==========================================
# AUTHENTICATION & ROLE-BASED ACCESS CONTROL (RBAC)
# ==========================================
USER_ROLES = {
    "admin123": "Admin",
    "cashier123": "Cashier",
    "inv123": "Inventory Manager",
    "acct123": "Accountant",
    "staff123": "Staff"
}

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["user_role"] = None

if not st.session_state["authenticated"]:
    st.title("🔐 SmartStock - Enterprise Authentication")
    entered_passcode = st.text_input("Enter Passcode to Access System:", type="password", key="login_pass")

    if st.button("Authenticate"):
        if entered_passcode in USER_ROLES:
            st.session_state["authenticated"] = True
            st.session_state["user_role"] = USER_ROLES[entered_passcode]
            st.rerun()
        else:
            st.error("Authentication Failed: Invalid Passcode.")
    st.stop()


# ==========================================
# UTILITY & REPORT GENERATION FUNCTIONS
# ==========================================
@st.cache_resource
def load_model():
    return joblib.load("inventory_forecasting_model.pkl")

try:
    model = load_model()
except Exception:
    st.error("Model file 'inventory_forecasting_model.pkl' not found.")
    st.stop()

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Purchase_Order")
    return output.getvalue()

def generate_pdf_report(df):
    pdf = fpdf.FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "SmartStock - Supplier Purchase Order", ln=True, align="C")
    pdf.ln(5)

    headers = ["SKU ID", "Item Name", "Current Stock", "Suggested Qty"]
    col_widths = [35, 75, 35, 40]

    pdf.set_font("Arial", "B", 10)
    for header, width in zip(headers, col_widths):
        pdf.cell(width, 8, header, border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial", "", 9)
    for _, row in df.iterrows():
        pdf.cell(col_widths[0], 8, str(row["SKU ID"]), border=1)
        pdf.cell(col_widths[1], 8, str(row["Item Name"]), border=1)
        pdf.cell(col_widths[2], 8, str(row["Current Stock"]), border=1, align="C")
        pdf.cell(col_widths[3], 8, str(row["Suggested Order Qty"]), border=1, align="C")
        pdf.ln()

    return pdf.output(dest="S").encode("latin1")

def generate_customer_invoice(cust_name, cust_phone, cart_items, discount, net_total):
    pdf = fpdf.FPDF()
    pdf.add_page()

    pdf.set_font("Arial", "B", 18)
    pdf.cell(0, 10, "SMARTSTOCK - CUSTOMER INVOICE", ln=True, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, "Nawabshah Wholesale Depot & Distribution Hub", ln=True, align="C")
    pdf.ln(8)

    current_time_str = datetime.now().strftime("%d-%b-%Y %I:%M:%S %p")
    issued_by = st.session_state.get("user_role", "Staff")

    pdf.set_font("Arial", "B", 10)
    pdf.cell(100, 6, f"Customer Name: {cust_name if cust_name else 'Walk-in Customer'}", ln=False)
    pdf.cell(0, 6, f"Date & Time: {current_time_str}", ln=True, align="R")
    
    pdf.cell(100, 6, f"Phone Number: {cust_phone if cust_phone else 'N/A'}", ln=False)
    pdf.cell(0, 6, f"Billed By: {issued_by}", ln=True, align="R")
    pdf.ln(6)

    headers = ["Item Name", "Qty", "Unit Price (PKR)", "Total (PKR)"]
    col_widths = [75, 20, 40, 45]

    pdf.set_font("Arial", "B", 10)
    for header, width in zip(headers, col_widths):
        pdf.cell(width, 8, header, border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial", "", 10)
    subtotal = 0.0
    for item in cart_items:
        item_total = item["qty"] * item["unit_price"]
        subtotal += item_total
        pdf.cell(col_widths[0], 8, str(item["item_name"])[:35], border=1)
        pdf.cell(col_widths[1], 8, str(item["qty"]), border=1, align="C")
        pdf.cell(col_widths[2], 8, f"{item['unit_price']:,.2f}", border=1, align="R")
        pdf.cell(col_widths[3], 8, f"{item_total:,.2f}", border=1, align="R")
        pdf.ln()

    pdf.ln(5)
    pdf.set_font("Arial", "", 11)
    pdf.cell(135, 6, "Subtotal:", align="R")
    pdf.cell(45, 6, f"PKR {subtotal:,.2f}", align="R", ln=True)

    pdf.cell(135, 6, "Discount:", align="R")
    pdf.cell(45, 6, f"- PKR {discount:,.2f}", align="R", ln=True)

    pdf.set_font("Arial", "B", 12)
    pdf.cell(135, 8, "Net Amount Payable:", align="R")
    pdf.cell(45, 8, f"PKR {net_total:,.2f}", align="R", ln=True)

    pdf.ln(15)
    pdf.set_font("Arial", "I", 9)
    pdf.cell(0, 5, "Thank you for doing business with SmartStock Wholesale!", align="C", ln=True)

    return pdf.output(dest="S").encode("latin1")

def get_total_deductions_by_item():
    engine = get_db_connection()
    query = "SELECT item_name, SUM(quantity_sold) as total_sold FROM sales_history GROUP BY item_name"
    try:
        if engine:
            deductions = pd.read_sql_query(query, engine)
        else:
            conn = sqlite3.connect("inventory.db")
            deductions = pd.read_sql_query(query, conn)
            conn.close()
    except Exception:
        deductions = pd.DataFrame(columns=["item_name", "total_sold"])
    return dict(zip(deductions["item_name"], deductions["total_sold"]))


def send_low_stock_email(item_name, current_stock, reorder_point, suggested_qty, sender_email, app_password, receiver_email):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = f"🚨 LOW STOCK ALERT: {item_name}"

        body = f"""
        SmartStock - Automated Inventory Alert

        Attention Admin,

        Item '{item_name}' has dropped below its safe stock threshold!

        • Current Stock: {current_stock} units
        • Re-Order Point Threshold: {reorder_point} units
        • Suggested Order Quantity: {suggested_qty} units

        Please log in to SmartStock Dashboard to generate and approve the purchase order.
        """
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, app_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Failed to send email alert: {e}")
        return False


# ==========================================
# DATA LOADING & INVENTORY ENGINE
# ==========================================
try:
    df = pd.read_csv("wholesale_sales_processed.csv")
    df["Date"] = pd.to_datetime(df["Date"])
except Exception:
    st.error("Data file 'wholesale_sales_processed.csv' missing!")
    st.stop()


# ==========================================
# SIDEBAR CONTROL PANEL & RBAC NAVIGATION
# ==========================================
current_role = st.session_state['user_role']
st.sidebar.title("📦 SmartStock Enterprise")
st.sidebar.caption(f"Role: **{current_role}**")

if st.sidebar.button("Logout", key="logout_btn"):
    st.session_state["authenticated"] = False
    st.session_state["user_role"] = None
    st.rerun()

st.sidebar.markdown("---")

# Granular Role-Based Menu Mapping
ROLE_PERMISSIONS = {
    "Admin": ["🧾 Billing & Invoicing", "💰 Financial Analytics", "🛒 Purchase Orders", "📈 Demand Forecasts", "⚠️ Dead Stock Detector", "📜 Sales Records"],
    "Cashier": ["🧾 Billing & Invoicing"],
    "Inventory Manager": ["🛒 Purchase Orders", "📈 Demand Forecasts", "⚠️ Dead Stock Detector"],
    "Accountant": ["💰 Financial Analytics", "📜 Sales Records"],
    "Staff": ["🧾 Billing & Invoicing", "📜 Sales Records"]
}

menu_options = ROLE_PERMISSIONS.get(current_role, ["🧾 Billing & Invoicing"])
selected_page = st.sidebar.radio("Navigation Menu", menu_options)

st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ Inventory Parameters")
lead_time = st.sidebar.slider("Lead Time (Days)", 1, 14, 3)
service_level = st.sidebar.selectbox("Protection Level", ["95% (Recommended)", "90% (Moderate)", "99% (Maximum)"])
z_score = {"95% (Recommended)": 1.65, "90% (Moderate)": 1.28, "99% (Maximum)": 2.33}.get(service_level, 1.65)

if current_role in ["Admin", "Inventory Manager"]:
    with st.sidebar.expander("📧 Email Alert Settings"):
        enable_email = st.checkbox("Enable Alerts", value=True)
        sender_email = st.text_input("Sender Gmail", value=os.getenv("SENDER_EMAIL", ""), placeholder="gmail@domain.com")
        app_password = st.text_input("App Password", value=os.getenv("APP_PASSWORD", ""), type="password")
        receiver_email = st.text_input("Receiver Email", value=os.getenv("RECEIVER_EMAIL", ""), placeholder="admin@domain.com")
else:
    enable_email = False
    sender_email, app_password, receiver_email = "", "", ""

features = ["Unit_Price", "Day_Of_Week", "Day_Of_Month", "Month", "Is_Weekend", "Quantity_Lag_1", "Quantity_Lag_7", "Quantity_Lag_14", "Quantity_Lag_30", "Rolling_Mean_7", "Rolling_Std_7", "Rolling_Mean_14", "Rolling_Std_14", "Rolling_Mean_30", "Rolling_Std_30"]
latest_date = df["Date"].max()
latest_data = df[df["Date"] == latest_date].copy()
latest_data["Predicted_Demand"] = model.predict(latest_data[features])

live_deductions = get_total_deductions_by_item()
inventory_list = []

for sku in latest_data["SKU_ID"].unique():
    item_data = df[df["SKU_ID"] == sku].sort_values("Date")
    item_info = latest_data[latest_data["SKU_ID"] == sku].iloc[0]
    item_name = item_info["Item_Name"]

    std_sales = item_data["Quantity_Sold"].tail(30).std()
    avg_sales = item_data["Quantity_Sold"].tail(30).mean()

    safety_stock = int(z_score * std_sales * np.sqrt(lead_time))
    reorder_point = int((avg_sales * lead_time) + safety_stock)

    np.random.seed(hash(sku) % 1000)
    base_stock = np.random.randint(low=int(reorder_point * 0.3), high=int(reorder_point * 2.0))

    deducted_qty = live_deductions.get(item_name, 0.0)
    current_stock = max(0, int(base_stock - deducted_qty))

    needs_order = current_stock <= reorder_point
    order_qty = max(0, (reorder_point * 2) - current_stock) if needs_order else 0
    is_dead = avg_sales < 5 and current_stock > 50

    unit_price = float(item_info["Unit_Price"])
    cost_price = float(unit_price * 0.75)
    stock_cost_value = current_stock * cost_price
    stock_sales_value = current_stock * unit_price
    projected_profit = stock_sales_value - stock_cost_value

    inventory_list.append({
        "SKU ID": sku,
        "Item Name": item_name,
        "Category": item_info["Category"],
        "Current Stock": current_stock,
        "Cost Price (PKR)": round(cost_price, 2),
        "Selling Price (PKR)": round(unit_price, 2),
        "Stock Cost Value (PKR)": round(stock_cost_value, 2),
        "Est. Sales Value (PKR)": round(stock_sales_value, 2),
        "Est. Gross Profit (PKR)": round(projected_profit, 2),
        "Re-Order Point": reorder_point,
        "Safety Stock": safety_stock,
        "Predicted Daily Demand": int(item_info["Predicted_Demand"]),
        "Status": "🚨 ORDER NOW" if needs_order else "✅ OK",
        "Suggested Order Qty": order_qty,
        "Dead Stock Warning": "⚠️ CAPITAL BLOCKED" if is_dead else "Clear",
    })

inventory_df = pd.DataFrame(inventory_list)


def record_sale_transaction(item, qty, price, cust_name):
    engine = get_db_connection()
    current_timestamp = datetime.now()
    recorded_by = st.session_state.get("user_role", "Staff")

    if engine:
        # PostgreSQL Transaction Recording
        with engine.connect() as conn:
            conn.execute(text("""
                INSERT INTO sales_history (date, item_name, quantity_sold, unit_price, customer_name, recorded_by)
                VALUES (:date, :item, :qty, :price, :cust, :user)
            """), {"date": current_timestamp, "item": item, "qty": qty, "price": price, "cust": cust_name, "user": recorded_by})
            conn.commit()
    else:
        # SQLite Transaction Recording Fallback
        conn = sqlite3.connect("inventory.db")
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO sales_history (date, item_name, quantity_sold, unit_price, customer_name, recorded_by)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (current_timestamp.strftime("%Y-%m-%d %H:%M:%S"), item, qty, price, cust_name, recorded_by))
        conn.commit()
        conn.close()

    item_row = inventory_df[inventory_df["Item Name"] == item]
    if not item_row.empty:
        curr_stk = item_row["Current Stock"].values[0] - qty
        reorder_pt = item_row["Re-Order Point"].values[0]
        sug_qty = item_row["Suggested Order Qty"].values[0]

        if curr_stk <= reorder_pt and enable_email and sender_email and app_password and receiver_email:
            send_low_stock_email(item, curr_stk, reorder_pt, sug_qty, sender_email, app_password, receiver_email)


# ==========================================
# DASHBOARD MODULES & VIEW ROUTING
# ==========================================
total_capital = inventory_df["Stock Cost Value (PKR)"].sum()
total_sales_value = inventory_df["Est. Sales Value (PKR)"].sum()
total_expected_profit = inventory_df["Est. Gross Profit (PKR)"].sum()
profit_margin_pct = (total_expected_profit / total_sales_value * 100) if total_sales_value > 0 else 0

st.title("📦 SmartStock Enterprise")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Capital Blocked", f"PKR {total_capital:,.0f}")
m2.metric("Inventory Value", f"PKR {total_sales_value:,.0f}")
m3.metric("Est. Gross Profit", f"PKR {total_expected_profit:,.0f}")
m4.metric("Avg Profit Margin", f"{profit_margin_pct:.1f}%")

st.markdown("---")

# Module 1: Billing & Invoicing
if selected_page == "🧾 Billing & Invoicing":
    st.subheader("🧾 Outward Billing & Customer Invoice Generator")

    if "billing_cart" not in st.session_state:
        st.session_state["billing_cart"] = []

    c1, c2 = st.columns(2)
    with c1:
        cust_name = st.text_input("Customer Name", placeholder="e.g. Ali Traders")
        cust_phone = st.text_input("Phone Number", placeholder="0300-XXXXXXX")
    with c2:
        st.text_input("Transaction Date & Time", value=datetime.now().strftime("%Y-%m-%d %I:%M %p"), disabled=True)

    st.markdown("##### 🛒 Add Products To Bill")
    col_item, col_qty, col_add = st.columns([3, 2, 1])

    with col_item:
        selected_item = st.selectbox("Select Product", df["Item_Name"].unique())
        selected_row = inventory_df[inventory_df["Item Name"] == selected_item]
        unit_price = selected_row["Selling Price (PKR)"].values[0] if not selected_row.empty else 100.0
        st.caption(f"Selling Price: **PKR {unit_price:,.2f}**")

    with col_qty:
        quantity = st.number_input("Quantity", min_value=1.0, step=1.0, value=1.0)

    with col_add:
        st.write(" ")
        st.write(" ")
        if st.button("➕ Add", use_container_width=True):
            st.session_state["billing_cart"].append({
                "item_name": selected_item,
                "qty": quantity,
                "unit_price": unit_price,
                "total": quantity * unit_price
            })
            st.success("Added!")

    if st.session_state["billing_cart"]:
        st.markdown("---")
        st.markdown("##### Selected Items Cart:")
        cart_df = pd.DataFrame(st.session_state["billing_cart"])
        st.dataframe(cart_df[["item_name", "qty", "unit_price", "total"]], use_container_width=True)

        col_disc, col_btn = st.columns([2, 1])
        with col_disc:
            discount = st.number_input("Overall Discount (PKR)", min_value=0.0, step=10.0, value=0.0)
            subtotal = sum(i["total"] for i in st.session_state["billing_cart"])
            net_total = max(0.0, subtotal - discount)
            st.markdown(f"### Payable Amount: **PKR {net_total:,.2f}**")

        with col_btn:
            st.write(" ")
            st.write(" ")
            if st.button("🗑 Clear Cart"):
                st.session_state["billing_cart"] = []
                st.rerun()

        if st.button("✅ Complete Sale & Generate PDF Invoice", type="primary"):
            for item in st.session_state["billing_cart"]:
                record_sale_transaction(item["item_name"], item["qty"], item["unit_price"], cust_name)

            pdf_data = generate_customer_invoice(cust_name, cust_phone, st.session_state["billing_cart"], discount, net_total)
            st.success("Transaction recorded successfully with exact timestamp!")

            st.download_button(
                label="📄 Download Printable PDF Invoice",
                data=pdf_data,
                file_name=f"Invoice_{cust_name if cust_name else 'Customer'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf"
            )
            st.session_state["billing_cart"] = []

# Module 2: Financial Analytics
elif selected_page == "💰 Financial Analytics":
    st.subheader("💰 Profit Margin & Inventory Capital Breakdown")
    fin_cols = ["SKU ID", "Item Name", "Category", "Current Stock", "Cost Price (PKR)", "Selling Price (PKR)", "Stock Cost Value (PKR)", "Est. Gross Profit (PKR)"]
    st.dataframe(inventory_df[fin_cols], use_container_width=True)

    fig_margin = px.bar(
        inventory_df, x="Item Name", y="Est. Gross Profit (PKR)", color="Category",
        title="Projected Profit Breakdown by Product"
    )
    st.plotly_chart(fig_margin, use_container_width=True)

# Module 3: Purchase Orders
elif selected_page == "🛒 Purchase Orders":
    st.subheader("🛒 Supplier Re-Order Recommendations")
    order_now_df = inventory_df[inventory_df["Status"] == "🚨 ORDER NOW"]

    if not order_now_df.empty:
        st.dataframe(order_now_df, use_container_width=True)
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📊 Export Excel Purchase Order", convert_df_to_excel(order_now_df), "Purchase_Order.xlsx")
        with col2:
            st.download_button("📄 Export PDF Purchase Order", generate_pdf_report(order_now_df), "Purchase_Order.pdf")
    else:
        st.success("All stock levels are optimal!")

    st.markdown("---")
    st.subheader("📋 Full Inventory Master Table")
    st.dataframe(inventory_df, use_container_width=True)

# Module 4: Demand Forecasts
elif selected_page == "📈 Demand Forecasts":
    st.subheader("📈 Historical Demand & Forecast Analytics")
    selected_sku = st.selectbox("Select Product to Analyze:", df["Item_Name"].unique())
    sku_data = df[df["Item_Name"] == selected_sku].sort_values("Date").tail(90)
    fig = px.line(sku_data, x="Date", y="Quantity_Sold", title=f"90-Day Demand Trend for {selected_sku}")
    st.plotly_chart(fig, use_container_width=True)

# Module 5: Dead Stock Detector
elif selected_page == "⚠️ Dead Stock Detector":
    st.subheader("⚠️ Dead Stock & Blocked Capital Detector")
    dead_stock_df = inventory_df[inventory_df["Dead Stock Warning"] != "Clear"]
    if not dead_stock_df.empty:
        st.warning("Low-velocity stock detected! Capital blocked in unsold inventory.")
        st.dataframe(dead_stock_df, use_container_width=True)
    else:
        st.success("No dead stock detected!")

# Module 6: Sales Records
elif selected_page == "📜 Sales Records":
    st.subheader("📜 Recorded Sales Transactions")
    engine = get_db_connection()
    query = "SELECT id, date, customer_name, item_name, quantity_sold, unit_price, recorded_by FROM sales_history ORDER BY id DESC LIMIT 100"
    
    if engine:
        df_db = pd.read_sql_query(query, engine)
    else:
        conn = sqlite3.connect("inventory.db")
        df_db = pd.read_sql_query(query, conn)
        conn.close()

    if not df_db.empty:
        st.dataframe(df_db, use_container_width=True)
    else:
        st.info("No recorded transactions found.")

st.markdown("---")
st.caption("⚡ SmartStock Enterprise System")