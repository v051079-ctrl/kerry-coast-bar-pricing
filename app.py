"""
Kerry Coast Hotel - Spirits Master Catalog & Procurement Intelligence System
=============================================================================
Complete standalone Python/Streamlit web application for bar stock & price audit.
Fully compatible with Streamlit Cloud & local execution.

Includes:
1. Master Catalog (PLU Touch Office, Categories, Volumes, Irish Measure calculations, CRUD)
2. Supplier Prices (Wholesale distributors, FOC "Buy X Get Y Free" deals, Optic cost)
3. Price Comparison (Live multi-supplier benchmark, lowest bottle/measure cost, savings)
4. Invoice Check (Delivery invoice audit, contract price variance, overcharge claim generator)
5. Import / Export (Consolidated 4-sheet Excel Kerry_Coast_Spirits_Master_Catalog.xlsx)
"""

import os
import io
from datetime import datetime
import pandas as pd
import streamlit as st

# ==============================================================================
# STREAMLIT PAGE CONFIGURATION
# ==============================================================================
st.set_page_config(
    page_title="Kerry Coast Hotel - Spirits Management",
    page_icon="🍸",
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_FILENAME = "Kerry_Coast_Spirits_Master_Catalog.xlsx"
SHEET_MASTER = "Master Catalog"
SHEET_PRICES = "Supplier Prices"
SHEET_COMPARISON = "Price Comparison"
SHEET_INVOICE = "Invoice Check"

# Irish optics & bar serving standards
STANDARD_MEASURES = {
    "Whiskey": 35.5,
    "Gin": 35.5,
    "Vodka": 35.5,
    "Rum": 35.5,
    "Brandy/Cognac": 35.5,
    "Tequila": 35.5,
    "Liqueur": 50.0,
    "Vermouth": 50.0,
    "Wine & Sparkling": 175.0,
    "Beer & Cider": 568.0,
    "Soft Drinks & Mixers": 200.0,
}

# ==============================================================================
# INITIAL DATASETS (Embedded seed data for instant startup on Streamlit Cloud)
# ==============================================================================
DEFAULT_MASTER_DATA = [
    {"PLU": "121", "Product Name": "PADDY", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "122", "Product Name": "POWER / POWERS GOLD LABEL", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "123", "Product Name": "JAMESON", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "124", "Product Name": "JAMESON BLACK BARREL", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "125", "Product Name": "BUSHMILLS RED / ORIGINAL", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "126", "Product Name": "BLACK BUSH", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "127", "Product Name": "GREENSPOT", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "128", "Product Name": "YELLOW SPOT", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "132", "Product Name": "RED BREAST 12YR", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "161", "Product Name": "GORDONS GIN", "Category": "Gin", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "164", "Product Name": "DINGLE GIN", "Category": "Gin", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "168", "Product Name": "BOMBAY SAPPHIRE", "Category": "Gin", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "191", "Product Name": "SMIRNOFF VODKA", "Category": "Vodka", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "194", "Product Name": "TITOS", "Category": "Vodka", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "211", "Product Name": "HENNESSY", "Category": "Brandy/Cognac", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "221", "Product Name": "BACARDI", "Category": "Rum", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "231", "Product Name": "JACK DANIELS", "Category": "Whiskey", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "261", "Product Name": "JOSE CUERVO BLANCO", "Category": "Tequila", "Typical Volume (ml)": 700, "Measure Size (ml)": 35.5, "Active": "YES"},
    {"PLU": "271", "Product Name": "JAGERMEISTER", "Category": "Liqueur", "Typical Volume (ml)": 700, "Measure Size (ml)": 50.0, "Active": "YES"},
    {"PLU": "274", "Product Name": "SAMBUCA", "Category": "Liqueur", "Typical Volume (ml)": 700, "Measure Size (ml)": 50.0, "Active": "YES"},
    {"PLU": "291", "Product Name": "BAILEYS", "Category": "Liqueur", "Typical Volume (ml)": 700, "Measure Size (ml)": 50.0, "Active": "YES"},
    {"PLU": "300", "Product Name": "DISARONNO", "Category": "Liqueur", "Typical Volume (ml)": 700, "Measure Size (ml)": 50.0, "Active": "YES"},
    {"PLU": "301", "Product Name": "VALENTIA VERMOUTH", "Category": "Vermouth", "Typical Volume (ml)": 700, "Measure Size (ml)": 50.0, "Active": "YES"},
    {"PLU": "1508", "Product Name": "APEROL", "Category": "Liqueur", "Typical Volume (ml)": 700, "Measure Size (ml)": 50.0, "Active": "YES"}
]

DEFAULT_PRICES_DATA = [
    {"Supplier": "Cliffords C&C", "Product Name": "Jameson 1L x 6", "Pack Size": 6, "Volume (ml)": 1000, "Price per Case (€)": 174.95, "Linked PLU": "123", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Sive", "Product Name": "Jameson 1L x 6", "Pack Size": 6, "Volume (ml)": 1000, "Price per Case (€)": 193.50, "Linked PLU": "123", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Classic Drinks", "Product Name": "Jameson 1L x 6 (Promo)", "Pack Size": 6, "Volume (ml)": 1000, "Price per Case (€)": 182.00, "Linked PLU": "123", "FOC Buy": 5, "FOC Free": 1, "In Stock": True},
    {"Supplier": "Cliffords C&C", "Product Name": "Gordon's Gin 1L x 12", "Pack Size": 12, "Volume (ml)": 1000, "Price per Case (€)": 285.95, "Linked PLU": "161", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Classic Drinks", "Product Name": "Gordons Gin 1L x 12", "Pack Size": 12, "Volume (ml)": 1000, "Price per Case (€)": 298.00, "Linked PLU": "161", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Sive", "Product Name": "Gordons Gin 1L x 12", "Pack Size": 12, "Volume (ml)": 1000, "Price per Case (€)": 318.00, "Linked PLU": "161", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Cliffords C&C", "Product Name": "Smirnoff Vodka 1L x 12", "Pack Size": 12, "Volume (ml)": 1000, "Price per Case (€)": 258.90, "Linked PLU": "191", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Classic Drinks", "Product Name": "Smirnoff Vodka 1L x 12", "Pack Size": 12, "Volume (ml)": 1000, "Price per Case (€)": 262.00, "Linked PLU": "191", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Classic Drinks", "Product Name": "Hennessy 700ml x 12", "Pack Size": 12, "Volume (ml)": 700, "Price per Case (€)": 359.00, "Linked PLU": "211", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Cliffords C&C", "Product Name": "Hennessy 700ml x 12", "Pack Size": 12, "Volume (ml)": 700, "Price per Case (€)": 339.00, "Linked PLU": "211", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Classic Drinks", "Product Name": "Bacardi 1L x 6", "Pack Size": 6, "Volume (ml)": 1000, "Price per Case (€)": 151.00, "Linked PLU": "221", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Cliffords C&C", "Product Name": "Bacardi 1L x 6", "Pack Size": 6, "Volume (ml)": 1000, "Price per Case (€)": 162.95, "Linked PLU": "221", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Sive", "Product Name": "Bacardi 1L x 6", "Pack Size": 6, "Volume (ml)": 1000, "Price per Case (€)": 149.70, "Linked PLU": "221", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Classic Drinks", "Product Name": "Jack Daniels 700ml x 6", "Pack Size": 6, "Volume (ml)": 700, "Price per Case (€)": 138.00, "Linked PLU": "231", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Cliffords C&C", "Product Name": "Jack Daniels 700ml x 6", "Pack Size": 6, "Volume (ml)": 700, "Price per Case (€)": 146.95, "Linked PLU": "231", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Classic Drinks", "Product Name": "Baileys 1L x 12", "Pack Size": 12, "Volume (ml)": 1000, "Price per Case (€)": 201.00, "Linked PLU": "291", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Cliffords C&C", "Product Name": "Baileys 1L x 12", "Pack Size": 12, "Volume (ml)": 1000, "Price per Case (€)": 220.00, "Linked PLU": "291", "FOC Buy": 0, "FOC Free": 0, "In Stock": True},
    {"Supplier": "Cliffords C&C", "Product Name": "Aperol 700ml x 6", "Pack Size": 6, "Volume (ml)": 700, "Price per Case (€)": 74.95, "Linked PLU": "1508", "FOC Buy": 0, "FOC Free": 0, "In Stock": True}
]

# ==============================================================================
# STATE MANAGEMENT
# ==============================================================================
def init_session_state():
    """Initializes app state with Master Catalog and Supplier Quotes."""
    if "master_df" not in st.session_state:
        df_m = pd.DataFrame(DEFAULT_MASTER_DATA)
        df_m["Measures in Bottle"] = (df_m["Typical Volume (ml)"] / df_m["Measure Size (ml)"]).round(2)
        st.session_state.master_df = df_m

    if "prices_df" not in st.session_state:
        st.session_state.prices_df = pd.DataFrame(DEFAULT_PRICES_DATA)

    if "invoice_items" not in st.session_state:
        st.session_state.invoice_items = [
            {"PLU": "123", "Product Name": "Jameson Irish Whiskey 1L", "Qty Units": 3, "Bottles in Unit": 6, "Invoice Price": 182.00},
            {"PLU": "161", "Product Name": "Gordon's London Dry Gin 1L", "Qty Units": 2, "Bottles in Unit": 12, "Invoice Price": 298.00},
            {"PLU": "191", "Product Name": "Smirnoff Red Label Vodka 1L", "Qty Units": 4, "Bottles in Unit": 12, "Invoice Price": 275.00},
            {"PLU": "211", "Product Name": "Hennessy VS Cognac 700ml", "Qty Units": 1, "Bottles in Unit": 12, "Invoice Price": 369.00}
        ]

init_session_state()

# ==============================================================================
# CORE BUSINESS FORMULAS
# ==============================================================================
def calculate_effective_bottle_price(row: pd.Series) -> float:
    """Calculates bottle cost including FOC 'Buy X Get Y Free' deals."""
    case_price = float(row.get("Price per Case (€)", 0.0))
    pack_size = int(row.get("Pack Size", 1))
    if pack_size <= 0:
        pack_size = 1

    standard_bottle = case_price / pack_size
    foc_buy = int(row.get("FOC Buy", 0) or 0)
    foc_free = int(row.get("FOC Free", 0) or 0)

    if foc_buy > 0 and foc_free > 0:
        effective = standard_bottle * (foc_buy / (foc_buy + foc_free))
        return round(effective, 2)
    return round(standard_bottle, 2)


def get_enriched_prices() -> pd.DataFrame:
    """Enriches supplier quotes with measures and cost per measure."""
    df_p = st.session_state.prices_df.copy()
    if df_p.empty:
        return df_p

    df_p["Effective Bottle Price (€)"] = df_p.apply(calculate_effective_bottle_price, axis=1)

    # Lookup measure size from master catalog
    df_m = st.session_state.master_df
    m_dict = df_m.set_index("PLU")["Measure Size (ml)"].to_dict() if not df_m.empty else {}

    def get_measures_for_quote(row):
        plu_str = str(row.get("Linked PLU", ""))
        measure_size = m_dict.get(plu_str, 35.5)
        volume = float(row.get("Volume (ml)", 700))
        return round(volume / measure_size, 2) if measure_size > 0 else 19.72

    df_p["Measures in Bottle"] = df_p.apply(get_measures_for_quote, axis=1)
    df_p["Cost per Measure (€)"] = (df_p["Effective Bottle Price (€)"] / df_p["Measures in Bottle"]).round(2)
    return df_p


def generate_comparison_df() -> pd.DataFrame:
    """Builds side-by-side comparison matrix for all PLUs across suppliers."""
    df_m = st.session_state.master_df
    df_p = get_enriched_prices()

    records = []
    for _, m in df_m.iterrows():
        plu = str(m["PLU"])
        name = m.get("Product Name", "")
        cat = m.get("Category", "")
        vol = m.get("Typical Volume (ml)", 700)
        msize = m.get("Measure Size (ml)", 35.5)
        measures = m.get("Measures in Bottle", round(vol / msize, 2))

        quotes = df_p[df_p["Linked PLU"].astype(str) == plu] if not df_p.empty else pd.DataFrame()
        if not quotes.empty:
            min_bottle = quotes["Effective Bottle Price (€)"].min()
            best_supplier = quotes.loc[quotes["Effective Bottle Price (€)"].idxmin()]["Supplier"]
            max_bottle = quotes["Effective Bottle Price (€)"].max()
            savings = round(max_bottle - min_bottle, 2)
            savings_pct = round((savings / max_bottle) * 100, 1) if max_bottle > 0 else 0.0
            lowest_measure = round(min_bottle / measures, 2) if measures > 0 else min_bottle

            records.append({
                "PLU": plu,
                "Product Name": name,
                "Category": cat,
                "Volume (ml)": vol,
                "Measures/Btl": measures,
                "Lowest Bottle (€)": min_bottle,
                "Lowest Measure (€)": lowest_measure,
                "Cheapest Supplier": best_supplier,
                "Highest Bottle (€)": max_bottle,
                "Savings per Bottle (€)": savings,
                "Savings (%)": savings_pct,
                "Quotes Count": len(quotes)
            })
        else:
            records.append({
                "PLU": plu,
                "Product Name": name,
                "Category": cat,
                "Volume (ml)": vol,
                "Measures/Btl": measures,
                "Lowest Bottle (€)": None,
                "Lowest Measure (€)": None,
                "Cheapest Supplier": "No quotes",
                "Highest Bottle (€)": None,
                "Savings per Bottle (€)": 0.0,
                "Savings (%)": 0.0,
                "Quotes Count": 0
            })
    return pd.DataFrame(records)


def export_full_excel_workbook() -> bytes:
    """Exports all 4 sheets into memory buffer for instant download."""
    output = io.BytesIO()
    df_m = st.session_state.master_df.copy()
    df_p = get_enriched_prices()
    df_c = generate_comparison_df()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_m.to_excel(writer, sheet_name=SHEET_MASTER, index=False)
        df_p.to_excel(writer, sheet_name=SHEET_PRICES, index=False)
        df_c.to_excel(writer, sheet_name=SHEET_COMPARISON, index=False)
        pd.DataFrame(st.session_state.invoice_items).to_excel(writer, sheet_name=SHEET_INVOICE, index=False)
        
    return output.getvalue()


# ==============================================================================
# SIDEBAR
# ==============================================================================
with st.sidebar:
    st.markdown("## 🏨 Kerry Coast Hotel")
    st.caption("Spirits Master Catalog & Price Management")
    st.divider()

    st.metric("Master PLU Catalog", len(st.session_state.master_df))
    distributors = st.session_state.prices_df["Supplier"].nunique() if not st.session_state.prices_df.empty else 0
    st.metric("Active Suppliers", distributors)
    st.metric("Price Quotes in DB", len(st.session_state.prices_df))

    st.divider()
    st.markdown("**Download Consolidated Workbook:**")
    st.download_button(
        label="📥 Download Kerry_Coast Excel (.xlsx)",
        data=export_full_excel_workbook(),
        file_name=EXCEL_FILENAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()
    st.markdown("**Irish Standard Measure Guide:**")
    st.markdown("- Irish Standard Shot: **35.5 ml**")
    st.markdown("- Liqueur / Aperitif: **50.0 ml**")
    st.markdown("- Standard Spirit 700ml: **19.72 measures**")
    st.markdown("- Standard Spirit 1000ml: **28.17 measures**")


# ==============================================================================
# MAIN NAVIGATION TABS
# ==============================================================================
st.title("🍸 Bar Spirits Procurement & Price Audit")

tab_master, tab_prices, tab_comparison, tab_invoice, tab_import_export = st.tabs([
    "📋 1. Master Catalog",
    "🏢 2. Supplier Prices",
    "📊 3. Price Comparison",
    "🧾 4. Invoice Check",
    "🔄 5. Import / Export"
])


# ------------------------------------------------------------------------------
# TAB 1: MASTER CATALOG
# ------------------------------------------------------------------------------
with tab_master:
    st.subheader("📋 TouchOffice PLU Master Catalog")
    st.caption("Official bar inventory registry: PLU codes, volume in ml, optic measure size, and measures per bottle.")

    col_s1, col_s2, col_s3 = st.columns([3, 2, 2])
    search_q = col_s1.text_input("🔍 Search PLU or Product Name:", key="m_search")
    categories = ["All Categories"] + sorted(st.session_state.master_df["Category"].dropna().unique().tolist())
    cat_filter = col_s2.selectbox("Filter by Category:", categories, key="m_cat")
    active_only = col_s3.checkbox("Active products only (YES)", value=True, key="m_active")

    df_view = st.session_state.master_df.copy()
    if search_q:
        mask = (
            df_view["PLU"].astype(str).str.contains(search_q, case=False, na=False) |
            df_view["Product Name"].astype(str).str.contains(search_q, case=False, na=False)
        )
        df_view = df_view[mask]
    if cat_filter != "All Categories":
        df_view = df_view[df_view["Category"] == cat_filter]
    if active_only and "Active" in df_view.columns:
        df_view = df_view[df_view["Active"].astype(str).str.upper() == "YES"]

    st.dataframe(
        df_view.style.format({
            "Typical Volume (ml)": "{:,.0f} ml",
            "Measure Size (ml)": "{:,.1f} ml",
            "Measures in Bottle": "{:,.2f}"
        }),
        use_container_width=True,
        hide_index=True
    )

    # Add / Edit Master Item Form
    with st.expander("➕ Add New Item to Master Catalog"):
        with st.form("add_master_item_form"):
            c1, c2 = st.columns(2)
            f_plu = c1.text_input("TouchOffice PLU (Our Code)*", placeholder="e.g. 150")
            f_name = c2.text_input("Product Name (EPOS)*", placeholder="e.g. TULLAMORE DEW 12 YO")
            f_cat = c1.selectbox("Category", list(STANDARD_MEASURES.keys()), index=0)
            f_vol = c2.number_input("Typical Volume (ml)", min_value=50, max_value=5000, value=700, step=50)
            default_measure = STANDARD_MEASURES.get(f_cat, 35.5)
            f_measure = c1.number_input("Measure Size (ml)", min_value=1.0, max_value=500.0, value=float(default_measure), step=0.5)
            f_active = c2.selectbox("Active in Bar", ["YES", "NO"], index=0)

            submitted = st.form_submit_button("Save Item to Catalog", type="primary")
            if submitted:
                clean_plu = f_plu.strip()
                clean_name = f_name.strip()
                if not clean_plu or not clean_name:
                    st.error("Please fill in both PLU and Product Name!")
                elif clean_plu in st.session_state.master_df["PLU"].astype(str).values:
                    st.error(f"PLU '{clean_plu}' already exists in catalog! Each product must have a unique PLU.")
                else:
                    new_item = {
                        "PLU": clean_plu,
                        "Product Name": clean_name,
                        "Category": f_cat,
                        "Typical Volume (ml)": int(f_vol),
                        "Measure Size (ml)": float(f_measure),
                        "Measures in Bottle": round(f_vol / f_measure, 2),
                        "Active": f_active
                    }
                    st.session_state.master_df = pd.concat([st.session_state.master_df, pd.DataFrame([new_item])], ignore_index=True)
                    st.success(f"Product '{clean_name}' (PLU {clean_plu}) successfully added!")
                    st.rerun()


# ------------------------------------------------------------------------------
# TAB 2: SUPPLIER PRICES
# ------------------------------------------------------------------------------
with tab_prices:
    st.subheader("🏢 Supplier Wholesale Prices")
    st.caption("Case prices from distributors (Cliffords C&C, Classic Drinks, Sive) with FOC deal support and optic cost calculations.")

    df_p_enriched = get_enriched_prices()

    col_p1, col_p2 = st.columns([3, 2])
    search_price = col_p1.text_input("🔍 Search quotes by product or supplier:", key="p_search")
    sup_options = ["All Suppliers"] + sorted(df_p_enriched["Supplier"].unique().tolist()) if not df_p_enriched.empty else ["All"]
    sup_filter = col_p2.selectbox("Filter Supplier:", sup_options, key="p_sup")

    df_p_filtered = df_p_enriched.copy()
    if search_price:
        mask = (
            df_p_filtered["Supplier"].astype(str).str.contains(search_price, case=False, na=False) |
            df_p_filtered["Product Name"].astype(str).str.contains(search_price, case=False, na=False) |
            df_p_filtered["Linked PLU"].astype(str).str.contains(search_price, case=False, na=False)
        )
        df_p_filtered = df_p_filtered[mask]
    if sup_filter != "All Suppliers":
        df_p_filtered = df_p_filtered[df_p_filtered["Supplier"] == sup_filter]

    st.dataframe(
        df_p_filtered.style.format({
            "Price per Case (€)": "{:,.2f} €",
            "Effective Bottle Price (€)": "{:,.2f} €",
            "Cost per Measure (€)": "{:,.2f} €",
            "Measures in Bottle": "{:,.2f}"
        }),
        use_container_width=True,
        hide_index=True
    )

    # Add Supplier Quote Form
    with st.expander("➕ Add Supplier Price Quote"):
        with st.form("add_supplier_quote_form"):
            c_a, c_b = st.columns(2)
            plu_items = [f"{r['PLU']} - {r['Product Name']}" for _, r in st.session_state.master_df.iterrows()]
            selected_master = c_a.selectbox("Link to Master PLU*", plu_items)
            sup_name = c_b.selectbox("Supplier*", ["Cliffords C&C", "Classic Drinks", "Sive", "Custom Supplier"])
            if sup_name == "Custom Supplier":
                sup_name = c_b.text_input("Enter Supplier Name", "New Supplier")

            item_title = c_a.text_input("Supplier Description", placeholder="e.g. Jameson 1L x 6 Case")
            case_price = c_b.number_input("Case Price (€)*", min_value=1.0, value=150.0, step=1.0)
            pack_bottles = c_a.number_input("Bottles per Case*", min_value=1, value=6, step=1)
            vol_bottle = c_b.number_input("Bottle Volume (ml)*", min_value=50, value=1000, step=50)

            st.markdown("**Free-of-Charge (FOC) Deal (Optional):**")
            cf1, cf2 = st.columns(2)
            foc_b = cf1.number_input("Buy X Cases", min_value=0, value=0, step=1)
            foc_f = cf2.number_input("Get Y Free Cases", min_value=0, value=0, step=1)

            if st.form_submit_button("Add Quote to Database", type="primary"):
                sel_plu = selected_master.split(" - ")[0]
                new_quote = {
                    "Supplier": sup_name,
                    "Product Name": item_title or selected_master.split(" - ")[1],
                    "Pack Size": int(pack_bottles),
                    "Volume (ml)": int(vol_bottle),
                    "Price per Case (€)": float(case_price),
                    "Linked PLU": sel_plu,
                    "FOC Buy": int(foc_b),
                    "FOC Free": int(foc_f),
                    "In Stock": True
                }
                st.session_state.prices_df = pd.concat([st.session_state.prices_df, pd.DataFrame([new_quote])], ignore_index=True)
                st.success("Price quote successfully saved!")
                st.rerun()


# ------------------------------------------------------------------------------
# TAB 3: PRICE COMPARISON
# ------------------------------------------------------------------------------
with tab_comparison:
    st.subheader("📊 Price Comparison Matrix")
    st.caption("Side-by-side benchmark of all wholesale distributors per PLU. Automatically reveals lowest bottle price, optic cost, and cost savings.")

    df_comp = generate_comparison_df()

    col_c1, col_c2 = st.columns([3, 2])
    comp_search = col_c1.text_input("🔍 Search Comparison:", key="c_search")
    savings_only = col_c2.checkbox("Show items with active price savings (> €0.00)", value=False)

    df_comp_view = df_comp.copy()
    if comp_search:
        df_comp_view = df_comp_view[
            df_comp_view["PLU"].astype(str).str.contains(comp_search, case=False, na=False) |
            df_comp_view["Product Name"].astype(str).str.contains(comp_search, case=False, na=False)
        ]
    if savings_only:
        df_comp_view = df_comp_view[df_comp_view["Savings per Bottle (€)"] > 0]

    # Quick High-Level Metrics
    total_items = len(df_comp_view)
    items_with_savings = len(df_comp_view[df_comp_view["Savings per Bottle (€)"] > 0])
    avg_savings = df_comp_view[df_comp_view["Savings per Bottle (€)"] > 0]["Savings per Bottle (€)"].mean() if items_with_savings > 0 else 0.0

    m1, m2, m3 = st.columns(3)
    m1.metric("Items Monitored", total_items)
    m2.metric("Items with Distributor Gap", items_with_savings)
    m3.metric("Avg Saving per Bottle", f"€{avg_savings:.2f}")

    st.dataframe(
        df_comp_view.style.format({
            "Lowest Bottle (€)": lambda x: f"{x:,.2f} €" if pd.notna(x) else "-",
            "Lowest Measure (€)": lambda x: f"{x:,.2f} €" if pd.notna(x) else "-",
            "Highest Bottle (€)": lambda x: f"{x:,.2f} €" if pd.notna(x) else "-",
            "Savings per Bottle (€)": lambda x: f"+{x:,.2f} €" if pd.notna(x) and x > 0 else "-",
            "Savings (%)": lambda x: f"{x:.1f}%" if pd.notna(x) and x > 0 else "-"
        }),
        use_container_width=True,
        hide_index=True
    )


# ------------------------------------------------------------------------------
# TAB 4: INVOICE CHECK
# ------------------------------------------------------------------------------
with tab_invoice:
    st.subheader("🧾 Delivery Invoice Audit")
    st.caption("Reconcile delivery invoices against contracted price lists upon arrival to detect overcharges and issue credit note claims.")

    inv_c1, inv_c2, inv_c3 = st.columns(3)
    available_sups = sorted(st.session_state.prices_df["Supplier"].unique().tolist()) if not st.session_state.prices_df.empty else ["Cliffords C&C"]
    inv_supplier = inv_c1.selectbox("Supplier on Invoice:", available_sups, key="audit_sup")
    inv_number = inv_c2.text_input("Invoice Number:", value="INV-2026-08892", key="audit_num")
    inv_date = inv_c3.date_input("Delivery Date:", datetime.today())

    df_p_current = get_enriched_prices()

    audit_rows = []
    total_invoiced = 0.0
    total_expected = 0.0
    total_overcharge = 0.0

    for item in st.session_state.invoice_items:
        plu_str = str(item["PLU"])
        qty = int(item["Qty Units"])
        pack = int(item["Bottles in Unit"])
        inv_price = float(item["Invoice Price"])
        
        cost_bottle_inv = round(inv_price / pack, 2) if pack > 0 else inv_price
        line_total_inv = round(qty * inv_price, 2)
        total_invoiced += line_total_inv

        match = df_p_current[
            (df_p_current["Supplier"].str.lower() == inv_supplier.lower()) &
            (df_p_current["Linked PLU"].astype(str) == plu_str)
        ]

        if not match.empty:
            expected_bottle = match["Effective Bottle Price (€)"].values[0]
            expected_line = round(qty * pack * expected_bottle, 2)
            total_expected += expected_line
            variance_bottle = round(cost_bottle_inv - expected_bottle, 2)
            variance_line = round(line_total_inv - expected_line, 2)

            if variance_bottle > 0.01:
                status = "🔴 OVERCHARGED"
                total_overcharge += variance_line
            elif variance_bottle < -0.01:
                status = "🟢 UNDERCHARGED"
            else:
                status = "✅ MATCHED"
        else:
            expected_bottle = None
            variance_bottle = 0.0
            variance_line = 0.0
            status = "⚪ NO CONTRACT PRICE"

        audit_rows.append({
            "PLU": plu_str,
            "Product Name": item["Product Name"],
            "Qty (Units)": qty,
            "Pack Size": f"x{pack}",
            "Invoice Unit (€)": inv_price,
            "Cost / Bottle (Inv) (€)": cost_bottle_inv,
            "Contract / Bottle (€)": expected_bottle if expected_bottle is not None else "-",
            "Variance / Bottle (€)": f"+{variance_bottle:.2f} €" if variance_bottle > 0 else f"{variance_bottle:.2f} €",
            "Total Line Variance (€)": variance_line,
            "Audit Status": status
        })

    # Summary metric tiles
    col_k1, col_k2, col_k3 = st.columns(3)
    col_k1.metric("Total Invoiced", f"€{total_invoiced:,.2f}")
    col_k2.metric("Contract Expected", f"€{total_expected:,.2f}")

    if total_overcharge > 0.01:
        col_k3.error(f"⚠️ OVERCHARGE DETECTED: +€{total_overcharge:,.2f}")
    else:
        col_k3.success("✅ All lines match contracted pricing!")

    st.dataframe(pd.DataFrame(audit_rows), use_container_width=True, hide_index=True)

    # Add line to invoice
    with st.expander("➕ Add Line Item to Invoice"):
        with st.form("add_inv_line_form"):
            ca, cb, cc, cd = st.columns(4)
            catalog_choices = [f"{r['PLU']} - {r['Product Name']}" for _, r in st.session_state.master_df.iterrows()]
            chosen_product = ca.selectbox("Item", catalog_choices)
            qty_in = cb.number_input("Qty Units", min_value=1, value=1)
            pack_in = cc.selectbox("Pack Size", [6, 12, 1, 24])
            inv_p = cd.number_input("Invoice Price per Unit (€)", min_value=1.0, value=120.0, step=1.0)

            if st.form_submit_button("Add Item to Invoice", type="primary"):
                st.session_state.invoice_items.append({
                    "PLU": chosen_product.split(" - ")[0],
                    "Product Name": chosen_product.split(" - ")[1],
                    "Qty Units": int(qty_in),
                    "Bottles in Unit": int(pack_in),
                    "Invoice Price": float(inv_p)
                })
                st.rerun()

    # Credit Note Generator
    if total_overcharge > 0.01:
        st.markdown("### ✉️ One-Click Credit Note Claim (Email Copy)")
        claim_body = f"""Subject: Kerry Coast Hotel - Invoice Discrepancy & Credit Note Request ({inv_number})

Dear {inv_supplier} Accounts Team,

Upon checking delivery invoice #{inv_number} dated {inv_date}, we identified price discrepancies exceeding our agreed wholesale rates.

Total Invoiced: €{total_invoiced:.2f}
Contract Expected: €{total_expected:.2f}
Total Overcharge: €{total_overcharge:.2f}

Please arrange an immediate Credit Note for the difference of €{total_overcharge:.2f}.

Kind regards,
Bar Management
Kerry Coast Hotel"""
        st.code(claim_body, language="text")


# ------------------------------------------------------------------------------
# TAB 5: IMPORT / EXPORT & SYNC
# ------------------------------------------------------------------------------
with tab_import_export:
    st.subheader("🔄 Excel Import, Export & Backup")
    st.caption("Directly synchronise data with Excel files and download consolidated reports.")

    col_io1, col_io2 = st.columns(2)

    with col_io1:
        st.markdown("### 📤 Export Data to Excel")
        st.write("Download the complete 4-sheet workbook with current Master Catalog, Supplier Prices, Comparison, and Invoices:")
        st.download_button(
            label="📥 Download Kerry_Coast_Spirits_Master_Catalog.xlsx",
            data=export_full_excel_workbook(),
            file_name=EXCEL_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    with col_io2:
        st.markdown("### 📥 Import Custom Excel")
        uploaded_file = st.file_uploader("Upload an Excel file (.xlsx):", type=["xlsx", "xls"])
        if uploaded_file is not None:
            if st.button("Apply Uploaded File", type="secondary", use_container_width=True):
                try:
                    df_up_m = pd.read_excel(uploaded_file, sheet_name=SHEET_MASTER)
                    if "Typical Volume (ml)" in df_up_m.columns and "Measure Size (ml)" in df_up_m.columns:
                        df_up_m["Measures in Bottle"] = (df_up_m["Typical Volume (ml)"] / df_up_m["Measure Size (ml)"]).round(2)
                    st.session_state.master_df = df_up_m

                    try:
                        df_up_p = pd.read_excel(uploaded_file, sheet_name=SHEET_PRICES)
                        st.session_state.prices_df = df_up_p
                    except Exception:
                        pass

                    st.success("Workbook imported successfully into application session!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error reading file: {e}")

    st.divider()
    if st.button("🔄 Reset to Default Kerry Coast Hotel Data"):
        st.session_state.master_df = pd.DataFrame(DEFAULT_MASTER_DATA)
        st.session_state.master_df["Measures in Bottle"] = (
            st.session_state.master_df["Typical Volume (ml)"] / st.session_state.master_df["Measure Size (ml)"]
        ).round(2)
        st.session_state.prices_df = pd.DataFrame(DEFAULT_PRICES_DATA)
        st.success("Database restored to Kerry Coast defaults!")
        st.rerun()
