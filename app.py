"""
Kerry Coast Hotel - Spirits Master Catalog & Procurement Intelligence System
=============================================================================
Complete standalone Python/Streamlit web application for bar stock & price audit.
Fully compatible with Streamlit Cloud & local execution.

Includes:
1. Master Catalog (Complete 76 PLU items directly from official TouchOffice POS database)
2. Supplier Prices (All 80 wholesale distributor price quotes from Classic Drinks, Cliffords C&C, Sive)
3. Price Comparison (Live multi-supplier benchmark, lowest bottle/measure cost, savings)
4. Invoice Check (Delivery invoice audit, contract price variance, overcharge claim generator)
5. Import / Export (Consolidated 4-sheet Excel Kerry_Coast_Spirits_Master_Catalog.xlsx)
"""

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

# Irish optics & bar serving standards (ml)
STANDARD_MEASURES = {
    "Spirits": 35.5,
    "Liqueurs & Aperitifs": 50.0,
    "Wine & Sparkling": 175.0,
    "Beer & Cider": 568.0,
    "Soft Drinks & Mixers": 200.0,
    "Syrups & Cordials": 25.0
}

# ==============================================================================
# COMPLETE 76-ITEM MASTER CATALOG & 80 SUPPLIER PRICE QUOTES
# ==============================================================================
DEFAULT_MASTER_DATA = [
    {"PLU":"121","Product Name":"PADDY","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"122","Product Name":"POWER / POWERS GOLD LABEL","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"123","Product Name":"JAMESON","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"124","Product Name":"JAMESON BLACK BARREL","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"125","Product Name":"BUSHMILLS RED / ORIGINAL","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"126","Product Name":"BLACK BUSH","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"127","Product Name":"GREENSPOT","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"128","Product Name":"YELLOW SPOT","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"129","Product Name":"RED SPOT","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"131","Product Name":"MIDDLETON","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"132","Product Name":"RED BREAST 12YR","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"133","Product Name":"RED BREAST 15YR","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"134","Product Name":"RED BREAST 21YR","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"135","Product Name":"SKELLIG WHISKEY","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"136","Product Name":"JAMESON BLACK BARREL (Alt PLU)","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"161","Product Name":"GORDONS GIN","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"162","Product Name":"GORDONS PINK GIN","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"163","Product Name":"DINGLE PINK GIN","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"164","Product Name":"DINGLE GIN","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"165","Product Name":"SKELLIG GIN","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"166","Product Name":"HENDRICKS","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"167","Product Name":"PORTMAGEE GIN","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"168","Product Name":"BOMBAY SAPPHIRE","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"169","Product Name":"GORDONS 00 GIN","Category":"Soft Drinks & Mixers","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"190","Product Name":"DINGLE VODKA","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"191","Product Name":"SMIRNOFF VODKA","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"192","Product Name":"ABSOLUT","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"193","Product Name":"GREY GOOSE","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"194","Product Name":"TITOS","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"195","Product Name":"VODKA MONSTER","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"211","Product Name":"HENNESSY","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"212","Product Name":"HENNESSY & PORT","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"213","Product Name":"REMY MARTIN","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"221","Product Name":"BACARDI","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"223","Product Name":"MALIBU","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"224","Product Name":"CAPTAIN MORGAN","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"231","Product Name":"JACK DANIELS","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"234","Product Name":"CANADIAN CLUB","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"235","Product Name":"SOUTHERN COMFORT","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"236","Product Name":"WOODFORD RESERVE","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"252","Product Name":"TEACHERS","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"253","Product Name":"JOHNNY WALKER RED","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"254","Product Name":"JOHNNY WALKER BLACK","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"261","Product Name":"JOSE CUERVO BLANCO","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"262","Product Name":"JOSE CUERVO REPOSADO","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"263","Product Name":"CASAMIGOS BLANCO","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"264","Product Name":"CASAMIGOS REPOSADO","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"272","Product Name":"TEQUILA (generic)","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"271","Product Name":"JAGERMEISTER","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"273","Product Name":"TEQUILA ROSE","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"274","Product Name":"SAMBUCA","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"275","Product Name":"BABY GUINNESS","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"276","Product Name":"AFTER SHOCK","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"277","Product Name":"APPLE SOURZ","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"289","Product Name":"CREME DE MENTHE","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"291","Product Name":"BAILEYS","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"292","Product Name":"TIA MARIA","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"293","Product Name":"KAHLUA","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"294","Product Name":"CAMPARI","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"295","Product Name":"PEACH SCHNAPPS","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"298","Product Name":"COINTREAU","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"299","Product Name":"GRAND MARNIER","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"300","Product Name":"DISARONNO","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"1502","Product Name":"TRIPLE SEC","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"1503","Product Name":"AMARETTO","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"1508","Product Name":"APEROL","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"296","Product Name":"MARTINI","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":750,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"297","Product Name":"MARTINI ROSSO","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":750,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"301","Product Name":"VALENTIA VERMOUTH","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"305","Product Name":"PORT","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":750,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"306","Product Name":"SHERRY","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":750,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"1504","Product Name":"VALENTIA VERMOUTH (stock)","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":700,"Measure Size (ml)":50,"Active":"YES"},
    {"PLU":"1501","Product Name":"VANILLA VODKA","Category":"Spirits","Typical Volume (ml)":700,"Measure Size (ml)":35.5,"Active":"YES"},
    {"PLU":"1505","Product Name":"GRENADINE","Category":"Syrups & Cordials","Typical Volume (ml)":700,"Measure Size (ml)":10,"Active":"YES"},
    {"PLU":"1506","Product Name":"ORANGE BITTER","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":200,"Measure Size (ml)":1,"Active":"YES"},
    {"PLU":"1507","Product Name":"BITTER","Category":"Liqueurs & Aperitifs","Typical Volume (ml)":200,"Measure Size (ml)":1,"Active":"YES"}
]

DEFAULT_PRICES_DATA = [
    {"Supplier":"Classic Drinks","Product Name":"BACARDI","Pack Size":6,"Volume (ml)":1000,"Price per Case (€)":151,"Linked PLU":"221","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"BAILEYS","Pack Size":12,"Volume (ml)":1000,"Price per Case (€)":201,"Linked PLU":"291","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"CAPTAIN MORGAN","Pack Size":12,"Volume (ml)":1000,"Price per Case (€)":282,"Linked PLU":"224","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"COINTREAU BOLS TRIPLE SEC","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":145,"Linked PLU":"298","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"DINGLE GIN","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":146,"Linked PLU":"164","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"DINGLE VODKA","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":123,"Linked PLU":"190","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"GORDONS GIN","Pack Size":12,"Volume (ml)":1000,"Price per Case (€)":298,"Linked PLU":"161","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"HENDRICKS","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":183,"Linked PLU":"166","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"HENNESSY","Pack Size":12,"Volume (ml)":700,"Price per Case (€)":359,"Linked PLU":"211","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"JACK DANIELS","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":138,"Linked PLU":"231","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"JAGERMEISTER","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":101,"Linked PLU":"271","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"MALIBU","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":79,"Linked PLU":"223","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"PADDY","Pack Size":6,"Volume (ml)":1000,"Price per Case (€)":142,"Linked PLU":"121","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"SMIRNOFF Vodka","Pack Size":12,"Volume (ml)":1000,"Price per Case (€)":262,"Linked PLU":"191","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"SOUTHERN COMFORT","Pack Size":12,"Volume (ml)":700,"Price per Case (€)":182,"Linked PLU":"235","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"TEQUILA","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":109,"Linked PLU":"272","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"TIA MARIA","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":92,"Linked PLU":"292","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Classic Drinks","Product Name":"Tequila Rose","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":94.5,"Linked PLU":"273","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"ABSOLUT VODKA","Pack Size":12,"Volume (ml)":1000,"Price per Case (€)":270,"Linked PLU":"192","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"ABSOLUT VODKA VANILLA","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":119,"Linked PLU":"1501","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"AFTERSHOCK RED","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":114.95,"Linked PLU":"276","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"APEROL","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":74.95,"Linked PLU":"1508","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"BACARDI","Pack Size":6,"Volume (ml)":1000,"Price per Case (€)":162.95,"Linked PLU":"221","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"BAILEYS","Pack Size":12,"Volume (ml)":1000,"Price per Case (€)":220,"Linked PLU":"291","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"BLACK BUSH","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":134.95,"Linked PLU":"126","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"BOMBAY SAPPHIRE","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":134.95,"Linked PLU":"168","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"BUSHMILLS WHITE","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":121.95,"Linked PLU":"125","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"CAMPARI","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":105,"Linked PLU":"294","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"CANADIAN CLUB","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":129.95,"Linked PLU":"234","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"CAPTAIN MORGAN SPICED","Pack Size":6,"Volume (ml)":1000,"Price per Case (€)":159.95,"Linked PLU":"224","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"CASAMIGOS BLANCO","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":279,"Linked PLU":"263","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"CASAMIGOS REPOSADO","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":295,"Linked PLU":"264","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"COINTREAU","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":155,"Linked PLU":"298","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"CR DE MENTHE GREEN BOLS","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":99,"Linked PLU":"289","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"DINGLE GIN","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":149.95,"Linked PLU":"164","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"DINGLE VODKA","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":126,"Linked PLU":"190","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"DISARONNO AMARETTO","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":129.95,"Linked PLU":"300","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"GORDONS GIN","Pack Size":12,"Volume (ml)":1000,"Price per Case (€)":285.95,"Linked PLU":"161","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"GORDONS PINK GIN","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":109.95,"Linked PLU":"162","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"GRAND MARNIER","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":154.95,"Linked PLU":"299","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"GREEN SPOT","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":219,"Linked PLU":"127","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"GRENADINE SYRUP MONIN","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":42.5,"Linked PLU":"1505","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"GREY GOOSE VODKA","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":199,"Linked PLU":"193","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"HARVEYS BRISTOL CREAM","Pack Size":6,"Volume (ml)":750,"Price per Case (€)":79.95,"Linked PLU":"306","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"HENDRICKS GIN","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":185,"Linked PLU":"166","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"HENNESSY BRANDY","Pack Size":12,"Volume (ml)":700,"Price per Case (€)":339,"Linked PLU":"211","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"JACK DANIELS","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":146.95,"Linked PLU":"231","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"JAGERMEISTER","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":106.95,"Linked PLU":"271","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"JAMESON","Pack Size":6,"Volume (ml)":1000,"Price per Case (€)":174.95,"Linked PLU":"123","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"JAMESON BLACK BARREL","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":205,"Linked PLU":"124","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"JOHNNIE WALKER BLACK","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":174.95,"Linked PLU":"254","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"JOHNNIE WALKER RED","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":119.95,"Linked PLU":"253","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"JOSE CUERVO SILVER","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":129.95,"Linked PLU":"261","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"JOSE CUERVO GOLD","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":129.95,"Linked PLU":"262","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"KAHLUA","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":109,"Linked PLU":"293","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"MALIBU","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":84.95,"Linked PLU":"223","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"MARTINI EXTRA DRY","Pack Size":6,"Volume (ml)":750,"Price per Case (€)":64.95,"Linked PLU":"296","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"MARTINI ROSSO","Pack Size":6,"Volume (ml)":750,"Price per Case (€)":64.95,"Linked PLU":"297","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"MIDLETON VERY RARE","Pack Size":1,"Volume (ml)":700,"Price per Case (€)":199,"Linked PLU":"131","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"PADDY","Pack Size":6,"Volume (ml)":1000,"Price per Case (€)":149.95,"Linked PLU":"121","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"PEACH SCHNAPPS ARCHERS","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":94.95,"Linked PLU":"295","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"POWERS","Pack Size":6,"Volume (ml)":1000,"Price per Case (€)":174.95,"Linked PLU":"122","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"REDBREAST 12 YO","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":269,"Linked PLU":"132","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"REDBREAST 15 YO","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":429,"Linked PLU":"133","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"REMY MARTIN VSOP","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":249,"Linked PLU":"213","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"SAMBUCA RAMAZZOTTI","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":109.95,"Linked PLU":"274","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"SANDEMAN RUBY PORT","Pack Size":6,"Volume (ml)":750,"Price per Case (€)":74.95,"Linked PLU":"305","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"SMIRNOFF RED","Pack Size":12,"Volume (ml)":1000,"Price per Case (€)":258.9,"Linked PLU":"191","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"SOUTHERN COMFORT","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":104.95,"Linked PLU":"235","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"SOURZ APPLE","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":69.95,"Linked PLU":"277","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"TEACHERS","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":114.95,"Linked PLU":"252","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"TEQUILA ROSE","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":99,"Linked PLU":"273","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"TIA MARIA","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":99,"Linked PLU":"292","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"TITOS HANDMADE VODKA","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":139,"Linked PLU":"194","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"WOODFORD RESERVE","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":179,"Linked PLU":"236","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Cliffords C&C","Product Name":"YELLOW SPOT","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":339,"Linked PLU":"128","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Sive","Product Name":"AFTERSHOCK RED 70CL","Pack Size":6,"Volume (ml)":700,"Price per Case (€)":122,"Linked PLU":"276","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Sive","Product Name":"BACARDI 1LTR","Pack Size":6,"Volume (ml)":1000,"Price per Case (€)":149.7,"Linked PLU":"221","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Sive","Product Name":"GORDONS GIN 1LTR","Pack Size":12,"Volume (ml)":1000,"Price per Case (€)":318,"Linked PLU":"161","FOC Buy":0,"FOC Free":0,"In Stock":True},
    {"Supplier":"Sive","Product Name":"JAMESON 1LTR","Pack Size":6,"Volume (ml)":1000,"Price per Case (€)":193.5,"Linked PLU":"123","FOC Buy":0,"FOC Free":0,"In Stock":True}
]

# ==============================================================================
# STATE INITIALIZATION
# ==============================================================================
def init_session_state():
    """Initializes app state with the full 76 Master Items and all Supplier Quotes."""
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
    """Builds side-by-side comparison matrix for all 76 PLUs across suppliers."""
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
    st.markdown("- Liqueurs & Aperitifs: **50.0 ml**")
    st.markdown("- Soft Drinks & Mixers: **200.0 ml**")
    st.markdown("- Syrups & Cordials: **25.0 ml**")
    st.markdown("- 700ml Spirit Bottle: **19.72 measures**")
    st.markdown("- 1000ml Spirit Bottle: **28.17 measures**")


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
    st.subheader(f"📋 TouchOffice PLU Master Catalog ({len(st.session_state.master_df)} items)")
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

    c_crud1, c_crud2 = st.columns(2)

    # 1. Add New Master Item
    with c_crud1:
        with st.expander("➕ Add New Item to Master Catalog"):
            with st.form("add_master_item_form"):
                f_plu = st.text_input("TouchOffice PLU (Till Code)*", placeholder="e.g. 501")
                f_name = st.text_input("Product Name (POS Description)*", placeholder="e.g. DINGLE WHISKEY")
                f_cat = st.selectbox("Category", list(STANDARD_MEASURES.keys()), index=0)
                f_vol = st.number_input("Typical Volume (ml)", min_value=50, max_value=5000, value=700, step=50)
                default_measure = STANDARD_MEASURES.get(f_cat, 35.5)
                f_measure = st.number_input("Measure Size (ml)", min_value=1.0, max_value=500.0, value=float(default_measure), step=0.5)
                f_active = st.selectbox("Active in Bar", ["YES", "NO"], index=0)

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

    # 2. Edit or Delete Existing Item
    with c_crud2:
        with st.expander("✏️ Edit / Delete Existing Item"):
            edit_options = [f"{r['PLU']} - {r['Product Name']}" for _, r in st.session_state.master_df.iterrows()]
            selected_to_edit = st.selectbox("Select Product to Modify:", edit_options, key="select_edit_m")
            if selected_to_edit:
                sel_plu = selected_to_edit.split(" - ")[0]
                current_row = st.session_state.master_df[st.session_state.master_df["PLU"].astype(str) == sel_plu].iloc[0]

                with st.form("edit_master_form"):
                    e_name = st.text_input("Product Name", value=current_row["Product Name"])
                    e_cat = st.selectbox(
                        "Category",
                        list(STANDARD_MEASURES.keys()),
                        index=list(STANDARD_MEASURES.keys()).index(current_row["Category"]) if current_row["Category"] in STANDARD_MEASURES else 0
                    )
                    e_vol = st.number_input("Typical Volume (ml)", min_value=50, max_value=5000, value=int(current_row["Typical Volume (ml)"]), step=50)
                    e_measure = st.number_input("Measure Size (ml)", min_value=1.0, max_value=500.0, value=float(current_row["Measure Size (ml)"]), step=0.5)
                    e_active = st.selectbox("Active", ["YES", "NO"], index=0 if current_row["Active"] == "YES" else 1)

                    col_b1, col_b2 = st.columns(2)
                    save_edit = col_b1.form_submit_button("Update Item", type="primary")
                    del_edit = col_b2.form_submit_button("Delete Item", type="secondary")

                    if save_edit:
                        idx = st.session_state.master_df[st.session_state.master_df["PLU"].astype(str) == sel_plu].index[0]
                        st.session_state.master_df.at[idx, "Product Name"] = e_name.strip()
                        st.session_state.master_df.at[idx, "Category"] = e_cat
                        st.session_state.master_df.at[idx, "Typical Volume (ml)"] = int(e_vol)
                        st.session_state.master_df.at[idx, "Measure Size (ml)"] = float(e_measure)
                        st.session_state.master_df.at[idx, "Measures in Bottle"] = round(e_vol / e_measure, 2)
                        st.session_state.master_df.at[idx, "Active"] = e_active
                        st.success(f"PLU {sel_plu} updated successfully!")
                        st.rerun()

                    if del_edit:
                        st.session_state.master_df = st.session_state.master_df[st.session_state.master_df["PLU"].astype(str) != sel_plu].reset_index(drop=True)
                        st.success(f"PLU {sel_plu} removed from catalog!")
                        st.rerun()


# ------------------------------------------------------------------------------
# TAB 2: SUPPLIER PRICES
# ------------------------------------------------------------------------------
with tab_prices:
    st.subheader(f"🏢 Wholesale Supplier Prices ({len(st.session_state.prices_df)} quotes)")
    st.caption("Case prices from distributors (Classic Drinks, Cliffords C&C, Sive) with FOC deal support and optic cost calculations.")

    df_p_enriched = get_enriched_prices()

    col_p1, col_p2 = st.columns([3, 2])
    search_price = col_p1.text_input("🔍 Search quotes by product, supplier, or PLU:", key="p_search")
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

    editable_cols = ["Price per Case (€)", "FOC Buy", "FOC Free"]
    disabled_cols = [col for col in df_p_filtered.columns if col not in editable_cols]

    edited_df = st.data_editor(
        df_p_filtered,
        key="supplier_prices_editor",
        disabled=disabled_cols,
        column_config={
            "Price per Case (€)": st.column_config.NumberColumn(
                "Price per Case (€)",
                help="Wholesale case price in Euros. Click to edit.",
                min_value=0.0,
                step=0.5,
                format="%.2f €"
            ),
            "FOC Buy": st.column_config.NumberColumn(
                "FOC Buy",
                help="Buy X Cases (set to 0 to disable promo)",
                min_value=0,
                step=1,
                format="%d"
            ),
            "FOC Free": st.column_config.NumberColumn(
                "FOC Free",
                help="Get Y Free Cases (set to 0 to disable promo)",
                min_value=0,
                step=1,
                format="%d"
            ),
            "Effective Bottle Price (€)": st.column_config.NumberColumn(
                "Effective Bottle Price (€)",
                help="Calculated bottle cost including FOC promo deals",
                format="%.2f €"
            ),
            "Cost per Measure (€)": st.column_config.NumberColumn(
                "Cost per Measure (€)",
                help="Optic shot cost calculated from bottle price & measure size",
                format="%.2f €"
            ),
            "Measures in Bottle": st.column_config.NumberColumn(
                "Measures in Bottle",
                format="%.2f"
            ),
        },
        use_container_width=True,
        hide_index=True
    )

    # Propagate edits back to st.session_state.prices_df
    has_changes = False
    if edited_df is not None:
        for idx in edited_df.index:
            if idx in st.session_state.prices_df.index:
                for col in editable_cols:
                    new_val = edited_df.at[idx, col]
                    curr_val = st.session_state.prices_df.at[idx, col]
                    if col == "Price per Case (€)":
                        new_val = round(float(new_val), 2) if pd.notna(new_val) else 0.0
                        curr_val = round(float(curr_val), 2) if pd.notna(curr_val) else 0.0
                    else:
                        new_val = int(new_val) if pd.notna(new_val) else 0
                        curr_val = int(curr_val) if pd.notna(curr_val) else 0

                    if new_val != curr_val:
                        st.session_state.prices_df.at[idx, col] = new_val
                        has_changes = True

    if has_changes:
        st.rerun()

    cp_crud1, cp_crud2 = st.columns(2)

    # Add Supplier Quote Form
    with cp_crud1:
        with st.expander("➕ Add Supplier Price Quote"):
            with st.form("add_supplier_quote_form"):
                plu_items = [f"{r['PLU']} - {r['Product Name']}" for _, r in st.session_state.master_df.iterrows()]
                selected_master = st.selectbox("Link to Master PLU*", plu_items)
                sup_name = st.selectbox("Supplier*", ["Classic Drinks", "Cliffords C&C", "Sive", "Other Supplier"])
                if sup_name == "Other Supplier":
                    sup_name = st.text_input("Enter Supplier Name", "New Supplier")

                item_title = st.text_input("Supplier Description", placeholder="e.g. Jameson 1L x 6 Case")
                case_price = st.number_input("Case Price (€)*", min_value=1.0, value=150.0, step=1.0)
                pack_bottles = st.number_input("Bottles per Case*", min_value=1, value=6, step=1)
                vol_bottle = st.number_input("Bottle Volume (ml)*", min_value=50, value=1000, step=50)

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

    # Delete Quote Form
    with cp_crud2:
        with st.expander("🗑️ Delete Price Quote"):
            if not st.session_state.prices_df.empty:
                quote_opts = [
                    f"#{i} | {r['Supplier']} - {r['Product Name']} (PLU {r['Linked PLU']}) - €{r['Price per Case (€)']:.2f}"
                    for i, r in st.session_state.prices_df.iterrows()
                ]
                sel_quote = st.selectbox("Select Quote to Remove:", quote_opts)
                if st.button("Delete Selected Quote", type="secondary"):
                    idx_to_del = int(sel_quote.split(" | ")[0].replace("#", ""))
                    st.session_state.prices_df = st.session_state.prices_df.drop(idx_to_del).reset_index(drop=True)
                    st.success("Quote removed!")
                    st.rerun()


# ------------------------------------------------------------------------------
# TAB 3: PRICE COMPARISON
# ------------------------------------------------------------------------------
with tab_comparison:
    st.subheader(f"📊 Price Comparison Matrix ({len(st.session_state.master_df)} products)")
    st.caption("Side-by-side benchmark of wholesale distributors per PLU. Automatically identifies lowest bottle price, measure cost, and potential savings.")

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

    # Metrics
    total_items = len(df_comp_view)
    items_with_savings = len(df_comp_view[df_comp_view["Savings per Bottle (€)"] > 0])
    avg_savings = df_comp_view[df_comp_view["Savings per Bottle (€)"] > 0]["Savings per Bottle (€)"].mean() if items_with_savings > 0 else 0.0

    m1, m2, m3 = st.columns(3)
    m1.metric("Products Monitored", total_items)
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
    available_sups = sorted(st.session_state.prices_df["Supplier"].unique().tolist()) if not st.session_state.prices_df.empty else ["Classic Drinks"]
    inv_supplier = inv_c1.selectbox("Supplier on Invoice:", available_sups, key="audit_sup")
    inv_number = inv_c2.text_input("Invoice Number:", value="INV-2026-08892", key="audit_num")
    inv_date = inv_c3.date_input("Delivery Date:", datetime.today())

    df_p_current = get_enriched_prices()

    audit_rows = []
    total_invoiced = 0.0
    total_expected = 0.0

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
            elif variance_bottle < -0.01:
                status = "🟢 UNDERCHARGED"
            else:
                status = "✅ MATCHED"
        else:
            expected_bottle = None
            total_expected += line_total_inv
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

    total_invoiced = round(total_invoiced, 2)
    total_expected = round(total_expected, 2)
    total_overcharge = round(total_invoiced - total_expected, 2)

    col_k1, col_k2, col_k3 = st.columns(3)
    col_k1.metric("Total Invoiced", f"€{total_invoiced:,.2f}")
    col_k2.metric("Contract Expected", f"€{total_expected:,.2f}")

    if total_overcharge > 0.01:
        col_k3.error(f"⚠️ OVERCHARGE DETECTED: +€{total_overcharge:,.2f}")
    else:
        col_k3.success("✅ All lines match contracted pricing!")

    if audit_rows:
        df_audit = pd.DataFrame(audit_rows)
    else:
        df_audit = pd.DataFrame(columns=[
            "PLU", "Product Name", "Qty (Units)", "Pack Size", "Invoice Unit (€)",
            "Cost / Bottle (Inv) (€)", "Contract / Bottle (€)", "Variance / Bottle (€)",
            "Total Line Variance (€)", "Audit Status"
        ])
    st.dataframe(df_audit, use_container_width=True, hide_index=True)

    if not st.session_state.invoice_items:
        st.info("ℹ️ Invoice has no line items. Use the form below to add drinks from the Master Catalog.")

    col_inv_a, col_inv_b = st.columns(2)

    with col_inv_a:
        with st.expander("➕ Add Line Item to Invoice", expanded=False):
            with st.form("add_inv_line_form"):
                ca, cb = st.columns(2)
                catalog_choices = [f"{r['PLU']} - {r['Product Name']}" for _, r in st.session_state.master_df.iterrows()]
                chosen_product = ca.selectbox("Select Product (Master Catalog)*", catalog_choices)
                qty_in = cb.number_input("Qty Units (Cases / Packs)*", min_value=1, value=1, step=1)

                cc, cd = st.columns(2)
                pack_in = cc.selectbox("Pack Size (Bottles in Unit)*", [6, 12, 1, 24])
                inv_p = cd.number_input("Invoice Price per Unit (€)*", min_value=0.5, value=120.0, step=1.0)

                if st.form_submit_button("Add Item to Invoice", type="primary", use_container_width=True):
                    st.session_state.invoice_items.append({
                        "PLU": chosen_product.split(" - ")[0],
                        "Product Name": chosen_product.split(" - ")[1],
                        "Qty Units": int(qty_in),
                        "Bottles in Unit": int(pack_in),
                        "Invoice Price": float(inv_p)
                    })
                    st.rerun()

    with col_inv_b:
        with st.expander("🗑️ Remove Lines / Clear Invoice", expanded=False):
            if st.session_state.invoice_items:
                line_choices = [
                    f"Line #{i + 1}: {item['Product Name']} (Qty: {item['Qty Units']}, Pack: x{item['Bottles in Unit']}, €{item['Invoice Price']:.2f})"
                    for i, item in enumerate(st.session_state.invoice_items)
                ]
                selected_line = st.selectbox("Select Line to Remove:", line_choices, key="sel_inv_line_rem")

                col_btn_rem, col_btn_clear = st.columns(2)
                if col_btn_rem.button("❌ Remove Selected Line", type="secondary", use_container_width=True):
                    line_idx = int(selected_line.split(":")[0].replace("Line #", "").strip()) - 1
                    if 0 <= line_idx < len(st.session_state.invoice_items):
                        popped_item = st.session_state.invoice_items.pop(line_idx)
                        st.success(f"Removed '{popped_item['Product Name']}' from invoice.")
                        st.rerun()

                if col_btn_clear.button("🗑️ Clear Entire Invoice", type="secondary", use_container_width=True):
                    st.session_state.invoice_items = []
                    st.success("All items cleared from invoice.")
                    st.rerun()
            else:
                st.caption("Invoice is currently empty — no lines to remove.")

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
        st.write(f"Download the complete 4-sheet workbook with current {len(st.session_state.master_df)} Master Items, {len(st.session_state.prices_df)} Supplier Prices, Comparison, and Invoices:")
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
    if st.button("🔄 Reset to Default 76-Item Kerry Coast Data"):
        st.session_state.master_df = pd.DataFrame(DEFAULT_MASTER_DATA)
        st.session_state.master_df["Measures in Bottle"] = (
            st.session_state.master_df["Typical Volume (ml)"] / st.session_state.master_df["Measure Size (ml)"]
        ).round(2)
        st.session_state.prices_df = pd.DataFrame(DEFAULT_PRICES_DATA)
        st.success(f"Database restored to Kerry Coast defaults ({len(DEFAULT_MASTER_DATA)} items)!")
        st.rerun()
