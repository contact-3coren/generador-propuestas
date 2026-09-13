import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date
import re
import os
import time
import tempfile
import io
import platform
import subprocess
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# 0. GESTIÓN DE SECRETOS Y SEGURIDAD
# ==========================================
def get_secret(key, default=""):
    try:
        return st.secrets[key]
    except:
        return default

PASSWORD_APP = get_secret("password_app", "3corenFachadas26")
ID_TEMPLATE = "1SLDdXgTWO4znUfRfUQaR5mpkqNHgs_oY"
ID_CARPETA_DIAGRAMAS = "1TeBilqGN01VHYF_hiIjCbnblvy1Pj5QS"

# ==========================================
# 1. CONFIGURACIÓN Y DISEÑO (CSS)
# ==========================================
st.set_page_config(page_title="Proposal Generator | 3COREN", layout="wide")

# --- BLOQUE DE CONTRASEÑA ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center; color: #0A1145;'>🔒 3COREN - Acceso Restringido</h2>", unsafe_allow_html=True)
    col_pwd1, col_pwd2, col_pwd3 = st.columns([3, 4, 3])
    with col_pwd2:
        pwd = st.text_input("Ingrese la contraseña:", type="password")
        if st.button("Entrar", type="primary", use_container_width=True):
            if pwd == PASSWORD_APP:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta.")
    st.stop()
# --- FIN BLOQUE DE CONTRASEÑA ---

st.markdown("""
    <style>
        .block-container { padding-top: 1rem; padding-bottom: 1rem; max-width: 95%; }
        div[data-testid="column"] { padding: 0px 5px; }
        .stTextInput label p, .stSelectbox label p, .stDateInput label p, .stTextArea label p { font-size: 14px !important; font-weight: 600 !important; color: #0A1145; }
        input, .stDateInput, textarea, .stSelectbox div[data-baseweb="select"] { font-size: 14px !important; }
        .bid-box { border-left: 5px solid #0A1145; padding: 15px; background-color: #f1f3f8; margin-bottom: 15px; border-radius: 0px 5px 5px 0px;}
        .child-box { border-left: 5px solid #FF3300; padding: 10px; margin-left: 40px; background-color: #ffffff; margin-bottom: 5px; border-radius: 0px 5px 5px 0px;}
        .mat-box { border-left: 5px solid #008000; padding: 15px; background-color: #f4fbf4; margin-bottom: 15px; border-radius: 0px 5px 5px 0px;}
        .scope-box { border-left: 3px solid #005500; padding: 10px; margin-left: 20px; background-color: #ffffff; margin-bottom: 5px;}
        .stCheckbox label p { font-size: 14px !important; font-weight: normal !important; color: #333; }
        .stCheckbox label { padding-right: 0px !important; margin-right: 0px !important; }
        hr { margin: 1.5em 0; border-top: 2px solid #ddd; }
        img { object-fit: contain !important; max-height: 80px !important; padding-top: 10px; }
        .stButton > button[kind="primary"] { background-color: #0056b3 !important; color: white !important; border-color: #0056b3 !important; }
        .stButton > button[kind="primary"]:hover { background-color: #004494 !important; border-color: #004494 !important; }
        .fluid-title-box { background-color: #e9ecef; padding: 10px; border-radius: 5px; font-weight: bold; color: #0A1145; margin-bottom: 15px; border: 1px solid #ccc;}
        .info-box-blue { background-color: #d9edf7; color: #31708f; padding: 10px; border-radius: 5px; margin-bottom: 15px; border: 1px solid #bce8f1; font-size: 15px;}
    </style>
""", unsafe_allow_html=True)

col_titulo, col_logo = st.columns([8, 2])
with col_titulo: st.title("Proposal Generator")
with col_logo: 
    try:
        st.image("3coren Logo.png")
    except:
        pass # Evita error si el logo no está en la nube aún

# ==========================================
# 2. CONEXIÓN A GOOGLE SHEETS Y DRIVE
# ==========================================
def get_credentials():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp_service_account" in st.secrets:
            return Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    except:
        pass
    return Credentials.from_service_account_file("credenciales.json", scopes=scopes)

def get_drive_service():
    creds = get_credentials()
    return build('drive', 'v3', credentials=creds)

@st.cache_resource
def conectar_sheets(retries=3):
    for attempt in range(retries):
        try:
            creds = get_credentials()
            gc = gspread.authorize(creds)
            
            ws_maestra = gc.open_by_url("https://docs.google.com/spreadsheets/d/1xepmy2zalJNVYdI24vMFDi-k6FkOuXG_CXY0e592WmI/edit").worksheet("CONSOLIDADO")
            wb_propuestas = gc.open_by_url("https://docs.google.com/spreadsheets/d/1pKwuuH64uC34kiAb0QIIIxf2j2VEWC7MWCWRXf6rbis/edit")
            
            try: ws_cond = wb_propuestas.worksheet("Condiciones")
            except: ws_cond = wb_propuestas.add_worksheet(title="Condiciones", rows="100", cols="10")
            
            return (ws_maestra, wb_propuestas.worksheet("General"), wb_propuestas.worksheet("Bids"), 
                    wb_propuestas.worksheet("Materials"), wb_propuestas.worksheet("Portafolio"), 
                    wb_propuestas.worksheet("Sistemas"), ws_cond)
        except Exception as e:
            if attempt < retries - 1: time.sleep(2)
            else:
                st.error(f"Error connecting to Sheets: {e}")
                st.stop()

try:
    ws_maestra, ws_gen, ws_bids, ws_mat, ws_port, ws_sist, ws_cond = conectar_sheets()
except Exception as e:
    st.stop()

def fetch_all_data(retries=3):
    for attempt in range(retries):
        try:
            return (pd.DataFrame(ws_maestra.get_all_records()), pd.DataFrame(ws_gen.get_all_records()),
                    pd.DataFrame(ws_bids.get_all_records()), pd.DataFrame(ws_mat.get_all_records()),
                    pd.DataFrame(ws_port.get_all_records()), pd.DataFrame(ws_sist.get_all_records()),
                    pd.DataFrame(ws_cond.get_all_records()))
        except Exception as e:
            if attempt < retries - 1: time.sleep(2)
            else:
                st.error(f"Connection error with Google Sheets (503). Service temporarily unavailable. Please reload the page in a few seconds. Detail: {e}")
                st.stop()

def cargar_datos_si_necesario():
    if 'df_maestra' not in st.session_state or 'df_cond' not in st.session_state:
        dfs = fetch_all_data()
        st.session_state.df_maestra, st.session_state.df_historial, st.session_state.df_bids = dfs[0], dfs[1], dfs[2]
        st.session_state.df_mat, st.session_state.df_port, st.session_state.df_sist, st.session_state.df_cond = dfs[3], dfs[4], dfs[5], dfs[6]

cargar_datos_si_necesario()

# ==========================================
# 3. FUNCIONES DE LÓGICA Y UTILIDADES
# ==========================================
U_SCOPE = ["SF", "LF", "Units", "m2"]
U_THICK = ["inches", "mm", "cm", "gauge"]
U_DIM = ["inches", "ft", "mm", "cm", "m"]

# --- LÓGICA DE DIAGRAMAS DESDE GOOGLE DRIVE ---
diagramas_disp = ["-- None --"]
try:
    if ID_CARPETA_DIAGRAMAS:
        drive_service = get_drive_service()
        results = drive_service.files().list(
            q=f"'{ID_CARPETA_DIAGRAMAS}' in parents and trashed=false",
            fields="files(id, name)"
        ).execute()
        items = results.get('files', [])
        for item in items:
            if item['name'].lower().endswith(('.png', '.jpg', '.jpeg')):
                diagramas_disp.append(item['name'])
                st.session_state[f"diag_id_{item['name']}"] = item['id']
except Exception as e:
    pass # Falla silenciosa si no hay ID o permisos, simplemente muestra "-- None --"

def buscar_dato(fila_dict, palabra_clave, excluir=None):
    if hasattr(fila_dict, 'to_dict'): fila_dict = fila_dict.to_dict()
    for col, val in fila_dict.items():
        col_limpia = str(col).replace('\n', ' ').replace('  ', ' ').strip().lower()
        if palabra_clave.lower() in col_limpia:
            if excluir and excluir.lower() in col_limpia: continue
            return str(val).strip()
    return ""

def limpiar_numero(val):
    if pd.isna(val) or val == "": return 0
    s_val = str(val).strip().replace(',', '')
    if '.' in s_val:
        s_val = s_val.split('.')[0]
    clean = re.sub(r'[^\d]', '', s_val)
    return int(clean) if clean else 0

def format_currency(val):
    num = limpiar_numero(val)
    return f"{num:,}" if num > 0 else ""

def format_currency_callback(key):
    val = st.session_state.get(key, "")
    st.session_state[key] = format_currency(val)

def limpiar_numero_decimal(val):
    clean = re.sub(r'[^\d.]', '', str(val).replace(',', ''))
    parts = clean.split('.')
    if len(parts) > 2: clean = parts[0] + '.' + ''.join(parts[1:])
    try: return float(clean) if clean else 0.0
    except: return 0.0

def format_decimal(val):
    if not str(val).strip(): return ""
    num = limpiar_numero_decimal(val)
    if num == 0 and str(val).strip() not in ["0", "0.", ".0", "0.0", "0.00"]: return ""
    formatted = f"{num:,.2f}"
    if formatted.endswith(".00"): return formatted[:-3]
    return formatted

def format_decimal_callback(key):
    val = st.session_state.get(key, "")
    st.session_state[key] = format_decimal(val)

def format_fluid_unit(val, unit):
    if not val: return ""
    if unit == "inches": return f"{val}\""
    if unit == "feet": return f"{val} ft"
    return f"{val} {unit}"

def extraer_clientes_avanzado(texto_col_h):
    clientes = {}
    if not texto_col_h: return clientes
    matches = re.findall(r'\[(.*?)\]:\s*([^\n\[]+)', str(texto_col_h))
    for comp, ems in matches:
        comp = comp.strip()
        em_list = [e.strip() for e in ems.split(',') if e.strip()]
        if comp not in clientes: clientes[comp] = set()
        clientes[comp].update(em_list)
    return {k: sorted(list(v)) for k, v in clientes.items()}

def validar_email(email):
    if not email or email == "-- Empty --" or email == "-- Type Manually --": return True
    return re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", email) is not None

def to_safe_date(val, default=None):
    if default is None: default = datetime.today().date()
    if val is None or val == "" or val is pd.NaT: return default
    if isinstance(val, date) and not isinstance(val, datetime): return val
    try:
        ts = pd.to_datetime(val, errors="coerce")
        if pd.isna(ts): return default
        return ts.date()
    except: return default

# ==========================================
# 4. ESTADO DE LA SESIÓN Y CARGA DE DATOS
# ==========================================
if 'next_bid_id' not in st.session_state: st.session_state.next_bid_id = 1
if 'next_child_id' not in st.session_state: st.session_state.next_child_id = 1
if 'next_mat_id' not in st.session_state: st.session_state.next_mat_id = 1
if 'next_scope_id' not in st.session_state: st.session_state.next_scope_id = 1
if 'current_loaded_id' not in st.session_state: st.session_state.current_loaded_id = None 

ids_maestra = st.session_state.df_maestra['ID_Proyecto'].dropna().astype(str).tolist()
ids_historial = st.session_state.df_historial['ID'].dropna().astype(str).tolist() if not st.session_state.df_historial.empty else []
lista_ids = list(dict.fromkeys(ids_maestra + ids_historial))[::-1] 
opciones_id = ["-- New Project (Manual) --"] + lista_ids

def get_defaults_db(id_sel):
    fila_m = st.session_state.df_maestra[st.session_state.df_maestra['ID_Proyecto'] == id_sel].iloc[0] if id_sel in ids_maestra else {}
    name = str(fila_m.get('Proyecto_Nombre', '')) if fila_m is not {} else ""
    inv = to_safe_date(fila_m.get('Fecha_invitación', '')) if fila_m is not {} else datetime.today().date()
    comp = to_safe_date(fila_m.get('Fecha_última_modificación', '')) if fila_m is not {} else datetime.today().date()
    return name, inv, comp

def al_cambiar_id():
    id_sel = st.session_state.get("selector_id", opciones_id[0])
    
    if st.session_state.get('current_loaded_id') == id_sel:
        return
    st.session_state.current_loaded_id = id_sel
    
    cargar_datos_si_necesario() 
    
    if 'next_bid_id' not in st.session_state: st.session_state.next_bid_id = 1
    if 'next_child_id' not in st.session_state: st.session_state.next_child_id = 1
    if 'next_mat_id' not in st.session_state: st.session_state.next_mat_id = 1
    if 'next_scope_id' not in st.session_state: st.session_state.next_scope_id = 1

    for k in ["def_calcs", "calc_val_input", "calc_opt"]:
        if k in st.session_state:
            del st.session_state[k]

    name_db, inv_db, comp_db = get_defaults_db(id_sel)
    
    st.session_state.proj_name = name_db
    st.session_state.ref_text = name_db 
    st.session_state.sec_text = ""
    st.session_state.gen_notes = ""
    st.session_state.opciones_rev = ["Original"]
    st.session_state.rev_seleccionada = "Original"
    st.session_state.prop_date = datetime.today().date()
    st.session_state.inv_date = inv_db
    st.session_state.comp_date = comp_db
    
    st.session_state.bids_list = [{'id': st.session_state.next_bid_id, 'type': 'BASE BID', 'name': '', 'amount_str': '', 'children': [], 'selected_sum_ids': []}]
    st.session_state.next_bid_id += 1
    st.session_state.mat_list = []
    
    st.session_state.cliente_selector = "-- Select --"
    st.session_state.cliente_editable = ""
    st.session_state.attention = ""
    for i in range(1, 5): st.session_state[f"e{i}_sel"] = "-- Empty --"; st.session_state[f"e{i}_man"] = ""
    
    df_cond = st.session_state.df_cond
    if not df_cond.empty:
        fila_c = df_cond.iloc[0]
        st.session_state.val_validity = str(fila_c.get('Validity', '30'))
        st.session_state.val_calcs = str(fila_c.get('Calcs', 'N'))
        st.session_state.val_mock = str(fila_c.get('Mock_Up', 'N'))
        st.session_state.val_tax = str(fila_c.get('Sales_Tax', 'N'))
        st.session_state.val_exec = to_safe_date(str(fila_c.get('Execution_Limit_Date', '')))
    else:
        st.session_state.val_validity = "30"; st.session_state.val_calcs = "N"; st.session_state.val_mock = "N"; st.session_state.val_tax = "N"
        st.session_state.val_exec = datetime.today().date()

    if id_sel and id_sel != "-- New Project (Manual) --":
        df_h = st.session_state.df_historial
        historial_id = df_h[df_h['ID'].astype(str) == str(id_sel)] if not df_h.empty else pd.DataFrame()
        
        if not historial_id.empty:
            revs_str = historial_id['Proposal'].astype(str).tolist()
            rev_nums = [int(re.findall(r'\d+', r)[0]) for r in revs_str if "Rev" in r and re.findall(r'\d+', r)]
            ultima_rev = f"Rev {max(rev_nums)}" if rev_nums else "Original"
            siguiente_rev = f"Rev {max(rev_nums) + 1}" if rev_nums else "Rev 1"
            
            st.session_state.opciones_rev = [ultima_rev, siguiente_rev]
            st.session_state.rev_seleccionada = ultima_rev
            
            fila_h = historial_id[historial_id['Proposal'] == ultima_rev].iloc[-1]
            
            st.session_state.ref_text = str(fila_h.get('Reference', name_db))
            st.session_state.sec_text = str(fila_h.get('Sections', ''))
            st.session_state.gen_notes = str(buscar_dato(fila_h, 'general notes'))
            st.session_state.inv_date = to_safe_date(fila_h.get('Invitation Date', inv_db))
            st.session_state.comp_date = to_safe_date(fila_h.get('Compliance Date', comp_db))
            
            cliente_db = buscar_dato(fila_h, 'customer')
            fila_m = st.session_state.df_maestra[st.session_state.df_maestra['ID_Proyecto'] == id_sel].iloc[0] if id_sel in ids_maestra else {}
            clientes_dict = extraer_clientes_avanzado(fila_m.get('Contactos_Envío', '')) if fila_m is not {} else {}
            
            st.session_state.cliente_selector = cliente_db if cliente_db in clientes_dict else "-- New (Manual) --"
            st.session_state.cliente_editable = cliente_db
            st.session_state.attention = buscar_dato(fila_h, 'attention')
            
            emails_base = clientes_dict.get(cliente_db, [])
            for i in range(1, 5):
                em_val = buscar_dato(fila_h, f'email {i}')
                if em_val:
                    if em_val in emails_base:
                        st.session_state[f"e{i}_sel"] = em_val
                        st.session_state[f"e{i}_man"] = ""
                    else:
                        st.session_state[f"e{i}_sel"] = "-- Type Manually --"
                        st.session_state[f"e{i}_man"] = em_val
            
            st.session_state.val_validity = str(buscar_dato(fila_h, 'validity')) or st.session_state.val_validity
            st.session_state.val_mock = str(buscar_dato(fila_h, 'mock up')) or st.session_state.val_mock
            st.session_state.val_tax = str(buscar_dato(fila_h, 'sales tax')) or st.session_state.val_tax
            
            st.session_state.val_calcs = str(buscar_dato(fila_h, 'calcs')) or st.session_state.val_calcs
            st.session_state.calc_val_input = format_currency(st.session_state.val_calcs) if st.session_state.val_calcs != "N" else ""
            
            df_b = st.session_state.df_bids
            if not df_b.empty:
                bids_hist = df_b[(df_b['ID'].astype(str) == str(id_sel)) & (df_b['Proposal'].astype(str) == str(ultima_rev))].copy()
                if not bids_hist.empty:
                    st.session_state.bids_list = []
                    bids_hist['Relation'] = bids_hist['Relation'].fillna(0).astype(str)
                    for rel, group in bids_hist.groupby('Relation'):
                        parent = group.iloc[0]
                        children = group.iloc[1:]
                        monto_p = buscar_dato(parent, 'amount')
                        bid_dict = {'id': st.session_state.next_bid_id, 'type': buscar_dato(parent, 'type').upper() or 'BASE BID', 'name': buscar_dato(parent, 'bid name'), 'amount_str': format_currency(monto_p), 'children': [], 'selected_sum_ids': []}
                        st.session_state.next_bid_id += 1
                        for _, child in children.iterrows():
                            monto_c = buscar_dato(child, 'amount')
                            bid_dict['children'].append({'id': st.session_state.next_child_id, 'type': buscar_dato(child, 'type').upper() or 'VA', 'name': buscar_dato(child, 'bid name'), 'amount_str': format_currency(monto_c)})
                            st.session_state.next_child_id += 1
                        st.session_state.bids_list.append(bid_dict)
            
            df_mat = st.session_state.df_mat
            if not df_mat.empty:
                mat_hist = df_mat[(df_mat['ID'].astype(str) == str(id_sel)) & (df_mat['Proposal'].astype(str) == str(ultima_rev))]
                if not mat_hist.empty:
                    mat_dict = {}
                    for _, row in mat_hist.iterrows():
                        full_item_name = buscar_dato(row, 'proposal item')
                        parts = full_item_name.split(" | Scope: ")
                        item_base = parts[0]
                        scope_desc = parts[1] if len(parts) > 1 else "General"
                        qty_val = buscar_dato(row, 'scope (qty)')
                        unit_val = buscar_dato(row, 'units (scope)') or 'SF'
                        
                        mat_key = f"{item_base}_{buscar_dato(row, 'material')}"
                        
                        if mat_key not in mat_dict:
                            mat_dict[mat_key] = {
                                'id': st.session_state.next_mat_id,
                                'item_name': item_base,
                                'mat_name': buscar_dato(row, 'material'), 'prev_name': buscar_dato(row, 'material'),
                                'scopes': [],
                                'thick': buscar_dato(row, 'thickness', excluir='units'), 'u_thick': buscar_dato(row, 'units (thickness)') or 'inches', 
                                't_thick': str(buscar_dato(row, 't_thick')).lower() == 'true',
                                'spec': buscar_dato(row, 'specification'), 
                                't_spec': str(buscar_dato(row, 't_spec')).lower() == 'true',
                                'width': buscar_dato(row, 'width', excluir='units'), 'u_width': buscar_dato(row, 'units (width)') or 'inches', 
                                't_width': str(buscar_dato(row, 't_width')).lower() == 'true',
                                'length': buscar_dato(row, 'lenght', excluir='units'), 'u_length': buscar_dato(row, 'units (lenght)') or 'inches', 
                                't_length': str(buscar_dato(row, 't_length')).lower() == 'true',
                                'brands': buscar_dato(row, 'brands'),
                                'colors': buscar_dato(row, 'color'),
                                'sys_sel': buscar_dato(row, 'system') or '-- New System --', 'sys_new': '', 
                                'diagram': buscar_dato(row, 'diagram') or '-- None --',
                                'custom_note': buscar_dato(row, 'custom note')
                            }
                            st.session_state.next_mat_id += 1
                            
                        mat_dict[mat_key]['scopes'].append({
                            'id': st.session_state.next_scope_id,
                            'desc': scope_desc,
                            'qty': format_currency(qty_val),
                            'unit': unit_val
                        })
                        st.session_state.next_scope_id += 1
                        
                    st.session_state.mat_list = list(mat_dict.values())

if 'selector_id' not in st.session_state:
    st.session_state.selector_id = opciones_id[0]
    al_cambiar_id()

# ==========================================
# 5. INTERFAZ - DATOS BÁSICOS
# ==========================================
st.markdown("### 1. Project Basic Data")

if st.session_state.get("selector_id") == "-- New Project (Manual) --":
    col_id1, col_id2 = st.columns([4, 8])
    with col_id1: st.selectbox("Project ID:", opciones_id, key="selector_id", on_change=al_cambiar_id)
    with col_id2: id_final = st.text_input("Manual ID:", key="id_manual")
else:
    st.selectbox("Project ID:", opciones_id, key="selector_id", on_change=al_cambiar_id)
    id_final = st.session_state.get("selector_id")

st.markdown(f"<div class='info-box-blue'><b>Project Name (DB):</b> {st.session_state.get('proj_name', '')}</div>", unsafe_allow_html=True)

col_inp1, col_inp2, col_inp3 = st.columns([0.6, 8.4, 3])
with col_inp1:
    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
    if st.button("🔄", key="res_ref", type="tertiary", help="Restore Reference"): 
        st.session_state.ref_text = st.session_state.get('proj_name', '')
        st.rerun()
with col_inp2: 
    val_ref = st.text_input("Reference (Editable):", value=st.session_state.get('ref_text', ''), key="ref_text", autocomplete="off")
with col_inp3: 
    val_prop = st.selectbox("Proposal (Version):", st.session_state.get('opciones_rev', ["Original"]), key="rev_seleccionada")

val_sec = st.text_area("Sections / Key Notes:", value=st.session_state.get('sec_text', ''), key="sec_text", height=68)

col_d1, col_d2, col_d3, col_d4, col_d5, col_d6, col_d7 = st.columns([0.6, 1.5, 0.6, 1.5, 0.6, 1.5, 3.7])
with col_d1:
    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
    if st.button("🔄", key="res_d1", type="tertiary"): st.session_state.prop_date = datetime.today().date(); st.rerun()
with col_d2: val_date = st.date_input("Proposal Date:", value=to_safe_date(st.session_state.get('prop_date')), key="prop_date")

with col_d3:
    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
    if st.button("🔄", key="res_d2", type="tertiary"): st.session_state.inv_date = get_defaults_db(id_final)[1]; st.rerun()
with col_d4: val_inv_date = st.date_input("Invitation Date:", value=to_safe_date(st.session_state.get('inv_date')), key="inv_date")

with col_d5:
    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
    if st.button("🔄", key="res_d3", type="tertiary"): st.session_state.comp_date = get_defaults_db(id_final)[2]; st.rerun()
with col_d6: val_comp_date = st.date_input("Compliance Date:", value=to_safe_date(st.session_state.get('comp_date')), key="comp_date")

st.divider()

# ==========================================
# 6. INTERFAZ - CLIENTE
# ==========================================
st.markdown("### 2. Customer & Emails")
fila_m = st.session_state.df_maestra[st.session_state.df_maestra['ID_Proyecto'] == id_final].iloc[0] if id_final in ids_maestra else {}
clientes_dict = extraer_clientes_avanzado(fila_m.get('Contactos_Envío', '')) if fila_m is not {} else {}

st.info(f"**Bidding Instructions (DB):**\n{str(fila_m.get('Instrucciones', 'No instructions')) if fila_m is not {} else 'N/A'}")

col_cli1, col_cli2 = st.columns(2)
with col_cli1:
    lista_nombres_clientes = ["-- Select --", "-- New (Manual) --"] + list(clientes_dict.keys())
    
    def al_cambiar_constructora():
        sel = st.session_state.get('cliente_selector', '-- Select --')
        st.session_state.cliente_editable = sel if sel not in ["-- Select --", "-- New (Manual) --"] else ""
        st.session_state.attention = ""
        for i in range(1, 5): st.session_state[f"e{i}_sel"] = "-- Empty --"; st.session_state[f"e{i}_man"] = ""

    sel_cust = st.selectbox("General Contractor:", lista_nombres_clientes, key="cliente_selector", on_change=al_cambiar_constructora)
    
    val_customer = st.text_input("Final Name to Use:", value=st.session_state.get('cliente_editable',''), key="cliente_editable", autocomplete="off")
    val_attention = st.text_input("Attention:", value=st.session_state.get('attention', ''), key="attention", autocomplete="off")

with col_cli2:
    emails_base = clientes_dict.get(sel_cust, [])
    def get_opciones(usados): return ["-- Empty --", "-- Type Manually --"] + [e for e in emails_base if e not in usados]
    
    e1_sel = st.selectbox("Email 1 (Main):", get_opciones([]), key="e1_sel")
    e1_final = st.text_input("✎ Type Email 1:", value=st.session_state.get('e1_man',''), key="e1_man", autocomplete="off") if e1_sel == "-- Type Manually --" else (e1_sel if e1_sel != "-- Empty --" else "")
    
    e2_sel = st.selectbox("Email 2 (CC):", get_opciones([e1_sel]), key="e2_sel")
    e2_final = st.text_input("✎ Type Email 2:", value=st.session_state.get('e2_man',''), key="e2_man", autocomplete="off") if e2_sel == "-- Type Manually --" else (e2_sel if e2_sel != "-- Empty --" else "")
    
    e3_sel = st.selectbox("Email 3 (CC):", get_opciones([e1_sel, e2_sel]), key="e3_sel")
    e3_final = st.text_input("✎ Type Email 3:", value=st.session_state.get('e3_man',''), key="e3_man", autocomplete="off") if e3_sel == "-- Type Manually --" else (e3_sel if e3_sel != "-- Empty --" else "")
    
    e4_sel = st.selectbox("Email 4 (CC):", get_opciones([e1_sel, e2_sel, e3_sel]), key="e4_sel")
    e4_final = st.text_input("✎ Type Email 4:", value=st.session_state.get('e4_man',''), key="e4_man", autocomplete="off") if e4_sel == "-- Type Manually --" else (e4_sel if e4_sel != "-- Empty --" else "")

st.divider()

# ==========================================
# 7. INTERFAZ - BIDS DINÁMICOS
# ==========================================
st.markdown("### 3. Pricing Structure")

def add_bid():
    st.session_state.bids_list.append({'id': st.session_state.next_bid_id, 'type': 'BASE BID', 'name': '', 'amount_str': '', 'children': [], 'selected_sum_ids': []})
    st.session_state.next_bid_id += 1

def add_child(bid_idx):
    st.session_state.bids_list[bid_idx]['children'].append({'id': st.session_state.next_child_id, 'type': 'VA', 'name': '', 'amount_str': ''})
    st.session_state.next_child_id += 1

def apply_sum_callback(b_id, t_sum):
    st.session_state[f"b_a_{b_id}"] = format_currency(t_sum)

available_dict = {}

for i, bid in enumerate(st.session_state.bids_list):
    st.markdown("<div class='bid-box'>", unsafe_allow_html=True)
    
    c1, c2, c3, c4, c5 = st.columns([2, 3, 2, 2, 1])
    
    type_options = ["BASE BID", "SUBTOTAL", "TOTAL"]
    type_idx = type_options.index(bid['type']) if bid['type'] in type_options else 0
    bid['type'] = c1.selectbox("Type", type_options, index=type_idx, key=f"b_t_{bid['id']}")
    
    bid['name'] = c2.text_input("Description", value=bid['name'], key=f"b_n_{bid['id']}")
    
    if f"b_a_{bid['id']}" not in st.session_state: 
        st.session_state[f"b_a_{bid['id']}"] = bid['amount_str']
    
    c3.text_input("Amount ($)", key=f"b_a_{bid['id']}", on_change=format_currency_callback, args=(f"b_a_{bid['id']}",))
    bid['amount_str'] = st.session_state[f"b_a_{bid['id']}"]
    
    if bid['type'] in ["SUBTOTAL", "TOTAL"]:
        with c4:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            with st.popover("🧮 Calculator"):
                st.markdown("**Select items to sum:**")
                if 'selected_sum_ids' not in bid: bid['selected_sum_ids'] = []
                
                new_selected = []
                total_sum = 0
                
                if not available_dict:
                    st.info("No previous items to sum.")
                else:
                    for item_key, item_data in available_dict.items():
                        is_checked = st.checkbox(
                            f"{item_data['name']} (${item_data['amount_str']})", 
                            value=(item_key in bid['selected_sum_ids']),
                            key=f"chk_{bid['id']}_{item_key}"
                        )
                        if is_checked:
                            new_selected.append(item_key)
                            total_sum += limpiar_numero(item_data['amount_str'])
                            
                    bid['selected_sum_ids'] = new_selected
                    
                    st.markdown(f"**Calculated Sum: {format_currency(total_sum)}**")
                    
                    st.button("✔️ Use this sum", key=f"apply_sum_{bid['id']}", on_click=apply_sum_callback, args=(bid['id'], total_sum))
    else:
        with c4:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            st.button("➕ Alt/VE", key=f"add_c_{bid['id']}", on_click=add_child, args=(i,))
            
    with c5:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        if st.button("🗑️", key=f"rem_b_{bid['id']}"): 
            st.session_state.bids_list.pop(i)
            st.rerun()
        
    available_dict[f"b_{bid['id']}"] = {'type': bid['type'], 'name': bid['name'] or f"Bid {bid['id']}", 'amount_str': bid['amount_str']}

    for j, child in enumerate(bid['children']):
        st.markdown("<div class='child-box'>", unsafe_allow_html=True)
        cc0, cc1, cc2, cc3, cc4 = st.columns([0.5, 2, 4, 2, 1])
        
        child_type_options = ["VA", "VE", "OPT"]
        child_type_idx = child_type_options.index(child['type']) if child['type'] in child_type_options else 0
        child['type'] = cc1.selectbox("Alt. Type", child_type_options, index=child_type_idx, key=f"c_t_{child['id']}")
        
        child['name'] = cc2.text_input("Alt. Description", value=child['name'], key=f"c_n_{child['id']}")
        
        if f"c_a_{child['id']}" not in st.session_state: st.session_state[f"c_a_{child['id']}"] = child['amount_str']
        cc3.text_input("Alt. Amount ($)", key=f"c_a_{child['id']}", on_change=format_currency_callback, args=(f"c_a_{child['id']}",))
        child['amount_str'] = st.session_state[f"c_a_{child['id']}"]
            
        if cc4.button("🗑️", key=f"rem_c_{child['id']}"): bid['children'].pop(j); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        
        available_dict[f"c_{child['id']}"] = {'type': child['type'], 'name': child['name'] or f"Alt {child['id']}", 'amount_str': child['amount_str']}

    st.markdown("</div>", unsafe_allow_html=True)

st.button("➕ Add New Bid/Total", on_click=add_bid, type="primary")

st.divider()

# ==========================================
# 8. INTERFAZ - MATERIALES INTELIGENTES
# ==========================================
st.markdown("### 4. Materials & Systems")

df_port = st.session_state.df_port
lista_portafolio = df_port['Material'].dropna().unique().tolist() if not df_port.empty else []

def add_material():
    st.session_state.mat_list.append({
        'id': st.session_state.next_mat_id, 'item_name': '', 'mat_name': '', 'prev_name': '',
        'scopes': [{'id': st.session_state.next_scope_id, 'desc': 'General', 'qty': '', 'unit': 'SF'}],
        'thick': '', 'u_thick': 'inches', 't_thick': False,
        'spec': '', 't_spec': False,
        'width': '', 'u_width': 'inches', 't_width': False,
        'length': '', 'u_length': 'inches', 't_length': False,
        'brands': '',
        'colors': '', 'sys_sel': '-- New System --', 'sys_new': '', 'diagram': '-- None --', 'custom_note': ''
    })
    st.session_state.next_mat_id += 1
    st.session_state.next_scope_id += 1

def add_scope(mat_idx):
    st.session_state.mat_list[mat_idx]['scopes'].append({'id': st.session_state.next_scope_id, 'desc': '', 'qty': '', 'unit': 'SF'})
    st.session_state.next_scope_id += 1

def remove_material(idx):
    st.session_state.mat_list.pop(idx)

def attr_change_callback(k_val, k_u, k_t, d_v, d_u, d_t, chk_key, is_decimal=False):
    if is_decimal: format_decimal_callback(k_val)
    val = st.session_state.get(k_val, "")
    u_val = st.session_state.get(k_u, "") if k_u else None
    t_val = st.session_state.get(k_t, False) if k_t else None
    
    match = (val == d_v) and (d_v != "")
    if k_u: match = match and (u_val == d_u)
    if k_t: match = match and (t_val == d_t)
    st.session_state[chk_key] = match

def reset_attr_callback(k_val, k_u, k_t, d_v, d_u, d_t, chk_key):
    st.session_state[k_val] = d_v
    if k_u: st.session_state[k_u] = d_u
    if k_t: st.session_state[k_t] = d_t
    st.session_state[chk_key] = (d_v != "")

for i, mat in enumerate(st.session_state.mat_list):
    st.markdown("<div class='mat-box'>", unsafe_allow_html=True)
    
    cm1, cm2, cm3 = st.columns([4, 4, 1])
    mat['item_name'] = cm1.text_input("1. Proposed Item Name (e.g., Building 4):", value=mat.get('item_name', ''), key=f"m_in_{mat['id']}")
    
    mat_options = ["-- New Material --"] + lista_portafolio
    mat_idx = mat_options.index(mat['mat_name']) if mat['mat_name'] in mat_options else 0
    
    mat_sel = cm2.selectbox("2. Material:", mat_options, index=mat_idx, key=f"m_sel_{mat['id']}")
    
    if mat_sel != mat['prev_name']:
        mat['prev_name'] = mat_sel
        if mat_sel != "-- New Material --":
            fila_p = df_port[df_port['Material'] == mat_sel].iloc[0]
            mat['mat_name'] = mat_sel
            mat['thick'] = str(buscar_dato(fila_p, 'thickness', excluir='units'))
            mat['u_thick'] = buscar_dato(fila_p, 'units (thickness)') or 'inches'
            mat['t_thick'] = str(buscar_dato(fila_p, 't_thick')).lower() == 'true'
            mat['spec'] = str(buscar_dato(fila_p, 'specification'))
            mat['t_spec'] = str(buscar_dato(fila_p, 't_spec')).lower() == 'true'
            mat['width'] = str(buscar_dato(fila_p, 'width', excluir='units'))
            mat['u_width'] = buscar_dato(fila_p, 'units (width)') or 'inches'
            mat['t_width'] = str(buscar_dato(fila_p, 't_width')).lower() == 'true'
            mat['length'] = str(buscar_dato(fila_p, 'lenght', excluir='units'))
            mat['u_length'] = buscar_dato(fila_p, 'units (lenght)') or 'inches'
            mat['t_length'] = str(buscar_dato(fila_p, 't_length')).lower() == 'true'
            mat['brands'] = str(buscar_dato(fila_p, 'brands'))
            mat['sys_sel'] = "-- New System --"
            mat['sys_new'] = ""
            mat['diagram'] = "-- None --"
            
            if f"s_d_{mat['id']}" in st.session_state: del st.session_state[f"s_d_{mat['id']}"]
            if f"def_diag_{mat['id']}" in st.session_state: del st.session_state[f"def_diag_{mat['id']}"]
        else:
            mat['mat_name'] = ""
        st.rerun()

    if mat_sel == "-- New Material --":
        mat['mat_name'] = cm2.text_input("New Material Name:", value=mat['mat_name'], key=f"m_n_{mat['id']}")
    else:
        mat['mat_name'] = mat_sel

    cm3.button("🗑️ Delete Material", key=f"rem_m_{mat['id']}", on_click=remove_material, args=(i,))

    st.markdown("**Attributes ( 🔄 = Restore to DB | 📌 = Include in Fluid Title | 💾 = Set as Default )**")
    
    fila_p = df_port[df_port['Material'] == mat['mat_name']].iloc[0] if not df_port.empty and mat['mat_name'] in df_port['Material'].values else None
    
    def get_db_val(col_name, exclude=None): return str(buscar_dato(fila_p, col_name, excluir=exclude)) if fila_p is not None else ""
    def get_db_bool(col_name): return str(buscar_dato(fila_p, col_name)).lower() == 'true' if fila_p is not None else False

    db_thick, db_u_thick, db_t_thick = get_db_val('thickness', 'units'), get_db_val('units (thickness)') or 'inches', get_db_bool('t_thick')
    db_width, db_u_width, db_t_width = get_db_val('width', 'units'), get_db_val('units (width)') or 'inches', get_db_bool('t_width')
    db_length, db_u_length, db_t_length = get_db_val('lenght', 'units'), get_db_val('units (lenght)') or 'inches', get_db_bool('t_length')
    db_spec, db_t_spec = get_db_val('specification'), get_db_bool('t_spec')
    db_brands = get_db_val('brands')

    def render_attr(cols, label, key_val, key_u, u_list, key_t, db_v, db_u, db_t, multiline=False):
        val_key = f"{key_val}_{mat['id']}"
        u_key = f"{key_u}_{mat['id']}" if u_list else None
        t_key = f"{key_t}_{mat['id']}" if key_t else None
        chk_key = f"def_{key_val}_{mat['id']}"

        if val_key not in st.session_state: st.session_state[val_key] = mat.get(key_val, '')
        if u_key and u_key not in st.session_state: st.session_state[u_key] = mat.get(key_u, u_list[0])
        if t_key and t_key not in st.session_state: st.session_state[t_key] = mat.get(key_t, False)
        
        if chk_key not in st.session_state: 
            match = (st.session_state[val_key] == db_v) and (db_v != "")
            if u_key: match = match and (st.session_state[u_key] == db_u)
            if t_key: match = match and (st.session_state[t_key] == db_t)
            st.session_state[chk_key] = match

        is_decimal = key_val in ['thick', 'width', 'length']

        c_in = cols[0]
        if u_list:
            c_u = cols[1]
            c_btn1 = cols[2]
            if key_t:
                c_btn2 = cols[3]
                c_btn3 = cols[4]
            else:
                c_btn3 = cols[3]
        else:
            c_btn1 = cols[1]
            if key_t:
                c_btn2 = cols[2]
                c_btn3 = cols[3]
            else:
                c_btn3 = cols[2]

        with c_in:
            if multiline:
                val = st.text_area(label, key=val_key, height=68, on_change=attr_change_callback, args=(val_key, u_key, t_key, db_v, db_u, db_t, chk_key, False))
            else:
                val = st.text_input(label, key=val_key, on_change=attr_change_callback, args=(val_key, u_key, t_key, db_v, db_u, db_t, chk_key, is_decimal))
        
        if u_list:
            with c_u:
                u_val = st.selectbox("Units", u_list, key=u_key, label_visibility="hidden", on_change=attr_change_callback, args=(val_key, u_key, t_key, db_v, db_u, db_t, chk_key, False))
        
        with c_btn1:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            st.button("🔄", key=f"ref_{key_val}_{mat['id']}", type="tertiary", on_click=reset_attr_callback, args=(val_key, u_key, t_key, db_v, db_u, db_t, chk_key))
        
        if key_t:
            with c_btn2:
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                t_val = st.checkbox("📌", key=t_key, on_change=attr_change_callback, args=(val_key, u_key, t_key, db_v, db_u, db_t, chk_key, False))
        
        with c_btn3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            mat[f'def_{key_val}'] = st.checkbox("💾", key=chk_key)
        
        mat[key_val] = st.session_state[val_key]
        if u_list: mat[key_u] = st.session_state[u_key]
        if key_t: mat[key_t] = st.session_state[t_key]

    cols_t = st.columns([2, 1.5, 0.6, 0.6, 0.6, 0.5, 4.2], gap="small")
    render_attr(cols_t[:5], "Thickness:", 'thick', 'u_thick', U_THICK, 't_thick', db_thick, db_u_thick, db_t_thick)
    
    cols_wl = st.columns([1.5, 1.2, 0.6, 0.6, 0.6, 0.5, 1.5, 1.2, 0.6, 0.6, 0.6, 0.5], gap="small")
    render_attr(cols_wl[0:5], "Width:", 'width', 'u_width', U_DIM, 't_width', db_width, db_u_width, db_t_width)
    render_attr(cols_wl[6:11], "Length:", 'length', 'u_length', U_DIM, 't_length', db_length, db_u_length, db_t_length)
    
    cols_s = st.columns([4, 0.6, 0.6, 0.6, 0.5, 3.7], gap="small")
    render_attr(cols_s[:4], "Specification:", 'spec', None, None, 't_spec', db_spec, None, db_t_spec, multiline=True)
    
    cols_b = st.columns([4, 0.6, 0.6, 0.6, 4.2], gap="small")
    render_attr(cols_b[:3], "Brands (Not in title):", 'brands', None, None, None, db_brands, None, None)

    parts = []
    if mat.get('t_thick') and mat.get('thick'): parts.append(format_fluid_unit(mat['thick'], mat['u_thick']))
    dim_parts = []
    if mat.get('t_width') and mat.get('width'): dim_parts.append(format_fluid_unit(mat['width'], mat['u_width']))
    if mat.get('t_length') and mat.get('length'): dim_parts.append(f"max {format_fluid_unit(mat['length'], mat['u_length'])}")
    if dim_parts: parts.append(" x ".join(dim_parts))
    if mat['mat_name']: parts.append(mat['mat_name'])
    if mat.get('t_spec') and mat.get('spec'): parts.append(mat['spec'])
    
    fluid_title = " - ".join(parts).strip()
    st.markdown(f"<div class='fluid-title-box'>3. Fluid Title: {fluid_title if fluid_title else '...'}</div>", unsafe_allow_html=True)

    st.markdown("📝 **Quantities / Areas (Scopes):**")
    for k, scope in enumerate(mat['scopes']):
        st.markdown("<div class='scope-box'>", unsafe_allow_html=True)
        cs1, cs2, cs3, cs4, cs_space = st.columns([3, 1.5, 1.5, 0.5, 3])
        
        vis = "visible" if k == 0 else "collapsed"
        scope['desc'] = cs1.text_input("Description (e.g., Fascias):", value=scope['desc'], key=f"sc_d_{scope['id']}", label_visibility=vis, placeholder="e.g., Fascias")
        
        if f"sc_q_{scope['id']}" not in st.session_state: st.session_state[f"sc_q_{scope['id']}"] = scope['qty']
        cs2.text_input("Scope (Qty):", key=f"sc_q_{scope['id']}", label_visibility=vis, on_change=format_currency_callback, args=(f"sc_q_{scope['id']}",))
        scope['qty'] = st.session_state[f"sc_q_{scope['id']}"]
        
        scope['unit'] = cs3.selectbox("Units:", U_SCOPE, index=U_SCOPE.index(scope['unit']) if scope['unit'] in U_SCOPE else 0, key=f"sc_u_{scope['id']}", label_visibility=vis)
        
        with cs4:
            if k == 0: st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if len(mat['scopes']) > 1:
                if st.button("🗑️", key=f"rem_sc_{scope['id']}"): 
                    mat['scopes'].pop(k); st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    if st.button("➕ Add another Area/Scope to this material", key=f"add_sc_{mat['id']}"): add_scope(i); st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    cp3, _ = st.columns([4, 6])
    mat['colors'] = cp3.text_input("Color(s):", value=mat['colors'], key=f"m_co_{mat['id']}")

    df_sist = st.session_state.df_sist
    sistemas_disp = df_sist[df_sist['Material'] == mat['mat_name']]['System'].tolist() if not df_sist.empty else []
    
    c_sys1, c_sys2, c_sys3, c_btn1, c_btn2, c_space = st.columns([2, 2, 2, 0.6, 0.6, 2.8], gap="small")
    sys_options = ["-- New System --", "-- None --"] + sistemas_disp
    
    if mat['sys_sel'] not in sys_options and mat['sys_sel'] != "":
        sys_options.append(mat['sys_sel'])
        
    if mat['sys_sel'] not in sys_options:
        mat['sys_sel'] = "-- New System --"
        
    sys_idx = sys_options.index(mat['sys_sel'])
    
    def sys_change_callback(k_sel, m_id):
        val = st.session_state.get(k_sel)
        if val is None:
            return 
            
        for m in st.session_state.mat_list:
            if m['id'] == m_id:
                m['sys_sel'] = val
                if val not in ["-- New System --", "-- None --"]:
                    try:
                        m['diagram'] = df_sist[(df_sist['Material'] == m['mat_name']) & (df_sist['System'] == val)]['Diagram'].iloc[0]
                    except:
                        m['diagram'] = "-- None --"
                else:
                    m['diagram'] = "-- None --"
                    m['sys_new'] = ""
                
                st.session_state[f"s_d_{m_id}"] = m['diagram']
                db_diag_val = df_sist[(df_sist['Material'] == m['mat_name']) & (df_sist['System'] == m['sys_new'])]['Diagram'].iloc[0] if not df_sist.empty and m['sys_new'] in df_sist['System'].values else ""
                st.session_state[f"def_diag_{m_id}"] = (m['diagram'] == db_diag_val) and (db_diag_val != "")
                break

    sys_sel = c_sys1.selectbox("System:", sys_options, index=sys_idx, key=f"s_sel_{mat['id']}", on_change=sys_change_callback, args=(f"s_sel_{mat['id']}", mat['id']))

    if mat['sys_sel'] == "-- New System --":
        mat['sys_new'] = c_sys2.text_input("System Name:", value=mat['sys_new'], key=f"s_n_{mat['id']}", placeholder="Type new system...")
    else:
        mat['sys_new'] = mat['sys_sel']
        c_sys2.text_input("System Name:", value=mat['sys_sel'], disabled=True, key=f"s_n_dis_{mat['id']}")
        
    idx_diag = diagramas_disp.index(mat['diagram']) if mat['diagram'] in diagramas_disp else 0
    
    def diag_change_callback(k_val, d_v, chk_key):
        val = st.session_state.get(k_val, "")
        st.session_state[chk_key] = (val == d_v) and (d_v != "")
        
    def reset_diag_callback(k_val, d_v, chk_key):
        st.session_state[k_val] = d_v
        st.session_state[chk_key] = (d_v != "")

    db_diag = df_sist[(df_sist['Material'] == mat['mat_name']) & (df_sist['System'] == mat['sys_new'])]['Diagram'].iloc[0] if not df_sist.empty and mat['sys_new'] in df_sist['System'].values else ""
    chk_diag_key = f"def_diag_{mat['id']}"
    
    if f"s_d_{mat['id']}" not in st.session_state: st.session_state[f"s_d_{mat['id']}"] = mat['diagram']
    if chk_diag_key not in st.session_state: st.session_state[chk_diag_key] = (st.session_state[f"s_d_{mat['id']}"] == db_diag) and (db_diag != "")

    mat['diagram'] = c_sys3.selectbox("Diagram:", diagramas_disp, key=f"s_d_{mat['id']}", on_change=diag_change_callback, args=(f"s_d_{mat['id']}", db_diag, chk_diag_key))
    
    with c_btn1:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        st.button("🔄", key=f"ref_diag_{mat['id']}", type="tertiary", on_click=reset_diag_callback, args=(f"s_d_{mat['id']}", db_diag, chk_diag_key))
        
    with c_btn2:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        mat['def_diag'] = st.checkbox("💾", key=chk_diag_key)

    c_note, _ = st.columns([6, 4])
    mat['custom_note'] = c_note.text_area("Custom Notes (Optional):", value=mat.get('custom_note', ''), height=68, key=f"m_cn_{mat['id']}", placeholder="Enter notes here. Use '-' for bullets.")
    st.markdown("</div>", unsafe_allow_html=True)

st.button("➕ Add New Material", on_click=add_material, type="primary")

st.divider()

# ==========================================
# 9. INTERFAZ - CONDICIONES ESPECIALES
# ==========================================
st.markdown("### 5. Special Conditions ( 🔄 = Restore to DB | 💾 = Set as Default )")

db_cond = st.session_state.df_cond.iloc[0] if not st.session_state.df_cond.empty else {}
db_val_validity = str(db_cond.get('Validity', '30'))
db_val_calcs = str(db_cond.get('Calcs', 'N'))
db_val_mock = str(db_cond.get('Mock_Up', 'N'))
db_val_tax = str(db_cond.get('Sales_Tax', 'N'))
db_val_exec = to_safe_date(str(db_cond.get('Execution_Limit_Date', '')))

def cond_change_callback(k_val, d_v, chk_key, is_currency=False):
    if is_currency: format_currency_callback(k_val)
    val = st.session_state.get(k_val, "")
    st.session_state[chk_key] = (val == d_v) and (d_v != "")

def reset_cond_callback(k_val, d_v, chk_key):
    st.session_state[k_val] = d_v
    st.session_state[chk_key] = (d_v != "")

def render_cond(cols, label, key_val, db_v, is_select=False, options=None):
    chk_key = f"def_{key_val}"
    if key_val not in st.session_state: st.session_state[key_val] = db_v
    if chk_key not in st.session_state: st.session_state[chk_key] = (st.session_state[key_val] == db_v) and (db_v != "")

    c1, c2, c3 = cols
    
    if is_select:
        val = c1.selectbox(label, options, key=key_val, on_change=cond_change_callback, args=(key_val, db_v, chk_key, False))
    else:
        val = c1.text_input(label, key=key_val, on_change=cond_change_callback, args=(key_val, db_v, chk_key, False))
    
    with c2:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        st.button("🔄", key=f"ref_{key_val}", type="tertiary", on_click=reset_cond_callback, args=(key_val, db_v, chk_key))
        
    with c3:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        return st.checkbox("💾", key=chk_key)

cols_r1 = st.columns([1.5, 0.6, 0.6, 0.5, 1.5, 0.6, 0.6, 0.5, 1.5, 0.6, 0.6, 0.4], gap="small")
def_val = render_cond(cols_r1[0:3], "Validity (Days):", 'val_validity', db_val_validity)
def_moc = render_cond(cols_r1[4:7], "Mock Up:", 'val_mock', db_val_mock, True, ["Y", "N"])
def_tax = render_cond(cols_r1[8:11], "Sales Tax:", 'val_tax', db_val_tax, True, ["Y", "N"])

st.markdown("<br>", unsafe_allow_html=True)

cols_r2 = st.columns([1, 1.5, 0.6, 0.6, 0.5, 1.5, 0.6, 0.6, 2.1], gap="small")

c_calc1, c_calc2, c_calc3, c_calc4 = cols_r2[0:4]

if "val_calcs" not in st.session_state:
    st.session_state.val_calcs = db_val_calcs

is_calcs_y = st.session_state.val_calcs != "N" and str(st.session_state.val_calcs).strip() != ""
chk_calc_key = "def_calcs"

def check_calc_match(current, db_val):
    c = str(current).strip().upper()
    d = str(db_val).strip().upper()
    if d == "": return False
    if c == "N" and d == "N": return True
    if c != "N" and d != "N" and c != "":
        return limpiar_numero(c) == limpiar_numero(d)
    return False

if chk_calc_key not in st.session_state: 
    st.session_state[chk_calc_key] = check_calc_match(st.session_state.val_calcs, db_val_calcs)

def on_calc_opt_change():
    opt = st.session_state.get("calc_opt", "N")
    if opt == "N":
        st.session_state.val_calcs = "N"
    else:
        if st.session_state.val_calcs == "N":
            st.session_state.val_calcs = ""
            if "calc_val_input" in st.session_state:
                st.session_state.calc_val_input = ""
    st.session_state[chk_calc_key] = check_calc_match(st.session_state.val_calcs, db_val_calcs)

calc_opt = c_calc1.selectbox("Calcs:", ["Y", "N"], index=0 if is_calcs_y else 1, key="calc_opt", on_change=on_calc_opt_change)

if calc_opt == "Y":
    if "calc_val_input" not in st.session_state:
        st.session_state.calc_val_input = format_currency(st.session_state.val_calcs) if st.session_state.val_calcs != "N" else ""

    def on_calc_val_change():
        format_currency_callback("calc_val_input")
        st.session_state.val_calcs = st.session_state.calc_val_input
        st.session_state[chk_calc_key] = check_calc_match(st.session_state.val_calcs, db_val_calcs)

    c_calc2.text_input("Amount ($):", key="calc_val_input", on_change=on_calc_val_change)
else:
    c_calc2.text_input("Amount ($):", value="N/A", disabled=True, key="calc_val_disabled")
    st.session_state.val_calcs = "N"

with c_calc3:
    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
    def reset_calcs():
        st.session_state.val_calcs = db_val_calcs
        if db_val_calcs != "N" and str(db_val_calcs).strip() != "":
            st.session_state.calc_opt = "Y"
            st.session_state.calc_val_input = format_currency(db_val_calcs)
        else:
            st.session_state.calc_opt = "N"
            if "calc_val_input" in st.session_state:
                st.session_state.calc_val_input = ""
        st.session_state[chk_calc_key] = check_calc_match(db_val_calcs, db_val_calcs)
    st.button("🔄", key="ref_calcs", type="tertiary", on_click=reset_calcs)

current_calc_final = st.session_state.val_calcs

with c_calc4:
    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
    def_cal = st.checkbox("💾", key=chk_calc_key)

c_exec1, c_exec2, c_exec3 = cols_r2[5:8]
chk_exec_key = "def_exec"

if "val_exec" not in st.session_state: st.session_state["val_exec"] = db_val_exec
if chk_exec_key not in st.session_state: st.session_state[chk_exec_key] = (st.session_state["val_exec"] == db_val_exec) and (db_val_exec != "")

val_exec = c_exec1.date_input("Execution Limit Date:", key="val_exec", on_change=cond_change_callback, args=("val_exec", db_val_exec, chk_exec_key, False))

with c_exec2:
    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
    st.button("🔄", key="ref_exec", type="tertiary", on_click=reset_cond_callback, args=("val_exec", db_val_exec, chk_exec_key))
    
with c_exec3:
    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
    def_exec = st.checkbox("💾", key=chk_exec_key)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("📝 **General Notes:**")
st.session_state.gen_notes = st.text_area("General Notes (Optional):", value=st.session_state.get('gen_notes', ''), height=100, placeholder="Enter general notes here. Use '-' for bullets.")

st.divider()

# ==========================================
# 10. GUARDADO Y GENERACIÓN (Nube / Local)
# ==========================================
st.markdown("### 6. Save & Generate Proposal")

col_btn1, col_btn2 = st.columns(2)
with col_btn1:
    btn_save = st.button("💾 Save Complete Proposal", type="primary", use_container_width=True)
with col_btn2:
    btn_save_gen = st.button("💾📄 Save and Generate Proposal", type="secondary", use_container_width=True)

if btn_save or btn_save_gen:
    
    correos_a_validar = [e1_final, e2_final, e3_final, e4_final]
    correos_validos = True
    for c in correos_a_validar:
        if not validar_email(c):
            correos_validos = False
            st.error(f"The email '{c}' is invalid. It must contain '@' and a domain.")
            break

    all_bids_complete = True
    bids_error_msg = ""
    for i, bid in enumerate(st.session_state.bids_list):
        if not bid['name'].strip() or limpiar_numero(bid['amount_str']) == 0:
            all_bids_complete = False
            bids_error_msg += f"- Row {i+1} (Level 1 - {bid['type']}): Missing description or amount.\n"
        for j, child in enumerate(bid['children']):
            if not child['name'].strip() or limpiar_numero(child['amount_str']) == 0:
                all_bids_complete = False
                bids_error_msg += f"- Row {i+1} (Alternative {j+1} - {child['type']}): Missing description or amount.\n"

    all_scopes_valid = True
    has_valid_scope = False
    for m in st.session_state.mat_list:
        for scope in m['scopes']:
            if not scope['desc'].strip() or limpiar_numero(scope['qty']) == 0:
                all_scopes_valid = False
            else:
                has_valid_scope = True

    if not val_ref.strip():
        st.error("❌ Reference cannot be empty.")
    elif not val_customer:
        st.error("❌ You must select or type a Customer.")
    elif not all_bids_complete:
        st.error(f"❌ There are incomplete Bids or Alternatives. Fill in description and amount in all created rows:\n{bids_error_msg}")
    elif not st.session_state.mat_list or not has_valid_scope or not all_scopes_valid:
        st.error("❌ All materials must have their areas (Scopes) with description and quantity > 0.")
    elif not correos_validos:
        pass 
    else:
        try:
            with st.spinner("Saving to database..."):
                # 1. GENERAL
                registros_gen = ws_gen.get_all_records()
                filas_borrar_gen = [idx for idx, r in enumerate(registros_gen, start=2) if str(r.get('ID', '')) == id_final and str(r.get('Proposal', '')) == val_prop and str(r.get('Customer', '')) == val_customer]
                for r_idx in reversed(filas_borrar_gen): ws_gen.delete_rows(r_idx)
                
                fila_gen = [id_final, st.session_state.get('proj_name', ''), val_prop, val_date.strftime("%m/%d/%Y"), val_customer, val_attention, 
                            e1_final.replace("-- Empty --", ""), e2_final.replace("-- Empty --", ""), e3_final.replace("-- Empty --", ""), e4_final.replace("-- Empty --", ""), 
                            val_ref, st.session_state.get('sec_text', ''), val_inv_date.strftime("%m/%d/%Y"), val_comp_date.strftime("%m/%d/%Y"), 
                            st.session_state['val_tax'], st.session_state['val_mock'], current_calc_final, st.session_state['val_validity'], val_exec.strftime("%m/%d/%Y"), st.session_state.gen_notes]
                ws_gen.append_row(fila_gen)
                
                # 2. BIDS
                registros_bids = ws_bids.get_all_records()
                filas_borrar_bids = [idx for idx, r in enumerate(registros_bids, start=2) if str(r.get('ID', '')) == id_final and str(r.get('Proposal', '')) == val_prop]
                for r_idx in reversed(filas_borrar_bids): ws_bids.delete_rows(r_idx)
                
                filas_bids = []
                rel = 1
                for b in st.session_state.bids_list:
                    if b['name'].strip():
                        filas_bids.append([id_final, val_prop, b['name'], b['type'], limpiar_numero(b['amount_str']), rel])
                        for c in b['children']:
                            if c['name'].strip():
                                filas_bids.append([id_final, val_prop, c['name'], c['type'], limpiar_numero(c['amount_str']), rel])
                        rel += 1
                if filas_bids: ws_bids.append_rows(filas_bids)
                
                # 3. MATERIALES
                registros_mat = ws_mat.get_all_records()
                filas_borrar_mat = [idx for idx, r in enumerate(registros_mat, start=2) if str(r.get('ID', '')) == id_final and str(r.get('Proposal', '')) == val_prop]
                for r_idx in reversed(filas_borrar_mat): ws_mat.delete_rows(r_idx)
                
                filas_mat = []
                for m in st.session_state.mat_list:
                    if m['mat_name'].strip():
                        sys_final = m['sys_new'] if m['sys_sel'] == "-- New System --" else (m['sys_sel'] if m['sys_sel'] != "-- None --" else "")
                        diag_final = m['diagram'] if m['diagram'] != "-- None --" else ""
                        
                        parts = []
                        if m.get('t_thick') and m.get('thick'): parts.append(format_fluid_unit(m['thick'], m['u_thick']))
                        dim_parts = []
                        if m.get('t_width') and m.get('width'): dim_parts.append(format_fluid_unit(m['width'], m['u_width']))
                        if m.get('t_length') and m.get('length'): dim_parts.append(f"max {format_fluid_unit(m['length'], m['u_length'])}")
                        if dim_parts: parts.append(" x ".join(dim_parts))
                        parts.append(m['mat_name'])
                        if m.get('t_spec') and m.get('spec'): parts.append(m['spec'])
                        final_fluid = " - ".join(parts).strip()
                        
                        for scope in m['scopes']:
                            combined_item_name = f"{m['item_name']} | Scope: {scope['desc']}" if scope['desc'] != 'General' else m['item_name']
                            
                            filas_mat.append([
                                id_final, val_prop, m['mat_name'], combined_item_name, final_fluid, 
                                limpiar_numero(scope['qty']), scope['unit'], 
                                m['thick'], m['u_thick'], m['spec'], m['width'], m['u_width'], m['length'], m['u_length'], 
                                m['brands'], m['colors'], sys_final, diag_final, m['custom_note'], 
                                m['t_thick'], m['t_spec'], m['t_width'], m['t_length']
                            ])
                        
                        # Actualizar Portafolio
                        try:
                            cell = ws_port.find(m['mat_name'], in_column=1)
                            if cell:
                                row_num = cell.row
                                current_row = ws_port.row_values(row_num)
                                while len(current_row) < 13: current_row.append("")
                                
                                current_row[1] = m['thick'] if m['def_thick'] else ""; current_row[2] = m['u_thick'] if m['def_thick'] else "inches"
                                current_row[3] = m['spec'] if m['def_spec'] else ""
                                current_row[4] = m['brands'] if m['def_brands'] else ""
                                current_row[5] = m['width'] if m['def_width'] else ""; current_row[6] = m['u_width'] if m['def_width'] else "inches"
                                current_row[7] = m['length'] if m['def_length'] else ""; current_row[8] = m['u_length'] if m['def_length'] else "inches"
                                current_row[9] = str(m['t_thick']) if m['def_thick'] else "False"
                                current_row[10] = str(m['t_spec']) if m['def_spec'] else "False"
                                current_row[11] = str(m['t_width']) if m['def_width'] else "False"
                                current_row[12] = str(m['t_length']) if m['def_length'] else "False"
                                
                                ws_port.update(f"A{row_num}:M{row_num}", [current_row])
                            else:
                                ws_port.append_row([
                                    m['mat_name'], 
                                    m['thick'] if m['def_thick'] else "", m['u_thick'] if m['def_thick'] else "inches", 
                                    m['spec'] if m['def_spec'] else "", 
                                    m['brands'] if m['def_brands'] else "",
                                    m['width'] if m['def_width'] else "", m['u_width'] if m['def_width'] else "inches", 
                                    m['length'] if m['def_length'] else "", m['u_length'] if m['def_length'] else "inches", 
                                    str(m['t_thick']) if m['def_thick'] else "False",
                                    str(m['t_spec']) if m['def_spec'] else "False",
                                    str(m['t_width']) if m['def_width'] else "False",
                                    str(m['t_length']) if m['def_length'] else "False"
                                ])
                        except: pass
                            
                        # Actualizar Sistemas
                        if sys_final:
                            try:
                                records_sist = ws_sist.get_all_records()
                                row_to_update = None
                                for idx, r in enumerate(records_sist, start=2):
                                    if str(r.get('Material', '')) == m['mat_name'] and str(r.get('System', '')) == sys_final:
                                        row_to_update = idx
                                        break
                                
                                if row_to_update:
                                    if m.get('def_diag'):
                                        ws_sist.update(f"C{row_to_update}", [[diag_final]])
                                else:
                                    ws_sist.append_row([m['mat_name'], sys_final, diag_final if m.get('def_diag') else "-- None --"])
                            except Exception as e: 
                                print(f"Error saving system: {e}")
                            
                if filas_mat: ws_mat.append_rows(filas_mat)
                
                # 4. GUARDAR CONDICIONES
                try: fila_c = ws_cond.row_values(2) if len(ws_cond.get_all_values()) > 1 else ["30", "N", "N", "N", ""]
                except: fila_c = ["30", "N", "N", "N", ""]
                
                while len(fila_c) < 5: fila_c.append("")
                
                fila_c[0] = st.session_state['val_validity'] if def_val else "30"
                fila_c[1] = current_calc_final if def_cal else "N"
                fila_c[2] = st.session_state['val_mock'] if def_moc else "N"
                fila_c[3] = st.session_state['val_tax'] if def_tax else "N"
                fila_c[4] = st.session_state['val_exec'].strftime("%m/%d/%Y") if st.session_state.get('def_exec') else ""
                
                ws_cond.clear()
                ws_cond.append_row(["Validity", "Calcs", "Mock_Up", "Sales_Tax", "Execution_Limit_Date"])
                ws_cond.append_row(fila_c)
                
                # Refrescar los datos en memoria
                dfs = fetch_all_data()
                st.session_state.df_maestra, st.session_state.df_historial, st.session_state.df_bids = dfs[0], dfs[1], dfs[2]
                st.session_state.df_mat, st.session_state.df_port, st.session_state.df_sist, st.session_state.df_cond = dfs[3], dfs[4], dfs[5], dfs[6]
                
                st.success(f"✅ Proposal saved successfully in Database!")
                st.toast("✅ Proposal saved successfully!", icon="✅")
                
            # ==========================================
            # GENERACIÓN DE DOCUMENTOS (NUBE / LOCAL)
            # ==========================================
            if btn_save_gen:
                try:
                    from docxtpl import DocxTemplate, InlineImage
                    from docx.shared import Mm
                    
                    temp_dir = tempfile.gettempdir()
                    template_path = os.path.join(temp_dir, "Template.docx")
                    
                    with st.spinner("Descargando plantilla y generando documento..."):
                        # Descargar plantilla desde Google Drive
                        drive_service = get_drive_service()
                        request = drive_service.files().get_media(fileId=ID_TEMPLATE)
                        fh = io.FileIO(template_path, 'wb')
                        downloader = MediaIoBaseDownload(fh, request)
                        done = False
                        while done is False:
                            status, done = downloader.next_chunk()
                            
                        context = {
                            'proposal_date': val_date.strftime("%B %d, %Y"),
                            'attention': val_attention,
                            'customer': val_customer,
                            'reference': val_ref,
                            'sections': st.session_state.get('sec_text', ''),
                            'bids': [],
                            'materiales': [],
                            'invitation_date': val_inv_date.strftime("%m/%d/%Y") if val_inv_date else "",
                            'compliance_date': val_comp_date.strftime("%m/%d/%Y") if val_comp_date else "",
                            'texto_calcs': f"Included (${current_calc_final})" if current_calc_final != "N" else "Not Included",
                            'texto_mockup': "Included" if st.session_state['val_mock'] == "Y" else "Not Included",
                            'validity': st.session_state['val_validity'],
                            'exec_limit': val_exec.strftime("%m/%d/%Y") if val_exec else "",
                            'texto_tax': "Included" if st.session_state['val_tax'] == "Y" else "Not Included",
                            'notas_generales': [n.strip().lstrip('-').strip() for n in st.session_state.get('gen_notes', '').split('\n') if n.strip()],
                        }
                        
                        for b in st.session_state.bids_list:
                            if b['name'].strip():
                                context['bids'].append({'descripcion': f"{b['type']}: {b['name']}", 'precio': f"${b['amount_str']}"})
                                for c in b['children']:
                                    if c['name'].strip():
                                        context['bids'].append({'descripcion': f"   {c['type']}: {c['name']}", 'precio': f"${c['amount_str']}"})
                                        
                        doc = DocxTemplate(template_path)
                        
                        for m in st.session_state.mat_list:
                            if m['mat_name'].strip():
                                mat_data = {
                                    'item_name': m['item_name'],
                                    'item_name_str': m['item_name'],
                                    'fluid_title_and_scope': "",
                                    'detalles': [],
                                    'notas': [n.strip().lstrip('-').strip() for n in m.get('custom_note', '').split('\n') if n.strip()],
                                    'imagen': "" 
                                }
                                
                                # Descargar diagrama si existe
                                if m.get('diagram') and m.get('diagram') != "-- None --":
                                    diag_id = st.session_state.get(f"diag_id_{m['diagram']}")
                                    if diag_id:
                                        img_path = os.path.join(temp_dir, m['diagram'])
                                        request_img = drive_service.files().get_media(fileId=diag_id)
                                        fh_img = io.FileIO(img_path, 'wb')
                                        downloader_img = MediaIoBaseDownload(fh_img, request_img)
                                        done_img = False
                                        while done_img is False:
                                            _, done_img = downloader_img.next_chunk()
                                        mat_data['imagen'] = InlineImage(doc, img_path, width=Mm(80))
                                
                                parts = []
                                if m.get('t_thick') and m.get('thick'): parts.append(format_fluid_unit(m['thick'], m['u_thick']))
                                dim_parts = []
                                if m.get('t_width') and m.get('width'): dim_parts.append(format_fluid_unit(m['width'], m['u_width']))
                                if m.get('t_length') and m.get('length'): dim_parts.append(f"max {format_fluid_unit(m['length'], m['u_length'])}")
                                if dim_parts: parts.append(" x ".join(dim_parts))
                                parts.append(m['mat_name'])
                                if m.get('t_spec') and m.get('spec'): parts.append(m['spec'])
                                fluid_title = " - ".join(parts).strip()
                                
                                scopes_str = ", ".join([f"{s['desc']}: {s['qty']} {s['unit']}" for s in m['scopes']])
                                mat_data['fluid_title_and_scope'] = f"{fluid_title} ({scopes_str})"
                                
                                if m.get('thick'): mat_data['detalles'].append(f"Thickness: {format_fluid_unit(m['thick'], m['u_thick'])}")
                                if m.get('spec'): mat_data['detalles'].append(f"Specification: {m['spec']}")
                                if m.get('brands'): mat_data['detalles'].append(f"Brands: {m['brands']}")
                                if m.get('colors'): mat_data['detalles'].append(f"Colors: {m['colors']}")
                                if m.get('sys_sel') and m.get('sys_sel') not in ["-- None --", "-- New System --"]: 
                                    mat_data['detalles'].append(f"System: {m['sys_sel']}")
                                elif m.get('sys_new'):
                                    mat_data['detalles'].append(f"System: {m['sys_new']}")
                                    
                                context['materiales'].append(mat_data)
                        
                        doc.render(context)
                        
                        # Lógica de nombres
                        if str(id_final).startswith("2026"):
                            base_id = str(id_final).replace("2026", "3COREN", 1)
                        else:
                            base_id = f"3COREN - {str(id_final)}"
                            
                        safe_id = re.sub(r'[\\/*?:"<>|]', "", base_id).strip()
                        safe_prop = re.sub(r'[\\/*?:"<>|]', "", str(val_prop)).strip()
                        base_filename = f"{safe_id} - {safe_prop}"
                        
                        docx_filepath = os.path.join(temp_dir, f"{base_filename}.docx")
                        pdf_filepath = os.path.join(temp_dir, f"{base_filename}.pdf")
                        
                        doc.save(docx_filepath)
                        st.success(f"✅ Documento Word generado exitosamente.")
                        
                        # Generación de PDF (Inteligente: Nube vs Local)
                        pdf_generado = False
                        if platform.system() == "Windows":
                            try:
                                from docx2pdf import convert
                                convert(docx_filepath, pdf_filepath)
                                pdf_generado = True
                                st.success(f"✅ Documento PDF generado exitosamente.")
                            except Exception as e:
                                st.warning(f"⚠️ Error generando PDF en Windows: {e}")
                        else:
                            try:
                                subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', docx_filepath, '--outdir', temp_dir], check=True)
                                pdf_generado = True
                                st.success(f"✅ Documento PDF generado exitosamente.")
                            except Exception as e:
                                st.warning(f"⚠️ Error generando PDF en la nube: {e}")
                        
                        # Botones de Descarga Directa
                        st.markdown("### 📥 Descargar Archivos")
                        col_d1, col_d2 = st.columns(2)
                        
                        with open(docx_filepath, "rb") as file:
                            col_d1.download_button(
                                label="📄 Descargar Propuesta (Word)", 
                                data=file, 
                                file_name=f"{base_filename}.docx", 
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                            
                        if pdf_generado:
                            with open(pdf_filepath, "rb") as file:
                                col_d2.download_button(
                                    label="📕 Descargar Propuesta (PDF)", 
                                    data=file, 
                                    file_name=f"{base_filename}.pdf", 
                                    mime="application/pdf",
                                    use_container_width=True
                                )
                                
                except ImportError:
                    st.error("❌ Required libraries for document generation are missing.")
                except Exception as e:
                    st.error(f"❌ Error generating document: {e}")

        except Exception as e: 
            st.error(f"Error saving: {e}")