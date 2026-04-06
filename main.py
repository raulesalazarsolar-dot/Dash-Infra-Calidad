import io
import base64
import json
import os
import time
import unicodedata
import pandas as pd
from urllib.parse import urlparse, unquote
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ==========================================
# 1. CONFIGURACIÓN GITHUB ACTIONS
# ==========================================
SITE_URL = "https://teams.wal-mart.com/sites/EquipoPlanificacin"
LIST_NAME = "Seguimiento Infraestructura"

# Credenciales seguras desde GitHub Secrets
USERNAME = os.environ.get("SP_USER")
PASSWORD = os.environ.get("SP_PASS") 

# Salida directa para GitHub Pages
OUTPUT_HTML = "index.html"

CACHE_FOLDER = "fotos_infra_cache"
if not os.path.exists(CACHE_FOLDER): os.makedirs(CACHE_FOLDER)

# ==========================================
# 2. UTILIDADES Y PROCESAMIENTO DE FOTOS
# ==========================================
def limpiar(val):
    if val is None: return ""
    s = str(val).strip()
    if s == "0" or s == "0.0": return "0"
    if s.lower() == "nan": return "" 
    return s.replace(".0", "")

def normalizar_texto(texto):
    if not texto: return ""
    s = str(texto).lower().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def formatear_fecha(texto_fecha):
    if not texto_fecha: return "--"
    try:
        s_fecha = str(texto_fecha)
        if "T" in s_fecha: return datetime.strptime(s_fecha.split("T")[0], "%Y-%m-%d").strftime("%d-%m-%Y")
        if isinstance(texto_fecha, datetime): return texto_fecha.strftime("%d-%m-%Y")
        if " " in s_fecha: return s_fecha.split(" ")[0]
        return s_fecha
    except: return str(texto_fecha)

def extraer_dato_seguro(properties, key):
    val = properties.get(key)
    if val is None: return ""
    if isinstance(val, dict): return str(val.get('Title') or val.get('Value') or val.get('Description') or "").strip()
    if isinstance(val, list): return ", ".join([str(v.get('Title') or v.get('Value') or "") for v in val])
    return limpiar(val)

def procesar_foto_attachment(ctx, item_id, json_raw):
    """Descarga la foto y la devuelve en base64 en su TAMAÑO ORIGINAL"""
    if not json_raw: return None
    try:
        data = json.loads(json_raw) if isinstance(json_raw, str) else json_raw
        filename = data.get("fileName")
        if not filename: return None
        rel_site = SITE_URL.replace("https://teams.wal-mart.com", "")
        url = f"{rel_site}/Lists/{LIST_NAME}/Attachments/{item_id}/{filename}"
        local = os.path.join(CACHE_FOLDER, f"ID_{item_id}_{filename.replace(' ', '_')}")
        
        if not os.path.exists(local) or os.path.getsize(local) == 0:
            with open(local, "wb") as f:
                ctx.web.get_file_by_server_relative_url(url).download(f).execute_query()
        
        if os.path.exists(local) and os.path.getsize(local) > 0:
            with open(local, "rb") as image_file:
                # Se lee directamente el archivo sin redimensionar ni usar PIL
                encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
                return f"data:image/jpeg;base64,{encoded_string}"
    except: return None
    return None

# ==========================================
# 3. GENERAR EXCEL CALIDAD
# ==========================================
def generar_excel_calidad_b64(db_json):
    data = []
    resp_str = "Michael Bahamodes (Mantenimiento)\nEvelio Revilla (Supervisor de producción)\nBarbara Gajardo (Gestión industrial)\nLeticia Sánchez (Prevenciòn de riesgo)\nNicolas Hermosilla (ISS)\nPilar Coronado (ISS)\nValentina Romo (Logística)"
    
    for key, item in db_json.items():
        if item.get('origen') == 'act' and "calidad" in str(item.get('clase', '')).lower():
            obs_full = item.get('observacion1', '')
            if item.get('observacion2', ''):
                obs_full += f"\n{item.get('observacion2', '')}"
                
            row = {
                "PLANTA": "MASAS", "Fecha de inspección": item.get('f_lev', ''), "Codigo": "R-ISO05-07", "PLANTA2": 2,
                "RESPONSABLES DE INSPECCIÓN (NOMBRES y ÁREAS)": resp_str, "ÁREA A INSPECCIONAR": item.get('ubicacion', ''),
                "Zonas de foco": "", "CUMPLIMIENTO": item.get('status', ''),
                "DESCRICIÓN DEL HALLAZGO": item.get('actividad', '') or item.get('titulo', ''),
                "INDICAR SI CORRESPONDE A TEMAS DE: LIMPIEZA Y ORDEN / MANTENIMIENTO / INFRAESTRUCTURA / EQUIPOS": "", 
                "CRITICIDAD (MENOR, MAYOR, CRITICA)": "", "RESPONSABLE DE CIERRE": "", "FECHA DE CIERRE": "", 
                "SAP": "", "OT": item.get('ot', ''), "ESTADO (ABIERTO/EN PROCESO/CERRADO)": "", 
                "Comentarios": obs_full, "TAG": item.get('tag', '') 
            }
            data.append(row)
            
    if not data: return None

    df = pd.DataFrame(data)
    cols_orden = ["PLANTA", "Fecha de inspección", "Codigo", "PLANTA2", "RESPONSABLES DE INSPECCIÓN (NOMBRES y ÁREAS)", "ÁREA A INSPECCIONAR", "Zonas de foco", "CUMPLIMIENTO", "DESCRICIÓN DEL HALLAZGO", "INDICAR SI CORRESPONDE A TEMAS DE: LIMPIEZA Y ORDEN / MANTENIMIENTO / INFRAESTRUCTURA / EQUIPOS", "CRITICIDAD (MENOR, MAYOR, CRITICA)", "RESPONSABLE DE CIERRE", "FECHA DE CIERRE", "SAP", "OT", "ESTADO (ABIERTO/EN PROCESO/CERRADO)", "Comentarios", "TAG"]
    for c in cols_orden:
        if c not in df.columns: df[c] = ""
    df = df[cols_orden] 
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Base Calidad')
    return base64.b64encode(output.getvalue()).decode()

# ==========================================
# 4. GENERADOR HTML 
# ==========================================
def generar_html_moderno(db_json, titulo_dashboard):
    if not db_json:
        print(f"⚠️ No hay datos para generar.")
        return None

    print(f"\n🔨 Construyendo archivo HTML...")
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    download_btn = ""
    b64_excel = generar_excel_calidad_b64(db_json)
    if b64_excel:
        download_btn = f'<div id="btn_dl_container"><a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="Base_Calidad.xlsx" class="seg-btn" style="text-decoration:none; display:flex; align-items:center; background:#dcfce7; color:#166534; border:1px solid #166534; border-radius:4px; padding:4px 12px; font-weight:bold; font-size:0.85rem;">📥 Descargar Calidad</a></div>'

    json_seguro = json.dumps(db_json).replace("</", "<\\/")

    full_html = f"""<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Dashboard - {titulo_dashboard}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <style>
        :root {{ --primary: #0f172a; --secondary: #334155; --accent: #2563eb; --bg: #f8fafc; --border: #e2e8f0; --text: #1e293b; --muted: #64748b; --success: #10b981; --warn: #f59e0b; --danger: #ef4444; --info: #3b82f6; }}
        * {{ box-sizing: border-box; outline: none; font-family: 'Segoe UI', system-ui, sans-serif; }}
        body {{ background: var(--bg); color: var(--text); margin: 0; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }}
        
        .top-bar {{ background: var(--primary); color: white; padding: 0 20px; height: 60px; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; z-index: 10; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
        .brand h2 {{ margin: 0; font-size: 1.2rem; display:flex; align-items:center; gap: 8px; }} 
        .brand span {{ opacity: 0.7; font-weight: 300; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.5px; }}
        
        .tabs-container {{ background: white; border-bottom: 1px solid var(--border); padding: 0 20px; flex-shrink: 0; display:flex; justify-content: space-between; z-index: 5; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
        .tabs-nav {{ display: flex; gap: 15px; }}
        .tab-btn {{ background: none; border: none; padding: 15px 5px; font-weight: 600; color: var(--muted); cursor: pointer; border-bottom: 3px solid transparent; transition: 0.2s; font-size: 0.95rem; }}
        .tab-btn:hover {{ color: var(--accent); }} .tab-btn.active {{ color: var(--accent); border-bottom-color: var(--accent); }}
        
        .app-layout {{ display: flex; height: calc(100vh - 110px); width: 100%; overflow: hidden; }}
        
        .col-filters {{ width: 280px; background: #fff; border-right: 1px solid var(--border); display: flex; flex-direction: column; flex-shrink: 0; z-index: 5; }}
        .filters-header {{ padding: 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--primary); font-size: 0.9rem; text-transform: uppercase; background: #f8fafc; }}
        .filters-body {{ flex: 1; overflow-y: auto; padding: 20px; min-height: 0; }} 
        .filters-footer {{ padding: 20px; border-top: 1px solid var(--border); background: #f8fafc; flex-shrink: 0; }}
        
        .f-group {{ margin-bottom: 15px; }}
        .f-group label {{ font-size: 0.75rem; font-weight: 700; color: var(--muted); display: block; margin-bottom: 6px; text-transform: uppercase; }}
        select, input {{ width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; color: var(--text); }}
        select:focus, input:focus {{ border-color: var(--accent); box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.1); }}
        .btn-clean {{ background: white; border: 1px solid var(--danger); color: var(--danger); padding: 10px; border-radius: 6px; cursor: pointer; font-weight: 700; transition: 0.2s; margin-top: 10px; width: 100%; text-transform: uppercase; font-size: 0.8rem; letter-spacing: 0.5px; }}
        .btn-clean:hover {{ background: var(--danger); color: white; }}

        .kpi-row-mini {{ display: flex; justify-content: space-between; margin-bottom: 15px; }}
        .kpi-box {{ text-align: center; }} .k-label {{ display: block; font-size: 0.7rem; color: var(--muted); font-weight: 700; }}
        .k-num {{ display: block; font-size: 1.3rem; font-weight: 800; color: var(--primary); }} .k-ok {{ color: var(--success); }} .k-pend {{ color: var(--danger); }}
        .prog-title {{ display: flex; justify-content: space-between; font-size: 0.75rem; font-weight: 700; color: var(--muted); margin-bottom: 6px; }}
        .progress-bar-container {{ width: 100%; height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden; }}
        .progress-bar-fill {{ height: 100%; background: var(--success); width: 0%; transition: width 1s cubic-bezier(0.4, 0, 0.2, 1); }}
        
        .col-list {{ width: 380px; background: #fff; border-right: 1px solid var(--border); display: flex; flex-direction: column; flex-shrink: 0; }}
        .list-header {{ padding: 20px; border-bottom: 1px solid var(--border); font-weight: 600; background: #f8fafc; color: var(--secondary); font-size: 0.9rem; flex-shrink: 0; display:flex; flex-direction:column; gap:12px; }}
        .list-scroll-area {{ flex: 1; overflow-y: auto; min-height: 0; }}
        
        .list-item {{ padding: 15px 20px; border-bottom: 1px solid var(--border); cursor: pointer; transition: 0.2s; border-left: 4px solid transparent; }}
        .list-item:hover {{ background: #f8fafc; }} .list-item.selected {{ background: #eff6ff; border-left-color: var(--accent); }}
        .li-top {{ display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 0.75rem; color: var(--muted); font-weight: 600; }}
        .li-title {{ font-weight: 700; font-size: 0.95rem; color: var(--primary); margin-bottom: 10px; line-height: 1.4; }}
        .li-btm {{ display: flex; justify-content: space-between; font-size: 0.75rem; align-items: center; }}
        
        .tag {{ padding: 4px 8px; border-radius: 4px; font-weight: 700; font-size: 0.7rem; letter-spacing: 0.3px; }}
        .st-ok {{ background: #dcfce7; color: #166534; }} .st-pend {{ background: #fee2e2; color: #991b1b; }} .st-prog {{ background: #e0f2fe; color: #075985; }} .st-proc {{ background: #fef3c7; color: #92400e; }}
        
        .col-detail {{ flex: 1; background: #f1f5f9; overflow-y: auto; padding: 40px; }}
        .empty-state {{ display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: var(--muted); opacity: 0.7; }}
        .detail-content {{ background: white; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); overflow: hidden; max-width: 1000px; margin: 0 auto; border: 1px solid var(--border); }}
        .detail-header {{ padding: 30px; border-bottom: 1px solid var(--border); background: #fff; }}
        .dh-top {{ display: flex; justify-content: space-between; margin-bottom: 15px; align-items:center; }}
        .detail-header h2 {{ margin: 0 0 5px 0; font-size: 1.6rem; color: var(--primary); }}
        
        @keyframes pulse-ring {{
            0% {{ box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }}
            70% {{ box-shadow: 0 0 0 10px rgba(16, 185, 129, 0); }}
            100% {{ box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }}
        }}
        .ticket-step-container {{ display: flex; justify-content: space-between; position: relative; width: 100%; max-width: 700px; margin: 0 auto; padding: 25px 0; }}
        .ticket-step-bg {{ position: absolute; top: 40px; left: 10%; right: 10%; height: 4px; background: #e2e8f0; z-index: 1; border-radius: 2px; }}
        .ticket-step-fill {{ position: absolute; top: 40px; left: 10%; height: 4px; background: var(--success); z-index: 2; transition: width 0.6s ease; border-radius: 2px; }}
        .ticket-step {{ display: flex; flex-direction: column; align-items: center; z-index: 3; width: 80px; text-align: center; }}
        .ticket-step-circle {{ width: 34px; height: 34px; border-radius: 50%; background: #fff; border: 3px solid #e2e8f0; display: flex; align-items: center; justify-content: center; font-size: 14px; font-weight: bold; color: var(--muted); margin-bottom: 8px; transition: all 0.4s ease; }}
        .ticket-step.active .ticket-step-circle {{ border-color: var(--success); background: var(--success); color: white; }}
        .ticket-step.current .ticket-step-circle {{ animation: pulse-ring 2s infinite; border-color: var(--success); background: var(--success); color: white; }}
        .ticket-step-label {{ font-size: 0.75rem; font-weight: 700; color: var(--muted); transition: color 0.4s ease; text-transform: uppercase; }}
        .ticket-step.active .ticket-step-label, .ticket-step.current .ticket-step-label {{ color: var(--primary); }}

        .data-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 25px; padding: 30px; background: #fff; border-bottom: 1px solid var(--border); }}
        .dg-item small {{ display: block; font-size: 0.7rem; color: var(--muted); font-weight: 700; margin-bottom: 6px; text-transform: uppercase; }}
        .dg-item strong {{ font-size: 1rem; color: var(--text); }}
        
        .obs-box {{ padding: 30px; border-bottom: 1px solid var(--border); }}
        .obs-box h4 {{ margin: 0 0 12px; color: var(--secondary); font-size: 0.9rem; text-transform: uppercase; }}
        .obs-box p {{ background: #f8fafc; padding: 20px; border-radius: 8px; border: 1px solid var(--border); margin: 0; line-height: 1.6; color: #334155; }}
        
        .gallery-section {{ padding: 30px; background: #f8fafc; display:flex; flex-direction: column; gap: 20px; }}
        .photo-card {{ flex: 1; text-align: center; border: 1px solid var(--border); border-radius: 8px; padding: 15px; background: #fff; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }}
        .pc-head {{ font-weight: 700; color: var(--secondary); margin-bottom: 10px; font-size: 0.85rem; text-transform: uppercase; }}
        .gal-img {{ max-width: 100%; max-height: 300px; object-fit: contain; border-radius: 8px; cursor: zoom-in; transition: transform 0.2s; }}
        .gal-img:hover {{ transform: scale(1.02); }}
        
        .carousel-wrapper {{ position: relative; height: 300px; width: 100%; display: flex; justify-content: center; align-items: center; }}
        .nav-btn {{ position: absolute; top: 50%; transform: translateY(-50%); background: rgba(0,0,0,0.5); color: white; border: none; padding: 8px 12px; cursor: pointer; border-radius: 50%; font-size: 1.2rem; transition: 0.2s; user-select: none; z-index: 10; }}
        .nav-btn:hover {{ background: rgba(0,0,0,0.8); }}
        .nav-prev {{ left: 5px; }} .nav-next {{ right: 5px; }}
        .img-counter {{ position: absolute; bottom: 5px; right: 5px; background: rgba(0,0,0,0.6); color: white; padding: 2px 6px; border-radius: 4px; font-size: 0.7rem; z-index: 10; }}

        .graficos-layout {{ flex: 1; padding: 30px; display: grid; grid-template-columns: repeat(2, 1fr); gap: 25px; overflow-y: auto; background: #f1f5f9; align-content:start; }}
        .chart-card {{ background: white; padding: 25px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; height: 400px; width: 100%; }}
        .chart-card.wide {{ grid-column: 1 / -1; height: 480px; }}
        .chart-title {{ font-size: 1rem; font-weight: 700; color: var(--secondary); margin-bottom: 15px; text-transform: uppercase; text-align: center; letter-spacing: 0.5px; }}
        .canvas-container {{ position: relative; flex: 1 1 auto; width: 100%; min-height: 0; }}
        
        .prio-flag {{ padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 0.75rem; }}
        .p-crit {{ background: #fee2e2; color: #dc2626; border: 1px solid #f87171; }}
        .p-alta {{ background: #ffedd5; color: #ea580c; border: 1px solid #fdba74; }}
        .p-med {{ background: #fef3c7; color: #d97706; border: 1px solid #fcd34d; }}
        .p-baja {{ background: #f1f5f9; color: #64748b; border: 1px solid #cbd5e1; }}
        
        .modal {{ display: none; position: fixed; z-index: 2000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(15, 23, 42, 0.85); align-items: center; justify-content: center; backdrop-filter: blur(4px); }}
        .modal img {{ max-width: 90%; max-height: 90vh; border-radius: 8px; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); }}
        
        #data_modal_content {{ background: white; width: 90%; max-width: 1200px; max-height: 85vh; border-radius: 12px; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); }}
        .dm-header {{ padding: 20px 25px; background: var(--primary); color: white; display: flex; justify-content: space-between; align-items: center; }}
        .dm-header h3 {{ margin: 0; font-size: 1.2rem; font-weight: 600; }}
        .dm-close {{ background: none; border: none; color: white; font-size: 1.8rem; cursor: pointer; opacity: 0.8; transition: 0.2s; line-height: 1; }}
        .dm-close:hover {{ opacity: 1; transform: scale(1.1); }}
        .dm-body {{ padding: 0; overflow-y: auto; flex: 1; background: var(--bg); }}
        .dm-table {{ width: 100%; border-collapse: collapse; background: white; font-size: 0.9rem; text-align: left; }}
        .dm-table th {{ background: #f8fafc; padding: 15px 20px; font-weight: 700; color: var(--secondary); border-bottom: 2px solid var(--border); position: sticky; top: 0; z-index: 10; text-transform: uppercase; font-size: 0.8rem; }}
        .dm-table td {{ padding: 15px 20px; border-bottom: 1px solid var(--border); color: var(--text); }}
        .dm-table tr {{ transition: background 0.2s; }}
        .dm-table tr:hover td {{ background: #eff6ff; cursor: pointer; }}
    </style>
</head>
<body>
    <div id="modal" class="modal" onclick="if(event.target===this) this.style.display='none'"><img id="modalImg"></div>
    <div id="data_modal" class="modal" onclick="if(event.target===this) this.style.display='none'">
        <div id="data_modal_content"></div>
    </div>

    <div class="top-bar">
        <div class="brand"><h2>🏭 Seguimiento de actividades <span>{titulo_dashboard}</span></h2></div>
        <div style="font-size:0.85rem; font-weight:600; opacity:0.9; display:flex; gap:15px; align-items:center;">
            {download_btn}
            <span>Actualizado: {fecha_actual}</span>
        </div>
    </div>
    
    <div class="tabs-container">
        <div class="tabs-nav">
            <button class="tab-btn active" onclick="setTab('all', this)" id="btn_tab_list">General</button>
            <button class="tab-btn" onclick="setTab('pendiente', this)">Pendientes</button>
            <button class="tab-btn" onclick="setTab('precierre', this)">Precierre</button>
            <button class="tab-btn" onclick="setTab('realizada', this)">Realizadas</button>
            <button class="tab-btn" onclick="setView('charts', this)">📊 Gráficos</button>
        </div>
    </div>
    
    <div class="app-layout">
        <div class="col-filters" id="main_filters">
            <div class="filters-header">🔍 Filtros Principales</div>
            <div class="filters-body" id="filters_dynamic"></div>
            <div class="filters-footer">
                <div class="kpi-row-mini">
                    <div class="kpi-box"><span class="k-label">TOTAL</span><span class="k-num" id="k_total">0</span></div>
                    <div class="kpi-box"><span class="k-label">CERRADAS</span><span class="k-num k-ok" id="k_ok">0</span></div>
                    <div class="kpi-box"><span class="k-label">PRECIERRE</span><span class="k-num" style="color:#f59e0b" id="k_pre">0</span></div>
                    <div class="kpi-box"><span class="k-label">PEND</span><span class="k-num k-pend" id="k_pend">0</span></div>
                </div>
                <div class="prog-title"><span>Cumplimiento Global</span><span id="k_perc">0%</span></div>
                <div class="progress-bar-container"><div id="bar_fill" class="progress-bar-fill"></div></div>
            </div>
        </div>

        <div id="view_list" style="display:flex; flex:1; overflow:hidden;">
            <div class="col-list">
                <div class="list-header" style="display: flex; flex-direction: column; gap: 10px;">
                    <div>📋 Registros</div>
                    <input type="text" id="search_input" placeholder="🔍 Buscar TAG o Título..." onkeyup="applyFilters()" style="width: 100%; padding: 8px 12px; border: 1px solid var(--border); border-radius: 6px; font-family: inherit; font-size: 0.8rem; outline: none; transition: border-color 0.2s;">
                </div>
                <div id="list_container" class="list-scroll-area"></div>
            </div>
            <div class="col-detail">
                <div id="empty_state" class="empty-state"><div style="font-size:4rem; margin-bottom:15px;">📋</div><h3 style="margin:0">Selecciona un registro</h3></div>
                <div id="detail_view" class="detail-content" style="display:none">
                    <div class="detail-header">
                        <div class="dh-top">
                            <div><span id="d_status" class="tag st-ok">STATUS</span></div>
                            <div id="d_prio_lbl">PRIO</div>
                        </div>
                        <h2 id="d_title">Título</h2>
                        <p style="color:var(--accent); font-weight: 600; font-size: 1.05rem; margin:0;" id="d_tag">TAG</p>
                    </div>
                    
                    <div id="ticket_progress_wrapper" style="display:none; padding: 10px 40px; background: #f8fafc; border-bottom: 1px solid var(--border);">
                        <div class="ticket-step-container">
                            <div class="ticket-step-bg"></div>
                            <div id="ticket_step_fill" class="ticket-step-fill"></div>
                            <div id="t_step_1" class="ticket-step"><div class="ticket-step-circle">🔍</div><div class="ticket-step-label">Hallazgo</div></div>
                            <div id="t_step_2" class="ticket-step"><div class="ticket-step-circle">👷</div><div class="ticket-step-label">Asignación</div></div>
                            <div id="t_step_3" class="ticket-step"><div class="ticket-step-circle">🔨</div><div class="ticket-step-label">Ejecutado</div></div>
                            <div id="t_step_4" class="ticket-step"><div class="ticket-step-circle">🏁</div><div class="ticket-step-label">Cierre</div></div>
                        </div>
                    </div>

                    <div class="data-grid" id="d_grid"></div>
                    
                    <div class="obs-box" id="box_obs1"><h4 id="lbl_obs_title">📝 Observación</h4><p id="d_obs">--</p></div>
                    <div class="obs-box" id="box_obs2" style="display:none;"><h4 id="lbl_obs_title2">📝 Observación 2</h4><p id="d_obs2">--</p></div>
                    
                    <div class="gallery-section" id="d_gallery_sec" style="display:none;">
                        <div style="display: flex; gap: 20px; width: 100%; justify-content: center; flex-wrap: wrap;">
                            <div class="photo-card" id="card_img_a" style="display:none;">
                                <div class="pc-head">📸 ANTES / EVIDENCIA</div>
                                <div id="d_img_a" class="carousel-wrapper"></div>
                            </div>
                            <div class="photo-card" id="card_img_d" style="display:none;">
                                <div class="pc-head">📸 DESPUÉS / CIERRE</div>
                                <div id="d_img_d" class="carousel-wrapper"></div>
                            </div>
                        </div>
                        <div class="photo-card" id="card_img_single" style="display:none; width: 100%;">
                            <div class="pc-head">📸 REGISTRO FOTOGRÁFICO</div>
                            <div id="d_img_single" class="carousel-wrapper"></div>
                        </div>
                    </div>

                </div>
            </div>
        </div>

        <div id="view_charts" class="graficos-layout" style="display:none;">
            <div class="chart-card wide"><div class="chart-title">Tendencia de Generación Mensual</div><div class="canvas-container"><canvas id="chart6"></canvas></div></div>
            <div class="chart-card"><div class="chart-title">Status General</div><div class="canvas-container"><canvas id="chart1"></canvas></div></div>
            <div class="chart-card"><div class="chart-title">Clasificación / Naturaleza</div><div class="canvas-container"><canvas id="chart2"></canvas></div></div>
            <div class="chart-card wide"><div class="chart-title">Gestión por Responsable</div><div class="canvas-container"><canvas id="chart3"></canvas></div></div>
            <div class="chart-card wide"><div class="chart-title">Pareto: Frecuencia de Hallazgos por Ubicación</div><div class="canvas-container"><canvas id="chart4"></canvas></div></div>
            <div class="chart-card wide"><div class="chart-title">Jack Knife: Frecuencia Total vs Críticos por Ubicación</div><div class="canvas-container"><canvas id="chart5"></canvas></div></div>
        </div>
    </div>

    <script>
    if (typeof ChartDataLabels !== 'undefined') {{
        Chart.register(ChartDataLabels);
        Chart.defaults.plugins.datalabels.display = false;
    }}

    const db = {json_seguro};
    const records = Object.values(db).sort((a,b) => b.id_real - a.id_real);
    const weeks = [...new Set(records.map(x=>x.semana).filter(x=>x!=="S/N"))].sort((a,b)=>{{ let na=parseInt(a), nb=parseInt(b); return (isNaN(na)||isNaN(nb)) ? a.localeCompare(b) : na-nb; }});
    
    const months = [...new Set(records.map(x => {{
        if (x.f_lev && x.f_lev !== '--' && x.f_lev.includes('-')) {{
            let p = x.f_lev.split('-');
            if(p.length >= 3) return p[1] + '-' + p[2];
        }}
        return 'Sin Fecha';
    }}))].filter(x => x !== 'Sin Fecha').sort();
    months.push('Sin Fecha');
    
    let appState = {{ statusFilter: 'all', view: 'list' }};
    let currentChartData = [];
    let chartInstances = {{}};
    let carousels = {{}};
    
    Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
    Chart.defaults.color = '#64748b';

    function buildFilters() {{
        const fDiv = document.getElementById('filters_dynamic');
        
        const createSelect = (id, label, options) => {{
            let sel = `<div class="f-group"><label>${{label}}</label><select id="${{id}}" onchange="applyFilters()">`;
            sel += `<option value="ALL">Todos</option>`;
            options.forEach(o => {{ if(o) sel += `<option value="${{o}}">${{o}}</option>`; }});
            sel += `</select></div>`;
            return sel;
        }};

        let html = '';
        html += `<div class="f-group"><label>Semana</label><div class="range-box"><select id="f_s1" onchange="applyFilters()"></select><span>a</span><select id="f_s2" onchange="applyFilters()"></select></div></div>`;
        html += createSelect('f_mes', 'Mes Levantamiento', months);
        html += createSelect('f_clase', 'Clase', [...new Set(records.map(x=>x.clase))].sort());
        html += createSelect('f_exec', 'Ejecutor', [...new Set(records.map(x=>x.ejecutor))].sort());
        html += createSelect('f_ubi', 'Ubicación / Área', [...new Set(records.map(x=>x.ubicacion))].sort());
        html += `<div class="f-group"><label>Prioridad</label><select id="f_prio" onchange="applyFilters()"><option value="ALL">Todas</option><option value="calavera">☠️ Muy Crítica</option><option value="0">🚩 Alta / Crítica</option><option value="1">⚠️ Media</option><option value="2">🟢 Baja</option></select></div>`;
        
        if (appState.view === 'charts') {{
            html += `<div class="f-group"><label>Estado</label><select id="f_status" onchange="applyFilters()"><option value="ALL">Todos</option><option value="ok">✅ Realizadas / Cerradas</option><option value="pend">⚠️ Pendientes / En Proceso</option></select></div>`;
        }}

        html += `<button class="btn-clean" onclick="resetFilters()">🔄 Limpiar Filtros</button>`;
        fDiv.innerHTML = html;

        const s1 = document.getElementById('f_s1');
        const s2 = document.getElementById('f_s2');
        if(s1 && s2) {{
            weeks.forEach((w, i) => {{ s1.add(new Option(w, i)); s2.add(new Option(w, i)); }});
            if(weeks.length > 0) s2.value = weeks.length - 1;
        }}
    }}

    function resetFilters() {{
        const searchInput = document.getElementById('search_input');
        if (searchInput) searchInput.value = '';
        document.querySelectorAll('.f-group select').forEach(sel => sel.value = "ALL");
        if(weeks.length > 0) {{
            document.getElementById('f_s1').value = 0;
            document.getElementById('f_s2').value = weeks.length - 1;
        }}
        applyFilters();
    }}

    function setTab(status, btn) {{
        appState.statusFilter = status;
        appState.view = 'list';
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        if(btn) btn.classList.add('active');
        
        document.getElementById('view_list').style.display = 'flex';
        document.getElementById('view_charts').style.display = 'none';

        buildFilters();
        applyFilters();
    }}

    function setView(view, btn) {{
        appState.view = view;
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        if(btn) btn.classList.add('active');
        
        document.getElementById('view_list').style.display = 'none';
        document.getElementById('view_charts').style.display = 'none';

        if (view === 'list') {{
            document.getElementById('view_list').style.display = 'flex';
            buildFilters();
            applyFilters();
        }} else if (view === 'charts') {{
            document.getElementById('view_charts').style.display = 'grid';
            buildFilters();
            applyFilters();
        }}
    }}

    function getFilteredData() {{
        const s1 = document.getElementById('f_s1') ? document.getElementById('f_s1').selectedIndex : 0;
        const s2 = document.getElementById('f_s2') ? document.getElementById('f_s2').selectedIndex : 999;
        
        const eEl = document.getElementById('f_exec');
        const eVal = eEl ? eEl.value : 'ALL';
        
        const uEl = document.getElementById('f_ubi');
        const uVal = uEl ? uEl.value : 'ALL';
        
        const pEl = document.getElementById('f_prio');
        const pVal = pEl ? pEl.value : 'ALL';
        
        const cEl = document.getElementById('f_clase');
        const cVal = cEl ? cEl.value : 'ALL';
        
        const stEl = document.getElementById('f_status');
        const stVal = stEl ? stEl.value : 'ALL';
        
        const searchEl = document.getElementById('search_input');
        const searchVal = searchEl ? searchEl.value.toLowerCase().trim() : '';

        const mEl = document.getElementById('f_mes');
        const mVal = mEl ? mEl.value : 'ALL';

        return records.filter(d => {{
            if (appState.view === 'charts') {{
                if (stVal !== 'ALL') {{
                    const isOk = (d.status === 'realizada' || d.status === 'cerrada');
                    if (stVal === 'ok' && !isOk) return false;
                    if (stVal === 'pendientes' && isOk) return false;
                }}
            }} else {{
                if (appState.statusFilter !== 'all') {{
                    if (appState.statusFilter === 'pendiente') {{
                        if (d.status !== 'pendiente' && d.status !== 'programado' && d.status !== 'tratando' && d.status !== 'abierta') return false;
                    }} else if (appState.statusFilter === 'realizada') {{
                        if (d.status !== 'realizada' && d.status !== 'cerrada') return false;
                    }} else {{
                        if (d.status !== appState.statusFilter) return false;
                    }}
                }}
            }}
            
            if (searchVal !== '') {{
                const text = `${{d.titulo}} ${{d.ot || ''}} ${{d.tag}}`.toLowerCase();
                if (!text.includes(searchVal)) return false;
            }}

            const wIdx = weeks.indexOf(d.semana);
            if (wIdx !== -1 && (wIdx < s1 || wIdx > s2)) return false;
            
            if (mVal !== 'ALL') {{
                let dMes = 'Sin Fecha';
                if(d.f_lev && d.f_lev !== '--' && d.f_lev