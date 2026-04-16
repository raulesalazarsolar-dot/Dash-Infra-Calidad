import io
import base64
import json
import os
import unicodedata
import pandas as pd
from urllib.parse import urlparse, unquote
from datetime import datetime
from zoneinfo import ZoneInfo
from PIL import Image

# Librerías de SharePoint
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# ==========================================
# 1. CONFIGURACIÓN (SHAREPOINT + GITHUB)
# ==========================================
SITE_URL = "https://teams.wal-mart.com/sites/EquipoPlanificacin"
LIST_NAME = "Seguimiento Infraestructura"

# ⚠️ SEGURIDAD GITHUB: Lee la clave desde "GitHub Secrets".
USERNAME = os.environ.get("SP_USERNAME", "r0r0noi@cl.wal-mart.com")
PASSWORD = os.environ.get("SP_PASSWORD", "fiXed.sPout+8")

OUTPUT_HTML = "index.html"

# ==========================================
# 2. UTILIDADES Y "SABUESO DE FOTOS"
# ==========================================
def limpiar(val):
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    if s == "0" or s == "0.0": return "0"
    if s.lower() == "nan": return "" 
    return s.replace(".0", "")

def normalizar_texto(texto):
    if not texto: return ""
    s = str(texto).lower().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def formatear_fecha(texto_fecha):
    if pd.isna(texto_fecha) or not texto_fecha or str(texto_fecha).strip() == "": 
        return "--"
    try:
        s = str(texto_fecha).strip().split("T")[0].split(" ")[0]
        s = s.replace("/", "-") 
        p = s.split("-")
        if len(p) == 3:
            if len(p[0]) == 4: return f"{p[2].zfill(2)}-{p[1].zfill(2)}-{p[0]}"
            else: return f"{p[0].zfill(2)}-{p[1].zfill(2)}-{p[2]}"
        return s
    except: return str(texto_fecha)

def descargar_foto_por_url(ctx, url):
    try:
        url = unquote(url)
        if url.startswith("http"): url = urlparse(url).path
        file_content = io.BytesIO()
        ctx.web.get_file_by_server_relative_url(url).download(file_content).execute_query()
        file_content.seek(0)
        if len(file_content.getvalue()) > 0:
            with Image.open(file_content) as img:
                if img.mode != "RGB": img = img.convert("RGB")
                img.thumbnail((400, 400))
                buf = io.BytesIO()
                img.save(buf, format='JPEG', quality=60)
                return f"data:image/jpeg;base64,{base64.b64encode(buf.getvalue()).decode('utf-8')}"
    except Exception: pass
    return None

def extraer_fotos_columna(ctx, p, col_name, item_id):
    imgs_b64 = []
    json_raw = p.get(col_name)
    if json_raw:
        try:
            data = json.loads(json_raw) if isinstance(json_raw, str) else json_raw
            if isinstance(data, dict): data = [data]
            if isinstance(data, list):
                for img_data in data:
                    if isinstance(img_data, dict):
                        url = img_data.get("serverRelativeUrl") or img_data.get("serverUrl") or img_data.get("Url")
                        filename = img_data.get("fileName")
                        b64 = None
                        if url: b64 = descargar_foto_por_url(ctx, url)
                        if not b64 and filename:
                            rel_site = SITE_URL.replace("https://teams.wal-mart.com", "")
                            url_adj = f"{rel_site}/Lists/{LIST_NAME}/Attachments/{item_id}/{filename}"
                            b64 = descargar_foto_por_url(ctx, url_adj)
                        if b64: imgs_b64.append(b64)
        except: pass
    return imgs_b64

# ==========================================
# 3. GENERAR EXCEL CALIDAD
# ==========================================
def generar_excel_calidad_b64(db_json):
    data = []
    resp_str = "Michael Bahamodes (Mantenimiento)\nEvelio Revilla (Supervisor de producción)\nBarbara Gajardo (Gestión industrial)\nLeticia Sánchez (Prevenciòn de riesgo)\nNicolas Hermosilla (ISS)\nPilar Coronado (ISS)\nValentina Romo (Logística)"
    for key, item in db_json.items():
        if item.get('origen') == 'act' and "calidad" in str(item.get('clase', '')).lower():
            obs_full = item.get('observacion1', '')
            if item.get('observacion2', ''): obs_full += f"\n{item.get('observacion2', '')}"
            row = {
                "PLANTA": "MASAS", "Fecha de inspección": item.get('f_lev', ''), "Codigo": "R-ISO05-07",
                "PLANTA2": 2, "RESPONSABLES DE INSPECCIÓN (NOMBRES y ÁREAS)": resp_str,
                "ÁREA A INSPECCIONAR": item.get('ubicacion', ''), "Zonas de foco": "", 
                "CUMPLIMIENTO": item.get('status', '').upper(), "DESCRICIÓN DEL HALLAZGO": item.get('actividad', '') or item.get('titulo', ''),
                "INDICAR SI CORRESPONDE A TEMAS DE: LIMPIEZA Y ORDEN / MANTENIMIENTO / INFRAESTRUCTURA / EQUIPOS": "", 
                "CRITICIDAD (MENOR, MAYOR, CRITICA)": "", "RESPONSABLE DE CIERRE": item.get('ejecutor', ''),  
                "FECHA DE CIERRE": item.get('f_cie', ''), "SAP": "", "OT": item.get('ot', ''), 
                "ESTADO (ABIERTO/EN PROCESO/CERRADO)": item.get('status', '').upper(), "Comentarios": obs_full, "TAG": item.get('tag', '') 
            }
            data.append(row)
    if not data: return None
    df = pd.DataFrame(data)
    cols_orden = ["PLANTA", "Fecha de inspección", "Codigo", "PLANTA2", "RESPONSABLES DE INSPECCIÓN (NOMBRES y ÁREAS)", "ÁREA A INSPECCIONAR", "Zonas de foco", "CUMPLIMIENTO", "DESCRICIÓN DEL HALLAZGO", "INDICAR SI CORRESPONDE A TEMAS DE: LIMPIEZA Y ORDEN / MANTENIMIENTO / INFRAESTRUCTURA / EQUIPOS", "CRITICIDAD (MENOR, MAYOR, CRITICA)", "RESPONSABLE DE CIERRE", "FECHA DE CIERRE", "SAP", "OT", "ESTADO (ABIERTO/EN PROCESO/CERRADO)", "Comentarios", "TAG"]
    for c in cols_orden:
        if c not in df.columns: df[c] = ""
    df = df[cols_orden] 
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Base Calidad')
    return base64.b64encode(output.getvalue()).decode()

# ==========================================
# 4. EXTRACCIÓN PRINCIPAL (SHAREPOINT API)
# ==========================================
def main():
    try:
        print("🚀 INICIANDO EXTRACCIÓN DIRECTA DESDE SHAREPOINT...")
        ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
        sp_list = ctx.web.lists.get_by_title(LIST_NAME)
        
        # CORREGIDO: ClaseMU -> ClaseM
        columnas_req = [
            "ID", "Title", "LinkTitle", "field_1", "field_2", "field_3", 
            "field_4", "field_5", "field_6", "field_7", "field_8", 
            "field_9", "field_10", "field_11", "field_12", "field_14", 
            "field_15", "Antes", "Despues", "Zona", "ClaseM", "Attachments"
        ]
        
        try: items = sp_list.items.select(columnas_req).expand(["AttachmentFiles"]).top(5000).get().execute_query()
        except Exception:
            items = sp_list.items.select(columnas_req).top(5000).get().execute_query()
            
        total_main = len(items)
        print(f"   ✅ Se descargaron {total_main} registros brutos.")
        
        db_act = {}
        for idx, item in enumerate(items):
            print(f"      ... Procesando OT {idx+1} de {total_main}", end='\r')
            p = item.properties
            item_id = int(p.get("ID", 0))
            
            # CORREGIDO: ClaseMU -> ClaseM
            clase_str = limpiar(p.get("ClaseM")).title() or "General"
            semana = limpiar(p.get("field_1"))
            f_lev = formatear_fecha(p.get("field_2"))
            f_cie = formatear_fecha(p.get("field_3"))
            act_str = limpiar(p.get("field_4"))
            ubicacion = limpiar(p.get("field_5"))
            ot = limpiar(p.get("field_7"))
            ejecutor = limpiar(p.get("field_9"))
            prio_raw = normalizar_texto(limpiar(p.get("field_10")))
            status_raw = normalizar_texto(limpiar(p.get("field_11")))
            tag_id = limpiar(p.get("LinkTitle"))
            obs1 = limpiar(p.get("field_14"))
            obs2 = limpiar(p.get("field_15"))

            # --- FILTRO DE CLASES ---
            clase_norm = normalizar_texto(clase_str)
            if not any(x in clase_norm for x in ["calidad", "sanitizacion", "infraestructura"]):
                continue 

            has_asignacion = bool(ejecutor and ejecutor.strip() and ejecutor.lower() != "sin asignar")
            has_ejecutado = any(k in status_raw for k in ['ok', 'listo', 'cerrad', 'realiza', 'complet'])
            is_calidad = "calidad" in clase_str.lower()

            if is_calidad:
                if has_ejecutado: status = "realizada"
                else: status = "pendiente"
            else:
                if has_ejecutado: status = "realizada"
                else: status = "pendiente"

            if "0" in prio_raw or "alta" in prio_raw or "1" in prio_raw: prio = "1"
            elif "media" in prio_raw or "2" in prio_raw: prio = "2"
            else: prio = "3"

            imgs_antes = extraer_fotos_columna(ctx, p, "Antes", item_id)
            imgs_despues = extraer_fotos_columna(ctx, p, "Despues", item_id)

            key_id = f"MTTO_{item_id}"
            db_act[key_id] = {
                "key_id": key_id, "id_real": item_id, "titulo": act_str or tag_id or f"OT #{item_id}", 
                "tag": tag_id, "semana": semana or "S/N", "ejecutor": ejecutor or "Sin Asignar",
                "prioridad": prio, "ubicacion": ubicacion or "Sin Ubicación",
                "ot": ot, "f_lev": f_lev, "f_cie": f_cie, "actividad": act_str, 
                "observacion1": obs1, "observacion2": obs2, "status": status, 
                "has_asignacion": has_asignacion, "clase": clase_str, "origen": "act", 
                "imgs_antes": imgs_antes, "imgs_despues": imgs_despues   
            }
            
        print(f"\n   ✅ Total Actividades mapeadas: {len(db_act)}")
        generar_html_moderno(db_act, "Panel Infraestructura")

    except Exception as e: 
        print(f"\n❌ Error Fatal: {e}")
        import traceback
        traceback.print_exc()

# ==========================================
# 5. GENERADOR HTML
# ==========================================
def generar_html_moderno(db_json, titulo_dashboard):
    if not db_json: return None
    fecha_actual = datetime.now(ZoneInfo("America/Santiago")).strftime("%d/%m/%Y %H:%M")
    
    download_btn = ""
    b64_excel = generar_excel_calidad_b64(db_json)
    if b64_excel:
        download_btn = f'<div id="btn_dl_container"><a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="Base_Calidad.xlsx" class="seg-btn" style="text-decoration:none; display:flex; align-items:center; background:#dcfce7; color:#166534; border:1px solid #166534; border-radius:4px; padding:4px 12px; font-weight:bold; font-size:0.85rem;">📥 Descargar Calidad</a></div>'

    json_seguro = json.dumps(db_json).replace("</", "<\\/")

    full_html = f"""<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Dashboard - {titulo_dashboard}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root {{ --primary: #0f172a; --secondary: #334155; --accent: #2563eb; --bg: #f8fafc; --border: #e2e8f0; --text: #1e293b; --muted: #64748b; --success: #10b981; --warn: #f59e0b; --danger: #ef4444; --info: #3b82f6; }}
        * {{ box-sizing: border-box; outline: none; font-family: 'Segoe UI', system-ui, sans-serif; }}
        body {{ background: transparent; color: var(--text); margin: 0; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }}
        .top-bar {{ background: var(--primary); color: white; padding: 0 20px; height: 60px; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; z-index: 10; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
        .brand h2 {{ margin: 0; font-size: 1.2rem; display:flex; align-items:center; gap: 8px; }} 
        .brand span {{ color: #fbbf24; font-weight: 700; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.5px; }}
        .tabs-container {{ background: #fff; border-bottom: 1px solid var(--border); padding: 0 20px; flex-shrink: 0; display:flex; justify-content: space-between; z-index: 5; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
        .tabs-nav {{ display: flex; gap: 15px; }}
        .tab-btn {{ background: none; border: none; padding: 15px 5px; font-weight: 600; color: var(--muted); cursor: pointer; border-bottom: 3px solid transparent; transition: 0.2s; font-size: 0.95rem; }}
        .tab-btn:hover {{ color: var(--accent); }} .tab-btn.active {{ color: var(--accent); border-bottom-color: var(--accent); }}
        .app-layout {{ display: flex; height: calc(100vh - 110px); width: 100%; overflow: hidden; }}
        .col-filters {{ width: 280px; background: #fff; border-right: 1px solid var(--border); display: flex; flex-direction: column; flex-shrink: 0; z-index: 5; }}
        .filters-header {{ padding: 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--primary); font-size: 0.9rem; text-transform: uppercase; background: #f8fafc; }}
        .filters-body {{ flex: 1; overflow-y: auto; padding: 20px; min-height: 0; }} 
        .filters-footer {{ padding: 20px; border-top: 1px solid var(--border); background: #f8fafc; flex-shrink: 0; }}
        .col-list {{ width: 380px; background: #fff; border-right: 1px solid var(--border); display: flex; flex-direction: column; flex-shrink: 0; }}
        .list-header {{ padding: 20px; border-bottom: 1px solid var(--border); font-weight: 600; background: #f8fafc; color: var(--secondary); font-size: 0.9rem; flex-shrink: 0; display:flex; flex-direction:column; gap:12px; }}
        .list-scroll-area {{ flex: 1; overflow-y: auto; min-height: 0; }}
        .col-detail {{ flex: 1; background: transparent; overflow-y: auto; padding: 40px; }}
        .graficos-layout {{ flex: 1; padding: 30px; display: grid; grid-template-columns: repeat(2, 1fr); gap: 25px; overflow-y: auto; background: transparent; align-content:start; }}
        .detail-content {{ background: white; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); overflow: hidden; max-width: 1000px; margin: 0 auto; border: 1px solid var(--border); }}
        .chart-card {{ background: white; padding: 25px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; height: 400px; width: 100%; }}
        .chart-card.wide {{ grid-column: 1 / -1; height: 480px; }}
        .f-group {{ margin-bottom: 15px; }}
        .f-group label {{ font-size: 0.75rem; font-weight: 700; color: var(--muted); display: block; margin-bottom: 6px; text-transform: uppercase; }}
        select, input {{ width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; color: var(--text); }}
        select:focus, input:focus {{ border-color: var(--accent); box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.1); }}
        .btn-clean {{ background: white; border: 1px solid var(--danger); color: var(--danger); padding: 10px; border-radius: 6px; cursor: pointer; font-weight: 700; transition: 0.2s; margin-top: 10px; width: 100%; text-transform: uppercase; font-size: 0.8rem; letter-spacing: 0.5px; }}
        .btn-clean:hover {{ background: var(--danger); color: white; }}
        .range-box {{ display: flex; align-items: center; gap: 8px; justify-content: space-between; }}
        .range-box select {{ width: 100%; padding: 6px 10px; font-size: 0.8rem; }}
        .range-box span {{ font-size: 0.85rem; color: var(--muted); font-weight: bold; text-transform: lowercase; }}
        .kpi-row-mini {{ display: flex; justify-content: space-between; margin-bottom: 15px; }}
        .kpi-box {{ text-align: center; }} .k-label {{ display: block; font-size: 0.7rem; color: var(--muted); font-weight: 700; }}
        .k-num {{ display: block; font-size: 1.3rem; font-weight: 800; color: var(--primary); }} .k-ok {{ color: var(--success); }} .k-pend {{ color: var(--danger); }}
        .prog-title {{ display: flex; justify-content: space-between; font-size: 0.75rem; font-weight: 700; color: var(--muted); margin-bottom: 6px; }}
        .progress-bar-container {{ width: 100%; height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden; }}
        .progress-bar-fill {{ height: 100%; background: var(--success); width: 0%; transition: width 1s cubic-bezier(0.4, 0, 0.2, 1); }}
        .list-item {{ padding: 15px 20px; border-bottom: 1px solid var(--border); cursor: pointer; transition: 0.2s; border-left: 4px solid transparent; }}
        .list-item:hover {{ background: #f8fafc; }} .list-item.selected {{ background: #eff6ff; border-left-color: var(--accent); }}
        .li-top {{ display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 0.75rem; color: var(--muted); font-weight: 600; }}
        .li-title {{ font-weight: 700; font-size: 0.95rem; color: var(--primary); margin-bottom: 10px; line-height: 1.4; }}
        .li-btm {{ display: flex; justify-content: space-between; font-size: 0.75rem; align-items: center; }}
        .tag {{ padding: 4px 8px; border-radius: 4px; font-weight: 700; font-size: 0.7rem; letter-spacing: 0.3px; }}
        .st-ok {{ background: #dcfce7; color: #166534; }} .st-pend {{ background: #fee2e2; color: #991b1b; }} .st-prog {{ background: #e0f2fe; color: #075985; }} .st-proc {{ background: #fef3c7; color: #92400e; }}
        .empty-state {{ display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: var(--muted); opacity: 0.7; }}
        .detail-header {{ padding: 30px; border-bottom: 1px solid var(--border); background: #fff; }}
        .dh-top {{ display: flex; justify-content: space-between; margin-bottom: 15px; align-items:center; }}
        .detail-header h2 {{ margin: 0 0 5px 0; font-size: 1.6rem; color: var(--primary); }}
        @keyframes pulse-ring {{ 0% {{ box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }} 70% {{ box-shadow: 0 0 0 10px rgba(16, 185, 129, 0); }} 100% {{ box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }} }}
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
        <div class="brand"><h2>🏭 Seguimiento Mantenimiento <span>{titulo_dashboard}</span></h2></div>
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
        <div style="display:flex; gap:10px;">
            <button onclick="descargarExcel()" class="btn-clean" style="margin: 0; padding: 8px 15px; width: auto; border-color: #10b981; color: #10b981; display: flex; align-items: center; gap: 8px;">
                <span style="font-size:1.2rem;">📊</span> Exportar Excel
            </button>
        </div>
    </div>
    
    <div class="app-layout">
        <div class="col-filters" id="main_filters">
            <div class="filters-header">
                <span>🔍 Filtros</span>
                <button onclick="resetFilters()" class="btn-clean" style="margin: 0; padding: 4px 8px; width: auto; font-size: 0.7rem; border-color: #ef4444; color: #ef4444; display: flex; align-items: center; gap: 4px; text-transform: none;">
                    🧹 Borrar
                </button>
            </div>
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
                    <input type="text" id="search_input" placeholder="🔍 Buscar TAG o Título..." onkeyup="applyFilters()" style="width: 100%; padding: 8px 12px; border: 1px solid var(--border); border-radius: 6px; font-family: inherit; font-size: 0.8rem; outline: none;">
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
        html += `<div class="f-group"><label>📅 Semana</label><div class="range-box"><select id="f_s1" onchange="applyFilters()"></select><span>a</span><select id="f_s2" onchange="applyFilters()"></select></div></div>`;
        html += createSelect('f_mes', '🗓️ Mes Levantamiento', months);
        html += createSelect('f_clase', '🏷️ Clase', [...new Set(records.map(x=>x.clase))].sort());
        html += createSelect('f_ubi', '📍 Ubicación / Área', [...new Set(records.map(x=>x.ubicacion))].sort());
        html += `<div class="f-group"><label>🚨 Prioridad / Criticidad</label><select id="f_prio" onchange="applyFilters()"><option value="ALL">Todas</option><option value="1">🚨 Crítica</option><option value="2">🟡 Mayor</option><option value="3">🟢 Menor</option></select></div>`;
        
        if (appState.view === 'charts') {{
            html += `<div class="f-group"><label>📋 Estado</label><select id="f_status" onchange="applyFilters()"><option value="ALL">Todos</option><option value="ok">✅ Realizadas / Cerradas</option><option value="pend">⚠️ Pendientes / En Proceso</option></select></div>`;
        }}

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
        const uEl = document.getElementById('f_ubi'); const uVal = uEl ? uEl.value : 'ALL';
        const pEl = document.getElementById('f_prio'); const pVal = pEl ? pEl.value : 'ALL';
        const cEl = document.getElementById('f_clase'); const cVal = cEl ? cEl.value : 'ALL';
        const stEl = document.getElementById('f_status'); const stVal = stEl ? stEl.value : 'ALL';
        const searchEl = document.getElementById('search_input'); const searchVal = searchEl ? searchEl.value.toLowerCase().trim() : '';
        const mEl = document.getElementById('f_mes'); const mVal = mEl ? mEl.value : 'ALL';

        return records.filter(d => {{
            if (appState.view === 'charts') {{
                if (stVal !== 'ALL') {{
                    const isOk = (d.status === 'realizada' || d.status === 'cerrada');
                    if (stVal === 'ok' && !isOk) return false;
                    if (stVal === 'pend' && isOk) return false;
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
                if(d.f_lev && d.f_lev !== '--' && d.f_lev.includes('-')) {{
                    let p = d.f_lev.split('-');
                    if(p.length >= 3) dMes = p[1] + '-' + p[2];
                }}
                if (dMes !== mVal) return false;
            }}

            let miClase = d.clase || '';
            if (cVal !== 'ALL' && miClase !== cVal) return false;
            if (uVal !== 'ALL' && d.ubicacion !== uVal) return false;
            if (pVal !== 'ALL' && d.prioridad !== pVal) return false;
            
            return true;
        }});
    }}

    function applyFilters() {{
        currentChartData = getFilteredData();
        let ok = 0; let pre = 0;
        currentChartData.forEach(d => {{ 
            if (d.status === 'realizada' || d.status === 'cerrada') ok++; 
            else if (d.status === 'precierre') pre++;
        }});
        const total = currentChartData.length;
        
        document.getElementById('k_total').innerText = total;
        document.getElementById('k_ok').innerText = ok;
        if(document.getElementById('k_pre')) document.getElementById('k_pre').innerText = pre;
        
        let pendCount = total - ok - pre;
        document.getElementById('k_pend').innerText = pendCount;
        
        let perc = total > 0 ? Math.round((ok/total)*100) : 0;
        document.getElementById('k_perc').innerText = perc + '%';
        const bar = document.getElementById('bar_fill');
        bar.style.width = perc + '%';
        bar.style.backgroundColor = perc > 80 ? '#10b981' : (perc > 40 ? '#f59e0b' : '#ef4444');

        if(appState.view === 'list') renderList(currentChartData);
        else drawCharts(currentChartData);
    }}

    function renderList(data) {{
        const container = document.getElementById('list_container');
        container.innerHTML = '';
        
        data.forEach(d => {{
            const item = document.createElement('div');
            item.className = 'list-item';
            item.onclick = function() {{ 
                renderDetail(d.key_id); 
                document.querySelectorAll('.list-item').forEach(i=>i.classList.remove('selected'));
                item.classList.add('selected');
            }};
            
            let stText = '⚠️ PEND'; let stClass = 'st-pend';
            if (d.status === 'realizada' || d.status === 'cerrada') {{ stText='✅ REALIZADA'; stClass='st-ok'; }}
            else if (d.status === 'precierre') {{ stText='🔍 PRECIERRE'; stClass='st-proc'; }}
            else if (d.status === 'programado') {{ stText='📅 PROG'; stClass='st-prog'; }}
            else if (d.status === 'en proceso' || d.status === 'tratando') {{ stText='🔨 PROCESO'; stClass='st-proc'; }}
            else {{ stText='📂 PEND'; stClass='st-pend'; }}
            
            let idDisplay = d.ot ? `OT: ${{d.ot}}` : (d.tag ? d.tag : '#' + d.id_real);
            
            item.innerHTML = `
                <div class="li-top"><span>${{idDisplay}}</span><span>Sem: ${{d.semana}}</span></div>
                <div class="li-title">${{d.titulo}}</div>
                <div class="li-btm">
                    <span class="tag ${{stClass}}">${{stText}}</span>
                    <span style="color:var(--muted); font-weight:700;">👷 ${{d.ejecutor.split(' ')[0]}}</span>
                </div>
            `;
            container.appendChild(item);
        }});
    }}

    function renderDetail(key) {{
        document.getElementById('empty_state').style.display='none';
        document.getElementById('detail_view').style.display='block';
        const d = db[key];
        
        document.getElementById('d_title').innerText = d.titulo;
        document.getElementById('d_tag').innerText = d.tag ? `TAG: ${{d.tag}}` : (d.ot ? `OT: ${{d.ot}}` : 'Sin ID');
        
        const stBadge = document.getElementById('d_status');
        if (d.status === 'realizada' || d.status === 'cerrada') {{ stBadge.innerText = '✅ REALIZADA'; stBadge.className = 'tag st-ok'; }}
        else if (d.status === 'precierre') {{ stBadge.innerText = '🔍 PRECIERRE'; stBadge.className = 'tag st-proc'; }}
        else if (d.status === 'programado') {{ stBadge.innerText = '📅 PROGRAMADA'; stBadge.className = 'tag st-prog'; }}
        else if (d.status === 'en proceso' || d.status === 'tratando') {{ stBadge.innerText = '🔨 EN PROCESO'; stBadge.className = 'tag st-proc'; }}
        else {{ stBadge.innerText = '⚠️ PENDIENTE'; stBadge.className = 'tag st-pend'; }}
        
        let pl = d.prioridad;
        if(pl==='1') pl='<span class="prio-flag p-crit">🚨 CRÍTICA</span>';
        else if(pl==='2') pl='<span class="prio-flag p-med">🟡 MAYOR</span>';
        else pl='<span class="prio-flag p-baja">🟢 MENOR</span>';
        document.getElementById('d_prio_lbl').innerHTML = pl;

        if (d.clase && d.clase.toLowerCase().includes('calidad')) {{
            document.getElementById('ticket_progress_wrapper').style.display = 'block';
            let stepsActive = 1; 
            if (d.has_asignacion) stepsActive = 2;
            if (d.has_ejecutado) stepsActive = 3;
            if (d.has_cierre) stepsActive = 4;
            let fillWidth = ((stepsActive - 1) / 3) * 80;
            document.getElementById('ticket_step_fill').style.width = fillWidth + '%';
            
            for(let i=1; i<=4; i++) {{
                let el = document.getElementById('t_step_' + i);
                el.classList.remove('active', 'current');
                if (i <= stepsActive) el.classList.add('active');
                if (i === stepsActive) el.classList.add('current');
            }}
        }} else {{
            document.getElementById('ticket_progress_wrapper').style.display = 'none';
        }}

        const grid = document.getElementById('d_grid');
        grid.innerHTML = '';
        const createItem = (label, val) => `<div class="dg-item"><small>${{label}}</small><strong>${{val||'--'}}</strong></div>`;
        grid.innerHTML += createItem('🏷️ Clase', d.clase);
        grid.innerHTML += createItem('👷 Responsable', d.ejecutor);
        grid.innerHTML += createItem('📍 Ubicación', d.ubicacion);
        grid.innerHTML += createItem('📅 Levantamiento', d.f_lev);
        grid.innerHTML += createItem('🏁 Cierre', d.f_cie);
        grid.innerHTML += createItem('📆 Semana', d.semana);
        if (d.ot) grid.innerHTML += createItem('⚙️ OT SAP', d.ot);
        
        document.getElementById('box_obs1').style.display = 'block';
        document.getElementById('lbl_obs_title').innerText = '📝 Observación';
        document.getElementById('d_obs').innerText = d.observacion1 || 'Sin observaciones';
        
        if (d.observacion2) {{
            document.getElementById('box_obs2').style.display = 'block';
            document.getElementById('lbl_obs_title2').innerText = '📝 Observación 2';
            document.getElementById('d_obs2').innerText = d.observacion2;
        }} else {{
            document.getElementById('box_obs2').style.display = 'none';
        }}
        
        if ((d.imgs_antes && d.imgs_antes.length > 0) || (d.imgs_despues && d.imgs_despues.length > 0)) {{
            document.getElementById('card_img_single').style.display = 'none';
            document.getElementById('card_img_a').style.display = 'block';
            document.getElementById('card_img_d').style.display = 'block';
            setupCarousel('d_img_a', d.imgs_antes || []);
            setupCarousel('d_img_d', d.imgs_despues || []);
            document.getElementById('d_gallery_sec').style.display = 'flex';
        }} else {{
            document.getElementById('d_gallery_sec').style.display = 'none';
        }}
    }}

    function setupCarousel(elementId, images) {{
        const container = document.getElementById(elementId);
        container.innerHTML = '';
        if (!images || images.length === 0) {{
            container.innerHTML = '<div style="height:100px; display:flex; align-items:center; justify-content:center; color:#ccc; font-style:italic; border: 1px dashed #cbd5e1; border-radius: 4px; width: 100%;">Sin evidencia</div>';
            return;
        }}
        carousels[elementId] = {{ idx: 0, imgs: images }};
        renderCarousel(elementId);
    }}

    function renderCarousel(id) {{
        const c = carousels[id];
        const container = document.getElementById(id);
        const currentSrc = c.imgs[c.idx];
        let navHtml = '';
        if (c.imgs.length > 1) {{
            navHtml = `
                <button class="nav-btn nav-prev" onclick="moveCarousel('${{id}}', -1)">❮</button>
                <button class="nav-btn nav-next" onclick="moveCarousel('${{id}}', 1)">❯</button>
                <div class="img-counter">${{c.idx + 1}} / ${{c.imgs.length}}</div>
            `;
        }}
        container.innerHTML = `<img src="${{currentSrc}}" class="gal-img" onclick="openModal(this.src)">${{navHtml}}`;
    }}

    window.moveCarousel = function(id, dir) {{
        const c = carousels[id];
        c.idx += dir;
        if (c.idx < 0) c.idx = c.imgs.length - 1;
        if (c.idx >= c.imgs.length) c.idx = 0;
        renderCarousel(id);
        event.stopPropagation();
    }}

    function openModal(src) {{
        document.getElementById('modalImg').src = src;
        document.getElementById('modal').style.display = 'flex';
    }}

    function showDataModal(title, filterFn) {{
        let html = `<div class="dm-header">
            <h3>📊 Desglose: ${{title}}</h3>
            <button class="dm-close" onclick="document.getElementById('data_modal').style.display='none'">&times;</button>
        </div>
        <div class="dm-body">
            <table class="dm-table">
                <thead><tr><th>ID / TAG</th><th>Semana</th><th>Título / Actividad</th><th>Responsable</th><th>Estado</th><th>Prioridad</th></tr></thead>
                <tbody>`;

        let found = false;
        currentChartData.forEach(d => {{
            if (filterFn(d)) {{
                found = true;
                let stColor = (d.status==='realizada' || d.status==='cerrada') ? '#166534' : (d.status==='pendiente' || d.status==='abierta' ? '#991b1b' : '#92400e');
                let idDisplay = d.ot ? d.ot : (d.tag ? d.tag : '#' + d.id_real);
                let pText = 'MENOR'; let pColor = '#64748b';
                if(d.prioridad==='1') {{ pText='🚨 CRÍTICA'; pColor='#dc2626'; }}
                else if(d.prioridad==='2') {{ pText='🟡 MAYOR'; pColor='#d97706'; }}

                html += `<tr onclick="document.getElementById('data_modal').style.display='none'; document.getElementById('btn_tab_list').click(); setTimeout(() => renderDetail('${{d.key_id}}'), 100);">
                    <td style="font-weight:700;">${{idDisplay}}</td>
                    <td>${{d.semana}}</td>
                    <td>${{d.titulo}}</td>
                    <td>${{d.ejecutor.split(' ')[0]}}</td>
                    <td style="color:${{stColor}}; font-weight:700; text-transform:uppercase;">${{d.status}}</td>
                    <td style="color:${{pColor}}; font-weight:700;">${{pText}}</td>
                </tr>`;
            }}
        }});

        if (!found) html += `<tr><td colspan="6" style="text-align:center; padding: 30px; color:var(--muted);">No hay registros para esta selección</td></tr>`;
        html += `</tbody></table></div>`;
        document.getElementById('data_modal_content').innerHTML = html;
        document.getElementById('data_modal').style.display = 'flex';
    }}

    function getFreshCanvas(id) {{
        const old = document.getElementById(id);
        if(!old) return null;
        const container = old.parentElement;
        container.innerHTML = `<canvas id="${{id}}"></canvas>`;
        return document.getElementById(id);
    }}

    function descargarExcel() {
        if (!currentChartData || currentChartData.length === 0) {
            alert("No hay datos para exportar.");
            return;
        }
        const datosExcel = currentChartData.map(d => ({
            "Levantamiento": d.f_lev, "Cierre": d.f_cie, "Actividad": d.actividad, "Clase": d.clase,
            "Zona": d.zona, "Ubicación": d.ubicacion, "OT": d.ot, "Ejecutor": d.ejecutor,
            "Status": d.status.toUpperCase(), "Semana": d.semana
        }));
        const worksheet = XLSX.utils.json_to_sheet(datosExcel);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Datos");
        XLSX.writeFile(workbook, `Reporte.xlsx`);
    }

    function drawCharts(data) {{
        if(!data || data.length === 0) return;
        const chartIds = ['chart1', 'chart2', 'chart4', 'chart5', 'chart6'];
        chartIds.forEach(id => {{ if (chartInstances[id]) {{ chartInstances[id].destroy(); chartInstances[id] = null; }} }});

        let stats = {{ ok:0, pend:0, pre:0, prog:0, loc:{{}}, wCounts:{{}}, cCounts:{{}}, mCounts:{{}} }};
        weeks.forEach(w => stats.wCounts[w] = {{total:0, ok:0, pre:0}});
        
        data.forEach(d => {{
            let isOk = (d.status === 'realizada' || d.status === 'cerrada');
            let isPre = (d.status === 'precierre');
            if(isOk) stats.ok++; else if(isPre) stats.pre++; else stats.pend++;
            
            let miClase = d.clase || 'General';
            stats.cCounts[miClase] = (stats.cCounts[miClase]||0)+1;

            const l = d.ubicacion || 'Sin Ubicación';
            if(!stats.loc[l]) stats.loc[l]={{total:0, critical:0}};
            stats.loc[l].total++;
            if(d.prioridad==='1') stats.loc[l].critical++;
            
            let catMensual = null; let sortKey = null;
            if (d.f_lev && d.f_lev !== '--' && d.f_lev.includes('-')) {{
                let p = d.f_lev.split('-'); 
                if (p.length >= 3) {{
                    let y = parseInt(p[2]); let m = p[1];
                    if (y < 2026) {{ catMensual = 'Arrastre ' + y; sortKey = y + '-00'; }} 
                    else {{ catMensual = m + '-' + y; sortKey = y + '-' + m; }}
                }}
            }}
            if (catMensual) {{
                if (!stats.mCounts[catMensual]) stats.mCounts[catMensual] = {{ total:0, ok:0, pre:0, sortKey: sortKey }};
                stats.mCounts[catMensual].total++;
                if(isOk) stats.mCounts[catMensual].ok++;
                if(isPre) stats.mCounts[catMensual].pre++;
            }}
        }});

        const commonOpts = {{ maintainAspectRatio:false, responsive:true, animation: {{ duration: 800, easing: 'easeOutQuart' }}, plugins: {{ datalabels: {{ display: false }}, legend: {{ labels: {{ usePointStyle: true, boxWidth: 8 }} }} }} }};
        
        let finalLabelsMeses = Object.keys(stats.mCounts).sort((a, b) => stats.mCounts[a].sortKey.localeCompare(stats.mCounts[b].sortKey));
        const c6DataLevantadas = finalLabelsMeses.map(m => stats.mCounts[m].total);
        const c6DataCerradas = finalLabelsMeses.map(m => stats.mCounts[m].ok);
        const c6DataPrecierre = finalLabelsMeses.map(m => stats.mCounts[m].pre);

        chartInstances['chart6'] = new Chart(getFreshCanvas('chart6'), {{
            type: 'line',
            data: {{
                labels: finalLabelsMeses,
                datasets: [
                    {{ label: 'Levantadas', data: c6DataLevantadas, borderColor: '#3b82f6', backgroundColor: 'rgba(59, 130, 246, 0.2)', borderWidth: 2, fill: true, tension: 0.4 }},
                    {{ label: 'Precierre', data: c6DataPrecierre, borderColor: '#f59e0b', backgroundColor: 'rgba(245, 158, 11, 0.2)', borderWidth: 2, fill: true, tension: 0.4 }},
                    {{ label: 'Cerradas', data: c6DataCerradas, borderColor: '#10b981', backgroundColor: 'rgba(16, 185, 129, 0.2)', borderWidth: 2, fill: true, tension: 0.4 }}
                ]
            }},
            options: {{ ...commonOpts, interaction: {{ mode: 'index', intersect: false }}, scales: {{ x: {{ grid: {{ display: false }} }}, y: {{ beginAtZero: true }} }} }}
        }});

        chartInstances['chart1'] = new Chart(getFreshCanvas('chart1'), {{ 
            type: 'doughnut', 
            data: {{ labels: ['Cerradas','Precierre','Pendientes'], datasets: [{{ data: [stats.ok, stats.pre, stats.pend], backgroundColor: ['#10b981','#f59e0b','#ef4444'], borderWidth: 2, borderColor: '#fff' }}] }}, 
            options: {{ ...commonOpts, cutout: '70%' }}
        }});
        
        chartInstances['chart2'] = new Chart(getFreshCanvas('chart2'), {{ 
            type: 'pie', 
            data: {{ labels:Object.keys(stats.cCounts), datasets:[{{ data:Object.values(stats.cCounts), backgroundColor:['#3b82f6','#8b5cf6','#ec4899','#14b8a6','#f97316','#d946ef','#f59e0b'] }}] }}, 
            options: commonOpts
        }});

        const sortedLocs = Object.entries(stats.loc).sort((a,b)=>b[1].total - a[1].total).slice(0, 20); 
        const labelsLoc = sortedLocs.map(x=>x[0]);
        const dataCounts = sortedLocs.map(x=>x[1].total);
        let sumAcc = 0; const totalHallazgos = dataCounts.reduce((a,b)=>a+b, 0);
        const dataAcumulado = dataCounts.map(count => {{ sumAcc += count; return totalHallazgos > 0 ? parseFloat(((sumAcc / totalHallazgos) * 100).toFixed(1)) : 0; }});

        chartInstances['chart4'] = new Chart(getFreshCanvas('chart4'), {{
            type: 'bar',
            data: {{
                labels: labelsLoc,
                datasets: [
                    {{ type: 'line', label: '% Acumulado', data: dataAcumulado, borderColor: '#ef4444', yAxisID: 'yPercentage' }},
                    {{ type: 'bar', label: 'Cantidad de Hallazgos', data: dataCounts, backgroundColor: 'rgba(59, 130, 246, 0.7)', yAxisID: 'yCount' }}
                ]
            }},
            options: {{
                ...commonOpts, interaction: {{ mode: 'index', intersect: false }}, 
                scales: {{
                    yCount: {{ type: 'linear', position: 'left' }},
                    yPercentage: {{ type: 'linear', position: 'right', min: 0, max: 100 }},
                    x: {{ ticks: {{ autoSkip: false, maxRotation: 45, minRotation: 45 }} }}
                }}
            }}
        }});

        const scatterData = Object.entries(stats.loc).filter(([name, stat]) => stat.total > 0).map(([name, stat]) => ({{ x: stat.total, y: stat.critical, label: name }}));
        chartInstances['chart5'] = new Chart(getFreshCanvas('chart5'), {{
            type: 'scatter',
            data: {{ datasets: [{{ label: 'Ubicaciones', data: scatterData, backgroundColor: 'rgba(239, 68, 68, 0.6)' }}] }},
            options: {{ ...commonOpts }}
        }});
    }}

    const initAntigravity = () => {{
        const canvas = document.createElement('canvas');
        canvas.id = 'antigravity-bg';
        document.body.prepend(canvas);
        const ctx = canvas.getContext('2d');
        canvas.style.position = 'fixed'; canvas.style.top = '0'; canvas.style.left = '0'; canvas.style.width = '100vw'; canvas.style.height = '100vh'; canvas.style.zIndex = '-1'; canvas.style.pointerEvents = 'none'; canvas.style.backgroundColor = '#f8fafc';
        let particles = []; const colors = ['#4285F4', '#EA4335', '#FBBC05', '#34A853', '#A0C3FF', '#FCA297'];
        let mouse = {{ x: null, y: null, radius: 120 }};
        window.addEventListener('mousemove', (e) => {{ mouse.x = e.x; mouse.y = e.y; }});
        window.addEventListener('mouseout', () => {{ mouse.x = undefined; mouse.y = undefined; }});
        window.addEventListener('resize', () => {{ canvas.width = window.innerWidth; canvas.height = window.innerHeight; initParticles(); }});
        class Particle {{
            constructor(x, y) {{ this.x = x; this.y = y; this.baseX = x; this.baseY = y; this.size = Math.random() * 2 + 1.5; this.color = colors[Math.floor(Math.random() * colors.length)]; this.density = (Math.random() * 20) + 2; }}
            draw() {{ ctx.fillStyle = this.color; ctx.beginPath(); ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2); ctx.closePath(); ctx.fill(); }}
            update() {{
                let dx = mouse.x - this.x; let dy = mouse.y - this.y; let distance = Math.sqrt(dx * dx + dy * dy);
                if (distance < mouse.radius) {{ let force = (mouse.radius - distance) / mouse.radius; this.x -= (dx / distance) * force * this.density; this.y -= (dy / distance) * force * this.density; }} 
                else {{ if (this.x !== this.baseX) this.x -= (this.x - this.baseX) / 15; if (this.y !== this.baseY) this.y -= (this.y - this.baseY) / 15; }}
                this.draw();
            }}
        }}
        function initParticles() {{ particles = []; canvas.width = window.innerWidth; canvas.height = window.innerHeight; let numberOfParticles = (canvas.width * canvas.height) / 7000; for (let i = 0; i < numberOfParticles; i++) particles.push(new Particle(Math.random() * canvas.width, Math.random() * canvas.height)); }}
        function animateParticles() {{ ctx.clearRect(0, 0, canvas.width, canvas.height); for (let i = 0; i < particles.length; i++) particles[i].update(); requestAnimationFrame(animateParticles); }}
        initParticles(); animateParticles();
    }};

    window.onload = () => {{
        buildFilters(); applyFilters(); initAntigravity();
    }};
    </script>
</body></html>
    """
    
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f: f.write(full_html)
    print(f"✅ REPORTE {titulo_dashboard} GUARDADO CON ÉXITO")

if __name__ == "__main__":
    main()
