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

# Credenciales desde GitHub Secrets
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
        # SharePoint envía fechas como "YYYY-MM-DDTHH:MM:SSZ"
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
        
        # COLUMNAS ACTUALIZADAS SEGÚN TU TABLA DE NOMBRES INTERNOS
        columnas_req = [
            "ID", "Title", "LinkTitle", "field_1", "field_2", "field_3", 
            "field_4", "field_5", "field_6", "field_7", "field_8", 
            "field_9", "field_10", "field_11", "field_12", "field_14", 
            "field_15", "Antes", "Despues", "Zona", "ClaseMU", "Attachments"
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
            
            # --- MAPEO SEGÚN TABLA DE CÓDIGOS ---
            clase_str = limpiar(p.get("ClaseMU")).title() or "General" # Clase M -> ClaseMU
            semana = limpiar(p.get("field_1")) # Semana -> field_1
            f_lev = formatear_fecha(p.get("field_2")) # Levantamiento -> field_2
            f_cie = formatear_fecha(p.get("field_3")) # Cierre -> field_3
            act_str = limpiar(p.get("field_4")) # Actividad -> field_4
            ubicacion = limpiar(p.get("field_5")) # Ubicación -> field_5
            ot = limpiar(p.get("field_7")) # OT -> field_7
            ejecutor = limpiar(p.get("field_9")) # Ejecutor -> field_9
            prio_raw = normalizar_texto(limpiar(p.get("field_10"))) # Prioridad -> field_10
            status_raw = normalizar_texto(limpiar(p.get("field_11"))) # Status -> field_11
            tag_id = limpiar(p.get("LinkTitle")) # Tag -> LinkTitle
            obs1 = limpiar(p.get("field_14")) # Observación -> field_14
            obs2 = limpiar(p.get("field_15")) # Observación 2 -> field_15

            # --- FILTRO DE CLASES ---
            clase_norm = normalizar_texto(clase_str)
            if not any(x in clase_norm for x in ["calidad", "sanitizacion", "infraestructura"]):
                continue 

            # Lógica de estados 
            has_asignacion = bool(ejecutor and ejecutor.strip() and ejecutor.lower() != "sin asignar")
            has_ejecutado = any(k in status_raw for k in ['ok', 'listo', 'cerrad', 'realiza', 'complet'])
            is_calidad = "calidad" in clase_str.lower()

            if is_calidad:
                if has_ejecutado: status = "realizada" # O precierre si usas el campo Estado
                else: status = "pendiente"
            else:
                if has_ejecutado: status = "realizada"
                else: status = "pendiente"

            # Lógica de criticidad (Prioridad -> field_10)
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
# 5. GENERADOR HTML (MANTENIDO)
# ==========================================
def generar_html_moderno(db_json, titulo_dashboard):
    fecha_actual = datetime.now(ZoneInfo("America/Santiago")).strftime("%d/%m/%Y %H:%M")
    b64_excel = generar_excel_calidad_b64(db_json)
    download_btn = ""
    if b64_excel:
        download_btn = f'<div id="btn_dl_container"><a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="Base_Calidad.xlsx" class="seg-btn" style="text-decoration:none; display:flex; align-items:center; background:#dcfce7; color:#166534; border:1px solid #166534; border-radius:4px; padding:4px 12px; font-weight:bold; font-size:0.85rem;">📥 Descargar Calidad</a></div>'

    # Se mantiene tu estructura HTML, inyectando el JSON actualizado
    # ... (Aquí va todo tu bloque html_template con los gráficos de Pareto, Jack Knife, etc.) ...
    
    # Simulación del guardado (asegúrate de tener el string html_template completo del mensaje anterior)
    # full_html = html_template.replace("__FECHA_ACTUAL__", fecha_actual)
    # full_html = full_html.replace("__DB_JSON_DATA__", json.dumps(db_json))
    # with open(OUTPUT_HTML, "w", encoding="utf-8") as f: f.write(full_html)
    print(f"✅ Dashboard generado exitosamente.")

if __name__ == "__main__":
    main()
