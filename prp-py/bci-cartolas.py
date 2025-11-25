import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
import time
import random
import os
import shutil
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import re

# --------------------------------------------------------------------------- #
# CONFIGURACIÓN GENERAL
# --------------------------------------------------------------------------- #
RUT               = "19913454k"
CLAVE             = "Prin2233"
URL_LOGIN         = "https://www.bci.cl/empresas"

# 🚩 TIEMPOS (10 HORAS DURA EL CICLO - TIEMPO DE ESPERA EN CADA INTERVALO 5 MINUTOS)
TIEMPO_TOTAL_HORAS       = 10   
INTERVALO_ESPERA_MINUTOS = 5    

# 🚩 RUTAS
ROOT_BCI          = r"C:\Users\bmolinac\Desktop\CARTOLAS BANCARIAS\BCI"
DOWNLOAD_DIR      = os.path.join(os.path.expanduser('~'), 'Downloads') 

EMPRESA_EXCLUIDA  = "CLINICA VESPUCIO SPA"

# --------------------------------------------------------------------------- #
# SELECTORES
# --------------------------------------------------------------------------- #
X_BTN_LOGIN     = (By.CLASS_NAME, "btn_login")
X_INPUT_RUT     = (By.ID, "rut_aux")
X_INPUT_CLAVE   = (By.ID, "clave_aux")
X_BTN_SUBMIT    = (By.XPATH, "//button[@type='submit']")

SEL_BOX_EMPRESA = (By.CSS_SELECTOR, ".box-grupo") 

X_DROPDOWN_EMPRESA_TRIGGER = (By.CSS_SELECTOR, "bci-select-search#selector-empresa mat-select")
X_OPCIONES_EMPRESA_ITEM    = (By.TAG_NAME, "mat-option")

X_MENU_CUENTAS       = (By.XPATH, "//div[@class='icon cuentas']/ancestor::mat-list-item")
X_MENU_CORRIENTE     = (By.XPATH, "//div[contains(text(),'Cuentas corrientes')]/ancestor::mat-list-item")
X_MENU_CARTOLA       = (By.XPATH, "//div[contains(text(),'Cartola Histórica')]/ancestor::mat-list-item")

ID_TABLA_CARTOLA     = "historical-balances-tabla-cartola"
X_TRES_PUNTOS        = ".//i[@id='historical-balances-ver-cartola']"
X_BOTON_EXCEL        = (By.XPATH, "//button[contains(., 'Descarga cartola Excel')]")
X_BOTON_PDF          = (By.XPATH, "//button[contains(., 'Descarga cartola PDF')]")
X_MODAL_CONTAINER     = (By.TAG_NAME, "mat-dialog-container")

X_SELECT_CUENTAS     = (By.CSS_SELECTOR, "mat-select#selector-search[aria-label='Cuenta selecionada'] .mat-select-trigger")
X_OPCIONES_CUENTA    = (By.CSS_SELECTOR, "mat-option")

# --------------------------------------------------------------------------- #
# UTILIDADES
# --------------------------------------------------------------------------- #
SPEED_FACTOR = 0.5

def esperar(min_seg=2, max_seg=4):
    time.sleep(random.uniform(max(min_seg * SPEED_FACTOR, 0.5), max(max_seg * SPEED_FACTOR, 0.5)))

def click_robusto(driver, elemento):
    """Intenta 3 métodos de clic para vencer a Angular"""
    try:
        ActionChains(driver).move_to_element(elemento).click().perform()
        return True
    except:
        try:
            elemento.click()
            return True
        except:
            try:
                driver.execute_script("arguments[0].click();", elemento)
                return True
            except:
                return False

def click_right_blank(driver):
    try: driver.execute_script("document.body.click();")
    except: pass

def click_viewer_close(driver):
    try: 
        driver.execute_script("var b=document.getElementById('viewer-hb-icono-cerrar'); if(b){b.click();}")
        click_right_blank(driver)
    except: pass

def limpiar_nombre(texto):
    return re.sub(r'[\\/*?:"<>|]', "", texto.strip()).strip()

def obtener_ruta_dinamica(empresa_nombre, cuenta_texto):
    empresa_clean = limpiar_nombre(empresa_nombre)
    texto_upper = cuenta_texto.upper()
    es_dolar = any(x in texto_upper for x in ["DOLAR", "DÓLAR", "USD", "EXTRANJERA"])
    subcarpeta_moneda = "USD" if es_dolar else "CLP"
    
    ruta_final = os.path.join(ROOT_BCI, empresa_clean, subcarpeta_moneda)
    if not os.path.exists(ruta_final):
        try: os.makedirs(ruta_final, exist_ok=True)
        except: pass
    return ruta_final

def existe_folio_en_archivo(nombre_archivo: str, folio_buscado: str) -> bool:
    """
    VERIFICACIÓN ESTRICTA DE FOLIO.
    Busca si el número de folio está presente en el nombre del archivo.
    Ahora compatible con el formato '105_25-11-2025.pdf'.
    """
    f = nombre_archivo.lower()
    fol = folio_buscado.lstrip('0')
    if not fol: return False
    
    # 1. Chequeo si el archivo EMPIEZA con el folio (ej: "105_...")
    if f.startswith(f"{fol}_") or f.startswith(f"{fol}."):
        return True
        
    # 2. Chequeo si el folio está aislado por símbolos en medio (ej: "Cartola_105_...")
    if re.search(rf"(?:^|[^0-9])0*{re.escape(fol)}(?![0-9])", f): 
        return True
        
    return False

def estado_folio_en_carpeta(dest_dir: str, folio: str) -> tuple[bool, bool]:
    pdf = False
    excel = False
    try:
        if not os.path.exists(dest_dir): return False, False
        for n in os.listdir(dest_dir):
            if existe_folio_en_archivo(n, folio):
                if n.lower().endswith('.pdf'): pdf = True
                if n.lower().endswith(('.xlsx', '.xls')): excel = True
        return pdf, excel
    except:
        return False, False

def archivo_descargado(before: set, extension: str, timeout: int = 30, since_ts: float = 0) -> str | None:
    ext_low = extension.lower()
    exts = [ext_low if ext_low.startswith('.') else f'.{ext_low}']
    if '.xlsx' in exts: exts.append('.xls')
    
    fin = time.time() + timeout
    while time.time() < fin:
        try:
            actuales = set(os.listdir(DOWNLOAD_DIR))
            nuevos   = actuales - before
            candidatos = [f for f in nuevos if any(f.lower().endswith(e) for e in exts)]
            
            if not candidatos and since_ts > 0:
                for f in actuales:
                    if any(f.lower().endswith(e) for e in exts):
                        if os.path.getmtime(os.path.join(DOWNLOAD_DIR, f)) >= since_ts:
                            candidatos.append(f)

            if candidatos:
                ruta = max((os.path.join(DOWNLOAD_DIR, f) for f in candidatos), key=os.path.getmtime)
                if ruta.endswith(".crdownload") or ruta.endswith(".tmp"):
                    time.sleep(1)
                    continue
                s1 = os.path.getsize(ruta)
                time.sleep(1)
                if s1 > 0 and s1 == os.path.getsize(ruta):
                    return ruta
        except: pass
        time.sleep(1)
    return None

def mover_renombrar(origen: str, empresa: str, fecha_dt: datetime, cuenta_tipo: str, extension: str, folio: str | None = None):
    # -------------------------------------------------------------
    # [CAMBIO] NOMBRE DE ARCHIVO: FOLIO_FECHA (ej: 105_25-11-2025)
    # -------------------------------------------------------------
    fecha_str = fecha_dt.strftime("%d-%m-%Y") # Formato con guiones para Windows
    
    ext = os.path.splitext(origen)[1].lower()
    if not ext: ext = extension if extension.startswith('.') else '.' + extension

    if folio and folio != "SINFOLIO":
        nuevo_nombre = f"{folio}_{fecha_str}{ext}"
    else:
        # Fallback si no hay folio
        nuevo_nombre = f"SINFOLIO_{fecha_str}{ext}"
    
    destino_dir = obtener_ruta_dinamica(empresa, cuenta_tipo)
    destino = os.path.join(destino_dir, nuevo_nombre)

    base, _ = os.path.splitext(destino)
    idx = 1
    while os.path.exists(destino):
        destino = f"{base}_{idx}{ext}"
        idx += 1

    try:
        shutil.move(origen, destino)
        print(f"      ★ GUARDADO EN: {destino}")
    except:
        try:
            shutil.copy2(origen, destino)
            os.remove(origen)
            print(f"      ★ COPIADO EN: {destino}")
        except: pass

# --------------------------------------------------------------------------- #
# INTERACCIÓN (JS / ROBUST CLICKS)
# --------------------------------------------------------------------------- #
def entrar_primera_empresa(driver, wait):
    print("   Iniciando sesión en primera empresa...")
    try:
        wait.until(EC.visibility_of_element_located(SEL_BOX_EMPRESA))
        tarjetas = driver.find_elements(*SEL_BOX_EMPRESA)
        if tarjetas:
            click_robusto(driver, tarjetas[0])
            esperar(5)
            return True
    except: return False

def obtener_empresas_disponibles(driver, wait):
    print("   🔍 Escaneando dropdown de empresas...")
    try:
        drop = wait.until(EC.presence_of_element_located(X_DROPDOWN_EMPRESA_TRIGGER))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", drop)
        time.sleep(1)
        click_robusto(driver, drop)
        esperar(2)
        
        opciones = wait.until(EC.presence_of_all_elements_located(X_OPCIONES_EMPRESA_ITEM))
        lista_nombres = []
        for op in opciones:
            try:
                texto = op.find_element(By.CSS_SELECTOR, "p.texto").text.strip()
                if texto and texto not in lista_nombres:
                    lista_nombres.append(texto)
            except: pass
        
        ActionChains(driver).send_keys(u'\ue00c').perform() # ESC
        esperar(1)
        print(f"   📋 Empresas detectadas: {lista_nombres}")
        return lista_nombres
    except:
        click_right_blank(driver)
        return []

def cambiar_a_empresa(driver, wait, nombre_objetivo):
    print(f"   🔄 Cambiando contexto a: {nombre_objetivo}...")
    try:
        click_right_blank(driver)
        drop = wait.until(EC.presence_of_element_located(X_DROPDOWN_EMPRESA_TRIGGER))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", drop)
        time.sleep(1)
        click_robusto(driver, drop)
        esperar(1.5)
        
        xpath_opcion = f"//mat-option//p[contains(@class,'texto') and contains(text(), '{nombre_objetivo}')]/ancestor::mat-option"
        opcion = wait.until(EC.presence_of_element_located((By.XPATH, xpath_opcion)))
        click_robusto(driver, opcion)
        
        print("      Esperando actualización de datos...")
        esperar(5) 
        return True
    except:
        click_right_blank(driver)
        return False

def navegar_a_cartola_historica(driver, wait):
    print("   Navegando a menú Cartola Histórica...")
    try:
        m1 = wait.until(EC.element_to_be_clickable(X_MENU_CUENTAS))
        click_robusto(driver, m1)
        esperar(1)
        
        m2 = wait.until(EC.element_to_be_clickable(X_MENU_CORRIENTE))
        click_robusto(driver, m2)
        esperar(1)
        
        m3 = wait.until(EC.element_to_be_clickable(X_MENU_CARTOLA))
        click_robusto(driver, m3)
        esperar(5) 
        return True
    except: return False

# --------------------------------------------------------------------------- #
# LÓGICA DE PROCESAMIENTO
# --------------------------------------------------------------------------- #
def procesar_cuentas_actuales(driver, wait, empresa_nombre):
    cuentas_disponibles = []
    try:
        sel = wait.until(EC.presence_of_element_located(X_SELECT_CUENTAS))
        click_robusto(driver, sel)
        esperar(1)
        ops = wait.until(EC.presence_of_all_elements_located(X_OPCIONES_CUENTA))
        for o in ops:
            try: cuentas_disponibles.append(o.find_element(By.CSS_SELECTOR, "p.texto").text.strip())
            except: pass
        ActionChains(driver).send_keys(u'\ue00c').perform() 
        esperar(1)
    except: pass
    
    if not cuentas_disponibles: cuentas_disponibles = ["CUENTA ACTIVA"]

    for cuenta_nombre in cuentas_disponibles:
        print(f"      📂 Analizando cuenta: {cuenta_nombre}")
        
        # DEBUG: Mostrar dónde estamos buscando
        path_check = obtener_ruta_dinamica(empresa_nombre, cuenta_nombre)
        print(f"      (Revisando archivos en: {os.path.basename(path_check)})")

        if len(cuentas_disponibles) > 1 and cuenta_nombre != "CUENTA ACTIVA":
            try:
                sel = wait.until(EC.presence_of_element_located(X_SELECT_CUENTAS))
                click_robusto(driver, sel)
                esperar(1)
                xpath_c = f"//mat-option//*[contains(text(), '{cuenta_nombre}')]/ancestor::mat-option"
                op_cuenta = driver.find_element(By.XPATH, xpath_c)
                click_robusto(driver, op_cuenta)
                esperar(4) 
            except: continue

        try:
            tabla = wait.until(EC.presence_of_element_located((By.ID, ID_TABLA_CARTOLA)))
            filas = tabla.find_elements(By.TAG_NAME, "mat-row")
        except:
            print("      No se encontraron datos (tabla vacía).")
            continue

        for idx_fila, fila in enumerate(filas, start=1):
            try:
                # 1. LEER DATOS
                celdas = fila.find_elements(By.TAG_NAME, "mat-cell")
                folio_txt = celdas[0].text.strip() if len(celdas) > 0 else ""
                folio_num = re.sub(r"\D", "", folio_txt)
                fecha_txt = celdas[1].text.strip() if len(celdas) > 1 else ""
                fecha_dt  = datetime.strptime(fecha_txt, "%d/%m/%Y")
            except:
                fecha_dt = datetime.now()
                folio_num = f"SINFOLIO_{idx_fila}"

            # -------------------------------------------------------------
            # [CRÍTICO] VERIFICACIÓN INMEDIATA DE ARCHIVOS
            # -------------------------------------------------------------
            dest_dir = obtener_ruta_dinamica(empresa_nombre, cuenta_nombre)
            
            # Verificación precisa
            pdf_ok, excel_ok = estado_folio_en_carpeta(dest_dir, folio_num)
            
            # Si YA ESTÁN AMBOS, se salta toda la fila inmediatamente
            if folio_num and folio_num != "SINFOLIO" and pdf_ok and excel_ok:
                print(f"      [✓] Folio {folio_num} ya existe. Saltando.")
                continue 

            # Si llegamos aquí, falta algo. Abrimos menú.
            print(f"      ⬇️ Faltantes en folio {folio_num} (PDF:{'OK' if pdf_ok else 'NO'}, XLS:{'OK' if excel_ok else 'NO'}). Descargando...")
            try:
                puntos = fila.find_element(By.XPATH, X_TRES_PUNTOS)
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", puntos)
                click_robusto(driver, puntos)
                esperar(1)
            except: 
                click_right_blank(driver)
                continue

            # --- PDF ---
            if not pdf_ok:
                try:
                    btn = wait.until(EC.element_to_be_clickable(X_BOTON_PDF))
                    ts_start = time.time()
                    click_robusto(driver, btn)
                    
                    ruta = archivo_descargado(set(os.listdir(DOWNLOAD_DIR)), ".pdf", timeout=5, since_ts=ts_start)
                    
                    if not ruta:
                        try:
                            wait_visor = WebDriverWait(driver, 5)
                            try: btn_visor = wait_visor.until(EC.element_to_be_clickable((By.ID, "viewer-hb-boton-descargar")))
                            except: btn_visor = wait_visor.until(EC.element_to_be_clickable((By.ID, "hb-boton-descargar")))
                            click_robusto(driver, btn_visor)
                            click_viewer_close(driver)
                        except: pass
                        ruta = archivo_descargado(set(os.listdir(DOWNLOAD_DIR)), ".pdf", timeout=20, since_ts=ts_start)

                    if ruta:
                        mover_renombrar(ruta, empresa_nombre, fecha_dt, cuenta_nombre, "pdf", folio_num)
                    else:
                        print(f"      ⚠️ No se pudo descargar PDF {folio_num}")

                    if not excel_ok:
                        try: 
                            puntos = fila.find_element(By.XPATH, X_TRES_PUNTOS)
                            click_robusto(driver, puntos)
                            esperar(1)
                        except: pass
                except Exception as e: 
                    print(f"Error PDF: {e}")
            else:
                print(f"      [i] PDF ya existe.")

            # --- EXCEL ---
            if not excel_ok:
                try:
                    btn = wait.until(EC.element_to_be_clickable(X_BOTON_EXCEL))
                    ts_start = time.time()
                    click_robusto(driver, btn)
                    
                    try: 
                        wait_modal = WebDriverWait(driver, 4)
                        wait_modal.until(EC.presence_of_element_located(X_MODAL_CONTAINER))
                        btn_modal = driver.find_element(By.ID, "hb-boton-descargar")
                        click_robusto(driver, btn_modal)
                        click_viewer_close(driver)
                    except: pass
                    
                    ruta = archivo_descargado(set(os.listdir(DOWNLOAD_DIR)), ".xlsx", timeout=30, since_ts=ts_start)
                    if not ruta: ruta = archivo_descargado(set(os.listdir(DOWNLOAD_DIR)), ".xls", timeout=10, since_ts=ts_start)
                    
                    if ruta:
                        mover_renombrar(ruta, empresa_nombre, fecha_dt, cuenta_nombre, "xlsx", folio_num)
                    else:
                        print(f"      ⚠️ No se pudo descargar Excel {folio_num}")
                except Exception as e:
                    print(f"Error Excel: {e}")
            else:
                print(f"      [i] Excel ya existe.")
            
            click_right_blank(driver)

# --------------------------------------------------------------------------- #
# CICLO Y MAIN
# --------------------------------------------------------------------------- #
def ejecutar_ciclo_completo():
    opts = uc.ChromeOptions()
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
        "plugins.always_open_pdf_externally": True 
    }
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--start-maximized")
    
    driver = None
    try:
        driver = uc.Chrome(options=opts)
        wait = WebDriverWait(driver, 20)
        
        print("--- CONECTANDO AL BANCO ---")
        driver.get(URL_LOGIN)
        esperar()
        wait.until(EC.element_to_be_clickable(X_BTN_LOGIN)).click()
        esperar(2)
        driver.find_element(*X_INPUT_RUT).send_keys(RUT)
        driver.find_element(*X_INPUT_CLAVE).send_keys(CLAVE)
        driver.find_element(*X_BTN_SUBMIT).click()
        esperar(6)

        if entrar_primera_empresa(driver, wait):
            if navegar_a_cartola_historica(driver, wait):
                empresas = obtener_empresas_disponibles(driver, wait)
                for emp in empresas:
                    if emp == EMPRESA_EXCLUIDA: continue
                    print(f"\n>>> EMPRESA: {emp}")
                    if cambiar_a_empresa(driver, wait, emp):
                        procesar_cuentas_actuales(driver, wait, emp)
            else:
                print("Error: No se cargó el módulo de Cartola Histórica.")
    except Exception as e:
        print(f"Error en ciclo: {e}")
    finally:
        if driver:
            print("Cerrando sesión del ciclo...")
            try: driver.quit()
            except: pass

if __name__ == "__main__":
    inicio = datetime.now()
    fin = inicio + timedelta(hours=TIEMPO_TOTAL_HORAS)
    
    print(f"🤖 BOT BCI INICIADO - INTERVALO 5 MINUTOS")
    print(f"⏱️ Duración: {TIEMPO_TOTAL_HORAS} horas (Hasta {fin.strftime('%H:%M')})")
    
    ciclo = 1
    while datetime.now() < fin:
        print(f"\n[{datetime.now().strftime('%H:%M')}] ▶ CICLO #{ciclo}")
        ejecutar_ciclo_completo()
        
        if datetime.now() < fin:
            print(f"\n💤 Esperando {INTERVALO_ESPERA_MINUTOS} mins...")
            time.sleep(INTERVALO_ESPERA_MINUTOS * 60)
            ciclo += 1
            
    print("\n✅ TAREA COMPLETADA.")