import asyncio
import os
import pandas as pd
from datetime import datetime
from playwright.async_api import async_playwright, Page, BrowserContext, expect
import subprocess
import json

try:
    with open('Utils/credentials.json', 'r', encoding='utf-8') as archivo_json:
        credenciales = json.load(archivo_json)
        USUARIO = credenciales['usuario']
        PASSWORD = credenciales['password']
except FileNotFoundError:
    print("Error: No se encontró el archivo 'Utils/credentials.json'. Asegúrate de crearlo.")
    exit()
except KeyError as e:
    print(f"Error: Falta la clave {e} en el archivo JSON.")
    exit()

DOCUMENTS_UPLOAD = 'Documentos'
INPUT_FILE = 'Control Oficios Alta.xlsx'
FILE_EXITOS = 'Resultados/Resultados_Exitosos.xlsx'
FILE_ERRORES = 'Resultados/Reporte_Errores.xlsx'
TEMP_EXITOS_CSV = 'Resultados/temp_exitos_backup.csv'
TEMP_ERRORES_CSV = 'Resultados/temp_errores_backup.csv'

BATCH_GUARDADO = 5

URL_LOGIN = "https://acprod.intranet.com.mx/mbom_mx_ws/mbom_mx_web/PortalLogon"
URL_ALTA_OFICIO = "https://acprod.intranet.com.mx/boixp_mx_web/boixp_mx_web/servlet/ServletOperacionWeb?OPERACION=VGOMX002&LOCALE=es_ES&DATOS_ENTRADA.FLUJO_LANZAR=GOMXFL10010"

# TODO: Paso Alta Persona (cuando ya tenga creado el FOLIO SUGO)
def paso_alta_persona():
   return

# TODO: Paso Finalizar Captura Básica (Cuando ya este el paso Alta persona)
def paso_finalizar_captura():
    return

async def obtener_estado_sesion(browser):
    context = await browser.new_context(ignore_https_errors=True, viewport={'width': 1280, 'height': 800})
    page = await context.new_page()
    try:
        await page.goto(URL_LOGIN, wait_until="domcontentloaded")
        await asyncio.sleep(2)
        await page.fill(".name", USUARIO)
        await page.fill(".pass", PASSWORD)
        await asyncio.sleep(1)
        
        async with context.expect_page() as page_info:
            if await page.locator("//p[@onclick='validaCampos()']").is_visible():
                await page.evaluate("""() => validaCampos()""")
            else:
                await page.keyboard.press('Enter')

        print(f"Usuario {USUARIO} autenticado")
        
        popup = await page_info.value
        await popup.wait_for_load_state()
        await asyncio.sleep(3) 
        
        storage = await context.storage_state()
        await context.close()
        return storage
    except Exception:
        return None

def estandarizar_fechas(fecha):
    meses_es = {
        "ene": "01", "feb": "02", "mar": "03", "abr": "04", "may": "05", "jun": "06",
        "jul": "07", "ago": "08", "sep": "09", "oct": "10", "nov": "11", "dic": "12"
    }
    fecha = str(fecha).strip().lower()
    if pd.isnull(fecha) or fecha == 'nat' or fecha == '':
        return ''
    for mes, num in meses_es.items():
        if mes in fecha:
            fecha = fecha.replace(mes, num)
            break
    formatos = [
        "%Y-%m-%d",  
        "%d-%m-%Y",  
        "%d/%m/%Y",  
        "%Y/%m/%d",  
        "%d-%m-%y",  
        "%d-%m-%Y",  
        "%d %m %Y",
        "%m/%d/%Y",
        "%d/%m/%Y %H:%M:%S",  
        "%d/%m/%y %H:%M:%S",  
        "%d-%m-%Y %H:%M:%S",  
        "%Y-%m-%d %H:%M:%S",  
        "%d-%m-%y %H:%M:%S",  
    ]
    for formato in formatos:
        try:
            fecha_obj = datetime.strptime(fecha, formato)
            return fecha_obj.strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""

async def registro_alta_oficio(dato, resultados, page: Page, context: BrowserContext, contador, intento=1, max_intentos=3):
    if intento > max_intentos:
        print(f"{contador} - {dato['Oficio Autoridad']}: Falló después de {max_intentos} intentos\n")
        resultados.append([dato['Oficio Autoridad'], "ERROR", "ERROR"])
        return

    try:
        await page.goto(URL_ALTA_OFICIO, wait_until="domcontentloaded", timeout=60000)
        print(f"{contador} - {dato['Oficio Autoridad']}:")

        # En cada selección del campo realiza un postback
        async with page.expect_navigation():
          if dato['Origen Folio'].strip():
            try: 
              await page.select_option('#listOF', label=dato['Origen Folio'].strip(), timeout=2000)
            except Exception:
              print(f"    [ALERT] '{dato['Origen Folio'].strip()}' no existe en Origen Folio")
        
        async with page.expect_navigation():
          if dato['Región'].strip():
            try:
              await page.select_option('#listRegion', label=dato['Región'].strip(), timeout=2000)
            except Exception:
              print(f"    [ALERT] '{dato['Región'].strip()}' no existe en Región")
        
        async with page.expect_navigation():
          if dato['Plaza'].strip():
            try:  
              await page.select_option('#listPlaza', label=dato['Plaza'].strip(), timeout=2000)
            except Exception:
              print(f"    [ALERT] '{dato['Plaza'].strip()}' no existe en Plaza")
        
        await page.select_option('#listOficio', label=dato['Tipo de Oficio'])
        
        if dato['Autoridad Específica'].strip():
          async with context.expect_page() as page_autoridades:
              await page.evaluate(""" () => AutoE(); """)
          
          pantalla_lista_autoridades = await page_autoridades.value
          await pantalla_lista_autoridades.wait_for_load_state()

          td_autoridad_seleccionada = None
          td_autoridad_seleccionada = pantalla_lista_autoridades.locator(f"//td[contains(text(), '{dato['Autoridad Específica']}')]").first

          if await td_autoridad_seleccionada.count() > 0:
              checkbox = td_autoridad_seleccionada.locator("input[type='radio']").first
              await checkbox.check()
              await asyncio.sleep(0.5)
              await pantalla_lista_autoridades.evaluate(""" () => enviaDatos(); """)
          else:
              print(f"    [ALERT] '{dato['Autoridad Específica']}' no existe en Autoridad Específica")
              await pantalla_lista_autoridades.close()
        
        if dato['Autoridad'].strip():
          try:
              await page.select_option('#listAutoridad', label=dato['Autoridad'].strip(), timeout=2000)
          except Exception:
              print(f"    [ALERT] '{dato['Autoridad'].strip()}' no existe en Autoridad")

        await page.fill('#ofiCNBV', dato['Oficio Autoridad'])
        await page.fill('#expediente', dato['Expediente Autoridad'])
        await page.fill('#subAutoridad', dato['Sub Autoridad'])
        await page.fill('#ofiSUB', dato['Oficio Sub Autoridad'])
        await page.fill('#expedienteSub', dato['Expediente sub autoridad'])

        await page.evaluate(f"document.getElementById('fechaFin').value = '{dato['Fecha Recepción']}'")
        await page.fill('#plazo', str(dato['Plazo (Días)']))
        
        if dato['Abogado Solicitante'].strip():
          try:
            await page.select_option('#cbAbogado', label=dato['Abogado Solicitante'].strip(), timeout=2000)
          except Exception:
            print(f"    [ALERT] '{dato['Abogado Solicitante'].strip()}' no existe en Abogado Solicitante")

        if dato['Empresa'].strip():
          try:
            await page.select_option('#cbEmpresa', label=dato['Empresa'].strip(), timeout=2000)
          except Exception:
            print(f"    [ALERT] '{dato['Empresa'].strip()}' no existe en Empresa")

        async def aceptar_alerta(dialog):
            await dialog.accept()
            
        page.once("dialog", aceptar_alerta)

        await page.evaluate(""" () => guardar(); """)

        elemento_folio = page.locator("#folioSolicitud")
        await elemento_folio.wait_for(state="visible", timeout=15000)

        folio_sugo = await elemento_folio.input_value()
        folio_sugo = str(folio_sugo).strip()

        print(f"    [OK] Proceso completado exitosamente")
        print(f"    [FOLIO SUGO] {folio_sugo}\n")
        
        resultados.append([dato['Oficio Autoridad'], "Correcto", folio_sugo])
        await asyncio.sleep(1000)
        

    except Exception as e:
        print(f"    [ERROR] Error en intento {intento} - {str(e)}")
        print(f"    [MESSAGE] Intentando de nuevo...")

        await asyncio.sleep(1)
        
        await registro_alta_oficio(dato, resultados, page, context, contador, intento + 1, max_intentos)


async def main():
    for f in [TEMP_EXITOS_CSV, TEMP_ERRORES_CSV, FILE_EXITOS, FILE_ERRORES]:
        if os.path.exists(f):
            try: os.remove(f)
            except: pass

    try:
        df_input = pd.read_excel(INPUT_FILE, dtype=str, engine='openpyxl')
        columns = [
            "Origen Folio","Región","Plaza","Tipo de Oficio",
            "Autoridad Específica","Autoridad","Oficio Autoridad",
            "Expediente Autoridad","Sub Autoridad","Oficio Sub Autoridad",
            "Expediente sub autoridad","Fecha Recepción","Plazo (Días)",
            "Abogado Solicitante","Empresa","Documento"
        ]
        if df_input.shape[1] == len(columns):
            df_input.columns = columns
            # Realizamos limpieza de datos y estandarización de fechas
            df_input['Fecha Recepción'] = df_input['Fecha Recepción'].apply(estandarizar_fechas)
            df_input['Fecha Recepción'] = pd.to_datetime(df_input['Fecha Recepción'], errors='coerce').dt.strftime('%d/%m/%Y')
            df_input['Plazo (Días)'] = df_input['Plazo (Días)'].astype(str).str.extract('(\d+)').fillna(0).astype(int)
            # Convertimos a string para evitar problemas con NaN o formatos mixtos
            df_input = df_input.fillna('').astype(str)
            df_input = df_input.applymap(lambda x: x.strip() if isinstance(x, str) else x)

            oficios_totales = df_input.to_dict(orient='records')
        else:
            print(f"Error: El número de columnas en {INPUT_FILE} no coincide con el formato esperado.")
            print(f"Columnas esperadas: {len(columns)}, Columnas encontradas: {df_input.shape[1]}")
            return

    except Exception as e:
        print(f"Error leyendo Excel: {e}")
        return
    
    print("\n========= BOT (Alta Oficio - SUGO) =========")
    print(f"Iniciando [{len(oficios_totales)} oficios | 1 worker]")
    print(f"Respaldo automático cada {BATCH_GUARDADO} oficios por worker.")

    resultados = []
    await asyncio.sleep(2)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False) 
        storage_state = await obtener_estado_sesion(browser)
        
        if not storage_state:
            print("Falló el login.")
            await browser.close()
            return
            
        context = await browser.new_context(storage_state=storage_state, ignore_https_errors=True)

        for contador, dato in enumerate(oficios_totales, start=1):
            page = await context.new_page()
            await registro_alta_oficio(dato, resultados, page, context, contador)
            await page.close()
        
    # Exportar resultados a Excel
    if resultados:
        df_resultados = pd.DataFrame(resultados, columns=['Oficio Autoridad', 'Resultado', 'Folio SUGO'])
        df_resultados.to_excel(FILE_EXITOS, index=False)
        print(f"Resultados exitosos guardados en {FILE_EXITOS}")

def instalar_navegadores():
    import sys
    from playwright._impl._cli import main as playwright_cli
    
    print("Verificando/Instalando navegador Chromium... (esto puede tardar un par de minutos la primera vez)")
    
    original_argv = sys.argv.copy()
    
    sys.argv = ["", "install", "chromium"]
    
    try:
        playwright_cli()
    except SystemExit:
        pass 
    finally:
        sys.argv = original_argv

if __name__ == "__main__":
    instalar_navegadores() 
    asyncio.run(main())