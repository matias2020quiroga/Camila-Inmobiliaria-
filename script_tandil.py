import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# --- CONFIGURACIÓN DE EXCEL ---
FILA_INICIAL = 4       # Fila donde empiezan los datos
COL_NOMBRE = 1         # Columna A: Nombre
COL_DIRECCION = 2      # Columna B: Dirección
COL_SERVICIOS = 3      # Columna C: Cuenta Servicios Sanitarios

EXCEL_FILE = "propiedades.xlsx"
TXT_OUTPUT = "resultados_servicios_sanitarios.txt"

URL_TANDIL = "https://www.autogestion.tandil.gov.ar/apex/f?p=114:101"

def consultar_servicios_sanitarios():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active

    # Opciones de Chrome para entorno automatizado
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")  # Se ejecuta de fondo sin abrir ventana visible
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(5)

    with open(TXT_OUTPUT, "w", encoding="utf-8") as f_out:
        f_out.write("=== REPORTE DE DEUDAS - SERVICIOS SANITARIOS ===\n")
        f_out.write(f"Fecha de consulta: {time.strftime('%d/%m/%Y %H:%M')}\n")
        f_out.write("=" * 55 + "\n\n")

        try:
            for row in range(FILA_INICIAL, sheet.max_row + 1):
                nombre = str(sheet.cell(row=row, column=COL_NOMBRE).value or "Sin Nombre").strip()
                direccion = str(sheet.cell(row=row, column=COL_DIRECCION).value or "Sin Dirección").strip()
                val_servicio = sheet.cell(row=row, column=COL_SERVICIOS).value

                if val_servicio is None or str(val_servicio).strip() == "":
                    continue

                num_cuenta = str(int(val_servicio) if isinstance(val_servicio, float) else val_servicio).strip()

                # 1. Cargar formulario principal
                driver.get(URL_TANDIL)

                # 2. Seleccionar "SERVICIOS SANITARIOS"
                dropdown_elem = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "select"))
                )
                select_tasa = Select(dropdown_elem)
                for option in select_tasa.options:
                    if "sanitar" in option.text.lower():
                        select_tasa.select_by_visible_text(option.text)
                        break

                # 3. Ingresar número de cuenta
                input_cuenta = driver.find_element(By.CSS_SELECTOR, "input[type='text']")
                input_cuenta.clear()
                input_cuenta.send_keys(num_cuenta)

                # 4. Clic en Consultar
                btn_buscar = driver.find_element(By.XPATH, "//button[contains(text(), 'Consultar') or contains(text(), 'Buscar')]")
                btn_buscar.click()

                time.sleep(2)

                # 5. Si existe el menú lateral "Consultas", hacer clic
                try:
                    btn_consultas = driver.find_element(By.XPATH, "//*[contains(text(), 'Consultas')]")
                    btn_consultas.click()
                    time.sleep(2)
                except Exception:
                    pass  # Si ya está en la vista directa, continúa

                # 6. Recorrer la tabla de "Consulta de detalle de deuda"
                detalles_deuda = []
                monto_total_sanitar = 0.0

                try:
                    # Buscar las filas de la tabla de deudas
                    filas_tabla = driver.find_elements(By.XPATH, "//tr")
                    
                    for fila in filas_tabla:
                        texto_fila = fila.text
                        # Filtrar únicamente las filas de la tasa SERV.SANITAR
                        if "SERV.SANITAR" in texto_fila:
                            columnas = fila.find_elements(By.TAG_NAME, "td")
                            if len(columnas) >= 6:
                                cuota = columnas[3].text.strip()
                                vencimiento = columnas[4].text.strip()
                                # Tomar la columna "Actualizado" (o "Importe" según convenga)
                                actualizado_str = columnas[6].text.strip().replace(".", "").replace(",", ".")
                                
                                try:
                                    monto_float = float(actualizado_str)
                                    monto_total_sanitar += monto_float
                                except ValueError:
                                    monto_float = 0.0

                                detalles_deuda.append(f"  • Cuota {cuota} (Venc: {vencimiento}): ${columnas[6].text.strip()}")

                except Exception as ex_tabla:
                    detalles_deuda.append(f"  • No se pudo leer la tabla: {ex_tabla}")

                # 7. Formatear la salida
                if detalles_deuda:
                    resumen_deuda = "\n".join(detalles_deuda)
                    linea_monto = f"Total Deuda SERV.SANITAR: ${monto_total_sanitar:,.2f}\n{resumen_deuda}"
                else:
                    linea_monto = "Sin Deuda en SERV.SANITAR / No Encontrado"

                # 8. Escribir resultado
                registro = (
                    f"Cliente: {nombre}\n"
                    f"Dirección: {direccion}\n"
                    f"N° Cuenta: {num_cuenta}\n"
                    f"{linea_monto}\n"
                    f"-------------------------------------------------------\n"
                )
                f_out.write(registro)
                f_out.flush()

        finally:
            driver.quit()

if __name__ == "__main__":
    consultar_servicios_sanitarios()
