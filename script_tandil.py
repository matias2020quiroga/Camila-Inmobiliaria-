import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# --- CONFIGURACIÓN DE COLUMNAS DEL EXCEL ---
FILA_INICIAL = 4       # Fila donde empiezan los datos

# Asumiendo la estructura del Excel:
# Columna A: Nombre / Cliente
# Columna B: Dirección / Inmueble
# Columna C: Servicios Sanitarios (Cuenta)
# Columna D: Retributiva de Servicios (Cuenta)

COL_NOMBRE = 1
COL_DIRECCION = 2
COL_SERVICIOS = 3
COL_RETRIBUTIVA = 4

EXCEL_FILE = "propiedades.xlsx"
TXT_OUTPUT = "resultados_tandil.txt"

def consultar_deuda():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active

    # Iniciar navegador Chrome
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    
    url = "https://www.autogestion.tandil.gov.ar/apex/f?p=114:101"

    # Preparar el archivo TXT para escribir los resultados
    with open(TXT_OUTPUT, "w", encoding="utf-8") as f_out:
        f_out.write("=== REPORTE DE DEUDAS - TANDIL ===\n")
        f_out.write(f"Fecha de consulta: {time.strftime('%d/%m/%Y %H:%M')}\n")
        f_out.write("=" * 40 + "\n\n")

        try:
            for row in range(FILA_INICIAL, sheet.max_row + 1):
                nombre = sheet.cell(row=row, column=COL_NOMBRE).value or "Sin Nombre"
                direccion = sheet.cell(row=row, column=COL_DIRECCION).value or "Sin Dirección"
                val_servicio = sheet.cell(row=row, column=COL_SERVICIOS).value
                val_retributiva = sheet.cell(row=row, column=COL_RETRIBUTIVA).value

                num_cuenta = None
                tipo_tasa = None

                # Determinar cuál columna contiene el número de cuenta
                if val_servicio is not None and str(val_servicio).strip() != "":
                    num_cuenta = str(int(val_servicio) if isinstance(val_servicio, float) else val_servicio).strip()
                    tipo_tasa = "SERVICIOS SANITARIOS"
                elif val_retributiva is not None and str(val_retributiva).strip() != "":
                    num_cuenta = str(int(val_retributiva) if isinstance(val_retributiva, float) else val_retributiva).strip()
                    tipo_tasa = "RETRIBUTIVA DE SERVICIOS"
                else:
                    continue  # Fila vacía, pasa a la siguiente

                driver.get(url)

                # 1. Seleccionar la tasa en el desplegable
                dropdown_elem = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "select"))
                )
                select_tasa = Select(dropdown_elem)

                # Buscar la opción correspondiente en el menú desplegable
                opcion_encontrada = False
                for option in select_tasa.options:
                    if tipo_tasa.lower() in option.text.lower():
                        select_tasa.select_by_visible_text(option.text)
                        opcion_encontrada = True
                        break

                # 2. Ingresar la cuenta en el campo de texto
                input_cuenta = driver.find_element(By.CSS_SELECTOR, "input[type='text']")
                input_cuenta.clear()
                input_cuenta.send_keys(num_cuenta)

                # 3. Hacer clic en Consultar
                btn_buscar = driver.find_element(By.XPATH, "//button[contains(text(), 'Consultar') or contains(text(), 'Buscar')]")
                btn_buscar.click()

                time.sleep(3)  # Pausa para la carga de datos

                # 4. Extraer el monto de la respuesta
                try:
                    monto_elem = driver.find_element(By.XPATH, "//*[contains(text(), '$')]")
                    monto_texto = monto_elem.text
                except Exception:
                    monto_texto = "Sin Deuda / No Encontrado"

                # 5. Escribir registro detallado en el TXT
                resultado_linea = (
                    f"Cliente: {nombre}\n"
                    f"Dirección: {direccion}\n"
                    f"Tasa: {tipo_tasa}\n"
                    f"N° Cuenta: {num_cuenta}\n"
                    f"Monto Deuda: {monto_texto}\n"
                    f"----------------------------------------\n"
                )
                
                f_out.write(resultado_linea)
                f_out.flush()  # Escribe en disco de inmediato
                print(f"Procesado: {nombre} - Cuenta {num_cuenta} -> {monto_texto}")

        finally:
            driver.quit()
            print("Proceso completado. Resultados guardados en resultados_tandil.txt")

if __name__ == "__main__":
    consultar_deuda()
