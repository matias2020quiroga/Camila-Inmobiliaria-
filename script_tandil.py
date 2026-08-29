import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

FILA_INICIAL = 4       # Fila donde empiezan los datos
COL_NOMBRE = 1         # Columna A
COL_DIRECCION = 2      # Columna B
COL_SERVICIOS = 3      # Columna C

EXCEL_FILE = "propiedades.xlsx"
TXT_OUTPUT = "resultados_servicios_sanitarios.txt"

URL_TANDIL = "https://www.autogestion.tandil.gov.ar/apex/f?p=114:101"

def consultar_servicios_sanitarios():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
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

                driver.get(URL_TANDIL)

                # 1. Seleccionar tipo de tasa
                dropdown_elem = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "select"))
                )
                select_tasa = Select(dropdown_elem)
                for option in select_tasa.options:
                    if "sanitar" in option.text.lower():
                        select_tasa.select_by_visible_text(option.text)
                        break

                # 2. Ingresar número de cuenta
                input_cuenta = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
                )
                input_cuenta.clear()
                input_cuenta.send_keys(num_cuenta)

                # 3. Hacer clic en consultar mediante selector amplio o tecla ENTER
                try:
                    btn = driver.find_element(By.XPATH, "//*[contains(@class, 'button') or contains(text(), 'Consultar') or contains(text(), 'Buscar') or contains(@id, 'BTN') or contains(@id, 'SUBMIT')]")
                    btn.click()
                except Exception:
                    input_cuenta.send_keys(Keys.ENTER)

                time.sleep(3)

                # 4. Intentar ingresar a la sección de Consultas/Detalle si aplica
                try:
                    btn_consultas = driver.find_element(By.XPATH, "//*[contains(text(), 'Consultas')]")
                    btn_consultas.click()
                    time.sleep(2)
                except Exception:
                    pass

                # 5. Mapear tabla de deudas para SERV.SANITAR
                detalles_deuda = []
                monto_total_sanitar = 0.0

                try:
                    filas_tabla = driver.find_elements(By.XPATH, "//tr")
                    for fila in filas_tabla:
                        texto_fila = fila.text
                        if "SERV.SANITAR" in texto_fila:
                            columnas = fila.find_elements(By.TAG_NAME, "td")
                            if len(columnas) >= 6:
                                cuota = columnas[3].text.strip()
                                vencimiento = columnas[4].text.strip()
                                actualizado_raw = columnas[6].text.strip()
                                actualizado_clean = actualizado_raw.replace(".", "").replace(",", ".")
                                
                                try:
                                    monto_float = float(actualizado_clean)
                                    monto_total_sanitar += monto_float
                                except ValueError:
                                    pass

                                detalles_deuda.append(f"  • Cuota {cuota} (Venc: {vencimiento}): ${actualizado_raw}")
                except Exception as ex_tabla:
                    detalles_deuda.append(f"  • Error en lectura de tabla: {ex_tabla}")

                if detalles_deuda:
                    resumen = "\n".join(detalles_deuda)
                    linea_monto = f"Total Deuda SERV.SANITAR: ${monto_total_sanitar:,.2f}\n{resumen}"
                else:
                    linea_monto = "Sin Deuda en SERV.SANITAR / No Encontrado"

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
