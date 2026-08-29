import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

EXCEL_FILE = "propiedades.xlsx"
TXT_OUTPUT = "resultados_servicios_sanitarios.txt"
URL_TANDIL = "https://www.autogestion.tandil.gov.ar/apex/f?p=114:101"

def formatear_texto(val):
    if val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str

def consultar_servicios_sanitarios():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        sheet = wb.active
    except Exception as e:
        print(f"Error al abrir Excel: {e}")
        return

    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(5)

    with open(TXT_OUTPUT, "w", encoding="utf-8") as f_out:
        f_out.write("=== REPORTE DE DEUDAS - SERVICIOS SANITARIOS ===\n")
        f_out.write(f"Fecha de consulta: {time.strftime('%d/%m/%Y %H:%M')}\n")
        f_out.write("=" * 55 + "\n\n")

        try:
            for row in range(1, sheet.max_row + 1):
                val_col1 = sheet.cell(row=row, column=1).value
                val_col2 = sheet.cell(row=row, column=2).value
                val_col3 = sheet.cell(row=row, column=3).value

                str_col1 = formatear_texto(val_col1)
                str_col3 = formatear_texto(val_col3)

                # Ignorar encabezados o filas vacías
                if not str_col3 or str_col3.lower() in ["servicio", "nro. cuenta", "cuenta", "none"]:
                    continue

                usuario = str_col1 if str_col1 and str_col1.lower() != "usuario" else f"Fila {row}"
                direccion = formatear_texto(val_col2) or "Sin Dirección"
                num_cuenta = str_col3

                # 1. Ingresar a la página principal
                driver.get(URL_TANDIL)
                time.sleep(1.5)

                # 2. Completar Nro. Cuenta
                try:
                    input_cuenta = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
                    )
                    input_cuenta.clear()
                    input_cuenta.send_keys(num_cuenta)
                except Exception:
                    continue

                # 3. Seleccionar 'Tipo Deuda'
                try:
                    dropdown_elem = driver.find_element(By.TAG_NAME, "select")
                    select_tasa = Select(dropdown_elem)
                    for option in select_tasa.options:
                        if "sanitar" in option.text.lower():
                            select_tasa.select_by_visible_text(option.text)
                            break
                except Exception:
                    pass

                time.sleep(1)

                # 4. Clic en 'Iniciar Sesión'
                try:
                    btn_iniciar = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Iniciar') or contains(., 'Sesión')] | //input[@type='submit' or @type='button']"))
                    )
                    btn_iniciar.click()
                except Exception:
                    driver.execute_script("document.querySelector('button').click();")

                time.sleep(2.5)

                # 5. NAVEGACIÓN ROBUSTA A 'CONSULTA DE DEUDA'
                try:
                    # Intentar hacer clic desplegando el menú primero
                    consultas_elem = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Consultas')]"))
                    )
                    driver.execute_script("arguments[0].click();", consultas_elem)
                    time.sleep(1)

                    deuda_elem = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Consulta de deuda')]"))
                    )
                    driver.execute_script("arguments[0].click();", deuda_elem)
                    time.sleep(3)
                except Exception:
                    # Si falla el clic JS, buscar enlace <a> con el texto
                    try:
                        link_deuda = driver.find_element(By.PARTIAL_LINK_TEXT, "Consulta de deuda")
                        driver.get(link_deuda.get_attribute("href"))
                        time.sleep(3)
                    except Exception:
                        pass

                # 6. Leer la tabla de deudas y filtrar SERV.SANITAR
                detalles_deuda = []
                monto_total_sanitar = 0.0

                try:
                    # Esperar a que la tabla cargue
                    WebDriverWait(driver, 6).until(
                        EC.presence_of_element_located((By.TAG_NAME, "table"))
                    )
                    filas_tabla = driver.find_elements(By.XPATH, "//tr")
                    for fila in filas_tabla:
                        texto_fila = fila.text
                        if "SERV.SANITAR" in texto_fila:
                            columnas = fila.find_elements(By.TAG_NAME, "td")
                            if len(columnas) >= 6:
                                cuota = columnas[3].text.strip()
                                vencimiento = columnas[4].text.strip()
                                importe_raw = columnas[5].text.strip()
                                
                                importe_clean = importe_raw.replace(".", "").replace(",", ".")
                                try:
                                    monto_float = float(importe_clean)
                                    monto_total_sanitar += monto_float
                                except ValueError:
                                    pass

                                detalles_deuda.append(f"  • Cuota {cuota} (Venc: {vencimiento}): Importe = ${importe_raw}")
                except Exception as ex_tabla:
                    detalles_deuda.append(f"  • Error/No se cargó tabla: {ex_tabla}")

                # 7. Escribir datos al archivo de salida
                if detalles_deuda:
                    resumen = "\n".join(detalles_deuda)
                    linea_monto = f"Total Deuda SERV.SANITAR: ${monto_total_sanitar:,.2f}\n{resumen}"
                else:
                    linea_monto = "Sin Deuda en SERV.SANITAR / No Encontrado"

                registro = (
                    f"Usuario: {usuario}\n"
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
