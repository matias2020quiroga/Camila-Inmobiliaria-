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
            # Recorrer todas las filas desde la fila 2
            for row in range(2, sheet.max_row + 1):
                val_usuario = sheet.cell(row=row, column=1).value
                val_direccion = sheet.cell(row=row, column=2).value
                val_servicio = sheet.cell(row=row, column=3).value

                num_cuenta = formatear_texto(val_servicio)
                
                # Si no hay número de cuenta en la Columna C, salta la fila
                if not num_cuenta or num_cuenta.lower() == "none" or num_cuenta.lower() == "servicio":
                    continue

                usuario = formatear_texto(val_usuario) or f"Fila {row}"
                direccion = formatear_texto(val_direccion) or "Sin Dirección"

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

                time.sleep(3)

                # 5. Navegar a 'Consultas' -> 'Consulta de deuda'
                try:
                    btn_consultas = WebDriverWait(driver, 8).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Consultas')]"))
                    )
                    btn_consultas.click()
                    time.sleep(1.5)

                    btn_sub_deuda = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Consulta de deuda')]"))
                    )
                    btn_sub_deuda.click()
                    time.sleep(3)
                except Exception:
                    pass

                # 6. Leer tabla y filtrar SERV.SANITAR
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
                                importe_raw = columnas[5].text.strip()
                                
                                importe_clean = importe_raw.replace(".", "").replace(",", ".")
                                try:
                                    monto_float = float(importe_clean)
                                    monto_total_sanitar += monto_float
                                except ValueError:
                                    pass

                                detalles_deuda.append(f"  • Cuota {cuota} (Venc: {vencimiento}): Importe = ${importe_raw}")
                except Exception as ex_tabla:
                    detalles_deuda.append(f"  • Error en lectura de tabla: {ex_tabla}")

                # 7. Escribir registro
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
