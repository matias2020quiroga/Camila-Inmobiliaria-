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
            # Leer desde la fila 2 para ignorar la fila de títulos
            for row in range(2, sheet.max_row + 1):
                val_usuario = sheet.cell(row=row, column=1).value
                val_direccion = sheet.cell(row=row, column=2).value
                val_servicio = sheet.cell(row=row, column=3).value

                num_cuenta = formatear_texto(val_servicio)

                # Ignorar filas sin número de cuenta válido
                if not num_cuenta or num_cuenta.lower() in ["none", "servicio", ""]:
                    continue

                usuario = formatear_texto(val_usuario) or f"Fila {row}"
                direccion = formatear_texto(val_direccion) or "Sin Dirección"

                # 1. Ingresar a la página de login
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

                # 5. DESPLEGAR EL MENÚ ACORDEÓN 'CONSULTAS'
                try:
                    # Clic 1: Desplegar la flechita / menú 'Consultas'
                    elem_consultas = WebDriverWait(driver, 8).until(
                        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Consultas')]"))
                    )
                    # Forzar clic por evento directo para garantizar que abra la lista desplegable
                    driver.execute_script("arguments[0].click();", elem_consultas)
                    time.sleep(1.5)

                    # Clic 2: Hacer clic en el sub-ítem 'Consulta de deuda' que aparece desplegado
                    elem_sub_deuda = WebDriverWait(driver, 8).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Consulta de deuda')]"))
                    )
                    driver.execute_script("arguments[0].click();", elem_sub_deuda)
                    time.sleep(3.5)
                except Exception as ex_menu:
                    # Si ya estaba desplegado o falló la animación, intenta hacer clic directo en el enlace
                    try:
                        link_deuda = driver.find_element(By.XPATH, "//a[contains(., 'Consulta de deuda')]")
                        link_deuda.click()
                        time.sleep(3.5)
                    except Exception:
                        pass

                # 6. Leer la tabla de detalles de deuda
                detalles_deuda = []
                monto_total_sanitar = 0.0

                try:
                    # Esperar la aparición de las filas de datos
                    WebDriverWait(driver, 8).until(
                        EC.presence_of_element_located((By.XPATH, "//tr[contains(., 'SERV.SANITAR') or contains(., 'Tasa')]"))
                    )
                    filas_tabla = driver.find_elements(By.XPATH, "//tr")
                    for fila in filas_tabla:
                        texto_fila = fila.text
                        if "SERV.SANITAR" in texto_fila:
                            columnas = fila.find_elements(By.TAG_NAME, "td")
                            if len(columnas) >= 6:
                                cuota = columnas[3].text.strip()
                                vencimiento = columnas[4].text.strip()
                                importe_raw = columnas[5].text.strip()  # Toma la columna 'Importe'
                                
                                importe_clean = importe_raw.replace(".", "").replace(",", ".")
                                try:
                                    monto_float = float(importe_clean)
                                    monto_total_sanitar += monto_float
                                except ValueError:
                                    pass

                                detalles_deuda.append(f"  • Cuota {cuota} (Venc: {vencimiento}): Importe = ${importe_raw}")
                except Exception as ex_tabla:
                    detalles_deuda.append(f"  • Error en lectura de tabla: {ex_tabla}")

                # 7. Registrar resultado en el archivo .txt
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
