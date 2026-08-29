import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# --- CONFIGURACIÓN DE COLUMNAS DEL EXCEL ---
FILA_INICIAL = 2       # Fila 2 (la Fila 1 tiene los títulos)
COL_USUARIO = 1        # Columna A: usuario
COL_DIRECCION = 2      # Columna B: dirección
COL_SERVICIO = 3       # Columna C: Servicio (Nro. de cuenta)

EXCEL_FILE = "propiedades.xlsx"
TXT_OUTPUT = "resultados_servicios_sanitarios.txt"

URL_TANDIL = "https://www.autogestion.tandil.gov.ar/apex/f?p=114:101"

def formatear_cuenta(val):
    if val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str

def consultar_servicios_sanitarios():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active

    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(5)

    with open(TXT_OUTPUT, "w", encoding="utf-8") as f_out:
        f_out.write("=== REPORTE DE DEUDAS - SERVICIOS SANITARIOS ===\n")
        f_out.write(f"Fecha de consulta: {time.strftime('%d/%m/%Y %H:%M')}\n")
        f_out.write("=" * 55 + "\n\n")

        try:
            for row in range(FILA_INICIAL, sheet.max_row + 1):
                usuario_val = sheet.cell(row=row, column=COL_USUARIO).value
                # Si la fila entera está vacía (como la E), cortamos el bucle o saltamos
                if usuario_val is None or str(usuario_val).strip() == "":
                    continue

                usuario = str(usuario_val).strip()
                direccion = str(sheet.cell(row=row, column=COL_DIRECCION).value or "Sin Dirección").strip()
                val_servicio = sheet.cell(row=row, column=COL_SERVICIO).value

                num_cuenta = formatear_cuenta(val_servicio)

                # Si no hay número de cuenta válido, lo registra como omitido y sigue
                if not num_cuenta or num_cuenta == "None":
                    registro = (
                        f"Usuario: {usuario}\n"
                        f"Dirección: {direccion}\n"
                        f"N° Cuenta: (Vacío o No especificado)\n"
                        f"Omitido: No posee número de cuenta en el Excel.\n"
                        f"-------------------------------------------------------\n"
                    )
                    f_out.write(registro)
                    f_out.flush()
                    continue

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
                except Exception as e:
                    continue

                # 3. Seleccionar 'Tipo Deuda' buscando la opción de sanitarios
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

                # 4. Clic en 'Iniciar Sesión' con selector robusto por clase o texto
                try:
                    btn_iniciar = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Iniciar') or contains(., 'Sesión')] | //input[@type='submit' or @type='button']"))
                    )
                    btn_iniciar.click()
                except Exception:
                    # Intento alternativo por si cambia la estructura del botón
                    driver.execute_script("document.querySelector('button').click();")

                time.sleep(3)

                # 5. Desplegar menú 'Consultas' a la izquierda y hacer clic en 'Consulta de deuda'
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

                # 6. Leer la tabla y filtrar estrictamente SERV.SANITAR
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

                # 7. Guardar en el reporte .txt
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
