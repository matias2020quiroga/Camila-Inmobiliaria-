import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# --- CONFIGURACIÓN DE COLUMNAS DEL EXCEL (Basado en la nueva imagen) ---
FILA_INICIAL = 2       # Fila 2 (la Fila 1 contiene 'usuario', 'dirección', 'Servicio'...)
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
                usuario = str(sheet.cell(row=row, column=COL_USUARIO).value or "Sin Nombre").strip()
                direccion = str(sheet.cell(row=row, column=COL_DIRECCION).value or "Sin Dirección").strip()
                val_servicio = sheet.cell(row=row, column=COL_SERVICIO).value

                num_cuenta = formatear_cuenta(val_servicio)

                # Si la celda de la cuenta/servicio está vacía, pasa al siguiente
                if not num_cuenta:
                    continue

                # 1. Ingresar a la página principal
                driver.get(URL_TANDIL)
                time.sleep(1)

                # 2. Completar Nro. Cuenta / Patente / Código
                input_cuenta = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
                )
                input_cuenta.clear()
                input_cuenta.send_keys(num_cuenta)

                # 3. Seleccionar 'Tipo Deuda'
                dropdown_elem = driver.find_element(By.TAG_NAME, "select")
                select_tasa = Select(dropdown_elem)
                
                for option in select_tasa.options:
                    if "sanitar" in option.text.lower():
                        select_tasa.select_by_visible_text(option.text)
                        break

                time.sleep(1)

                # 4. Clic en 'Iniciar Sesión'
                btn_iniciar = driver.find_element(By.XPATH, "//button[contains(text(), 'Iniciar') or contains(text(), 'Sesión')] | //input[@type='button' or @type='submit']")
                btn_iniciar.click()
                time.sleep(3)

                # 5. Desplegar menú 'Consultas' a la izquierda
                try:
                    btn_consultas = WebDriverWait(driver, 8).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Consultas')]"))
                    )
                    btn_consultas.click()
                    time.sleep(1.5)

                    # 6. Hacer clic en sub-menú 'Consulta de deuda'
                    btn_sub_deuda = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Consulta de deuda')]"))
                    )
                    btn_sub_deuda.click()
                    time.sleep(3)
                except Exception:
                    pass  # En caso de ingresar directamente al reporte

                # 7. Leer la tabla y filtrar estrictamente SERV.SANITAR
                detalles_deuda = []
                monto_total_sanitar = 0.0

                try:
                    filas_tabla = driver.find_elements(By.XPATH, "//tr")
                    for fila in filas_tabla:
                        texto_fila = fila.text
                        # Validar si la fila contiene 'SERV.SANITAR'
                        if "SERV.SANITAR" in texto_fila:
                            columnas = fila.find_elements(By.TAG_NAME, "td")
                            # Estructura: Año(0), Tasa(1), Desc. Tasa(2), Cuota(3), Vencimiento(4), Importe(5), Actualizado(6)...
                            if len(columnas) >= 6:
                                cuota = columnas[3].text.strip()
                                vencimiento = columnas[4].text.strip()
                                importe_raw = columnas[5].text.strip()  # Toma la columna 'Importe' marcada en tu imagen
                                
                                # Convertir a float para acumular
                                importe_clean = importe_raw.replace(".", "").replace(",", ".")
                                try:
                                    monto_float = float(importe_clean)
                                    monto_total_sanitar += monto_float
                                except ValueError:
                                    pass

                                detalles_deuda.append(f"  • Cuota {cuota} (Venc: {vencimiento}): Importe = ${importe_raw}")
                except Exception as ex_tabla:
                    detalles_deuda.append(f"  • Error en lectura de tabla: {ex_tabla}")

                # 8. Dar formato e imprimir al reporte .txt
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
