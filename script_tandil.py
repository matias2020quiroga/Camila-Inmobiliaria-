import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# --- CONFIGURACIÓN DE COLUMNAS ---
COL_SERVICIOS = 3    # Columna C
COL_RETRIBUTIVA = 4   # Columna D
COL_MONTO = 5        # Columna E
FILA_INICIAL = 4     # Fila donde empiezan los encabezados/datos

EXCEL_FILE = "propiedades.xlsx"

def consultar_deuda():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active

    # Iniciar navegador Chrome
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    
    url = "https://www.autogestion.tandil.gov.ar/apex/f?p=114:101"

    try:
        for row in range(FILA_INICIAL, sheet.max_row + 1):
            val_servicio = sheet.cell(row=row, column=COL_SERVICIOS).value
            val_retributiva = sheet.cell(row=row, column=COL_RETRIBUTIVA).value

            num_cuenta = None
            tipo_tasa = None

            if val_servicio is not None and str(val_servicio).strip() != "":
                num_cuenta = str(int(val_servicio) if isinstance(val_servicio, float) else val_servicio).strip()
                tipo_tasa = "SERVICIOS SANITARIOS"
            elif val_retributiva is not None and str(val_retributiva).strip() != "":
                num_cuenta = str(int(val_retributiva) if isinstance(val_retributiva, float) else val_retributiva).strip()
                tipo_tasa = "RETRIBUTIVA DE SERVICIOS"
            else:
                continue

            driver.get(url)
            
            dropdown_elem = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "select"))
            )
            select_tasa = Select(dropdown_elem)
            
            for option in select_tasa.options:
                if tipo_tasa.lower() in option.text.lower():
                    select_tasa.select_by_visible_text(option.text)
                    break

            input_cuenta = driver.find_element(By.CSS_SELECTOR, "input[type='text']")
            input_cuenta.clear()
            input_cuenta.send_keys(num_cuenta)

            btn_buscar = driver.find_element(By.XPATH, "//button[contains(text(), 'Consultar') or contains(text(), 'Buscar')]")
            btn_buscar.click()

            time.sleep(2)

            try:
                monto_elem = driver.find_element(By.XPATH, "//*[contains(text(), '$')]")
                monto_texto = monto_elem.text
            except Exception:
                monto_texto = "Sin Deuda / No Encontrado"

            sheet.cell(row=row, column=COL_MONTO).value = monto_texto
            
            wb.save(EXCEL_FILE)
            print(f"Fila {row}: Cuenta {num_cuenta} ({tipo_tasa}) -> Monto: {monto_texto}")

    finally:
        driver.quit()
        wb.save(EXCEL_FILE)
        print("Proceso finalizado.")

if __name__ == "__main__":
    consultar_deuda()
