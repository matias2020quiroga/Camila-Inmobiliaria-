import openpyxl
import requests
from bs4 import BeautifulSoup
import time

# --- CONFIGURACIÓN DE COLUMNAS (Excel) ---
FILA_INICIAL = 4       # Fila donde empiezan los datos
COL_NOMBRE = 1         # Columna A: Nombre / Cliente
COL_DIRECCION = 2      # Columna B: Dirección / Inmueble
COL_SERVICIOS = 3      # Columna C: Cuenta Servicios Sanitarios

EXCEL_FILE = "propiedades.xlsx"
TXT_OUTPUT = "resultados_servicios_sanitarios.txt"

URL_TANDIL = "https://www.autogestion.tandil.gov.ar/apex/f?p=114:101"

def consultar_servicios_sanitarios():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb.active
    except Exception as e:
        print(f"Error al abrir {EXCEL_FILE}: {e}")
        return

    with open(TXT_OUTPUT, "w", encoding="utf-8") as f_out:
        f_out.write("=== REPORTE DE DEUDAS - SERVICIOS SANITARIOS ===\n")
        f_out.write(f"Fecha de consulta: {time.strftime('%d/%m/%Y %H:%M')}\n")
        f_out.write("=" * 50 + "\n\n")

        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })

        for row in range(FILA_INICIAL, sheet.max_row + 1):
            nombre = str(sheet.cell(row=row, column=COL_NOMBRE).value or "Sin Nombre").strip()
            direccion = str(sheet.cell(row=row, column=COL_DIRECCION).value or "Sin Dirección").strip()
            val_servicio = sheet.cell(row=row, column=COL_SERVICIOS).value

            if val_servicio is None or str(val_servicio).strip() == "":
                continue  # Salta filas sin cuenta de Servicios Sanitarios

            num_cuenta = str(int(val_servicio) if isinstance(val_servicio, float) else val_servicio).strip()

            try:
                # 1. Obtener página inicial
                resp_get = session.get(URL_TANDIL, timeout=10)
                soup = BeautifulSoup(resp_get.text, 'html.parser')

                # 2. Preparar consulta enviando número de cuenta y tipo de tasa fijo
                payload = {
                    "P101_TIPO_IMPUESTO": "SERVICIOS SANITARIOS",
                    "P101_CUENTA": num_cuenta
                }

                resp_post = session.post(URL_TANDIL, data=payload, timeout=10)
                soup_res = BeautifulSoup(resp_post.text, 'html.parser')

                # 3. Buscar el texto con el monto/deuda
                monto_texto = "Sin Deuda / No Encontrado"
                for elem in soup_res.find_all(text=True):
                    if "$" in elem:
                        monto_texto = elem.strip()
                        break

            except Exception as ex:
                monto_texto = f"Error de consulta: {ex}"

            # 4. Guardar en el TXT
            linea = (
                f"Cliente: {nombre}\n"
                f"Dirección: {direccion}\n"
                f"Tasa: SERVICIOS SANITARIOS\n"
                f"N° Cuenta: {num_cuenta}\n"
                f"Monto Deuda: {monto_texto}\n"
                f"----------------------------------------\n"
            )
            f_out.write(linea)
            f_out.flush()

        print("Consulta finalizada. Se generó 'resultados_servicios_sanitarios.txt'")

if __name__ == "__main__":
    consultar_servicios_sanitarios()
