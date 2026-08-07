#!/usr/bin/env python3
"""
Automatizador de envío de asistencia a Google Forms
Ensamble de música - Proyectos Estudiantiles PGP 2026

Uso:
    python3 enviar_asistencia.py <URL_FORMS>
"""

import sys
import os
import re
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
import time

# Constantes
PROYECTO = "Ensamble The Crumbs"
TIPO_USUARIO = "Estudiante de pregrado"
EXCEL_FILE = "formato.xlsx"
LOG_FILE = "asistencia_log.txt"

class ValidadorDatos:
    """Valida datos del Excel antes de enviar"""
    
    @staticmethod
    def validar_fecha(fecha_str):
        if not fecha_str:
            return False, "Fecha vacía"
        
        formatos = ["%d/%m", "%d/%m/%Y", "%Y-%m-%d"]
        for fmt in formatos:
            try:
                datetime.strptime(str(fecha_str).strip(), fmt)
                return True, None
            except:
                continue
        
        return False, f"Formato de fecha inválido: {fecha_str}"
    
    @staticmethod
    def validar_correo(correo):
        if not correo:
            return False, "Correo vacío"
        
        correo = str(correo).strip()
        if not re.match(r"^[\w\.-]+@unal\.edu\.co$", correo):
            return False, f"Correo inválido (debe ser @unal.edu.co): {correo}"
        
        return True, None
    
    @staticmethod
    def validar_documento(doc):
        if not doc:
            return False, "Documento vacío"
        
        if not str(doc).strip().isdigit():
            return False, f"Documento debe ser numérico: {doc}"
        
        return True, None
    
    @staticmethod
    def validar_texto(texto, campo_nombre):
        if not texto or (isinstance(texto, str) and not texto.strip()):
            return False, f"{campo_nombre} vacío"
        return True, None

class LectorExcel:
    """Lee y procesa el archivo Excel"""
    
    def __init__(self, ruta_excel):
        if not Path(ruta_excel).exists():
            raise FileNotFoundError(f"No se encontró: {ruta_excel}")
        
        self.ruta = ruta_excel
        self.datos = []
        self.nombre_estudiante = None
        self.leer()
    
    def leer(self):
        wb = load_workbook(self.ruta)
        ws = wb.active
        
        nombre_cell = ws['E2']
        if not nombre_cell.value:
            raise ValueError("❌ La casilla E2 está vacía. Debe contener el nombre del estudiante.")
        
        self.nombre_estudiante = str(nombre_cell.value).strip()
        
        headers = {}
        for col_idx, cell in enumerate(ws[2], 1):
            if cell.value:
                headers[cell.value.strip()] = col_idx
        
        if not all(h in headers for h in ["Fecha", "Lugar", "Nombre del Evento", "Correo"]):
            raise ValueError("Excel debe tener columnas: Fecha, Lugar, Nombre del Evento, Correo")
        
        col_fecha = headers["Fecha"]
        col_lugar = headers["Lugar"]
        col_evento = headers["Nombre del Evento"]
        col_asistencia = 5  # Columna E
        col_correo = headers["Correo"]
        col_documento = headers.get("Documento", None)
        
        for row_idx in range(3, ws.max_row + 1):
            fecha = ws.cell(row_idx, col_fecha).value
            lugar = ws.cell(row_idx, col_lugar).value
            evento = ws.cell(row_idx, col_evento).value
            asistio = ws.cell(row_idx, col_asistencia).value
            correo = ws.cell(row_idx, col_correo).value
            doc = ws.cell(row_idx, col_documento).value if col_documento else None
            
            if asistio and str(asistio).strip().lower() == "si":
                self.datos.append({
                    "fecha": fecha,
                    "lugar": lugar,
                    "evento": evento,
                    "correo": correo,
                    "documento": doc,
                    "fila": row_idx
                })
    
    def obtener_nombre_estudiante(self):
        return self.nombre_estudiante
    
    def obtener_registros(self):
        return self.datos

class AutomatizadorForms:
    """Automatiza el envío al Google Forms con Selenium adaptado para formularios multisección"""
    
    def __init__(self, url_forms):
        self.url_forms = url_forms
        self.driver = None
        self.registros_enviados = 0
        self.registros_fallidos = 0
        self.errores = []
        self._iniciar_driver()
    
    def _iniciar_driver(self):
        chrome_options = ChromeOptions()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
        except Exception as e:
            print(f"❌ Error al iniciar Chrome: {e}")
            print("   Asegúrate de tener Chrome instalado y chromedriver disponible")
            print("   Instala: pip install webdriver-manager")
            sys.exit(1)
    
    def _formato_fecha_digits(self, fecha):
        """Convierte la fecha a dígitos continuos (DDMMYYYY) para inputs de fecha de Google Forms"""
        if not fecha:
            return ""
        fecha_str = str(fecha).strip()
        
        match = re.match(r"^(\d{2})/(\d{2})(?:/(\d{4}))?$", fecha_str)
        if match:
            d, m, y = match.groups()
            if not y:
                y = "2026"
            return f"{d}{m}{y}"
        
        try:
            dt = datetime.strptime(fecha_str, "%Y-%m-%d")
            return dt.strftime("%d%m%Y")
        except:
            pass
        
        return re.sub(r"\D", "", fecha_str)

    def _llenar_campo_por_texto(self, texto_pregunta, valor):
        """Ubica el bloque de pregunta por su título y escribe en su input/textarea"""
        if valor is None:
            valor = ""
        texto_lower = texto_pregunta.lower()
        xpath_pregunta = (
            "//div[contains(@role, 'listitem') or contains(@class, 'geS5ne')]"
            "[.//span[contains("
            "translate(normalize-space(string(.)), "
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚ', "
            "'abcdefghijklmnopqrstuvwxyzáéíóú'), "
            f"\"{texto_lower}\")]]"
        )
        container = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located((By.XPATH, xpath_pregunta))
        )
        inp = container.find_element(By.XPATH, ".//input | .//textarea")
        inp.clear()
        inp.send_keys(str(valor))

    def _seleccionar_dropdown_por_texto(self, texto_pregunta, valor_opcion):
        """Abre un menú desplegable de Google Forms y selecciona una opción"""
        texto_lower = texto_pregunta.lower()
        xpath_pregunta = (
            "//div[contains(@role, 'listitem') or contains(@class, 'geS5ne')]"
            "[.//span[contains("
            "translate(normalize-space(string(.)), "
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚ', "
            "'abcdefghijklmnopqrstuvwxyzáéíóú'), "
            f"\"{texto_lower}\")]]"
        )
        container = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located((By.XPATH, xpath_pregunta))
        )
        
        dropdown = container.find_element(By.XPATH, ".//div[@role='listbox']")
        dropdown.click()
        time.sleep(1)
        
        opcion = WebDriverWait(self.driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, f"//div[@role='option']//span[contains(text(), '{valor_opcion}')]"))
        )
        opcion.click()
        time.sleep(0.5)

    def _click_boton(self, texto_boton):
        """Hace clic en un botón de Google Forms (div role='button') por su etiqueta de texto"""
        xpath_btn = f"//div[@role='button'][.//span[contains(text(), '{texto_boton}')]]"
        btn = WebDriverWait(self.driver, 8).until(
            EC.element_to_be_clickable((By.XPATH, xpath_btn))
        )
        btn.click()
        time.sleep(2)

    def enviar_registro(self, registro):
        import traceback
        try:
            self.driver.get(self.url_forms)
            time.sleep(3)

            # --- SECCIÓN 1: Correo ---
            self._llenar_campo_por_texto("correo", registro["correo"])
            self._click_boton("Siguiente")

            # --- SECCIÓN 2: Información del Evento ---
            # 1. Nombre del proyecto
            self._seleccionar_dropdown_por_texto("Nombre del proyecto", PROYECTO)
            
            # 2. Capacitación - Taller - Reunión - Evento (Nombre del evento)
            self._llenar_campo_por_texto("Capacitación", registro["evento"])
            
            # 3. Fecha
            fecha_digits = self._formato_fecha_digits(registro["fecha"])
            self._llenar_campo_por_texto("Fecha", fecha_digits)
            
            # 4. Lugar / Enlace
            self._llenar_campo_por_texto("Lugar / Enlace", registro["lugar"])
            
            # 5. Tipo de usuario
            self._seleccionar_dropdown_por_texto("Tipo de usuario", TIPO_USUARIO)
            
            self._click_boton("Siguiente")

            # --- SECCIÓN 3: Confirmación de Identidad ---
            # 1. Documento de identidad
            if registro.get("documento"):
                self._llenar_campo_por_texto("Documento de identidad", registro["documento"])
            
            # 2. Correo UNAL
            self._llenar_campo_por_texto("Correo UNAL", registro["correo"])

            # Enviar formulario
            self._click_boton("Enviar")
            time.sleep(2)
            return True

        except Exception as e:
            print(f"    💥 EXCEPCIÓN: {type(e).__name__}: {e}")
            traceback.print_exc()
            try:
                screenshot_path = f"error_fila_{registro['fila']}.png"
                self.driver.save_screenshot(screenshot_path)
                print(f"    📸 Screenshot guardado: {screenshot_path}")
            except:
                pass
            self.errores.append(f"Fila {registro['fila']}: {str(e)}")
            return False
    
    def procesar_lote(self, registros, nombre_estudiante):
        print(f"\n📋 Procesando {len(registros)} registros...\n")
        
        for i, registro in enumerate(registros, 1):
            print(f"[{i}/{len(registros)}] Enviando: {registro['evento']} ({registro['fecha']})...")
            
            if self.enviar_registro(registro):
                self.registros_enviados += 1
                print(f"    ✅ Enviado exitosamente")
            else:
                self.registros_fallidos += 1
                print(f"    ❌ Error al enviar")
            
            time.sleep(1)
        
        self._generar_log(nombre_estudiante)
        self.cerrar()
    
    def _generar_log(self, nombre_estudiante):
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            f.write("=" * 60 + "\n")
            f.write("LOG DE ENVÍO - ASISTENCIA A REUNIONES\n")
            f.write(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")
            
            f.write(f"Estudiante: {nombre_estudiante}\n")
            f.write(f"Proyecto: {PROYECTO}\n")
            f.write(f"Tipo de usuario: {TIPO_USUARIO}\n\n")
            
            f.write(f"Registros enviados: {self.registros_enviados}\n")
            f.write(f"Registros fallidos: {self.registros_fallidos}\n")
            f.write(f"Total procesado: {self.registros_enviados + self.registros_fallidos}\n\n")
            
            if self.errores:
                f.write("ERRORES:\n")
                for error in self.errores:
                    f.write(f"  - {error}\n")
            else:
                f.write("✅ No se registraron errores\n")
        
        print(f"\n📄 Log guardado en: {LOG_FILE}")
    
    def cerrar(self):
        if self.driver:
            self.driver.quit()

def main():
    if len(sys.argv) < 2:
        print("Uso: python3 enviar_asistencia.py <URL_FORMS>")
        sys.exit(1)
    
    url_forms = sys.argv[1]
    
    print("\n" + "=" * 60)
    print("🎵 AUTOMATIZADOR DE ASISTENCIA - ENSAMBLE THE CRUMBS 🎵")
    print("=" * 60 + "\n")
    
    try:
        print("📁 Leyendo Excel...")
        lector = LectorExcel(EXCEL_FILE)
        nombre_estudiante = lector.obtener_nombre_estudiante()
        registros = lector.obtener_registros()
        
        print(f"👤 Estudiante: {nombre_estudiante}")
        
        if not registros:
            print(f"❌ No hay registros con asistencia (Sí) para {nombre_estudiante}")
            sys.exit(1)
        
        print(f"✅ Se encontraron {len(registros)} registros de asistencia")
        
        print("\n🔍 Validando datos...")
        validador = ValidadorDatos()
        
        for registro in registros:
            valido, msg = validador.validar_fecha(registro["fecha"])
            if not valido:
                print(f"❌ Fila {registro['fila']}: {msg}")
                sys.exit(1)
            
            valido, msg = validador.validar_correo(registro["correo"])
            if not valido:
                print(f"❌ Fila {registro['fila']}: {msg}")
                sys.exit(1)
            
            if registro.get("documento"):
                valido, msg = validador.validar_documento(registro["documento"])
                if not valido:
                    print(f"❌ Fila {registro['fila']}: {msg}")
                    sys.exit(1)
            
            valido, msg = validador.validar_texto(registro["evento"], "Nombre del Evento")
            if not valido:
                print(f"❌ Fila {registro['fila']}: {msg}")
                sys.exit(1)
            
            valido, msg = validador.validar_texto(registro["lugar"], "Lugar")
            if not valido:
                print(f"❌ Fila {registro['fila']}: {msg}")
                sys.exit(1)
        
        print("✅ Todos los datos son válidos\n")
        
        print(f"🌐 Abriendo Forms: {url_forms[:60]}...\n")
        automatizador = AutomatizadorForms(url_forms)
        automatizador.procesar_lote(registros, nombre_estudiante)
        
        print("\n" + "=" * 60)
        print("RESUMEN FINAL")
        print("=" * 60)
        print(f"👤 Estudiante: {nombre_estudiante}")
        print(f"✅ Enviados: {automatizador.registros_enviados}")
        print(f"❌ Fallidos: {automatizador.registros_fallidos}")
        print(f"📄 Log: {LOG_FILE}")
        print("=" * 60 + "\n")
    
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()