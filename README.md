# Automatizador de Asistencia - Ensamble The Crumbs

## ¿Qué hace este programa?

Lee tu archivo `formato.xlsx` y **envía automáticamente** tus registros de asistencia al Google Forms del proyecto.

**Requisitos:**
- Python 3.8+
- Google Chrome instalado en la computadora
- Tu archivo `formato.xlsx` lleno con tus datos

---

## 📋 Paso 1: Preparar tu Excel

Crea un archivo llamado **`formato.xlsx`** con esta estructura:

| | B: Fecha | C: Lugar | D: Nombre del Evento | E: [TU NOMBRE] | F: Correo | G: Documento |
|---|---------|----------|----------------------|-----------------|----------|------------|
| Fila 2 (headers) | Fecha | Lugar | Nombre del Evento | **ESCRIBE TU NOMBRE AQUÍ** | Correo | Documento |
| Fila 3 | 08/05 | Salon 208 CyT | Taller | Si | micorreo@unal.edu.co | 1010100110 |
| Fila 4 | 08-20/05 | Auditorio CyT | Capacitación PGP | Si | micorreo@unal.edu.co | 1010100110 |
| Fila 5 | 15/05 | Salon 208 CyT | Taller | No | micorreo@unal.edu.co | 1010100110 |

** MUY IMPORTANTE — Celda E2:**
- **E2 debe contener TU NOMBRE** (ejemplo: "Daniel Alfonso", "Felipe Gutierrez", etc.)
- El programa lee tu nombre de ahí y lo usa en el log

**Columnas (columna E en adelante):**
- ✅ **Columna E (Tu nombre)**: Escribe `Si` si asististe, `No` si no asististe
- ✅ **Correo**: Debe ser `xxxxx@unal.edu.co` (obligatorio)
- ✅ **Documento**: Debe ser números sin puntos ni comas (obligatorio)
- ✅ **Fecha**: Formato `DD/MM` o `DD/MM/YYYY` (obligatorio)
- ✅ **Lugar**: No puede estar vacío (obligatorio)
- ✅ **Nombre del Evento**: No puede estar vacío (obligatorio)

**Solo se envían registros donde tu nombre (columna E) = "Si".**

Si falta algún dato, **el programa te lo va a indicar y no enviará nada** hasta que lo corrijas.

---

## Paso 2: Instalar requisitos

### En Windows (PowerShell):
```powershell
pip install openpyxl selenium webdriver-manager
```

### En Mac/Linux (Terminal):
```bash
pip3 install openpyxl selenium webdriver-manager
```

---

## Paso 3: Ejecutar el programa

### Copia estos archivos en **la misma carpeta**:
- `enviar_asistencia.py` (el script)
- `formato.xlsx` (tu archivo con datos)

### En Windows (PowerShell):
```powershell
python enviar_asistencia.py "https://docs.google.com/forms/d/e/1FAIpQLSc...../viewform"
```

### En Mac/Linux (Terminal):
```bash
python3 enviar_asistencia.py "https://docs.google.com/forms/d/e/1FAIpQLSc...../viewform"
```

**Nota:** Reemplaza la URL con la del Forms que te proporcione el programa.

---

## ¿Qué pasa cuando ejecutas?

El programa:

1. ** Lee tu Excel** → Busca registros donde `Nombre_Estudiante = Sí`
2. ** Valida datos** → Verifica que todos los campos sean correctos
3. ** Abre el navegador** → Chrome se abre automáticamente (en segundo plano)
4. ** Completa el Forms** → Llena cada campo del formulario
5. ** Envía cada registro** → Un registro por sesión
6. ** Genera un log** → Crea `asistencia_log.txt` con los resultados

**Ejemplo de salida:**

```
============================================================
🎵 AUTOMATIZADOR DE ASISTENCIA - ENSAMBLE THE CRUMBS 🎵
============================================================

📁 Leyendo Excel...
✅ Se encontraron 5 registros

🔍 Validando datos...
✅ Todos los datos son válidos

🌐 Abriendo Forms...

📋 Procesando 5 registros...

[1/5] Enviando: Taller (08/05/2026)...
    ✅ Enviado exitosamente
[2/5] Enviando: Capacitación PGP (08-20/05/2026)...
    ✅ Enviado exitosamente
...

📄 Log guardado en: asistencia_log.txt

============================================================
RESUMEN FINAL
============================================================
✅ Enviados: 5
❌ Fallidos: 0
📄 Log: asistencia_log.txt
============================================================
```

---

## Solución de problemas

### "No se encontró: formato.xlsx"
**Solución:** Asegúrate que el archivo `formato.xlsx` esté en la **misma carpeta** que el script.

### "Error al iniciar Chrome"
**Solución:** 
1. Instala Chrome desde https://www.google.com/chrome
2. O ejecuta: `pip install webdriver-manager`

### "Correo inválido (debe ser @unal.edu.co)"
**Solución:** Tu correo debe terminar en `@unal.edu.co`, no otro dominio.

### "Documento debe ser numérico"
**Solución:** El documento no puede tener puntos ni comas. Usa solo números: `1010100110`

### "Fecha vacía" o "Formato de fecha inválido"
**Solución:** La fecha debe ser:
- `08/05` (día/mes)
- `08/05/2026` (día/mes/año)
- NO válido: `2026-05-08`

### El navegador se abre pero no llena nada
**Solución:** El Forms tiene una estructura diferente a la esperada. Avisa al coordinador.

---

## Ayuda

Si algo no funciona:
1. **Revisa el archivo `asistencia_log.txt`** → Busca el error específico
2. **Verifica tu Excel** → Asegúrate que todos los datos estén correctos
3. **Contacta al coordinador** → Si persiste el error

---

## Tips

- ✅ Llena el Excel en orden: Fecha → Lugar → Evento → Sí/No → Correo → Documento
- ✅ Si solo asisten en algunas sesiones, **marca las otras con "No"** (no serán enviadas)
- ✅ Ejecuta el programa **cerca del límite de plazo** (no hay urgencia de hacerlo inmediatamente)
- ✅ Si algo se engulle, solo ejecuta el programa de nuevo — detecta qué ya fue enviado

---

**Versión:** 1.0  
**Última actualización:** Agosto 2026  
**Proyecto:** Ensamble The Crumbs - Proyectos Estudiantiles PGP
