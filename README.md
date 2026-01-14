# Generador de Códigos QR desde Excel

Proyecto en Python que permite:
- Leer contactos desde un archivo Excel (debe existir un archivo inicial llamado contactos.xlsx, debes crearlo tu mismo)
- Generar códigos QR en formato vCard
- Insertarlos en una nueva hoja del mismo Excel
- Visualización vertical para fácil navegación (se terminara generando un nuevo archivo con el resultado)

## 📋 Requisitos

- Python 3.9 o superior
- pandas
- qrcode
- openpyxl
- pillow

## 📦 Instalación

1. Clonar el repositorio:
```bash
git clone https://github.com/tu_usuario/qr-contactos-excel.git
```
2. entrar al proyecto
```bash
cd qr-contactos-excel
```
3. Crear entorno virtual:
```bash
python -m venv venv
```

4. Activar entorno virtual:
```bash
venv\Scripts\activate
```
5. Instalar dependencias:
```bash
pip install -r requirements.txt
```

## 📁 Estructura del proyecto
```
qr-contactos-excel/
├── main.py
├── contactos.xlsx (este archivo no va estar inicialmente, debes crearlo y poner tu constactos)
├── README.md
├── .gitignore
└── venv/
```
## 📊 Estructura obligatoria del archivo Excel
El script **espera un archivo Excel** llamado `contactos.xlsx` con una hoja llamada: `Contactos`

### 🔑 Reglas importantes

- La **primera fila** debe contener los encabezados
- Los nombres de las columnas deben escribirse **exactamente así**
- No deben existir filas vacías entre contactos
- La columna `telefono` debe estar en formato **Texto** (no número)

---

### 📋 Columnas requeridas

| Columna   | Tipo recomendado | Descripción |
|----------|------------------|-------------|
| nombre   | Texto            | Nombre del contacto |
| apellido | Texto            | Apellido del contacto |
| telefono | Texto            | Número de celular (incluya indicativo si aplica) |
| email    | Texto            | Correo electrónico |

📌 Ejemplo válido:

| nombre | apellido | telefono   | email |
|------|---------|------------|-------|
| Juan | Pérez   | 3001234567 | juan@email.com |
| Ana  | Gómez   | 3109876543 | ana@email.com |

---

### ⚠️ Errores comunes

❌ Cambiar nombres de columnas (`Nombre`, `tel`, etc.)  
❌ Teléfono como número (pierde ceros o formato)  
❌ Hoja con nombre distinto a `Contactos`

---

### 🧠 Personalización

Si deseas usar otros nombres de columnas o agregar campos como `empresa`,
debes modificar la función:

```python
fila_a_vcard(fila)
```

## ▶️ Uso
Coloca el archivo contactos.xlsx en la raíz del proyecto y ejecuta:
```bash
python main.py
```
Se generará un archivo: `contactos_con_qr.xlsx` abre el archivo y podras ver tus codigos QR generados

