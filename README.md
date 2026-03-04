# 🚀 AutoInput

> Automatización inteligente de ingreso de datos en Windows mediante simulación de teclado a bajo nivel.

AutoInput es una herramienta de automatización para **Windows** que permite cargar datos desde archivos `.txt` y escribirlos automáticamente en otras aplicaciones (Excel, SAP, formularios web, sistemas legacy, etc.) simulando pulsaciones reales de teclado.

Ideal para flujos repetitivos de carga masiva de datos.

---

## 📌 Tabla de Contenidos

* [✨ Características](#-características)
* [🛠️ Requisitos](#️-requisitos)
* [📦 Instalación](#-instalación)
* [🚀 Uso](#-uso)
* [⚙️ Generar Ejecutable (.exe)](#️-generar-ejecutable-exe)
* [📂 Estructura del Proyecto](#-estructura-del-proyecto)
* [🧠 Arquitectura](#-arquitectura)
* [⚠️ Consideraciones](#️-consideraciones)
* [📄 Licencia](#-licencia)

---

## ✨ Características

### 📄 Lectura Masiva

* Carga archivos `.txt`
* Procesamiento línea por línea
* Ideal para grandes volúmenes de datos

### 🪟 Gestión de Ventanas

* Detecta aplicaciones abiertas
* Permite acoplar automáticamente la ventana destino a la derecha
* Supervisión visual durante la ejecución

### 📊 Modo Excel Inteligente

* Detecta automáticamente si el destino es Excel
* Formatea celdas como texto
* Previene errores de redondeo y pérdida de ceros a la izquierda

### 🎛️ Control Total en Tiempo Real

* Pausar
* Continuar
* Cancelar proceso

### ⏱️ Velocidad Ajustable

* Control mediante slider
* Permite adaptar la velocidad según la aplicación destino

---

## 🛠️ Requisitos

| Requisito         | Versión               |
| ----------------- | --------------------- |
| Sistema Operativo | Windows 10 o superior |
| Python            | 3.8+                  |

### 📚 Librerías necesarias

```bash
pip install pywin32 customtkinter
```

---

## 📦 Instalación

1. Clona el repositorio:

```bash
git clone https://github.com/tu-usuario/AutoInput.git
cd AutoInput
```

2. Instala dependencias:

```bash
pip install -r requirements.txt
```

> Si no tienes `requirements.txt`, usa:
> `pip install pywin32 customtkinter`

---

## 🚀 Uso

Ejecuta el archivo principal:

```bash
python main.py
```

### 🖥️ Flujo en la Interfaz

1. Haz clic en **"Seleccionar TXT"**
2. Selecciona la ventana de destino
3. Ajusta la velocidad
4. Marca "Limpiar campo" si es necesario
5. Presiona **"Comenzar"**

---

## ⚙️ Generar Ejecutable (.exe)

Para crear un ejecutable independiente:

### 1️⃣ Instalar PyInstaller

```bash
pip install pyinstaller
```

### 2️⃣ Generar el ejecutable

```bash
pyinstaller --noconsole --onefile --name "AutoInput" main.py
```

### 3️⃣ Resultado

El archivo `.exe` se generará en:

```
/dist/AutoInput.exe
```

---

## 📂 Estructura del Proyecto

```
AutoInput/
│
├── main.py                # Punto de entrada
├── app_ui.py              # Interfaz gráfica y control de hilos
├── automation_engine.py   # Lógica Win32 + COM Excel
└── README.md
```

---

## 🧠 Arquitectura

El proyecto está dividido en tres capas principales:

### 🎨 UI Layer (`app_ui.py`)

* Manejo de interfaz con `customtkinter`
* Gestión de estados
* Control de hilos
* Barra de progreso

### ⚙️ Automation Layer (`automation_engine.py`)

* Interacción con Win32 API
* Simulación de teclado a bajo nivel
* Control de ventanas
* Integración con objetos COM de Excel

### 🚀 Entry Point (`main.py`)

* Inicialización de la aplicación
* Inyección de dependencias
* Arranque del loop principal

---

## ⚠️ Consideraciones

* La aplicación simula pulsaciones reales de teclado.
* No se debe usar mientras se trabaja activamente en otra ventana.
* Se recomienda cerrar aplicaciones innecesarias durante la ejecución.
* Puede ser detectado como comportamiento automatizado por algunos sistemas corporativos.


