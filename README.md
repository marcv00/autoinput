Aquí tienes una **versión optimizada del README enfocada específicamente en el uso con *SirsiDynix WorkFlows***, manteniendo un tono más profesional y alineado con el objetivo real de la herramienta.

También eliminé partes que no aplican tanto (como Excel inteligente) y reforcé el caso de uso **bibliotecario / sistemas legacy tipo WorkFlows**.

---

````md
# 🚀 AutoInput Workflows

> Automatización segura y controlada de ingreso masivo de datos en **SirsiDynix WorkFlows** mediante simulación de teclado a bajo nivel.

**AutoInput Workflows** es una herramienta diseñada para **automatizar el ingreso repetitivo de datos en SirsiDynix WorkFlows**, permitiendo cargar grandes listas de identificadores (ISBN, códigos de barras, IDs de registros, etc.) desde archivos `.txt` y enviarlos automáticamente al campo activo del sistema.

La aplicación simula **pulsaciones reales de teclado**, garantizando compatibilidad con sistemas legacy que no permiten integraciones directas o APIs.

---

# 📌 Tabla de Contenidos

- [✨ Características](#-características)
- [🛠️ Requisitos](#️-requisitos)
- [📦 Instalación](#-instalación)
- [🚀 Uso](#-uso)
- [⚙️ Generar Ejecutable (.exe)](#️-generar-ejecutable-exe)
- [📂 Estructura del Proyecto](#-estructura-del-proyecto)
- [🧠 Arquitectura](#-arquitectura)
- [⚠️ Consideraciones](#️-consideraciones)


---

# ✨ Características

## 📄 Procesamiento Masivo de Datos

- Carga archivos `.txt`
- Procesamiento línea por línea
- Diseñado para grandes volúmenes de códigos o identificadores
- Ideal para operaciones repetitivas en WorkFlows

Ejemplos de uso:

- Ingreso de **ISBN**
- Ingreso de **códigos de barras**
- Procesamiento de **IDs de registros**
- Flujos de revisión masiva

---

## 🪟 Integración con SirsiDynix WorkFlows

- Detecta ventanas abiertas de WorkFlows
- Permite seleccionar la ventana destino
- Envía entradas directamente al campo activo

El sistema **no interactúa con la base de datos**, únicamente simula entradas de teclado tal como lo haría un operador humano.

## 🎛️ Automatización Segura

AutoInput incluye mecanismos para evitar errores durante la ejecución:

- Cuenta regresiva antes de iniciar
- Panel de instrucciones visible
- Control del delay entre entradas
- Confirmación visual del progreso

Esto permite al operador **posicionar el cursor correctamente antes de iniciar la automatización**.

---

## ⏱️ Control de Velocidad

La velocidad de ejecución puede ajustarse mediante un **slider de delay** entre cada entrada.

Esto permite adaptarse a:

- velocidad de respuesta del sistema
- carga del servidor
- estabilidad de WorkFlows

---

## 📊 Seguimiento del Proceso

Durante la ejecución se muestra:

- progreso del proceso
- número de items procesados
- tiempo total de ejecución
- estado actual del sistema

---

# 🛠️ Requisitos

| Requisito | Versión |
|--------|--------|
| Sistema Operativo | Windows 10 o superior |
| Python | 3.8+ |

### 📚 Librerías necesarias

```bash
pip install pywin32 customtkinter pillow
````

---

# 📦 Instalación

1. Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/AutoInput.git
cd AutoInput
```

2. Instalar dependencias

```bash
pip install -r requirements.txt
```

Si no existe `requirements.txt`:

```bash
pip install pywin32 customtkinter pillow
```

---

# 🚀 Uso

Ejecutar la aplicación:

```bash
python main.py
```

---

## 🖥️ Flujo de Uso

1️⃣ Seleccionar archivo `.txt` con los datos a procesar
2️⃣ Seleccionar la ventana de **SirsiDynix WorkFlows**
3️⃣ Ajustar el **delay entre entradas**
4️⃣ Confirmar si se desea **limpiar el campo antes de cada entrada**
5️⃣ Presionar **Comenzar**

---

## ⏳ Proceso de Inicio

Antes de iniciar la automatización, el sistema mostrará una **cuenta regresiva de 10 segundos** para permitir al operador:

* posicionar el cursor en el campo correcto
* verificar que WorkFlows esté listo
* evitar errores de entrada

---

# ⚙️ Generar Ejecutable (.exe)

Para distribuir la aplicación sin necesidad de instalar Python:

## 1️⃣ Instalar PyInstaller

```bash
pip install pyinstaller
```

---

## 2️⃣ Generar ejecutable

```bash
pyinstaller --noconsole --onefile --name "AutoInput" --add-data "refresh.png;." main.py
```

---

## 3️⃣ Resultado

El ejecutable se generará en:

```
/dist/AutoInputWorkflows.exe
```

---

# 📂 Estructura del Proyecto

```
AutoInput/
│
├── main.py
├── app_ui.py
├── automation_engine.py
├── refresh.png
└── README.md
```

---

# 🧠 Arquitectura

El proyecto está dividido en tres componentes principales.

---

## 🎨 UI Layer (`app_ui.py`)

Responsable de la interfaz gráfica.

Funciones principales:

* Interfaz con **CustomTkinter**
* Selección de archivos
* Selección de ventana objetivo
* Control de ejecución
* Barra de progreso
* Panel de instrucciones

---

## ⚙️ Automation Layer (`automation_engine.py`)

Contiene la lógica de automatización.

Responsabilidades:

* Interacción con **Win32 API**
* Detección de ventanas activas
* Simulación de teclado
* Envío de texto al sistema destino

Esta capa es independiente de la interfaz.

---

## 🚀 Entry Point (`main.py`)

Punto de inicio de la aplicación.

Responsabilidades:

* inicializar la interfaz
* lanzar el event loop principal
* cargar dependencias

---

# ⚠️ Consideraciones

* La herramienta **simula pulsaciones reales de teclado**.
* No debe utilizarse mientras se interactúa con otras aplicaciones.
* Se recomienda verificar que el **cursor esté correctamente posicionado** antes de iniciar.
* Siempre realizar pruebas con **archivos pequeños antes de ejecutar grandes volúmenes**.
