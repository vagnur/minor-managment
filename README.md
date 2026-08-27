# Minor Management Tools

Herramienta interna para la gestión operativa del Minor en Ciencia de Datos (USACH).

El sistema centraliza y automatiza procesos clave del Minor, reemplazando scripts dispersos por una aplicación modular con interfaz gráfica.

---

## Requisitos

- Python 3.10+
- Dependencias definidas en `requirements.txt`

Instalación de dependencias:

```bash
pip install -r requirements.txt
```

---

## Instalación y ejecución

Clonar el repositorio:

```bash
git clone https://github.com/vagnur/minor-managment.git
```

Acceder a la carpeta del proyecto:

```bash
cd minor-managment
```

Ejecutar la aplicación desde la raíz del proyecto:

```bash
python3 main.py
```

---

## Uso

Los procesos disponibles se encuentran organizados en el menú lateral de acuerdo con las distintas etapas de gestión del Minor:

- **Inicio de semestre**
- **Fin de semestre**
- **Utilidades**

Cada módulo incluye una guía que indica:

- qué proceso realiza;
- qué archivo o carpeta requiere como entrada;
- qué archivos genera como resultado;
- los pasos necesarios para su ejecución, cuando corresponde.

---

## Estructura del proyecto

```text
minor-managment/
├── app/
│   ├── core/          # Componentes y configuración general
│   ├── gui/           # Ventana principal y navegación
│   └── modules/       # Módulos funcionales
│
├── config/            # Archivos de configuración de los procesos
├── templates/         # Plantillas institucionales utilizadas por los módulos
├── main.py            # Punto de entrada de la aplicación
├── requirements.txt   # Dependencias de Python
└── CHANGELOG.md       # Registro de cambios del proyecto
```