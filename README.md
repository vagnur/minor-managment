# Minor Management Tools

Herramienta interna para gestionar los procesos operativos del **Minor en Ciencia de Datos**.

El objetivo del proyecto es centralizar distintos scripts utilizados en la coordinación del minor dentro de una sola aplicación con interfaz gráfica, modular y mantenible.

---

# Estado del proyecto

Versión actual: **v0.1**

Módulos implementados:

- ✅ Postulación
- ⏳ Aceptación
- ⏳ Filtrado
- ⏳ Actualización de datos
- ⏳ Inscripción
- ⏳ Procesamiento de notas
- ⏳ Identificación de finalizados

---

# Funcionalidad actual

El módulo **Postulación** permite:

- leer un archivo Excel consolidado de postulaciones
- validar columnas necesarias
- generar formularios Word personalizados
- crear automáticamente carpetas por carrera
- nombrar archivos según el estudiante

Los documentos se generan a partir de una **plantilla Word con marcadores**.

---

# Tecnologías utilizadas

- Python
- CustomTkinter
- Pandas
- python-docx
- openpyxl

---

# Estructura del proyecto

```
scripts_minor/

├── main.py
├── app/
│   ├── core/
│   │   ├── file_utils.py
│   │   ├── excel_utils.py
│   │   ├── validation_utils.py
│   │   └── docx_utils.py
│   │
│   ├── gui/
│   │   ├── main_window.py
│   │   ├── navigation.py
│   │   └── home_view.py
│   │
│   └── modules/
│       └── postulacion/
│           ├── ui.py
│           ├── service.py
│           └── config.py
│
├── config/
├── templates/
├── output/
└── README.md
```

---

# Instalación

Clonar el repositorio:

git clone https://github.com/vagnur/minor-managment.git
cd minor-managment

Crear entorno virtual:

python3 -m venv .venv
source .venv/bin/activate

Instalar dependencias:

pip install -r requirements.txt

---

# Ejecutar el programa

python main.py

---

# Flujo de uso del módulo Postulación

1. Descargar postulaciones desde Google Forms.
2. Revisar manualmente postulantes elegibles.
3. Definir secciones y horarios finales.
4. Consolidar información en un Excel final.
5. Ejecutar el módulo de postulación para generar los formularios.

---

# Objetivo del proyecto

Este sistema busca:

- reducir pasos manuales repetitivos
- centralizar herramientas de gestión
- evitar scripts hardcodeados
- facilitar mantenimiento y extensión del sistema

---

# Autor

**Gabriel Godoy**  
Coordinación Minor en Ciencia de Datos
