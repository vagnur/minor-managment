# Minor Management Tools

Herramienta interna para gestionar los procesos operativos del Minor en Ciencia de Datos.

El objetivo del proyecto es centralizar distintos scripts utilizados en la coordinación del minor dentro de una sola aplicación con interfaz gráfica, modular y mantenible.

---

# Estado del proyecto

Versión actual: v0.2 (en desarrollo)

Módulos implementados:

- Postulación
- Aceptación

Módulos planificados:

- Filtrado
- Actualización de datos
- Inscripción
- Procesamiento de notas
- Identificación de finalizados

---

# Funcionalidad actual

## Postulación
- lectura de Excel consolidado
- generación automática de formularios Word
- creación de carpetas por carrera
- validaciones de columnas

## Aceptación
- lectura de Excel de estudiantes aceptados
- generación automática del documento institucional de ingreso
- normalización de RUT
- heurística de separación de nombres
- mejora de formato de tabla DOCX

---

# Tecnologías utilizadas

- Python
- CustomTkinter
- Pandas
- python-docx
- openpyxl

---

# Estructura del proyecto

app/
core/
gui/
modules/
config/
output/