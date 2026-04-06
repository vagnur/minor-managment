# Minor Management Tools

Herramienta interna para la gestión operativa del Minor en Ciencia de Datos (USACH).

El sistema centraliza y automatiza procesos clave del Minor, reemplazando scripts dispersos por una aplicación modular con interfaz gráfica.

---

## Estado del proyecto

Versión actual: v0.3 (en desarrollo)

Módulos implementados:
- Postulación
- Aceptación
- Inscripción

Módulos en desarrollo / mejora:
- Filtrado
- Actualización de datos
- Procesamiento de notas
- Identificación de finalizados

---

## Descripción general

La herramienta permite procesar información proveniente de formularios (Google Forms) en formato Excel y generar automáticamente documentos institucionales en Word (.docx), reduciendo significativamente el trabajo manual de la coordinación del Minor.

Cada módulo responde a una etapa del flujo real del programa:
- Postulación → generación de formularios por estudiante
- Aceptación → consolidación de estudiantes aceptados
- Inscripción → formalización de asignaturas por estudiante

---

## Funcionalidades principales

### Postulación
- procesamiento de Excel consolidado desde Google Forms
- generación de formularios individuales en Word
- organización automática por carrera
- validación de estructura de datos

### Aceptación
- generación de documento institucional de aceptación
- normalización de nombres y RUT
- consolidación de estudiantes en tabla única
- compatibilidad con formato institucional

### Inscripción
- procesamiento por asignatura desde Excel con múltiples hojas
- generación de formularios de inscripción en Word
- soporte para distintas estructuras (caso especial TAAA)
- manejo dinámico de horarios (cátedra / laboratorio)
- generación automática de tablas con formato consistente

---

## Uso general

1. Exportar respuestas de Google Forms a Excel  
2. Seleccionar el módulo correspondiente en la aplicación  
3. Configurar parámetros (semestre, fecha, asignaturas)  
4. Ejecutar el proceso  
5. Revisar documentos generados en carpeta de salida  

---

## Requisitos

- Python 3.10+
- Dependencias principales:
  - pandas
  - python-docx
  - openpyxl
  - customtkinter

Instalación sugerida:

```bash
pip install -r requirements.txt