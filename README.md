# Reto 1: Instalaciones Eléctricas - Panadería Industrial

Proyecto de ingeniería para el dimensionado, diseño y justificación documental de la instalación eléctrica de baja tensión (BT) para una **Panadería Industrial**, cumpliendo con el **REBT (Reglamento Electrotécnico para Baja Tensión)** y sus Instrucciones Técnicas Complementarias (ITC-BT).

---

## 📁 Contenido del Repositorio

- **`Plantilla_Proyecto_Instalacion_Electrica_Panaderia.docx`**: Plantilla maestra oficial en formato Word para el desarrollo del proyecto, estructurada según UNE 157001, con portada reglamentaria, índice, fórmulas teóricas explicadas paso a paso, tablas de cálculo, pliego, presupuesto y checklist para calificación 10.
- **`Receptores_panaderia_v2.xlsx`**: Hoja de cálculo con el listado de receptores, potencias unitarias y totales, tensiones (230 V / 400 V), factores de potencia ($\cos \varphi$), intensidades nominales y notas de cálculo.
- **`Nave 2.dwg`**: Archivo de dibujo técnico en AutoCAD proporcionado por el profesor como base geométrica para el dimensionado y distribución de la nave.
- **`PROYECTO Año anterior.pdf`**: Memoria técnica y proyecto de ejemplo de cursos anteriores como referencia de formato y estructura.
- **`TABAJO DOCUMENTAL  SOBRE INSTALACIONES ELECTRICAS.pdf`**: Requisitos del reto y especificaciones académicas del trabajo.
- **`ajustes gemini/`**: Carpeta reservada para la configuración, prompts y scripts de asistencia con IA / Antigravity:
  - **`GEMINI.md`**: Directrices técnicas, normativas REBT y rol de ingeniería para que Antigravity trabaje como copiloto técnico.
  - **`generate_master_word.py`**: Script de automatización que construye y actualiza la plantilla de Word.

---

## 🚀 Guía Rápida para tu Compañero (Configuración con Antigravity)

Para trabajar juntos de forma sincronizada:

### 1. Clonar el repositorio
Abre una terminal (PowerShell, Git Bash o CMD) en la carpeta donde guardas tus proyectos de ingeniería y ejecuta:
```bash
git clone https://github.com/ferranio0/reto-1-instalaciones.git
cd reto-1-instalaciones
```

### 2. Abrir en Antigravity
1. Inicia **Antigravity** (la aplicación de escritorio o IDE).
2. Selecciona **Open Folder** (o `Ctrl + O` / `Cmd + O`) y elige la carpeta `reto-1-instalaciones` recién descargada.
3. ¡Listo! Antigravity reconocerá las directrices del archivo [`ajustes gemini/GEMINI.md`](ajustes%20gemini/GEMINI.md), conociendo de inmediato el contexto del reto, los receptores de la panadería y las normativas aplicables (REBT, ITC-BT-19, ITC-BT-29, ITC-BT-47, etc.).

---

## 🔄 Flujo de Trabajo en Equipo con Git

Para evitar conflictos y mantener el proyecto al día:

1. **Antes de empezar a trabajar**, descarga los últimos cambios:
   ```bash
   git pull origin main
   ```
2. **Para guardar tus avances**:
   ```bash
   git add .
   git commit -m "Descripción clara de lo que has añadido o modificado"
   git push origin main
   ```
3. **Recomendación para tareas grandes**: Crea una rama propia para trabajar en paralelo sin pisaros:
   ```bash
   git checkout -b feature/nombre-de-la-tarea
   # Trabajas, haces commits y luego:
   git push origin feature/nombre-de-la-tarea
   ```
