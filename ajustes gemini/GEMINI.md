# Contexto del Proyecto: Reto 1 - Instalaciones Eléctricas (Panadería Industrial)

Este repositorio contiene los documentos, tablas de cálculo y especificaciones técnicas para el dimensionado y diseño de la instalación eléctrica de baja tensión (BT) para una **Panadería Industrial**.

## Marco Normativo y Técnico
- **Reglamento Electrotécnico para Baja Tensión (REBT - RD 842/2002)** e Instrucciones Técnicas Complementarias (ITC-BT):
  - **ITC-BT-19**: Instalaciones interiores o receptoras (sección de conductores por criterio térmico y caída de tensión máx. admisible).
  - **ITC-BT-22 y 24**: Protecciones contra sobreintensidades (cortocircuitos y sobrecargas) y contra contactos directos e indirectos (diferenciales, puesta a tierra).
  - **ITC-BT-29**: Prescripciones particulares para instalaciones eléctricas en locales con riesgo de incendio o explosión (ATEX - polvo combustible de harina en zonas de silos y amasado).
  - **ITC-BT-30**: Instalaciones en locales de características especiales (húmedos, mojados, con riesgo de corrosión).
  - **ITC-BT-47**: Receptores - Motores eléctricos (intensidad de arranque, factor de 1.25 para el conductor del motor individual, etc.).
- **Normas UNE / IEC de aplicación** (UNE 20460-5-523 / HD 60364-5-52 para capacidades de corriente de cables según método de instalación y temperatura ambiente).

## Estructura de Carpetas y Datos
- **`planos/`**: Plano base de la nave en AutoCAD (`Nave 2.dwg`).
- **`calculos/`**: Hoja de cálculo (`Receptores_panaderia_v2.xlsx`) con el inventario de maquinaria y receptores (hornos, amasadoras, batidoras, silos, divisoras, climatización, alumbrado, etc.), sus potencias nominales, tensiones (230 V monofásica / 400 V trifásica), factores de potencia ($\cos \varphi$), intensidades calculadas y factores de simultaneidad y utilización.
- **`proyecto/`**: Documento oficial del proyecto en Word (`Plantilla_Proyecto_Instalacion_Electrica_Panaderia.docx`), estructurado según UNE 157001.
- **`referencias/`**: Documentación y guías académicas:
  - `PROYECTO Año anterior.pdf`: Memoria y proyecto de referencia de cursos anteriores.
  - `TABAJO DOCUMENTAL  SOBRE INSTALACIONES ELECTRICAS.pdf`: Requisitos del reto y especificaciones académicas.
- **`ajustes gemini/`**: Scripts de automatización y configuración técnica para Antigravity (`generate_master_word.py`, `inspect_excel.py`).

## Rol y Directrices para Antigravity en este Workspace
- **Perfil**: Actuar como Ingeniero/a Especialista en Instalaciones Eléctricas de Baja Tensión.
- **Idioma**: Español técnico de ingeniería.
- **Rigor en Cálculos**:
  - Explicar siempre las fórmulas utilizadas (e.g., $I = \frac{P}{\sqrt{3} \cdot V \cdot \cos\varphi}$ para trifásica, $I = \frac{P}{V \cdot \cos\varphi}$ para monofásica).
  - Justificar los criterios de cálculo de sección de conductores:
    1. Criterio de intensidad máxima admisible (criterio térmico según tipo de aislamiento PVC/XLPE y método de instalación).
    2. Criterio de caída de tensión máxima admisible ($\Delta V \le 3\%$ para alumbrado, $\Delta V \le 5\%$ para otros usos desde el cuadro general, o según ITC-BT-19).
    3. Criterio de cortocircuito (en conductores de pequeña sección o largas distancias).
  - Prestar especial atención a las zonas con atmósfera potencialmente explosiva por polvo de harina (ITC-BT-29 / ATEX) y locales húmedos (cámaras de fermentación/lavado).
