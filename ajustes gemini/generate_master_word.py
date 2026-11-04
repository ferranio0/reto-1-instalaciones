# -*- coding: utf-8 -*-
"""
Generador de la Plantilla Maestra en Word (.docx) para el Proyecto de Instalación Eléctrica
de Baja Tensión para una Panadería Industrial (UPV - Reto 1 Instalaciones).
Actualizado con:
- Estructura UNE 157001 completa.
- Índice General, Índice de Tablas e Índice de Figuras.
- Sección completa y rigurosa de Antecedentes y Objeto según el modelo "Nave 2" y el Excel v2.
- Anexo oficial de Fichas de Materiales.
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_document():
    doc = docx.Document()
    
    # Configuración de márgenes estándar (2.5 cm en todos los lados)
    for section in doc.sections:
        section.top_margin = Inches(0.98)     # ~2.5 cm
        section.bottom_margin = Inches(0.98)
        section.left_margin = Inches(0.98)
        section.right_margin = Inches(0.98)
        section.different_first_page_header_footer = True
        
        # Encabezado
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("PROYECTO DE INSTALACIÓN ELÉCTRICA BT — PANADERÍA INDUSTRIAL | RETO 1")
        hrun.font.name = "Calibri"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(100, 116, 139)
        
        # Pie de página
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        frun1 = fp.add_run("ETSID - Universidad Politécnica de Valencia | Grado en Ingeniería Mecánica")
        frun1.font.name = "Calibri"
        frun1.font.size = Pt(8.5)
        frun1.font.color.rgb = RGBColor(100, 116, 139)

    # Estilos de párrafos y fuentes
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Funciones de apoyo XML
    def set_cell_background(cell, hex_color):
        tcPr = cell._tc.get_or_add_tcPr()
        for child in list(tcPr):
            if child.tag.endswith('shd'):
                tcPr.remove(child)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        tcPr.append(shd)

    def set_cell_margins(cell, top=120, bottom=120, left=180, right=180):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'  <w:top w:w="{top}" w:type="dxa"/>'
            f'  <w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'  <w:left w:w="{left}" w:type="dxa"/>'
            f'  <w:right w:w="{right}" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

    def set_table_borders(table, color="CBD5E0", sz="4", val="single"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:left w:val="none"/>'
            f'  <w:right w:val="none"/>'
            f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def set_callout_border(cell, border_color="1B365D", sz="24"):
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'  <w:top w:val="none"/>'
            f'  <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{border_color}"/>'
            f'  <w:bottom w:val="none"/>'
            f'  <w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)

    def add_callout(title, text, box_type="tip"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        
        if box_type == "tip":
            bg_color = "EBF4FF"
            border_color = "1B365D"
            prefix = "💡 GUÍA DE EXCELENCIA (PARA SACAR UN 10):"
            tcolor = RGBColor(27, 54, 93)
        elif box_type == "warning":
            bg_color = "FEF2F2"
            border_color = "DC2626"
            prefix = "⚠️ REQUISITO CRÍTICO DE EVALUACIÓN (EVITAR EL CERO):"
            tcolor = RGBColor(220, 38, 38)
        else:
            bg_color = "F8FAFC"
            border_color = "475569"
            prefix = "📝 INSTRUCCIONES DE DESARROLLO Y REDACCIÓN:"
            tcolor = RGBColor(71, 85, 105)
            
        set_cell_background(cell, bg_color)
        set_callout_border(cell, border_color, sz="24")
        set_cell_margins(cell, top=140, bottom=140, left=200, right=160)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        run_p = p.add_run(f"{prefix} {title}\n")
        run_p.bold = True
        run_p.font.name = "Calibri"
        run_p.font.size = Pt(9.5)
        run_p.font.color.rgb = tcolor
        
        run_t = p.add_run(text)
        run_t.font.name = "Calibri"
        run_t.font.size = Pt(9.0)
        run_t.font.italic = True
        run_t.font.color.rgb = RGBColor(51, 65, 85)
        
        pa = doc.add_paragraph()
        pa.paragraph_format.space_before = Pt(0)
        pa.paragraph_format.space_after = Pt(4)

    def add_formula_card(title, formula, description, variables):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        set_cell_background(cell, "F1F5F9")
        set_callout_border(cell, "0D5C75", sz="16")
        set_cell_margins(cell, top=120, bottom=120, left=180, right=160)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        
        r_title = p.add_run(f"📌 {title}\n")
        r_title.bold = True
        r_title.font.name = "Calibri"
        r_title.font.size = Pt(9.5)
        r_title.font.color.rgb = RGBColor(13, 92, 117)
        
        r_desc = p.add_run(f"{description}\n\n")
        r_desc.font.name = "Calibri"
        r_desc.font.size = Pt(9.0)
        
        r_formula = p.add_run(f"    {formula}\n\n")
        r_formula.bold = True
        r_formula.font.name = "Consolas"
        r_formula.font.size = Pt(10)
        r_formula.font.color.rgb = RGBColor(15, 23, 42)
        
        r_vars_head = p.add_run("Donde:\n")
        r_vars_head.bold = True
        r_vars_head.font.name = "Calibri"
        r_vars_head.font.size = Pt(8.5)
        
        for var, desc in variables:
            r_var = p.add_run(f"  • {var}: ")
            r_var.bold = True
            r_var.font.name = "Calibri"
            r_var.font.size = Pt(8.5)
            r_d = p.add_run(f"{desc}\n")
            r_d.font.size = Pt(8.5)
            
        pa = doc.add_paragraph()
        pa.paragraph_format.space_before = Pt(0)
        pa.paragraph_format.space_after = Pt(4)

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(27, 54, 93)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(13, 92, 117)
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(51, 65, 85)
        return p

    # ==========================================
    # 1. PORTADA NORMALIZADA UNE 157001
    # ==========================================
    p_upv = doc.add_paragraph()
    p_upv.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_upv = p_upv.add_run("UNIVERSIDAD POLITÉCNICA DE VALENCIA\n")
    r_upv.bold = True
    r_upv.font.size = Pt(13)
    r_upv.font.color.rgb = RGBColor(71, 85, 105)
    r_esc = p_upv.add_run("ESCUELA TÉCNICA SUPERIOR DE INGENIERÍA DEL DISEÑO\n")
    r_esc.bold = True
    r_esc.font.size = Pt(11)
    r_esc.font.color.rgb = RGBColor(100, 116, 139)
    r_asig = p_upv.add_run("GRADO EN INGENIERÍA MECÁNICA — ASIGNATURA: INSTALACIONES ELÉCTRICAS (RETO 1)")
    r_asig.font.size = Pt(10)
    r_asig.font.color.rgb = RGBColor(100, 116, 139)
    
    doc.add_paragraph().paragraph_format.space_before = Pt(36)

    # Recuadro con el título principal
    tbl_title = doc.add_table(rows=1, cols=1)
    tbl_title.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_title.autofit = False
    cell_t = tbl_title.cell(0, 0)
    cell_t.width = Inches(6.5)
    set_cell_background(cell_t, "1B365D")
    set_cell_margins(cell_t, top=260, bottom=260, left=240, right=240)
    
    p_t = cell_t.paragraphs[0]
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t1 = p_t.add_run("PROYECTO TÉCNICO DE INSTALACIÓN ELÉCTRICA EN BAJA TENSIÓN PARA UNA INDUSTRIA DE PANIFICACIÓN Y BOLLERÍA\n")
    r_t1.bold = True
    r_t1.font.size = Pt(15.5)
    r_t1.font.color.rgb = RGBColor(255, 255, 255)
    
    r_t2 = p_t.add_run("CON INFRAESTRUCTURA DE RECARGA DE VEHÍCULOS ELÉCTRICOS (ITC-BT-52) Y SISTEMA DE GENERACIÓN SOLAR FOTOVOLTAICA (ITC-BT-40)")
    r_t2.font.size = Pt(11.0)
    r_t2.font.color.rgb = RGBColor(226, 232, 240)
    
    doc.add_paragraph().paragraph_format.space_before = Pt(30)
    
    # Cuadro de metadatos de autoría y firma digital
    tbl_meta = doc.add_table(rows=5, cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_meta.autofit = False
    set_table_borders(tbl_meta, color="CBD5E0")
    
    col_widths = [Inches(2.2), Inches(4.3)]
    meta_rows = [
        ("AUTORES DEL PROYECTO:", "[Nombre y Apellidos Alumno 1] — Grupo ____\n[Nombre y Apellidos Alumno 2] — Grupo ____\n[Nombre y Apellidos Alumno 3] — Grupo ____"),
        ("TUTOR / DEPARTAMENTO:", "Departamento de Ingeniería Eléctrica — UPV"),
        ("CONVOCATORIA Y CURSO:", "Curso Académico 2026 / 2027"),
        ("ESTRUCTURA DOCUMENTAL:", "Doc. 1: Memoria Descriptiva | Doc. 2: Cálculos Justificativos\nDoc. 3: Pliego de Condiciones | Doc. 4: Presupuesto | Doc. 5: Planos A3 | Doc. 6: Bibliografía | Doc. 7: Anexos"),
        ("FIRMA DIGITAL REGLAMENTARIA:", "Espacio reservado para el estampillado y firma digital del PDF\n(Requisito obligatorio según enunciado oficial del Reto)")
    ]
    
    for i, (label, val) in enumerate(meta_rows):
        row = tbl_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_margins(c0, top=80, bottom=80, left=120, right=120)
        set_cell_margins(c1, top=80, bottom=80, left=120, right=120)
        set_cell_background(c0, "F8FAFC")
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.size = Pt(8.5)
        r0.font.color.rgb = RGBColor(71, 85, 105)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(val)
        r1.font.size = Pt(8.5)
        if "firma digital" in val.lower():
            r1.font.italic = True
            r1.font.color.rgb = RGBColor(148, 163, 184)
            
    doc.add_page_break()

    # ==========================================
    # 2. CONTROL DE VERSIONES Y CAMBIOS
    # ==========================================
    add_h2("CONTROL DE REVISIONES Y REGISTRO DE MODIFICACIONES")
    
    p_ver_desc = doc.add_paragraph("Para asegurar la trazabilidad del trabajo en equipo según la norma UNE 157001, cada cambio relevante en el documento debe registrarse en la siguiente tabla:")
    
    tbl_rev = doc.add_table(rows=4, cols=4)
    tbl_rev.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_rev, color="CBD5E0")
    rev_widths = [Inches(1.0), Inches(1.3), Inches(1.8), Inches(2.4)]
    
    rev_headers = ["Revisión", "Fecha", "Responsable", "Descripción de los Cambios"]
    for j, h in enumerate(rev_headers):
        c = tbl_rev.rows[0].cells[j]
        c.width = rev_widths[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, top=80, bottom=80, left=100, right=100)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    rev_data = [
        ("Rev. 0.1", "Noviembre 2026", "Equipo de Trabajo", "Creación de la plantilla maestra estructurada según UNE 157001."),
        ("Rev. 0.2", "Noviembre 2026", "Equipo de Trabajo", "Redacción de Antecedentes y Objeto según Nave 2 y Receptores v2 (159,54 kW)."),
        ("Rev. 1.0", "[Pendiente]", "Todo el Equipo", "Integración de cálculos térmicos finales, planos A3, presupuesto y firma digital.")
    ]
    for i, data in enumerate(rev_data):
        row = tbl_rev.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]
            c.width = rev_widths[j]
            set_cell_margins(c, top=70, bottom=70, left=100, right=100)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # 3. ÍNDICES REGLAMENTARIOS (GENERAL, TABLAS, FIGURAS)
    # ==========================================
    add_h2("ÍNDICE GENERAL DEL PROYECTO")
    
    add_callout(
        "CÓMO ACTUALIZAR EL ÍNDICE EN MICROSOFT WORD",
        "Esta plantilla utiliza títulos automáticos (Título 1, Título 2, Título 3). Cuando hayáis terminado de redactar o modificar apartados, haced clic derecho sobre el índice de Word y pulsad 'Actualizar campos' (o seleccionad el índice y presionad la tecla F9) para actualizar automáticamente los números de página.",
        box_type="tip"
    )
    
    toc_items = [
        ("RESUMEN EJECUTIVO (ABSTRACT)", "Pág. 5"),
        ("ÍNDICE DE TABLAS", "Pág. 6"),
        ("ÍNDICE DE FIGURAS Y ESQUEMAS", "Pág. 6"),
        ("DOCUMENTO Nº 1: MEMORIA DESCRIPTIVA", "Pág. 7"),
        ("  1.1 Antecedentes", "Pág. 7"),
        ("  1.2 Objeto del Proyecto", "Pág. 7"),
        ("  1.3 Características del Edificio y Geometría de la Nave (Modelo Nave 2)", "Pág. 8"),
        ("  1.4 Entorno Exterior y Zona de Aparcamiento con Puntos de Recarga", "Pág. 9"),
        ("  1.5 Características de la Actividad y Previsión de Potencia (159,54 kW)", "Pág. 9"),
        ("  1.6 Clasificación Reglamentaria de los Locales (ATEX Zona 22, Húmedos, Hornos)", "Pág. 10"),
        ("  1.7 Suministro Eléctrico e Instalación de Enlace (CGP, LGA, Contadores, DI)", "Pág. 12"),
        ("  1.8 Arquitectura del Sistema de Distribución y Cuadros Eléctricos", "Pág. 13"),
        ("  1.9 Red de Puesta a Tierra y Protección Equipotencial (ITC-BT-18)", "Pág. 15"),
        ("  1.10 Alumbrado Normal, de Emergencia y Señalización", "Pág. 16"),
        ("DOCUMENTO Nº 2: CÁLCULOS JUSTIFICATIVOS", "Pág. 18"),
        ("  2.1 Criterios Normativos y Metodología de Cálculo", "Pág. 18"),
        ("  2.2 Tabla Maestra de Cálculo de Circuitos y Líneas de la Panadería", "Pág. 21"),
        ("  2.3 Justificación de la Protección contra Contactos Indirectos y Red de Tierra", "Pág. 23"),
        ("  2.4 Justificación de la Protección contra Sobretensiones (ITC-BT-23)", "Pág. 25"),
        ("  2.5 Compensación de Reactiva y Filtrado de Armónicos en Obrador", "Pág. 26"),
        ("  2.6 Infraestructura para Recarga de Vehículos Eléctricos (ITC-BT-52)", "Pág. 27"),
        ("  2.7 Dimensionado del Sistema Solar Fotovoltaico en Cubierta (ITC-BT-40)", "Pág. 28"),
        ("DOCUMENTO Nº 3: PLIEGO DE CONDICIONES Y SELECCIÓN DE MATERIALES", "Pág. 30"),
        ("  3.1 Normativa Técnica y Reglamentación Aplicable", "Pág. 30"),
        ("  3.2 Especificación Justificada de Aparamenta y Materiales de Referencia", "Pág. 31"),
        ("  3.3 Prescripciones de Montaje y Tendido de Canalizaciones", "Pág. 34"),
        ("  3.4 Pruebas Reglamentarias y Ensayos de Puesta en Servicio (ITC-BT-05)", "Pág. 35"),
        ("DOCUMENTO Nº 4: PRESUPUESTO", "Pág. 37"),
        ("  4.1 Cuadro de Precios Unitarios", "Pág. 37"),
        ("  4.2 Presupuesto Desglosado por Capítulos", "Pág. 38"),
        ("  4.3 Resumen General del Presupuesto (PEM, GG, BI, IVA)", "Pág. 39"),
        ("DOCUMENTO Nº 5: ÍNDICE DE PLANOS", "Pág. 40"),
        ("DOCUMENTO Nº 6: BIBLIOGRAFÍA Y REFERENCIAS NORMATIVAS (UPV / IEEE)", "Pág. 42"),
        ("DOCUMENTO Nº 7: ANEXO DE FICHAS TÉCNICAS DE MATERIALES", "Pág. 44"),
        ("ANEXO: CHECKLIST DE AUTOEVALUACIÓN PARA OBTENER CALIFICACIÓN 10", "Pág. 46")
    ]
    
    tbl_toc = doc.add_table(rows=len(toc_items), cols=2)
    tbl_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_toc, color="E2E8F0")
    for idx, (item, page) in enumerate(toc_items):
        r = tbl_toc.rows[idx]
        c0, c1 = r.cells[0], r.cells[1]
        c0.width, c1.width = Inches(5.3), Inches(1.2)
        set_cell_margins(c0, top=40, bottom=40, left=80, right=80)
        set_cell_margins(c1, top=40, bottom=40, left=80, right=80)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        run0 = p0.add_run(item)
        run0.font.size = Pt(8.5)
        if item.startswith("DOCUMENTO") or item.startswith("ÍNDICE") or item.startswith("RESUMEN") or item.startswith("ANEXO"):
            run0.bold = True
            run0.font.color.rgb = RGBColor(27, 54, 93)
            set_cell_background(c0, "F8FAFC")
            set_cell_background(c1, "F8FAFC")
            
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_after = Pt(0)
        run1 = p1.add_run(page)
        run1.font.size = Pt(8.5)
        run1.font.color.rgb = RGBColor(100, 116, 139)
        
    doc.add_page_break()

    # ==========================================
    # ÍNDICE DE TABLAS E ÍNDICE DE FIGURAS
    # ==========================================
    add_h2("ÍNDICE DE TABLAS")
    doc.add_paragraph("Relación ordenada de tablas técnicas contenidas en el proyecto según la norma UNE 157001:")
    
    tablas_list = [
        ("Tabla 1", "Registro y Control de Revisiones del Proyecto", "Pág. 3"),
        ("Tabla 2", "Superficies Útiles y Construidas de la Nave 2 (Planta Baja y Altillo)", "Pág. 8"),
        ("Tabla 3", "Clasificación de Dependencias y Tipología de Riesgo REBT / ATEX", "Pág. 11"),
        ("Tabla 4", "Previsión de Potencias por Zonas y Subcuadros de Distribución", "Pág. 14"),
        ("Tabla 5", "Tabla Maestra de Dimensionado de Circuitos y Líneas Eléctricas", "Pág. 21"),
        ("Tabla 6", "Justificación de Resistencia de Tierra y Tensiones de Defecto", "Pág. 24"),
        ("Tabla 7", "Especificación de Aparamenta y Materiales con Fabricantes de Referencia", "Pág. 32"),
        ("Tabla 8", "Resumen de Presupuesto por Capítulos de Obra", "Pág. 38"),
        ("Tabla 9", "Resumen General Económico de la Instalación (PEM, PEC e IVA)", "Pág. 39"),
        ("Tabla 10", "Relación de Planos Oficiales del Proyecto (Formato A3 Fondo Blanco)", "Pág. 40")
    ]
    
    tbl_tab = doc.add_table(rows=len(tablas_list)+1, cols=3)
    tbl_tab.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_tab, color="E2E8F0")
    
    # Headers
    h_c0, h_c1, h_c2 = tbl_tab.rows[0].cells[0], tbl_tab.rows[0].cells[1], tbl_tab.rows[0].cells[2]
    h_c0.width, h_c1.width, h_c2.width = Inches(1.2), Inches(4.3), Inches(1.0)
    set_cell_background(h_c0, "1B365D"); set_cell_background(h_c1, "1B365D"); set_cell_background(h_c2, "1B365D")
    h_c0.paragraphs[0].add_run("Número").bold = True; h_c0.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255); h_c0.paragraphs[0].runs[0].font.size = Pt(8.5)
    h_c1.paragraphs[0].add_run("Denominación de la Tabla").bold = True; h_c1.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255); h_c1.paragraphs[0].runs[0].font.size = Pt(8.5)
    h_c2.paragraphs[0].add_run("Página").bold = True; h_c2.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255); h_c2.paragraphs[0].runs[0].font.size = Pt(8.5); h_c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    for idx, (num, desc, pg) in enumerate(tablas_list):
        r = tbl_tab.rows[idx+1]
        c0, c1, c2 = r.cells[0], r.cells[1], r.cells[2]
        c0.width, c1.width, c2.width = Inches(1.2), Inches(4.3), Inches(1.0)
        set_cell_margins(c0, top=40, bottom=40, left=70, right=70)
        set_cell_margins(c1, top=40, bottom=40, left=70, right=70)
        set_cell_margins(c2, top=40, bottom=40, left=70, right=70)
        if idx % 2 == 1:
            set_cell_background(c0, "F8FAFC"); set_cell_background(c1, "F8FAFC"); set_cell_background(c2, "F8FAFC")
        p0 = c0.paragraphs[0]; p0.paragraph_format.space_after = Pt(0); r0 = p0.add_run(num); r0.bold = True; r0.font.size = Pt(8.5)
        p1 = c1.paragraphs[0]; p1.paragraph_format.space_after = Pt(0); r1 = p1.add_run(desc); r1.font.size = Pt(8.5)
        p2 = c2.paragraphs[0]; p2.paragraph_format.space_after = Pt(0); p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; r2 = p2.add_run(pg); r2.font.size = Pt(8.5); r2.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    add_h2("ÍNDICE DE FIGURAS Y ESQUEMAS")
    doc.add_paragraph("Relación de figuras, croquis y esquemas técnicos exigidos en el enunciado del reto:")
    
    figuras_list = [
        ("Figura 1", "Diagrama de Bloques del Flujo Productivo de la Panadería Industrial", "Pág. 7"),
        ("Figura 2", "Geometría en Planta en 'L' de la Nave 2 y Delimitación de Altillo", "Pág. 8"),
        ("Figura 3", "Zonificación ATEX (Zona 22) por Polvo Combustible de Harina en Silos", "Pág. 10"),
        ("Figura 4", "Esquema de Distribución Jerárquica: CGP, LGA, Centralización y Subcuadros", "Pág. 13"),
        ("Figura 5", "Curvas Características de Disparo de Magnetotérmicos Curva C y Curva D", "Pág. 19"),
        ("Figura 6", "Esquema de Conexión a Tierra TT y Bucle de Defecto a Masa", "Pág. 23"),
        ("Figura 7", "Coordinación de Descargadores de Sobretensión Tipo 1+2 y Tipo 2", "Pág. 25"),
        ("Figura 8", "Triángulo de Potencias y Compensación con Batería Desintonizada", "Pág. 26"),
        ("Figura 9", "Esquema de Modulación Dinámica de Recarga de VE con Sistema SPL (ITC-BT-52)", "Pág. 27")
    ]
    
    tbl_fig = doc.add_table(rows=len(figuras_list)+1, cols=3)
    tbl_fig.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_fig, color="E2E8F0")
    
    # Headers
    h_c0, h_c1, h_c2 = tbl_fig.rows[0].cells[0], tbl_fig.rows[0].cells[1], tbl_fig.rows[0].cells[2]
    h_c0.width, h_c1.width, h_c2.width = Inches(1.2), Inches(4.3), Inches(1.0)
    set_cell_background(h_c0, "1B365D"); set_cell_background(h_c1, "1B365D"); set_cell_background(h_c2, "1B365D")
    h_c0.paragraphs[0].add_run("Número").bold = True; h_c0.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255); h_c0.paragraphs[0].runs[0].font.size = Pt(8.5)
    h_c1.paragraphs[0].add_run("Título de la Figura / Esquema").bold = True; h_c1.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255); h_c1.paragraphs[0].runs[0].font.size = Pt(8.5)
    h_c2.paragraphs[0].add_run("Página").bold = True; h_c2.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255); h_c2.paragraphs[0].runs[0].font.size = Pt(8.5); h_c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    for idx, (num, desc, pg) in enumerate(figuras_list):
        r = tbl_fig.rows[idx+1]
        c0, c1, c2 = r.cells[0], r.cells[1], r.cells[2]
        c0.width, c1.width, c2.width = Inches(1.2), Inches(4.3), Inches(1.0)
        set_cell_margins(c0, top=40, bottom=40, left=70, right=70)
        set_cell_margins(c1, top=40, bottom=40, left=70, right=70)
        set_cell_margins(c2, top=40, bottom=40, left=70, right=70)
        if idx % 2 == 1:
            set_cell_background(c0, "F8FAFC"); set_cell_background(c1, "F8FAFC"); set_cell_background(c2, "F8FAFC")
        p0 = c0.paragraphs[0]; p0.paragraph_format.space_after = Pt(0); r0 = p0.add_run(num); r0.bold = True; r0.font.size = Pt(8.5)
        p1 = c1.paragraphs[0]; p1.paragraph_format.space_after = Pt(0); r1 = p1.add_run(desc); r1.font.size = Pt(8.5)
        p2 = c2.paragraphs[0]; p2.paragraph_format.space_after = Pt(0); p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; r2 = p2.add_run(pg); r2.font.size = Pt(8.5); r2.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_page_break()

    # ==========================================
    # 4. RESUMEN EJECUTIVO (ABSTRACT)
    # ==========================================
    add_h1("RESUMEN EJECUTIVO / ABSTRACT")
    
    doc.add_paragraph(
        "El presente proyecto técnico tiene por objeto el diseño, dimensionamiento, justificación y cálculo "
        "de la instalación eléctrica de baja tensión para una nave industrial dedicada a la panificación y bollería "
        "industrial, situada en el término municipal de Paterna (Valencia). La instalación parte desde la red de distribución "
        "en baja tensión de la compañía suministradora (i-DE Redes Eléctricas Inteligentes, S.A.U.), acometiendo a la Caja "
        "General de Protección (CGP) e incluyendo la Línea General de Alimentación (LGA), equipo de medida fiscal en baja tensión, "
        "Derivación Individual (DI) y Cuadro General de Mando y Protección (CGMP) junto a cinco subcuadros de distribución sectorizada."
    )
    
    doc.add_paragraph(
        "La potencia nominal total instalada suma un total de 159,54 kW (con una capacidad ampliada de 174,34 kW "
        "considerando la reserva del tercer cargador de vehículo eléctrico), cumpliendo rigurosamente con la franja exigida "
        "de entre 120 y 190 kW. La potencia máxima simultánea prevista de cálculo se establece en 138,20 kW tras aplicar "
        "los factores de simultaneidad y utilización característicos de la maquinaria de obrador (hornos de convección Unox, "
        "amasadoras de espiral Bongard, batidoras, cámaras de fermentación Sveba Dahlen, cámaras frigoríficas Intarcon y climatización Mitsubishi)."
    )
    
    doc.add_paragraph(
        "La instalación adopta un esquema de conexión a tierra tipo TT, garantizando la protección contra contactos directos e indirectos "
        "mediante interruptores diferenciales coordinados con una resistencia de puesta a tierra calculada inferior a 1,50 Ω, muy por debajo "
        "del límite reglamentario de 24 V fijado para locales húmedos. Como elementos de excelencia técnica y sostenibilidad, se incluye la "
        "clasificación ATEX para atmósferas explosivas por polvo combustible de harina (Zona 22 según ITC-BT-29), una infraestructura de recarga "
        "inteligente de vehículos eléctricos con modulación de potencia SPL (ITC-BT-52) para 28 plazas de aparcamiento, una instalación solar "
        "fotovoltaica en cubierta para autoconsumo de 25,2 kWp (ITC-BT-40), y una batería automática de condensadores con reactancias de "
        "desintonización para la mitigación de armónicos."
    )
    
    add_h2("Technical Abstract (English)")
    p_en = doc.add_paragraph(
        "This engineering project covers the comprehensive design and electrical sizing of a low-voltage electrical "
        "installation for an industrial bakery and pastry production facility located in Paterna (Valencia, Spain). The electrical "
        "infrastructure operates at 400 V / 230 V (50 Hz) under a TT grounding scheme, with a total installed nominal power of "
        "159.54 kW (expandable to 174.34 kW) and a diversified peak demand of 138.20 kW, fully compliant with the 120–190 kW regulatory "
        "requirement. Key engineering highlights include compliance with the Spanish Low Voltage Electrotechnical Regulations (REBT), "
        "specialized ATEX zoning for combustible flour dust areas (Zone 22, ITC-BT-29), electric vehicle charging infrastructure with "
        "dynamic load balancing (ITC-BT-52) across 28 parking bays, a 25.2 kWp rooftop photovoltaic generation plant for self-consumption "
        "(ITC-BT-40), and an automatic detuned capacitor bank for power factor correction and harmonic distortion mitigation."
    )
    p_en.runs[0].font.italic = True
    
    doc.add_page_break()

    # ==========================================
    # 5. DOCUMENTO Nº 1: MEMORIA DESCRIPTIVA
    # ==========================================
    add_h1("DOCUMENTO Nº 1: MEMORIA DESCRIPTIVA")
    
    add_h2("1.1 Antecedentes")
    doc.add_paragraph(
        "El presente proyecto surge por iniciativa de la mercantil Panificadora Levante S.L., entidad dedicada a la "
        "elaboración y distribución a gran escala de productos de panadería, masas congeladas y bollería industrial "
        "(clasificada en el Código Nacional de Actividades Económicas bajo el CNAE 1071)."
    )
    doc.add_paragraph(
        "La promotora proyecta la construcción e implantación de un nuevo centro de producción en una parcela de uso industrial, "
        "situada en el término municipal de Paterna (Valencia). Dicha parcela cuenta con acceso directo a la red de distribución "
        "pública en baja tensión gestionada por la empresa distribuidora (i-DE Redes Eléctricas Inteligentes, S.A.U.)."
    )
    doc.add_paragraph(
        "La necesidad de redactar este proyecto técnico de ingeniería responde al cumplimiento de los preceptos legales establecidos "
        "en la Ley 21/1992 de Industria y en el Real Decreto 842/2002, de 2 de agosto, por el que se aprueba el Reglamento Electrotécnico "
        "para Baja Tensión (REBT). De acuerdo con el artículo 18 del REBT y las Instrucciones Técnicas Complementarias aplicables, la "
        "instalación requiere obligatoriamente de Proyecto de Instalación Eléctrica visado y tramitado ante el órgano competente de la "
        "Comunidad Autónoma para la obtención de la Autorización de Explotación y el Certificado de Instalación Eléctrica (CIE)."
    )
    
    add_h2("1.2 Objeto del Proyecto")
    doc.add_paragraph("El objeto del presente proyecto de ingeniería comprende los siguientes fines técnicos y administrativos:")
    
    obj_points = [
        ("1. Justificación y Dimensionado Eléctrico:", "Diseñar, dimensionar y justificar técnicamente la totalidad de la infraestructura eléctrica en baja tensión (400 V trifásica y 230 V monofásica, a frecuencia industrial de 50 Hz) para la nueva industria de panificación."),
        ("2. Definición de Redes y Protecciones:", "Definir los trazados de canalizaciones, líneas de distribución, cuadros eléctricos sectorizados, dispositivos de mando, corte y protección (frente a sobrecargas, cortocircuitos, contactos directos/indirectos y sobretensiones), así como la red perimetral de puesta a tierra."),
        ("3. Clasificación de Ambientes Especiales:", "Clasificar reglamentariamente los recintos singulares del proceso fabril, con especial atención a las zonas con riesgo de atmósfera explosiva por polvo combustible de harina (Zona 22 ATEX según ITC-BT-29), locales húmedos/mojados de fermentación y lavado (ITC-BT-30) y áreas de alta temperatura (batería de hornos)."),
        ("4. Sostenibilidad y Movilidad Eléctrica:", "Integrar la infraestructura exterior dedicada a la recarga inteligente de vehículos eléctricos (ITC-BT-52) y el sistema de generación solar fotovoltaica para autoconsumo en cubierta (ITC-BT-40)."),
        ("5. Marco Contractual y Legalización:", "Servir de base técnica y documental vinculante para la ejecución material de las obras, la petición de suministro a la compañía distribuidora y la legalización de la industria ante la Dirección General de Industria de la Generalitat Valenciana.")
    ]
    for o_tit, o_txt in obj_points:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(2)
        r0 = p.add_run(o_tit + " ")
        r0.bold = True
        r0.font.size = Pt(9.0)
        r0.font.color.rgb = RGBColor(13, 92, 117)
        r1 = p.add_run(o_txt)
        r1.font.size = Pt(9.0)

    add_h2("1.3 Características del Edificio y Geometría de la Nave (Modelo Nave 2)")
    doc.add_paragraph(
        "El inmueble se implanta como una nave industrial exenta de tipología estructural metálica porticada, cimentación "
        "por zapatas aisladas enlazadas por vigas de atado y cerramiento exterior de paneles prefabricados de hormigón de 20 cm "
        "de espesor, tal y como se refleja en el plano técnico Nave 2.dwg."
    )
    doc.add_paragraph(
        "A partir del modelo geométrico 'Nave 2' facilitado por el Área de Ingeniería Mecánica y conforme a los condicionantes de "
        "diseño especificados (longitudes principales múltiplos de 0,5 m y altillo de al menos 150 m² sin superar el 30% de la superficie "
        "de planta), se establecen las siguientes cotas y superficies de cálculo:"
    )
    
    tbl_geom = doc.add_table(rows=6, cols=3)
    tbl_geom.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_geom, color="CBD5E0")
    g_widths = [Inches(2.5), Inches(1.5), Inches(2.5)]
    g_headers = ["Zona / Elemento Constructivo", "Superficie Útil", "Cotas y Justificación de Diseño"]
    for j, h in enumerate(g_headers):
        c = tbl_geom.rows[0].cells[j]
        c.width = g_widths[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(255, 255, 255)
        
    geom_data = [
        ("Nave Principal (Planta Baja)", "602,00 m²", "Geometría en 'L' (28,0 m x 24,0 m con retranqueo 14,0 m x 10,0 m). Cotas múltiplos de 0,5 m."),
        ("Altillo o Planta Elevada", "156,00 m² (25,91 %)", "Oficinas, control y vestuarios. Cumple >= 150 m² y <= 30% de la superficie de PB."),
        ("Superficie Total Construida", "758,00 m²", "Suma de Planta Baja + Planta Altillo."),
        ("Altura Libre en Obrador", "7,00 m", "Altura libre bajo cerchas metálicas para ventilación y puente grúa/maquinaria."),
        ("Altura Libre en Altillo", "3,00 m", "Altura libre bajo falso techo continuo desmontable en oficinas.")
    ]
    for i, data in enumerate(geom_data):
        row = tbl_geom.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]; c.width = g_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=80, right=80)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val); r.font.size = Pt(8.5)
            if j == 0 and "Total" in val:
                r.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_h2("1.4 Entorno Exterior y Zona de Aparcamiento con Puntos de Recarga")
    doc.add_paragraph(
        "Alrededor de la nave se dispone de una parcela perimetral urbanizada de 950 m² dotada de:\n"
        "1. Viales de rodadura de 6,0 m de anchura para permitir el radio de giro de camiones de suministro de materias primas a granel (harina) y vehículos de reparto refrigerado.\n"
        "2. Área de aparcamiento exterior con una dotación total de 28 plazas normalizadas (cumpliendo la condición de entre 20 y 40 plazas exigida en el enunciado).\n"
        "3. Puntos de Recarga de Vehículo Eléctrico (ITC-BT-52): Dotación de 4 plazas equipadas con cargadores inteligentes Schneider EVlink Pro AC en Modo 3 (7,4 kW / 32 A en 230 V) con conector Tipo 2 y sistema de modulación dinámica de potencia (SPL).\n"
        "4. Alumbrado exterior: 13 luminarias viarias LED de alta eficiencia tipo Philips CoreLine Malaga BRP102 (IP65, IK08)."
    )

    add_h2("1.5 Características de la Actividad y Previsión de Potencia (159,54 kW)")
    doc.add_paragraph(
        "La instalación eléctrica se dimensiona para dar respuesta al proceso integral de panificación y pastelería industrial, "
        "caracterizado por una demanda intensiva de energía térmica en hornos de cocción continua, fuerza motriz en motores de "
        "amasado y agitación con variadores de frecuencia, producción de frío industrial en cámaras y acondicionamiento ambiental.\n\n"
        "Conforme al inventario detallado de maquinaria y receptores justificado en la hoja de cálculo Receptores_panaderia_v2.xlsx, "
        "la potencia nominal instalada suma un total de 159,54 kW (y 174,34 kW considerando la previsión de ampliación del circuito "
        "de recarga C3), situándose en el rango reglamentario exigido de entre 120 y 190 kW."
    )

    add_h2("1.6 Clasificación Reglamentaria de los Locales (ATEX Zona 22, Húmedos, Hornos)")
    
    add_callout(
        "JUSTIFICACIÓN ATEX Y LOCALES ESPECIALES",
        "La harina de trigo es un polvo combustible (Clase St 1, Kst ~ 100-130 bar·m/s). Clasificar la zona de descarga y silos como ZONA 22 e indicar el tipo de protección de los motores (Ex tc IIIC T125°C IP65) es un factor decisivo de excelencia para alcanzar el 10.",
        box_type="tip"
    )
    
    clasif_items = [
        ("a) Locales con Riesgo de Incendio o Explosión (ITC-BT-29 / ATEX):",
         "La zona de almacenamiento de harina en silo y la tolva de vertido de las amasadoras se clasifican como ZONA 22 (atmósfera explosiva en forma de nube de polvo combustible no previsible en servicio normal o solo durante cortos periodos). Todo el material eléctrico en esta zona dispondrá de marcado ATEX categoría 3D, modo de protección por envolvente contra polvo Ex tc IIIC T125°C y grado de protección mínimo IP65."),
        ("b) Locales Húmedos y Mojados (ITC-BT-30):",
         "La sala de fermentación (humedad relativa > 75%), la zona de lavado de carros/bandejas y las cámaras frigoríficas se clasifican como locales mojados. Se exige aparamenta y envolventes estancas con grado mínimo IPX4 / IPX5, canalizaciones en tubo rígido blindado curvable en caliente y tensión de contacto límite admisible UL = 24 V para el dimensionado de tierras."),
        ("c) Locales a Temperatura Elevada:",
         "El recinto perimetral a la batería de cuatro hornos de convección presenta temperaturas superiores a 35 °C. Se prescriben conductores con aislamiento termoestable de alta temperatura (silicona o XLPE especial) y aplicación de los factores de corrección por temperatura ambiente ft reglamentarios."),
        ("d) Locales de Pública Concurrencia y Vías de Evacuación (ITC-BT-28):",
         "El despacho anexo de venta de pan y los pasillos de evacuación de planta baja y altillo requieren conductores no propagadores del incendio y con emisión de humos y opacidad reducida (cables CPR clase Cca-s1b,d1,a1) y alumbrado de emergencia garantizado."),
        ("e) Infraestructura de Recarga de Vehículos Eléctricos (ITC-BT-52):",
         "Zona exterior de aparcamiento dotada de estación de recarga con cuatro puntos Modo 3, protegidos con diferenciales Tipo A súper inmunizados y vigilancia de corrientes de fuga en continua (6 mA DC) o Tipo B, con modulación dinámica de carga mediante Sistema de Protección de Línea (SPL).")
    ]
    for c_title, c_desc in clasif_items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        r_ct = p.add_run(c_title + "\n")
        r_ct.bold = True; r_ct.font.size = Pt(9.0); r_ct.font.color.rgb = RGBColor(13, 92, 117)
        r_cd = p.add_run(c_desc); r_cd.font.size = Pt(9.0)

    add_h2("1.7 Suministro Eléctrico e Instalación de Enlace (CGP, LGA, Contadores, DI)")
    doc.add_paragraph(
        "La instalación de enlace conecta la red de distribución pública en baja tensión de la empresa "
        "distribuidora (i-DE Redes Eléctricas Inteligentes) con la instalación interior de la industria:\n"
        "• Acometida Eléctrica: Derivación subterránea en BT desde la red mallada en 400/230 V a 50 Hz hasta la fachada de la industria.\n"
        "• Caja General de Protección (CGP - ITC-BT-13): Ubicada en la fachada exterior, aloja 3 bases tripolares con cortacircuitos fusibles tipo cuchilla BUC tamaño 2 de calibre 250 A clase gG y borne neutro seccionable.\n"
        "• Línea General de Alimentación (LGA - ITC-BT-14): Conductores unipolares de cobre bajo tubo protector enterrado, aislamiento termoestable RZ1-K 0,6/1 kV CPR Cca-s1b,d1,a1, con caída de tensión inferior al 0,5%.\n"
        "• Equipo de Medida (ITC-BT-16): Módulo de medida indirecta en baja tensión con transformadores de intensidad 250/5 A y contador fiscal multifunción con módem de telegestión.\n"
        "• Derivación Individual (DI - ITC-BT-15): Une la centralización de contadores con el Cuadro General (CGMP). Caída de tensión máxima admisible de 1,0%."
    )

    add_h2("1.8 Arquitectura del Sistema de Distribución y Cuadros Eléctricos")
    doc.add_paragraph(
        "La instalación se estructura de forma jerárquica mediante un Cuadro General y cinco subcuadros sectorizados:\n"
        "1. CGMP (Cuadro General de Mando y Protección): Alberga el IGA de 4 polos 250 A (Pdc = 25 kA), el protector contra sobretensiones transitorias y permanentes, la acometida a la batería de condensadores y las cabeceras de los subcuadros.\n"
        "2. SC-1 (Subcuadro Obrador y Fuerza): Alimenta los 4 hornos Unox de 10,3 kW, amasadoras Bongard 130 kg y 80 kg, batidora Saturne 60 L, divisoras, laminadora y silos.\n"
        "3. SC-2 (Subcuadro Climatización y Frío): Alimenta los 3 equipos Mitsubishi PLA-RP140BA y las cámaras frigoríficas Intarcon.\n"
        "4. SC-3 (Subcuadro Alumbrado y Tomas): Alimenta las regletas LED Philips WT120C IP65, downlights DN065B, tomas de corriente schuko/CETAC y puertas motorizadas.\n"
        "5. SC-4 (Subcuadro Recarga VE - ITC-BT-52): Alimenta los 4 cargadores de 7,4 kW del aparcamiento exterior con modulación SPL.\n"
        "6. SC-5 (Subcuadro Generación Fotovoltaica - ITC-BT-40): Cuadro de interconexión AC del inversor solar trifásico de 25 kW."
    )

    add_h2("1.9 Red de Puesta a Tierra y Protección Equipotencial (ITC-BT-18)")
    doc.add_paragraph(
        "Se proyecta un anillo perimetral enterrado en la cimentación de la nave con conductor de cobre desnudo de 35 mm², "
        "complementado por 6 picas verticales de acero cobrizado de 2 metros hincadas en arquetas de registro con puente de prueba "
        "seccionable. Toda la estructura metálica, envolventes de hornos y amasadoras, y bandejas portacables se conectan a la barra equipotencial principal (BEP)."
    )

    add_h2("1.10 Alumbrado Normal, de Emergencia y Señalización")
    doc.add_paragraph(
        "El alumbrado de producción se resuelve mediante luminarias estancas LED Philips CoreLine Waterproof WT120C IP65 de 41 W, "
        "y downlights Philips DN065B en oficinas. El alumbrado de emergencia (ITC-BT-28 / UNE-EN 60598-2-22) emplea bloques autónomos "
        "Legrand URA21 LED PLUS con autonomía de 1 hora, garantizando 5 lux en cuadros y 1 lux en vías de evacuación."
    )

    doc.add_page_break()

    # ==========================================
    # 6. DOCUMENTO Nº 2: CÁLCULOS JUSTIFICATIVOS
    # ==========================================
    add_h1("DOCUMENTO Nº 2: CÁLCULOS JUSTIFICATIVOS")
    
    add_callout(
        "EL NÚCLEO TÉCNICO DEL TRABAJO — CÓMO SACAR UN 10",
        "El profesor sanciona expresamente: 'Insertar, o copiar y pegar íntegramente los informes de los programas, se califica con un cero en esta parte'. "
        "Para obtener la máxima calificación, en primer lugar se deben declarar y explicar las ecuaciones teóricas y criterios de dimensionado; después presentar la tabla maestra de circuitos y justificar explícitamente: "
        "1) Protección contra contactos indirectos y valor de Rtierra, 2) Protección contra sobretensiones (permanentes y transitorias), y 3) Motores y factor 1,25 (ITC-BT-47).",
        box_type="warning"
    )
    
    add_h2("2.1 Criterios Normativos y Metodología de Cálculo")
    
    add_formula_card(
        "Cálculo de la Intensidad Nominal y de Servicio (Ib)",
        "Para circuitos trifásicos (400 V) y monofásicos (230 V) en corriente alterna senoidal:",
        "Trifásica:  Ib = P / (√3 · Un · cos φ)       |       Monofásica:  Ib = P / (Un · cos φ)",
        [
            ("P", "Potencia activa demandada por el circuito en vatios (W) o kilovatios (kW)."),
            ("Un", "Tensión nominal entre fases (400 V en trifásica) o fase-neutro (230 V en monofásica)."),
            ("cos φ", "Factor de potencia del receptor (según catálogo o ficha técnica; ej. 0,99 en hornos, 0,85 en motores)."),
            ("Ib", "Intensidad de diseño o servicio en amperios (A).")
        ]
    )
    
    add_formula_card(
        "Sobredimensionado Obligatorio para Motores (ITC-BT-47)",
        "Los conductores que alimentan a un motor individual deben estar dimensionados para una intensidad mínima del 125% de su corriente nominal a plena carga:",
        "Ib_motor = 1,25 · In_motor",
        [
            ("In_motor", "Intensidad nominal a plena carga indicada en la placa del motor o ficha técnica."),
            ("Ib_motor", "Intensidad de cálculo para la elección de la sección del conductor (A)."),
            ("Justificación", "Compensar las sobreintensidades prolongadas durante el arranque y variaciones mecánicas de carga.")
        ]
    )
    
    add_formula_card(
        "Criterio Térmico o de Capacidad de Corriente (Iz >= In >= Ib)",
        "La corriente máxima admisible por el conductor (Iz) debe ser igual o superior al calibre de la protección (In) y a la corriente de cálculo (Ib):",
        "Iz = Iz0 · ft · fa · fe  >=  In  >=  Ib",
        [
            ("Iz0", "Intensidad admisible base del cable según UNE-HD 60364-5-52 (aislamiento XLPE 90°C o PVC 70°C)."),
            ("ft", "Factor de corrección por temperatura ambiente distinta de 40°C al aire o 25°C enterrado."),
            ("fa", "Factor de corrección por agrupamiento de varios circuitos en la misma canalización."),
            ("fe", "Factor de corrección por enterramiento o resistividad térmica del terreno.")
        ]
    )
    
    add_formula_card(
        "Criterio de Caída de Tensión Máxima Admisible (ITC-BT-19)",
        "Cálculo de la sección requerida para no superar los límites porcentuales reglamentarios (3% alumbrado, 5% otros usos):",
        "Trifásica:  S = (√3 · L · P) / (γ · ΔU · Un)   |   Monofásica:  S = (2 · L · P) / (γ · ΔU · Un)",
        [
            ("L", "Longitud de la línea en metros (m)."),
            ("P", "Potencia transmitida en vatios (W)."),
            ("γ (conductividad)", "Para cobre a temperatura máxima de servicio: γ = 44 m/(Ω·mm²) para XLPE (90°C); γ = 48 m/(Ω·mm²) para PVC (70°C)."),
            ("ΔU", "Caída de tensión máxima en voltios (V). ΔU = e% · Un / 100."),
            ("Límites ITC-BT-19", "LGA <= 0,5% | DI <= 1,0% | Circuitos interiores: Alumbrado <= 3,0%, Fuerza/Otros usos <= 5,0%.")
        ]
    )
    
    add_formula_card(
        "Criterio de Cortocircuito y Esfuerzo Térmico en Conductores",
        "Verificación de que la energía específica pasante del cortocircuito no destruye el aislamiento del cable antes de que dispare la protección:",
        "Icc² · t <= k² · S²   ==>   S_min = (√Icc² · t) / k",
        [
            ("Icc", "Corriente de cortocircuito en el punto de la instalación (A)."),
            ("t", "Tiempo de disparo del interruptor magnetotérmico ante cortocircuito (s; típicamente < 0,1 s)."),
            ("k", "Constante del material: k = 143 para conductor de cobre con aislamiento termoestable XLPE/EPR; k = 115 para PVC."),
            ("S", "Sección transversal del conductor en mm².")
        ]
    )

    doc.add_page_break()

    # ==========================================
    # 2.2 TABLA MAESTRA DE CIRCUITOS Y LÍNEAS
    # ==========================================
    add_h2("2.2 Tabla Maestra de Cálculo de Circuitos y Líneas de la Panadería")
    
    circ_headers = ["Línea / Circuito", "Receptor / Destino", "Tensión", "P (kW)", "cos φ", "L (m)", "Ib (A)", "Sección (mm²)", "ΔU (%)", "Protección Asignada"]
    circ_widths = [Inches(0.6), Inches(1.4), Inches(0.5), Inches(0.5), Inches(0.4), Inches(0.4), Inches(0.5), Inches(0.7), Inches(0.5), Inches(1.0)]
    
    circ_data = [
        ("LGA", "Acometida CGP a Contadores", "400 V", "138,2", "0,93", "25", "214,4", "3x150+1x70", "0,35 %", "BUC 250A gG"),
        ("DI", "Contadores a CGMP", "400 V", "138,2", "0,93", "15", "214,4", "3x150+1x70", "0,21 %", "IGA 4P 250A"),
        ("L-HORN", "4x Hornos Unox 10,3 kW (41,2kW)", "400 V", "41,2", "0,99", "25", "60,1", "5G16", "0,84 %", "4P 63A Curva C"),
        ("L-AMAS1", "Amasadora Bongard 130 kg (6,8kW)", "400 V", "6,8", "0,85", "20", "14,4*", "5G4", "0,71 %", "4P 20A Curva D"),
        ("L-AMAS2", "Amasadora Bongard 80 kg (4,55kW)", "400 V", "4,55", "0,85", "20", "9,7*", "5G2.5", "0,76 %", "4P 16A Curva D"),
        ("L-BAT", "Batidora Saturne 60 L (3,1kW)", "400 V", "3,1", "0,85", "18", "6,6*", "5G2.5", "0,45 %", "4P 16A Curva D"),
        ("L-FERM", "2x Fermentadoras Sveba (8,0kW)", "400 V", "8,0", "1,00", "22", "11,5", "5G4", "0,61 %", "4P 20A Curva C"),
        ("L-FRIO1", "Cámara Congelación Intarcon (2x)", "400 V", "4,78", "0,85", "25", "12,1*", "5G4", "0,74 %", "4P 16A Curva D"),
        ("L-FRIO2", "Cámaras Conservación Intarcon", "400 V", "4,37", "0,85", "25", "18,6*", "5G6", "0,76 %", "4P 25A Curva D"),
        ("L-ENSAC", "Ensacadora Industrial", "400 V", "5,5", "0,92", "20", "10,8*", "5G2.5", "0,85 %", "4P 16A Curva D"),
        ("L-TRIT", "Trituradora de Pan", "400 V", "6,0", "0,95", "20", "11,4*", "5G4", "0,56 %", "4P 16A Curva D"),
        ("L-BOB", "Bobinadora Industrial", "400 V", "7,5", "0,97", "22", "13,9*", "5G4", "0,75 %", "4P 20A Curva D"),
        ("L-CLIM", "3x Cassettes Mitsubishi 5,2kW", "400 V", "15,6", "0,90", "35", "25,0", "5G6", "1,42 %", "4P 32A Curva C"),
        ("L-ALU1", "Alumbrado Obrador (Philips WT120C)", "400 V", "10,7", "0,90", "45", "17,1", "5G4", "1,88 %", "4P 20A Curva C"),
        ("L-ALU2", "Alumbrado Altillo (Philips DN065B)", "230 V", "0,92", "0,95", "27", "4,2", "3G1.5", "0,74 %", "2P 10A Curva C"),
        ("L-ALU3", "Alumbrado Exterior (Philips BRP102)", "230 V", "0,61", "0,98", "50", "2,7", "3G2.5", "0,53 %", "2P 10A Curva C"),
        ("L-TC1", "Tomas Corriente Obrador 16A", "230 V", "11,04", "1,00", "35", "16,0", "3G2.5", "2,18 %", "2P 16A Curva C"),
        ("L-PUER", "2x Puertas Hörmann WA300R", "230 V", "0,50", "0,85", "30", "1,6*", "3G1.5", "0,31 %", "2P 10A Curva C"),
        ("L-VE", "4x Cargadores VE Schneider 7,4kW", "400 V", "29,6", "1,00", "40", "42,7", "5G16", "1,04 %", "4P 50A Curva C / Dif B"),
        ("L-FV", "Inversor Solar Fotovoltaico 25kW", "400 V", "25,0", "1,00", "20", "36,1", "5G10", "0,70 %", "4P 40A Curva C")
    ]
    
    tbl_circ = doc.add_table(rows=len(circ_data)+1, cols=10)
    tbl_circ.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_circ, color="CBD5E0")
    
    for j, h in enumerate(circ_headers):
        c = tbl_circ.rows[0].cells[j]; c.width = circ_widths[j]
        set_cell_background(c, "1B365D"); set_cell_margins(c, top=60, bottom=60, left=50, right=50)
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h); r.bold = True; r.font.size = Pt(7.5); r.font.color.rgb = RGBColor(255, 255, 255)
        
    for i, data in enumerate(circ_data):
        row = tbl_circ.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]; c.width = circ_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=50, right=50)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val); r.font.size = Pt(7.5)
            if j == 0:
                r.bold = True
                
    p_notamot = doc.add_paragraph("(*) Intensidades de receptores con motor afectadas por el factor reglamentario 1,25 según la ITC-BT-47.")
    p_notamot.runs[0].font.size = Pt(8); p_notamot.runs[0].font.italic = True
    
    doc.add_page_break()

    # ==========================================
    # 2.3 JUSTIFICACIÓN DE CONTACTOS INDIRECTOS Y TIERRA
    # ==========================================
    add_h2("2.3 Justificación de la Protección contra Contactos Indirectos y Red de Tierra")
    
    add_formula_card(
        "Condición Reglamentaria de Desconexión en Esquema TT (ITC-BT-18 y 24)",
        "Para que la tensión de defecto generada en las masas metálicas no supere la tensión límite de contacto admisible UL:",
        "RA · IΔn <= UL   ===>   RA <= UL / IΔn",
        [
            ("RA", "Resistencia de puesta a tierra de las masas de la instalación en ohmios (Ω)."),
            ("IΔn", "Sensibilidad (corriente diferencial residual nominal) del interruptor diferencial en amperios (A)."),
            ("UL", "Tensión límite de contacto admisible: UL = 50 V en locales secos; UL = 24 V en locales mojados o húmedos (obrador y fermentación)."),
            ("Comprobación", "Con diferencial general de IΔn = 300 mA: RA <= 24 / 0,3 = 80,0 Ω. Con diferenciales de alta sensibilidad IΔn = 30 mA: RA <= 24 / 0,03 = 800,0 Ω.")
        ]
    )
    
    add_formula_card(
        "Resistencia de Difusión del Anillo y Picas Complementarias (ITC-BT-18)",
        "Fórmulas simplificadas para electrodos enterrados en terreno homogéneo (ρ = 100 Ω·m):",
        "R_anillo = (2 · ρ) / L       |       R_pica = ρ / l",
        [
            ("R_anillo", "Resistencia del conductor enterrado horizontalmente en anillo (L = 120 m): R_anillo = (2 · 100) / 120 = 1,67 Ω."),
            ("R_pica", "Resistencia de una pica vertical de l = 2 m: R_pica = 100 / 2 = 50,0 Ω."),
            ("R_total (anillo + 6 picas)", "Considerando la asociación en paralelo con factor de reducción por proximidad: R_tierra_calculada = 1,45 Ω."),
            ("Conclusión Justificativa", "El valor obtenido R_tierra = 1,45 Ω es extraordinariamente inferior al límite reglamentario más restrictivo de 80,0 Ω, garantizando que ninguna masa metálica alcance una tensión superior a 24 V bajo ningún defecto a tierra.")
        ]
    )

    add_h2("2.4 Justificación de la Protección contra Sobretensiones (ITC-BT-23)")
    doc.add_paragraph(
        "1. Nivel de Riesgo y Necesidad: Densidad de impactos de rayo Ng = 2 a 3 rayos/km²·año en comarca de l'Horta Nord. Protección obligatoria por alimentar procesos industriales continuos.\n"
        "2. Sobretensiones Transitorias: Protector Tipo 1+2 (Iimp = 12,5 kA / Imax = 50 kA, Up <= 1,5 kV) en CGMP, coordinado con Tipo 2 en subcuadros SC-1, SC-4 y SC-5.\n"
        "3. Sobretensiones Permanentes (POP): Dispositivo conforme a UNE-EN 50550 asociado al IGA con bobina de emisión de corriente que desconecta a > 275 V fase-neutro en < 0,2 s."
    )

    add_h2("2.5 Compensación de Reactiva y Filtrado de Armónicos en Obrador")
    add_formula_card(
        "Cálculo de la Potencia Reactiva de Compensación (Qc)",
        "Para elevar el factor de potencia medio desde cos φ1 = 0,85 hasta cos φ2 = 0,98:",
        "Qc = P_sim · (tan φ1 - tan φ2)",
        [
            ("P_sim", "Potencia simultánea activa: P_sim = 138,2 kW."),
            ("tan φ1 (cos φ = 0,85)", "tan φ1 = 0,620."),
            ("tan φ2 (cos φ = 0,98)", "tan φ2 = 0,203."),
            ("Qc calculada", "Qc = 138,2 · (0,620 - 0,203) = 138,2 · 0,417 = 57,63 kVAr."),
            ("Batería Adoptada", "Batería automática de 60 kVAr escalonada en (10+10+20+20 kVAr) con reactancias de desintonización a 189 Hz (p = 7%).")
        ]
    )

    add_h2("2.6 Infraestructura para Recarga de Vehículos Eléctricos (ITC-BT-52)")
    doc.add_paragraph(
        "Esquema 2 de la ITC-BT-52 con cuadro dedicado y Sistema de Protección de Línea (SPL) para modular la potencia de los 4 cargadores en función de la carga instantánea del obrador."
    )

    add_h2("2.7 Dimensionado del Sistema Solar Fotovoltaico en Cubierta (ITC-BT-40)")
    doc.add_paragraph(
        "Instalación fotovoltaica para autoconsumo de 25,2 kWp compuesta por 56 módulos monocristalinos de 450 Wp e inversor trifásico de 25 kW acoplado en BT."
    )

    doc.add_page_break()

    # ==========================================
    # 7. DOCUMENTO Nº 3: PLIEGO DE CONDICIONES Y MATERIALES
    # ==========================================
    add_h1("DOCUMENTO Nº 3: PLIEGO DE CONDICIONES Y SELECCIÓN DE MATERIALES")
    
    add_h2("3.1 Normativa Técnica y Reglamentación Aplicable")
    doc.add_paragraph(
        "Cumplimiento del REBT (RD 842/2002), UNE-HD 60364-5-52, UNE-EN 60898-1, UNE-EN 60947-2, UNE-EN 61008-1, UNE-EN 62423, UNE-EN 50575 (CPR), y UNE-EN 60079-0/31 (ATEX polvo)."
    )

    add_h2("3.2 Especificación Justificada de Aparamenta y Materiales de Referencia")
    
    mat_headers = ["Elemento / Función", "Fabricante y Gama", "Modelo de Referencia", "Norma de Cumplimiento", "Parámetros y Calibre"]
    mat_widths = [Inches(1.4), Inches(1.3), Inches(1.3), Inches(1.2), Inches(1.3)]
    mat_data = [
        ("Interruptor General Automático", "Schneider Electric", "ComPact NSX250F", "UNE-EN 60947-2", "4P, 250 A, Pdc = 36 kA, MicroLogic 2.2"),
        ("Magnetotérmicos Modulares", "Schneider Electric", "Acti9 iC60N", "UNE-EN 60898-1", "Curvas C y D, Pdc = 10 kA (10A a 63A)"),
        ("Diferenciales Súper Inmunizados", "Schneider Electric", "Acti9 iID Clase A SI", "UNE-EN 61008-1", "In = 40A / 63A, IΔn = 30mA y 300mA"),
        ("Diferenciales Recarga VE", "Schneider Electric", "Acti9 iID Clase B", "UNE-EN 62423", "4P, 40A, IΔn = 30mA (detección AC+DC)"),
        ("Protector Sobretensiones CGMP", "Schneider Electric", "Acti9 iPRD1 12.5r", "UNE-EN 61643-11", "Tipo 1+2, Iimp = 12,5 kA, Up <= 1,5 kV"),
        ("Protector Permanentes (POP)", "Schneider Electric", "iQuick PRD + Combi SPU", "UNE-EN 50550", "Disparo < 0,2s sobre IGA a 275V"),
        ("Cables de Distribución Libre Hal.", "Prysmian Group", "Afumex Class RZ1-K", "UNE-EN 50575", "0,6/1 kV, CPR Cca-s1b,d1,a1, XLPE 90°C"),
        ("Cables Resistentes al Fuego", "Prysmian Group", "Afumex FOC Sz1-K AS+", "UNE-EN 50200", "0,6/1 kV, PH120 (circuitos de emergencia)"),
        ("Luminarias LED Industriales", "Signify (Philips)", "CoreLine Waterproof WT120C", "UNE-EN 60598-1", "41 W, 4.000 lm, 4000K, IP65, IK08"),
        ("Luminarias LED Oficinas", "Signify (Philips)", "DN065B LED20S/830", "UNE-EN 60598-1", "23 W, 2.000 lm, 3000K, IP20, empotrar"),
        ("Luminarias de Emergencia", "Legrand", "URA21 LED PLUS 660611PL", "UNE-EN 60598-2-22", "Autonomía 1h, LED no permanente, IP42/IP65"),
        ("Canalizaciones Bajo Tubo", "AISCAN", "Tubo Rígido Blindado CR", "UNE-EN 61386-21", "Grado 5557, resistencia compresión > 4000 N")
    ]
    
    tbl_mat = doc.add_table(rows=len(mat_data)+1, cols=5)
    tbl_mat.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_mat, color="CBD5E0")
    for j, h in enumerate(mat_headers):
        c = tbl_mat.rows[0].cells[j]; c.width = mat_widths[j]
        set_cell_background(c, "1B365D"); set_cell_margins(c, top=60, bottom=60, left=60, right=60)
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8); r.font.color.rgb = RGBColor(255, 255, 255)
        
    for i, data in enumerate(mat_data):
        row = tbl_mat.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]; c.width = mat_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=60, right=60)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val); r.font.size = Pt(7.5)
            if j == 0:
                r.bold = True

    add_h2("3.3 Prescripciones de Montaje y Tendido de Canalizaciones")
    doc.add_paragraph("Instalación en bandejas de acero galvanizado en caliente a 3,5 m de altura o tubo rígido blindado IK10. Pasos de cables en zona ATEX con prensaestopas metálicos Ex d/e IP65.")

    add_h2("3.4 Pruebas Reglamentarias y Ensayos de Puesta en Servicio (ITC-BT-05)")
    doc.add_paragraph("Continuidad de conductores de protección, resistencia de aislamiento (R >= 1 MΩ), resistencia de tierra con telurómetro y comprobación de tiempo y corriente de disparo diferencial.")

    doc.add_page_break()

    # ==========================================
    # 8. DOCUMENTO Nº 4: PRESUPUESTO
    # ==========================================
    add_h1("DOCUMENTO Nº 4: PRESUPUESTO")
    
    add_h2("4.1 Resumen de Presupuesto por Capítulos")
    cap_headers = ["Capítulo", "Denominación del Capítulo", "Importe (€)"]
    cap_widths = [Inches(1.2), Inches(3.8), Inches(1.5)]
    cap_data = [
        ("Capítulo 1", "Acometida, CGP, Enlace y Cuadro General (CGMP)", "14.850,00 €"),
        ("Capítulo 2", "Subcuadros de Distribución Sectorizados (SC-1 a SC-5)", "11.420,00 €"),
        ("Capítulo 3", "Líneas de Distribución, Bandejas y Tubos Protectores", "18.630,00 €"),
        ("Capítulo 4", "Instalación de Alumbrado Normal y de Emergencia LED", "8.950,00 €"),
        ("Capítulo 5", "Red de Puesta a Tierra y Sistema Equipotencial", "3.480,00 €"),
        ("Capítulo 6", "Infraestructura de Recarga de Vehículos Eléctricos (ITC-BT-52)", "9.200,00 €"),
        ("Capítulo 7", "Sistema Solar Fotovoltaico en Cubierta 25 kWp (ITC-BT-40)", "19.500,00 €"),
        ("Capítulo 8", "Batería de Condensadores con Reactancias de Desintonización", "4.650,00 €"),
        ("Capítulo 9", "Ensayos, Medidas Reglamentarias, Certificación y Legalización", "2.800,00 €")
    ]
    
    tbl_cap = doc.add_table(rows=len(cap_data)+1, cols=3)
    tbl_cap.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_cap, color="CBD5E0")
    for j, h in enumerate(cap_headers):
        c = tbl_cap.rows[0].cells[j]; c.width = cap_widths[j]
        set_cell_background(c, "1B365D"); set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(255, 255, 255)
        if j == 2: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
    for i, data in enumerate(cap_data):
        row = tbl_cap.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]; c.width = cap_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=80, right=80)
            if i % 2 == 1: set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val); r.font.size = Pt(8.5)
            if j == 2: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; r.bold = True

    add_h2("4.2 Resumen General del Presupuesto")
    res_headers = ["Concepto Económico", "Porcentaje", "Importe Resultante (€)"]
    res_widths = [Inches(3.5), Inches(1.5), Inches(1.5)]
    res_data = [
        ("PRESUPUESTO DE EJECUCIÓN MATERIAL (PEM)", "Suma Capítulos", "93.480,00 €"),
        ("Gastos Generales de Estructura", "13,00 %", "12.152,40 €"),
        ("Beneficio Industrial del Contratista", "6,00 %", "5.608,80 €"),
        ("SUBTOTAL PRESUPUESTO DE EJECUCIÓN POR CONTRATA (PEC)", "PEM + GG + BI", "111.241,20 €"),
        ("Impuesto sobre el Valor Añadido (IVA)", "21,00 %", "23.360,65 €"),
        ("PRESUPUESTO TOTAL GENERAL DE LA INSTALACIÓN", "TOTAL PROYECTO", "134.601,85 €")
    ]
    
    tbl_res = doc.add_table(rows=len(res_data)+1, cols=3)
    tbl_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_res, color="CBD5E0")
    for j, h in enumerate(res_headers):
        c = tbl_res.rows[0].cells[j]; c.width = res_widths[j]
        set_cell_background(c, "1B365D"); set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(255, 255, 255)
        if j == 2: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
    for i, data in enumerate(res_data):
        row = tbl_res.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]; c.width = res_widths[j]
            set_cell_margins(c, top=60, bottom=60, left=80, right=80)
            if "TOTAL" in data[0] or "SUBTOTAL" in data[0] or "PEM" in data[0]:
                set_cell_background(c, "EBF4FF")
            elif i % 2 == 1: set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val); r.font.size = Pt(8.5)
            if "TOTAL" in data[0] or "SUBTOTAL" in data[0] or "PEM" in data[0]:
                r.bold = True
                if "TOTAL GENERAL" in data[0]:
                    r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(27, 54, 93)
            if j == 2: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p_letra = doc.add_paragraph("Asciende el presente Presupuesto General a la cantidad de CIENTO TREINTA Y CUATRO MIL SEISCIENTOS UN EUROS CON OCHENTA Y CINCO CÉNTIMOS (134.601,85 €), I.V.A. incluido.")
    p_letra.runs[0].font.size = Pt(8.5); p_letra.runs[0].font.italic = True
    
    doc.add_page_break()

    # ==========================================
    # 9. DOCUMENTO Nº 5: ÍNDICE DE PLANOS
    # ==========================================
    add_h1("DOCUMENTO Nº 5: ÍNDICE DE PLANOS Y REQUISITOS TÉCNICOS")
    
    add_callout(
        "PENALIZACIÓN CRÍTICA EN PLANOS (APARTADO 4 Y COMENTARIOS)",
        "El profesor especifica expresamente en las bases:\n"
        "• 'Los planos han de estar formateados a tamaño A3, en el cual han de ser legibles. Han de tener casillero y estar correctamente grafiados y acotados.'\n"
        "• 'Figuras y planos con fondo negro (o demasiado oscuro), califican el trabajo con un cero.'\n"
        "Aseguraos de exportar vuestros planos de AutoCAD SIEMPRE con fondo blanco limpio.",
        box_type="warning"
    )
    
    tbl_planos = doc.add_table(rows=5, cols=4)
    tbl_planos.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_planos, color="CBD5E0")
    pl_widths = [Inches(1.0), Inches(2.5), Inches(1.5), Inches(1.5)]
    pl_headers = ["Nº Plano", "Título del Plano", "Formato Normalizado", "Escala de Representación"]
    for j, h in enumerate(pl_headers):
        c = tbl_planos.rows[0].cells[j]; c.width = pl_widths[j]
        set_cell_background(c, "1B365D"); set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(255, 255, 255)
        
    pl_data = [
        ("Plano Nº 1", "Situación, Emplazamiento y Acometida Eléctrica", "A3 (Fondo Blanco)", "1:500 / 1:1000"),
        ("Plano Nº 2", "Distribución en Planta, Maquinaria y Canalizaciones (Nave 2)", "A3 (Fondo Blanco)", "1:100"),
        ("Plano Nº 3", "Esquema Unifilar General del CGMP y Subcuadros", "A3 (Fondo Blanco)", "Sin escala (Normalizado)"),
        ("Plano Nº 4", "Red de Puesta a Tierra y Red Equipotencial", "A3 (Fondo Blanco)", "1:100")
    ]
    for i, data in enumerate(pl_data):
        row = tbl_planos.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]; c.width = pl_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=80, right=80)
            if i % 2 == 1: set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val); r.font.size = Pt(8.5)
            if j == 0: r.bold = True

    doc.add_page_break()

    # ==========================================
    # 10. BIBLIOGRAFÍA Y REFERENCIAS NORMATIVAS (UPV)
    # ==========================================
    add_h1("DOCUMENTO Nº 6: BIBLIOGRAFÍA Y REFERENCIAS (FORMATO UPV / IEEE)")
    
    bib_items = [
        ("[1] MINISTERIO DE CIENCIA Y TECNOLOGÍA.", "Real Decreto 842/2002, de 2 de agosto, por el que se aprueba el Reglamento Electrotécnico para Baja Tensión (REBT) e Instrucciones Técnicas Complementarias (ITC) BT 01 a BT 52. Boletín Oficial del Estado (BOE), núm. 224, 18 de septiembre de 2002."),
        ("[2] MINISTERIO DE INDUSTRIA, ENERGÍA Y TURISMO.", "Guía Técnica de Aplicación del Reglamento Electrotécnico para Baja Tensión. Madrid: Dirección General de Industria y de la Pequeña y Mediana Empresa, 2019."),
        ("[3] ASOCIACIÓN ESPAÑOLA DE NORMALIZACIÓN (UNE).", "UNE-HD 60364-5-52:2014. Instalaciones eléctricas de baja tensión. Parte 5-52: Selección e instalación de equipos eléctricos. Canalizaciones. Madrid: AENOR, 2014."),
        ("[4] ASOCIACIÓN ESPAÑOLA DE NORMALIZACIÓN (UNE).", "UNE 157001:2014. Criterios generales para la elaboración formal de los documentos que constituyen un proyecto técnico. Madrid: AENOR, 2014."),
        ("[5] ASOCIACIÓN ESPAÑOLA DE NORMALIZACIÓN (UNE).", "UNE-EN 50575:2015/A1:2016. Cables de energía, control y comunicación. Requisitos de reacción al fuego (Reglamento de Productos de Construcción - CPR). Madrid: AENOR, 2016."),
        ("[6] SCHNEIDER ELECTRIC.", "Guía de diseño de instalaciones eléctricas según normas internacionales CEI. Rueil-Malmaison (Francia): Schneider Electric Industries SAS, 2024."),
        ("[7] PRYSMIAN GROUP.", "Catálogo General de Cables para Baja Tensión y Guía CPR. Vilanova i la Geltrú: Prysmian Cables Spain S.A., 2023."),
        ("[8] SIGNIFY IBERIA (PHILIPS LIGHTING).", "Manual Técnico de Iluminación Industrial y Catálogo de Luminarias LED CoreLine. Madrid: Signify, 2024."),
        ("[9] MITSUBISHI ELECTRIC IBERIA.", "Ficha Técnica y Manual de Instalación Climatización Cassettes Serie Mr. Slim (PLA-RP140BA). Sant Cugat del Vallès: Mitsubishi Electric, 2023."),
        ("[10] GARCÍA TRASANCOS, J.", "Instalaciones Eléctricas de Enlace y Centros de Transformación. 6ª ed. Madrid: Ediciones Paraninfo, 2020.")
    ]
    for num, cit in bib_items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.first_line_indent = Inches(-0.4)
        p.paragraph_format.space_after = Pt(4)
        r0 = p.add_run(num + " "); r0.bold = True; r0.font.size = Pt(8.5); r0.font.color.rgb = RGBColor(27, 54, 93)
        r1 = p.add_run(cit); r1.font.size = Pt(8.5)

    doc.add_page_break()

    # ==========================================
    # 11. ANEXO DE FICHAS TÉCNICAS DE MATERIALES (EXIGIDO)
    # ==========================================
    add_h1("DOCUMENTO Nº 7: ANEXO DE FICHAS TÉCNICAS DE MATERIALES")
    
    add_callout(
        "REQUISITO EXPRESO DEL ENUNCIADO",
        "El profesor indica en el apartado de comentarios: 'El trabajo se presentará en un único documento en pdf. Incorporará portada, índice, memoria descriptiva, planos, fuentes/bibliografía, listado de figuras, y ANEXO DE FICHAS DE MATERIALES'. "
        "A continuación se sintetizan los equipos reales seleccionados de las fichas técnicas consultadas en Receptores_panaderia_v2.xlsx.",
        box_type="tip"
    )
    
    fichas_headers = ["Equipo / Receptor", "Fabricante y Modelo Comercial", "Potencia y Tensión", "Características Destacadas y Fuente"]
    fichas_widths = [Inches(1.5), Inches(1.8), Inches(1.2), Inches(2.0)]
    fichas_data = [
        ("Horno de Convección (4 uds)", "Unox BakerLux SHOP.Pro Master XEFT-06EU-EMRV", "10,3 kW c/u (400 V 3F+N)", "cos φ = 0,99, In = 15,0 A, Imáx = 17,5 A. Panel digital. (Ficha Unox)."),
        ("Amasadora de Espiral 130 kg", "Bongard Spiral EVO 200", "6,8 kW (400 V 3F+T)", "Motor espiral + artesa. cos φ = 0,85. Factor 1,25 -> 14,4 A. (Ficha Bongard)."),
        ("Amasadora de Espiral 80 kg", "Bongard Spiral EVO 110", "4,55 kW (400 V 3F+T)", "cos φ = 0,85. Factor 1,25 -> 9,7 A. (Ficha Bongard)."),
        ("Batidora Planetaria 60 L", "Bongard Saturne EVO 60", "3,1 kW (400 V 3F+T)", "Variador de velocidad electrónico. cos φ = 0,85. (Ficha Bongard)."),
        ("Cámara de Fermentación (2 uds)", "Sveba Dahlen F500 HIK", "4,0 kW c/u (400 V 3F+N)", "Carga resistiva con control estricto de humedad. cos φ = 1,0. (Ficha Sveba)."),
        ("Cámara Congelación (-20°C)", "Intarcon BCV-NG-3 096 (2 uds)", "2,39 kW c/u (400 V 3F+N)", "Compresor hermético. Imáx = 12,1 A de ficha. (Ficha Intarcon)."),
        ("Cámara Conservación (0°C)", "Intarcon MCV-NY-3 108", "2,48 kW (400 V 3F+N)", "Compresor de media temperatura. Imáx = 18,6 A de ficha. (Ficha Intarcon)."),
        ("Climatización Cassette (3 uds)", "Mitsubishi Electric PLA-RP140BA", "5,21 kW c/u (400 V 3F+N)", "Tecnología Inverter. Imáx = 14,0 A de catálogo. (Ficha Mitsubishi)."),
        ("Luminarias LED Industriales", "Philips CoreLine WT120C LED40S/840", "41 W (230 V 1F+N)", "4.000 lm, IP65, IK08. Estancas para obrador. (Ficha Signify)."),
        ("Luminarias LED Downlight", "Philips DN065B LED20S/830", "23 W (230 V 1F+N)", "2.000 lm, IP20 para falso techo de altillo. (Ficha Signify)."),
        ("Luminarias de Emergencia", "Legrand URA21 LED PLUS 660611PL", "2,3 W (230 V 1F+N)", "100 lm, autonomía 3h, autotest. (Ficha Legrand)."),
        ("Puntos de Recarga VE (4 uds)", "Schneider EVlink Pro AC EVB3S07N4A", "7,4 kW c/u (230 V 32 A)", "Toma T2S, detección fugas 6 mA DC integrada. (Ficha Schneider)."),
        ("Puertas Motorizadas (2 uds)", "Hörmann WA300R S4/SE", "0,25 kW c/u (230 V 1F+N)", "Motor para persiana enrollable con cuadro IP65. (Ficha Hörmann).")
    ]
    
    tbl_fich = doc.add_table(rows=len(fichas_data)+1, cols=4)
    tbl_fich.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_fich, color="CBD5E0")
    for j, h in enumerate(fichas_headers):
        c = tbl_fich.rows[0].cells[j]; c.width = fichas_widths[j]
        set_cell_background(c, "1B365D"); set_cell_margins(c, top=60, bottom=60, left=70, right=70)
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.0); r.font.color.rgb = RGBColor(255, 255, 255)
        
    for i, data in enumerate(fichas_data):
        row = tbl_fich.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]; c.width = fichas_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=70, right=70)
            if i % 2 == 1: set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val); r.font.size = Pt(7.5)
            if j == 0: r.bold = True

    doc.add_page_break()

    # ==========================================
    # 12. ANEXO: CHECKLIST PARA OBTENER UN 10
    # ==========================================
    add_h1("ANEXO: CHECKLIST DE AUTOEVALUACIÓN PARA OBTENER UN 10")
    
    chk_headers = ["Aspecto Evaluado", "Requisito Exigido", "Criterio de Excelencia para el 10", "Estado"]
    chk_widths = [Inches(1.5), Inches(1.8), Inches(2.4), Inches(0.8)]
    chk_data = [
        ("Potencia de la Nave", "Entre 120 y 190 kW", "159,54 kW justificados con receptores comerciales reales en Receptores_panaderia_v2.xlsx", "[ X ] OK"),
        ("Superficie y Altillo", "Altillo <= 30% y >= 150 m²", "Nave 2 en 'L' (602 m² PB y 156 m² altillo = 25,91%), cotas múltiplos de 0,5 m", "[ X ] OK"),
        ("Aparcamiento y VE", "Entre 20 y 40 plazas", "28 plazas con 4 puntos Schneider EVlink Modo 3 y modulación SPL (ITC-BT-52)", "[ X ] OK"),
        ("Cálculos y Fórmulas", "No pegar pantallazos de software", "Ecuaciones teóricas analíticas declaradas antes de cada tabla maestra", "[ X ] OK"),
        ("Locales Especiales", "REBT general", "Clasificación rigurosa de Zona 22 ATEX por polvo de harina (ITC-BT-29) y locales mojados (ITC-BT-30)", "[ X ] OK"),
        ("Contactos Indirectos", "Justificación de Rtierra", "Cálculo analítico del anillo y picas demostrando RA = 1,45 Ω <= 24V / 0,3A = 80 Ω", "[ X ] OK"),
        ("Sobretensiones", "Justificación transitoria y permanente", "Descargadores Tipo 1+2 coordinados con Tipo 2 y protección POP UNE-EN 50550", "[ X ] OK"),
        ("Materiales y Marcas", "Indicar modelos y calibres", "Especificación completa con fabricantes de prestigio (Schneider, Prysmian, Philips)", "[ X ] OK"),
        ("Planos A3 Normalizados", "Fondo claro obligatorio", "Fondo blanco pulcro, cajetín oficial UPV acotado y grafiado profesional (fondo negro = cero)", "[ Pendiente ]"),
        ("Citas y Bibliografía", "Normas Biblioteca UPV", "10 citas en formato formal IEEE/ISO con normas UNE, REBT y catálogos técnicos", "[ X ] OK"),
        ("Firma Digital", "Documento único en PDF firmado", "Exportar a PDF y aplicar certificado digital oficial (DNIe / FNMT / ACCV)", "[ Pendiente ]")
    ]
    
    tbl_chk = doc.add_table(rows=len(chk_data)+1, cols=4)
    tbl_chk.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_chk, color="CBD5E0")
    for j, h in enumerate(chk_headers):
        c = tbl_chk.rows[0].cells[j]; c.width = chk_widths[j]
        set_cell_background(c, "1B365D"); set_cell_margins(c, top=60, bottom=60, left=70, right=70)
        p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.0); r.font.color.rgb = RGBColor(255, 255, 255)
        if j == 3: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
    for i, data in enumerate(chk_data):
        row = tbl_chk.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]; c.width = chk_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=70, right=70)
            if i % 2 == 1: set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val); r.font.size = Pt(7.5)
            if j == 0: r.bold = True
            if j == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER; r.bold = True; r.font.color.rgb = RGBColor(13, 92, 117)

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output_filename = os.path.join(base_dir, "proyecto", "Plantilla_Proyecto_Instalacion_Electrica_Panaderia.docx")
    doc.save(output_filename)
    print(f"Documento generado exitosamente como '{output_filename}'")

if __name__ == "__main__":
    create_document()
