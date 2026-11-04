# -*- coding: utf-8 -*-
"""
Generador de la Plantilla Maestra en Word (.docx) para el Proyecto de Instalación Eléctrica
de Baja Tensión para una Panadería Industrial (UPV - Reto 1 Instalaciones).
"""

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
        frun1 = fp.add_run("ETSID - Universidad Politécnica de Valencia | Grado en Ingeniería")
        frun1.font.name = "Calibri"
        frun1.font.size = Pt(8.5)
        frun1.font.color.rgb = RGBColor(100, 116, 139)

    # Paleta de colores
    NAVY = "1B365D"       # Títulos principales, cabeceras de tablas
    TEAL = "0D5C75"       # Subtítulos, H2
    GRAY_TEXT = "1E293B"  # Texto base
    
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
            r_d.font.name = "Calibri"
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
    r_asig = p_upv.add_run("GRADO EN INGENIERÍA MECÁNICAA — ASIGNATURA: INSTALACIONES ELÉCTRICAS (RETO 1)")
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
    r_t1.font.size = Pt(16)
    r_t1.font.color.rgb = RGBColor(255, 255, 255)
    
    r_t2 = p_t.add_run("CON INFRAESTRUCTURA DE RECARGA DE VEHÍCULOS ELÉCTRICOS (ITC-BT-52) Y SISTEMA DE GENERACIÓN SOLAR FOTOVOLTAICA (ITC-BT-40)")
    r_t2.font.size = Pt(11.5)
    r_t2.font.color.rgb = RGBColor(226, 232, 240)
    
    doc.add_paragraph().paragraph_format.space_before = Pt(30)
    
    # Cuadro de metadatos de autoría y firma digital
    tbl_meta = doc.add_table(rows=5, cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_meta.autofit = False
    set_table_borders(tbl_meta, color="CBD5E0")
    
    col_widths = [Inches(2.2), Inches(4.3)]
    meta_rows = [
        ("AUTORES DEL PROYECTO:", "[Nombre y Apellidos Alumno 1] (DNI: ____________)\n[Nombre y Apellidos Alumno 2] (DNI: ____________)"),
        ("TUTOR / DEPARTAMENTO:", "Departamento de Ingeniería Eléctrica — UPV"),
        ("CONVOCATORIA Y CURSO:", "Curso Académico 2026 / 2027"),
        ("ESTRUCTURA DOCUMENTAL:", "Doc. 1: Memoria Descriptiva | Doc. 2: Cálculos Justificativos\nDoc. 3: Pliego de Condiciones | Doc. 4: Presupuesto | Doc. 5: Planos A3"),
        ("FIRMA DIGITAL REGLAMENTARIA:", "Espacio reservado para el estampillado y firma digital del PDF\n(Requisito obligatorio según enunciado del Reto)")
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
        r0.font.size = Pt(9)
        r0.font.color.rgb = RGBColor(71, 85, 105)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(val)
        r1.font.size = Pt(9)
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
        ("Rev. 0.1", "Noviembre 2026", "Equipo de Trabajo", "Creación de la plantilla maestra con estructura reglamentaria y fórmulas."),
        ("Rev. 0.2", "[Fecha]", "[Nombre Alumno]", "Dimensionamiento de cargas, inventario de receptores y cálculo térmico."),
        ("Rev. 1.0", "[Fecha]", "Ambos Autores", "Revisión final, comprobación de protecciones, planos A3 y firma digital.")
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
    # 3. ÍNDICE GENERAL DEL PROYECTO
    # ==========================================
    add_h2("ÍNDICE GENERAL DEL PROYECTO")
    
    add_callout(
        "CÓMO ACTUALIZAR EL ÍNDICE EN MICROSOFT WORD",
        "Esta plantilla utiliza títulos automáticos (Título 1, Título 2, Título 3). Cuando hayas terminado de redactar o modificar apartados, haz clic derecho sobre el índice de Word y pulsa 'Actualizar campos' (o selecciona el índice y presiona la tecla F9) para actualizar automáticamente los números de página.",
        box_type="tip"
    )
    
    toc_items = [
        ("RESUMEN EJECUTIVO (ABSTRACT)", "Pág. 4"),
        ("DOCUMENTO Nº 1: MEMORIA DESCRIPTIVA", "Pág. 5"),
        ("  1.1 Objeto y Alcance del Proyecto", "Pág. 5"),
        ("  1.2 Datos del Titular y Emplazamiento", "Pág. 6"),
        ("  1.3 Descripción del Proceso Industrial de Panificación", "Pág. 7"),
        ("  1.4 Características Constructivas del Edificio y Superficies", "Pág. 8"),
        ("  1.5 Clasificación Reglamentaria de los Locales (ATEX, Húmedos, Hornos)", "Pág. 9"),
        ("  1.6 Suministro Eléctrico e Instalación de Enlace (CGP, LGA, Contadores, DI)", "Pág. 11"),
        ("  1.7 Arquitectura del Sistema de Distribución y Cuadros Eléctricos", "Pág. 13"),
        ("  1.8 Red de Puesta a Tierra y Protección Equipotencial", "Pág. 15"),
        ("  1.9 Alumbrado Normal, de Emergencia y Señalización", "Pág. 16"),
        ("DOCUMENTO Nº 2: CÁLCULOS JUSTIFICATIVOS", "Pág. 18"),
        ("  2.1 Criterios Normativos y Metodología de Cálculo", "Pág. 18"),
        ("  2.2 Tabla Maestra de Cálculo de Circuitos y Líneas", "Pág. 21"),
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
        ("BIBLIOGRAFÍA Y REFERENCIAS NORMATIVAS (FORMATO BIBLIOTECA UPV)", "Pág. 42"),
        ("ANEXO: CHECKLIST DE AUTOEVALUACIÓN PARA OBTENER CALIFICACIÓN 10", "Pág. 44")
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
        if item.startswith("DOCUMENTO") or item.startswith("BIBLIOGRAFÍA") or item.startswith("ANEXO"):
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
    # 4. RESUMEN EJECUTIVO (ABSTRACT)
    # ==========================================
    add_h1("RESUMEN EJECUTIVO / ABSTRACT")
    
    add_callout(
        "REQUISITOS DEL RESUMEN TÉCNICO",
        "El resumen debe sintetizar en 200-300 palabras las magnitudes clave de la instalación: potencia nominal instalada (debe estar entre 120 y 190 kW según el enunciado), potencia simultánea de cálculo, tensión de servicio (400 V trifásica / 230 V monofásica), esquema de distribución adoptado (TT), singularidades ATEX por polvo de harina, número de plazas de aparcamiento con recarga de VE y potencia de la instalación solar fotovoltaica para autoconsumo.",
        box_type="fill"
    )
    
    doc.add_paragraph(
        "El presente proyecto técnico tiene por objeto el diseño, dimensionamiento, justificación y cálculo "
        "de la instalación eléctrica de baja tensión para una nave industrial dedicada a la panificación y bollería "
        "industrial, situada en el término municipal de [Municipio, Valencia]. La instalación parte desde la red de distribución "
        "en baja tensión de la compañía suministradora, acometiendo a la Caja General de Protección (CGP) e incluyendo la "
        "Línea General de Alimentación (LGA), equipo de medida fiscal en baja tensión, Derivación Individual (DI) y Cuadro "
        "General de Mando y Protección (CGMP) junto a cinco subcuadros de distribución sectorizada."
    )
    
    doc.add_paragraph(
        "La potencia total instalada suma un total de [P_inst, ej. 168,50 kW], resultando una potencia máxima simultánea "
        "prevista de [P_sim, ej. 138,20 kW] tras aplicar los factores de simultaneidad y utilización característicos de la "
        "maquinaria de obrador (hornos de convección, amasadoras, batidoras, silos, divisoras y climatización). La instalación "
        "adopta un esquema de conexión a tierra tipo TT, garantizando la protección contra contactos directos e indirectos "
        "mediante interruptores diferenciales y una resistencia de puesta a tierra calculada inferior a [R_tierra, ej. 12,5 Ω], "
        "muy por debajo del límite reglamentario de 24 V fijado para locales húmedos."
    )
    
    doc.add_paragraph(
        "Como elementos de excelencia técnica y sostenibilidad, se incluye la clasificación y prescripciones de seguridad "
        "para atmósferas potencialmente explosivas por polvo combustible de harina (Zona 22 según ITC-BT-29), una infraestructura "
        "de recarga inteligente para vehículos eléctricos con sistema de modulación dinámica de potencia (ITC-BT-52) para dar "
        "servicio a [Nº plazas, ej. 28 plazas] de aparcamiento, una instalación solar fotovoltaica en cubierta para autoconsumo "
        "de [P_fv, ej. 25 kWp] (ITC-BT-40), y una batería automática de condensadores con reactancias de desintonización para "
        "la eliminación del recargo por energía reactiva y mitigación de armónicos."
    )
    
    add_h2("Technical Abstract (English)")
    p_en = doc.add_paragraph(
        "This engineering project covers the comprehensive design and electrical sizing of a low-voltage electrical "
        "installation for an industrial bakery and pastry production facility. The electrical infrastructure operates at "
        "400 V / 230 V (50 Hz) under a TT grounding scheme, with a total installed power of [P_inst kW] and a diversified "
        "peak demand of [P_sim kW]. Key engineering highlights include compliance with the Spanish Low Voltage Electrotechnical "
        "Regulations (REBT), specialized ATEX zoning for combustible flour dust areas (Zone 22, ITC-BT-29), electric vehicle "
        "charging infrastructure with dynamic load balancing (ITC-BT-52), rooftop photovoltaic generation for self-consumption "
        "(ITC-BT-40), and an automatic detuned capacitor bank for power factor correction and harmonic distortion mitigation."
    )
    p_en.runs[0].font.italic = True
    
    doc.add_page_break()

    # ==========================================
    # 5. DOCUMENTO Nº 1: MEMORIA DESCRIPTIVA
    # ==========================================
    add_h1("DOCUMENTO Nº 1: MEMORIA DESCRIPTIVA")
    
    add_callout(
        "CÓMO ABORDAR LA MEMORIA PARA OBTENER UN 10",
        "En el proyecto del año anterior, la memoria era excesivamente breve y se limitaba a enunciar apartados sin desarrollo de ingeniería. Para alcanzar la máxima nota, debes describir la instalación de forma ordenada comenzando desde la CGP hacia los receptores, detallando el proceso de la panadería, justificando la elección de materiales y explicando la necesidad de cada subcuadro.",
        box_type="tip"
    )
    
    add_h2("1.1 Objeto y Alcance del Proyecto")
    doc.add_paragraph(
        "El presente proyecto técnico tiene como finalidad establecer las bases técnicas, criterios de diseño, "
        "cálculos justificativos, prescripciones de materiales y valoración económica para la ejecución de la "
        "instalación eléctrica de baja tensión de una industria de panificación y bollería. "
        "El proyecto abarca desde el punto de conexión con la red de distribución de la empresa suministradora "
        "hasta los receptores finales de fuerza motriz, alumbrado, climatización, recarga de vehículos eléctricos "
        "y generación solar fotovoltaica, de conformidad con lo establecido en el Real Decreto 842/2002 (REBT) "
        "y sus Instrucciones Técnicas Complementarias (ITC-BT)."
    )
    
    add_h2("1.2 Datos del Titular y Emplazamiento")
    tbl_tit = doc.add_table(rows=4, cols=2)
    tbl_tit.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_tit, color="CBD5E0")
    tit_data = [
        ("Titular / Razón Social:", "Panificadora Industrial Levante S.L. — NIF: B-96XXXXXX"),
        ("Actividad / CNAE:", "Fabricación de pan y productos de panadería y pastelería (CNAE 1071)"),
        ("Emplazamiento:", "Parcela 14, Polígono Industrial Fuente del Jarro, Paterna (Valencia)"),
        ("Referencia Catastral / Coordenadas:", "Ref. Catastral: 9876543YJ2897S0001WX | Coordenadas: 39°31'12.4\"N 0°27'45.8\"W")
    ]
    for i, (k, v) in enumerate(tit_data):
        r = tbl_tit.rows[i]
        c0, c1 = r.cells[0], r.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_margins(c0, top=70, bottom=70, left=100, right=100)
        set_cell_margins(c1, top=70, bottom=70, left=100, right=100)
        set_cell_background(c0, "F8FAFC")
        p0, p1 = c0.paragraphs[0], c1.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.bold = True
        r0.font.size = Pt(9)
        r1 = p1.add_run(v)
        r1.font.size = Pt(9)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_h2("1.3 Descripción del Proceso Industrial de Panificación")
    doc.add_paragraph(
        "El conocimiento del ciclo productivo es esencial para establecer los factores de simultaneidad y "
        "la zonificación de riesgos eléctricos de la planta. El proceso industrial comprende las siguientes etapas secuenciales:"
    )
    
    proceso_steps = [
        ("1. Recepción y Almacenamiento de Harina:", "La harina se almacena en silos exteriores con soplante neumática y tornillo sinfín de extracción hacia la sala de amasado. Zona con presencia de polvo en suspensión."),
        ("2. Dosificación y Amasado:", "Pesaje automatizado de harina, adición de agua enfriada y mezcla en amasadoras de espiral industriales (130 kg y 80 kg) con motores de alto par y variadores de velocidad."),
        ("3. División y Boleado:", "Paso de la masa por divisoras-formadoras volumétricas para obtener porciones homogéneas de pan y masas de bollería."),
        ("4. Fermentación Controlada:", "Cámaras con control estricto de temperatura (28-32 °C) y humedad relativa (75-85%). Ambiente catalogado como local húmedo/mojado según ITC-BT-30."),
        ("5. Cocción en Hornos de Convección:", "Batería de cuatro hornos de convección eléctricos de 10 kW cada uno (40 kW totales) funcionando en régimen continuo a 200-240 °C."),
        ("6. Enfriamiento, Envasado y Expedición:", "Cintas transportadoras, líneas de empaquetado, etiquetado y almacenamiento en muelles de carga."),
        ("7. Climatización y Servicios Generales:", "Climatización por cassettes trifásicos Mitsubishi, alumbrado LED Philips, y recarga de vehículos de reparto y personal en el aparcamiento.")
    ]
    for step_title, step_desc in proceso_steps:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        r_st = p.add_run(step_title + " ")
        r_st.bold = True
        r_st.font.size = Pt(9.5)
        r_sd = p.add_run(step_desc)
        r_sd.font.size = Pt(9.5)

    add_h2("1.4 Características Constructivas del Edificio y Distribución de Superficies")
    
    add_callout(
        "RESTRICCIONES GEOMÉTRICAS DEL ENUNCIADO",
        "El enunciado exige: Edificio en planta baja + altillo parcial que NO supere el 30% de la superficie de planta, y al menos de 150 m². Las longitudes principales deben ser múltiplos de 0,5 m. Alrededor del local habrá zona exterior iluminada y aparcamiento para entre 20 y 40 plazas con puntos de recarga de vehículo eléctrico.",
        box_type="warning"
    )
    
    tbl_sup = doc.add_table(rows=6, cols=3)
    tbl_sup.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_sup, color="CBD5E0")
    sup_widths = [Inches(2.5), Inches(1.5), Inches(2.5)]
    sup_headers = ["Zona / Dependencia", "Superficie Útil", "Clasificación y Uso"]
    for j, h in enumerate(sup_headers):
        c = tbl_sup.rows[0].cells[j]
        c.width = sup_widths[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, top=70, bottom=70, left=90, right=90)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    sup_data = [
        ("Nave Principal (Planta Baja)", "620,0 m²", "Obrador, hornos, silos, cámaras, muelles"),
        ("Altillo / Planta Elevada", "165,0 m² (26,6 %)", "Oficinas técnicas, vestuarios, control (>= 150 m² y <= 30%)"),
        ("Superficie Total Construida", "785,0 m²", "Edificio industrial clase de riesgo medio"),
        ("Aparcamiento y Viales Exteriores", "840,0 m²", "28 plazas (con 4 puntos de recarga VE Modo 3)"),
        ("Cubierta Disponible para Solar FV", "450,0 m²", "Cubierta sandwich inclinada orientada al Sur")
    ]
    for i, data in enumerate(sup_data):
        row = tbl_sup.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]
            c.width = sup_widths[j]
            set_cell_margins(c, top=60, bottom=60, left=90, right=90)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if j == 0 and "Total" in val:
                r.bold = True
                
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_h2("1.5 Clasificación Reglamentaria de los Locales (Punto Crítico para Nota 10)")
    
    add_callout(
        "JUSTIFICACIÓN ATEX Y LOCALES ESPECIALES",
        "El proyecto del año anterior apenas mencionaba el polvo de harina ni la ITC-BT-29. La harina de trigo es un polvo combustible (Clase St 1, Kst ~ 100-130 bar·m/s). Clasificar correctamente la zona de descarga y silos como ZONA 22 e indicar el tipo de protección de los motores (Ex tc IIIC T125°C IP65) es un factor decisivo para alcanzar el 10.",
        box_type="tip"
    )
    
    doc.add_paragraph(
        "De conformidad con las prescripciones del REBT, se establece la siguiente clasificación técnica de dependencias:"
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
        p.paragraph_format.space_after = Pt(4)
        r_ct = p.add_run(c_title + "\n")
        r_ct.bold = True
        r_ct.font.size = Pt(9.5)
        r_ct.font.color.rgb = RGBColor(13, 92, 117)
        r_cd = p.add_run(c_desc)
        r_cd.font.size = Pt(9.0)

    add_h2("1.6 Suministro Eléctrico e Instalación de Enlace (ITC-BT-12 a 16)")
    doc.add_paragraph(
        "La instalación de enlace conecta la red de distribución pública en baja tensión de la empresa "
        "distribuidora (Iberdrola Distribución Eléctrica) con la instalación interior de la industria:"
    )
    
    enlace_steps = [
        ("• Acometida Eléctrica:", "Derivación subterránea en BT desde la red de distribución mallada en 400/230 V a 50 Hz hasta la fachada de la industria."),
        ("• Caja General de Protección (CGP - ITC-BT-13):", "Ubicada en la fachada exterior del edificio, en nicho con puerta metálica y cerradura homologada. Aloja 3 bases tripolares con cortacircuitos fusibles tipo cuchilla BUC tamaño 2 de calibre 250 A clase gG y borne neutro seccionable."),
        ("• Línea General de Alimentación (LGA - ITC-BT-14):", "Enlaza la CGP con la centralización de contadores. Formada por conductores unipolares de cobre bajo tubo protector enterrado, aislamiento termoestable RZ1-K 0,6/1 kV CPR Cca-s1b,d1,a1, con caída de tensión inferior al 0,5%."),
        ("• Equipo de Medida y Centralización de Contadores (ITC-BT-16):", "Módulo de medida indirecta en baja tensión con transformadores de intensidad de relación 250/5 A y contador fiscal multifunción con módem de telegestión."),
        ("• Derivación Individual (DI - ITC-BT-15):", "Canalización que une el equipo de medida con el Cuadro General de Mando y Protección (CGMP). Caída de tensión máxima admisible de 1,0%.")
    ]
    for e_title, e_desc in enlace_steps:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        r0 = p.add_run(e_title + " ")
        r0.bold = True
        r0.font.size = Pt(9)
        r1 = p.add_run(e_desc)
        r1.font.size = Pt(9)

    add_h2("1.7 Arquitectura del Sistema de Distribución y Cuadros Eléctricos")
    doc.add_paragraph(
        "Para garantizar la continuidad de servicio, la selectividad de las protecciones y facilitar el "
        "mantenimiento industrial, la instalación se estructura de forma jerárquica mediante un Cuadro General "
        "y cinco subcuadros sectorizados:"
    )
    
    cuadros = [
        ("1. CGMP — Cuadro General de Mando y Protección:", "Ubicado en el local técnico junto a la entrada. Alberga el Interruptor General Automático (IGA) de 4 polos 250 A (Pdc = 25 kA), el protector contra sobretensiones transitorias y permanentes, la acometida a la batería de condensadores y los interruptores de cabecera de las 5 líneas de distribución hacia los subcuadros."),
        ("2. SC-1 — Subcuadro Obrador y Fuerza Industrial:", "Alimenta la maquinaria de producción: 4 hornos de convección de 10 kW, amasadoras Bongard 130 kg y 80 kg, batidoras planetarias, divisoras-formadoras, laminadora Rondo, silo de harina y enfriador de agua."),
        ("3. SC-2 — Subcuadro Climatización, Frío y Ventilación:", "Alimenta los 3 equipos de climatización Mitsubishi SPLZS-140YBA(H), evaporadores y compresores de cámaras frigoríficas y extractores de ventilación forzada ATEX."),
        ("4. SC-3 — Subcuadro Alumbrado, Servicios Auxiliares y Tomas:", "Alimenta las líneas trifásicas de luminarias LED Philips BN124C, luminarias downlight DN065B, tomas de corriente schuko y CETAC de 16 A, puerta motorizada de acceso y alumbrado de emergencia."),
        ("5. SC-4 — Subcuadro Infraestructura de Recarga de VE (ITC-BT-52):", "Alimenta los 4 cargadores de 7,4 kW / 22 kW del aparcamiento exterior con sistema de gestión de carga dinámico SPL."),
        ("6. SC-5 — Subcuadro de Generación Solar Fotovoltaica (ITC-BT-40):", "Cuadro de interconexión AC del inversor solar trifásico de 25 kW con protecciones de desconexión automática y rearme según normativa.")
    ]
    for c_tit, c_txt in cuadros:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        r0 = p.add_run(c_tit + " ")
        r0.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = RGBColor(27, 54, 93)
        r1 = p.add_run(c_txt)
        r1.font.size = Pt(9)

    add_h2("1.8 Red de Puesta a Tierra y Protección Equipotencial (ITC-BT-18)")
    doc.add_paragraph(
        "Se proyecta una red de tierra mediante anillo perimetral enterrado en la cimentación de la nave "
        "con conductor de cobre desnudo de 35 mm² de sección, complementado por 6 picas verticales de acero "
        "cobrizado de 2 metros de longitud y 14 mm de diámetro hincadas en arquetas de registro con puente de prueba "
        "seccionable. Toda la estructura metálica de la nave, las envolventes metálicas de hornos y amasadoras, "
        "las tuberías metálicas de agua y las bandejas portacables se conectan a la barra equipotencial principal (BEP)."
    )

    add_h2("1.9 Alumbrado Normal, de Emergencia y Señalización")
    doc.add_paragraph(
        "El alumbrado normal se resuelve mediante regletas industriales LED Philips BN124C de 36 W montadas "
        "en trifásica para equilibrado de fases en obrador y luminarias downlight Philips DN065B en oficinas. "
        "El alumbrado de emergencia (ITC-BT-28 / UNE-EN 60598-2-22) garantiza una autonomía mínima de 1 hora "
        "proporcionando al menos 5 lux en cuadros eléctricos y puntos de primeros auxilios y 1 lux en vías de evacuación."
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
    
    # Tarjeta 1: Intensidad de Cálculo
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
    
    # Tarjeta 2: Criterio Motores ITC-BT-47
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
    
    # Tarjeta 3: Criterio Térmico (Intensidad Admisible UNE-HD 60364-5-52)
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
    
    # Tarjeta 4: Criterio de Caída de Tensión
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
    
    # Tarjeta 5: Criterio de Cortocircuito
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
    
    add_callout(
        "INCORPORACIÓN COMPLETA DE RECEPTORES",
        "Esta tabla integra todos los receptores dados en el enunciado del reto (hornos de 10 kW, climatizadores Mitsubishi, luminarias Philips, tomas, puerta) junto con las máquinas del obrador especificadas en Receptores_panaderia.xlsx (amasadoras Bongard, batidoras, silos, divisoras, laminadora, recarga VE y solar fotovoltaica). Completa las columnas de longitud, sección adoptada y calibres para tu entrega definitiva.",
        box_type="tip"
    )
    
    # Creamos la tabla con 10 columnas
    circ_headers = ["Línea / Circuito", "Receptor / Destino", "Tensión", "P (kW)", "cos φ", "L (m)", "Ib (A)", "Sección (mm²)", "ΔU (%)", "Protección Asignada"]
    circ_widths = [Inches(0.6), Inches(1.4), Inches(0.5), Inches(0.5), Inches(0.4), Inches(0.4), Inches(0.5), Inches(0.7), Inches(0.5), Inches(1.0)]
    
    # Datos de los circuitos
    circ_data = [
        ("LGA", "Acometida CGP a Contadores", "400 V", "138,2", "0,93", "25", "214,4", "3x150+1x70", "0,35 %", "BUC 250A gG"),
        ("DI", "Contadores a CGMP", "400 V", "138,2", "0,93", "15", "214,4", "3x150+1x70", "0,21 %", "IGA 4P 250A"),
        ("L-HORN", "4x Hornos Convección (10kW c/u)", "400 V", "40,0", "0,99", "25", "58,3", "5G16", "0,82 %", "4P 63A Curva C"),
        ("L-AMAS1", "Amasadora Bongard 130 kg", "400 V", "6,8", "0,85", "20", "14,5*", "5G4", "0,71 %", "4P 20A Curva D"),
        ("L-AMAS2", "Amasadora Bongard 80 kg", "400 V", "4,55", "0,85", "20", "9,7*", "5G2.5", "0,76 %", "4P 16A Curva D"),
        ("L-BAT", "Batidora Planetaria 60 L", "400 V", "1,5", "0,85", "18", "3,2*", "5G2.5", "0,22 %", "4P 10A Curva D"),
        ("L-SILO", "Silo Harina: Soplante + Sinfín", "400 V", "4,0", "0,85", "30", "8,5*", "5G4 (Ex)", "0,62 %", "4P 16A Curva D"),
        ("L-ENFR", "Enfriador de Agua Amasado", "400 V", "1,5", "0,85", "15", "3,2*", "5G2.5", "0,19 %", "4P 10A Curva C"),
        ("L-DIV", "2x Divisora-Formadora JAC", "400 V", "3,0", "0,85", "22", "6,4*", "5G2.5", "0,54 %", "4P 16A Curva D"),
        ("L-LAM", "Laminadora Rondo", "400 V", "1,4", "0,85", "20", "3,0*", "5G2.5", "0,23 %", "4P 10A Curva D"),
        ("L-CLIM", "3x Cassettes Mitsubishi 14kW", "400 V", "29,2", "0,90", "35", "46,8", "5G10", "1,45 %", "4P 50A Curva C"),
        ("L-ALU1", "Alumbrado Obrador 1 (20x36W)", "400 V", "0,72", "0,95", "45", "1,1", "5G2.5", "0,18 %", "4P 10A Curva C"),
        ("L-ALU2", "Alumbrado Obrador 2 (20x36W)", "400 V", "0,72", "0,95", "45", "1,1", "5G2.5", "0,18 %", "4P 10A Curva C"),
        ("L-ALU3", "Alumbrado Hornos/Silos (20x36W)", "400 V", "0,72", "0,95", "45", "1,1", "5G2.5", "0,18 %", "4P 10A Curva C"),
        ("L-ALU4", "Alumbrado Oficinas (6x23W)", "230 V", "0,14", "0,95", "27", "0,6", "3G1.5", "0,21 %", "2P 10A Curva C"),
        ("L-TC1", "Tomas Corriente Obrador 16A", "230 V", "3,68", "0,95", "35", "16,0", "3G2.5", "2,18 %", "2P 16A Curva C"),
        ("L-TC2", "Tomas Corriente Servicios 16A", "230 V", "3,68", "0,95", "40", "16,0", "3G2.5", "2,49 %", "2P 16A Curva C"),
        ("L-PUER", "Puerta de Acceso Motorizada", "230 V", "0,25", "0,80", "30", "1,4*", "3G1.5", "0,26 %", "2P 10A Curva C"),
        ("L-VE", "Recarga Vehículo Eléctrico (4x)", "400 V", "29,6", "0,98", "40", "43,6", "5G16", "0,98 %", "4P 50A Curva C / Dif B"),
        ("L-FV", "Inversor Solar Fotovoltaico", "400 V", "25,0", "1,00", "20", "36,1", "5G10", "0,65 %", "4P 40A Curva C")
    ]
    
    tbl_circ = doc.add_table(rows=len(circ_data)+1, cols=10)
    tbl_circ.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_circ, color="CBD5E0")
    
    for j, h in enumerate(circ_headers):
        c = tbl_circ.rows[0].cells[j]
        c.width = circ_widths[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, top=60, bottom=60, left=50, right=50)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(7.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    for i, data in enumerate(circ_data):
        row = tbl_circ.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]
            c.width = circ_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=50, right=50)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(7.5)
            if j == 0:
                r.bold = True
                
    p_notamot = doc.add_paragraph("(*) Intensidades de receptores con motor afectadas por el factor reglamentario 1,25 según la ITC-BT-47.")
    p_notamot.runs[0].font.size = Pt(8)
    p_notamot.runs[0].font.italic = True
    
    doc.add_page_break()

    # ==========================================
    # 2.3 JUSTIFICACIÓN DE CONTACTOS INDIRECTOS Y TIERRA
    # ==========================================
    add_h2("2.3 Justificación de la Protección contra Contactos Indirectos y Red de Tierra")
    
    add_callout(
        "JUSTIFICACIÓN EXIGIDA EN EL ENUNCIADO (APARTADO 3.2.1)",
        "El enunciado pide explícitamente: '3.2.1 Justificar la protección contra contactos indirectos. Justificar la resistencia de la tierra de protección de la instalación'. En el esquema TT, la seguridad se basa en la desconexión automática mediante interruptores diferenciales coordinados con el valor de la resistencia de tierra de las masas (RA <= UL / IΔn).",
        box_type="warning"
    )
    
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
    
    doc.add_paragraph(
        "Cálculo de la resistencia de difusión teórica de la red de tierra proyectada:\n"
        "Considerando un terreno con resistividad media estimada ρ = 100 Ω·m (terreno arcilloso/pedregoso de zona industrial Levante), "
        "la resistencia del anillo perimetral cerrado de L = 120 m de longitud enterrado a 0,8 m de profundidad se calcula según la fórmula del REBT:"
    )
    
    add_formula_card(
        "Resistencia de Difusión del Anillo y Picas Complementarias (ITC-BT-18)",
        "Fórmulas simplificadas para electrodos enterrados en terreno homogéneo:",
        "R_anillo = (2 · ρ) / L       |       R_pica = ρ / l",
        [
            ("R_anillo", "Resistencia del conductor enterrado horizontalmente en anillo: R_anillo = (2 · 100) / 120 = 1,67 Ω."),
            ("R_pica", "Resistencia de una pica vertical de l = 2 m: R_pica = 100 / 2 = 50,0 Ω."),
            ("R_total (anillo + 6 picas)", "Considerando la asociación en paralelo con factor de reducción por proximidad: R_tierra_calculada = 1,45 Ω."),
            ("Conclusión Justificativa", "El valor obtenido R_tierra = 1,45 Ω es extraordinariamente inferior al límite reglamentario más restrictivo de 80,0 Ω, garantizando que ninguna masa metálica alcance una tensión superior a 24 V bajo ningún defecto a tierra.")
        ]
    )

    add_h2("2.4 Justificación de la Protección contra Sobretensiones (ITC-BT-23)")
    
    add_callout(
        "JUSTIFICACIÓN EXIGIDA EN EL ENUNCIADO (APARTADO 3.2.2)",
        "El enunciado exige: '3.2.2 Justificar la protección contra sobretensiones'. Se debe justificar la doble protección: sobretensiones transitorias (rayos y conmutaciones en red según ITC-BT-23) y sobretensiones permanentes (rotura o corte del neutro según UNE-EN 50550).",
        box_type="warning"
    )
    
    doc.add_paragraph(
        "De acuerdo con la ITC-BT-23, se evalúa la necesidad y tipología de descargadores:"
    )
    
    st_points = [
        ("1. Nivel de Riesgo y Necesidad:", "La comarca de l'Horta Nord (Valencia) presenta una densidad media de impactos de rayo sobre el suelo Ng = 2 a 3 rayos/km²·año. Al tratarse de una nave industrial alimentada por red de distribución y con maquinaria electrónica y variadores, la protección es reglamentariamente obligatoria."),
        ("2. Protección contra Sobretensiones Transitorias:", "Se instala en el CGMP un protector combinado Tipo 1+2 (Iimp = 12,5 kA / Imax = 50 kA por polo) con nivel de protección Up <= 1,5 kV, coordinado aguas abajo con protectores Tipo 2 (In = 20 kA) en los subcuadros SC-1 (obrador), SC-4 (vehículo eléctrico) y SC-5 (fotovoltaica)."),
        ("3. Protección contra Sobretensiones Permanentes (POP):", "Se instala dispositivo conforme a la norma UNE-EN 50550 asociado al Interruptor General Automático (IGA) con bobina de disparo por emisión de corriente, que desconecta la instalación en menos de 0,2 segundos si la tensión entre fase y neutro supera los 275 V por pérdida del conductor neutro de la compañía.")
    ]
    for st_tit, st_txt in st_points:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        r0 = p.add_run(st_tit + " ")
        r0.bold = True
        r0.font.size = Pt(9)
        r1 = p.add_run(st_txt)
        r1.font.size = Pt(9)

    add_h2("2.5 Compensación de Reactiva y Filtrado de Armónicos en Obrador")
    
    add_callout(
        "ELEMENTO DIFERENCIADOR DE EXCELENCIA (PARA NOTA 10)",
        "La maquinaria del obrador (amasadoras, batidoras, laminadoras, compresores) tiene motores asíncronos y variadores de frecuencia que degradan el cos φ global de la panadería a valores de ~ 0,85 y generan distorsión armónica (THD). Diseñar la batería con reactancias antifiltro desintonizadas (p = 7%, fr = 189 Hz) previene sobrecalentamientos y resonancias.",
        box_type="tip"
    )
    
    add_formula_card(
        "Cálculo de la Potencia Reactiva de Compensación (Qc)",
        "Para elevar el factor de potencia medio de la planta desde cos φ1 = 0,85 hasta cos φ2 = 0,98 (eliminación total de penalizaciones por energía reactiva):",
        "Qc = P_sim · (tan φ1 - tan φ2)",
        [
            ("P_sim", "Potencia simultánea activa de la nave: P_sim = 138,2 kW."),
            ("tan φ1 (cos φ = 0,85)", "tan φ1 = 0,620."),
            ("tan φ2 (cos φ = 0,98)", "tan φ2 = 0,203."),
            ("Qc calculada", "Qc = 138,2 · (0,620 - 0,203) = 138,2 · 0,417 = 57,63 kVAr."),
            ("Batería Adoptada", "Batería automática de condensadores con reactancias de desintonización p = 7% (189 Hz) de 60 kVAr escalonada en (10 + 10 + 20 + 20 kVAr) con regulador microprocesado varimétrico.")
        ]
    )

    add_h2("2.6 Infraestructura para Recarga de Vehículos Eléctricos (ITC-BT-52)")
    doc.add_paragraph(
        "Se adopta el Esquema 2 de la ITC-BT-52 (instalación con cuadro principal y línea específica dedicada hacia los puntos de recarga). "
        "Para dar servicio a 4 puntos de recarga con tomas Tipo 2 (Mennekes) en Modo 3 de recarga inteligente, se implementa un "
        "Sistema de Protección de Línea (SPL) que monitoriza en tiempo real la potencia total demandada por el obrador y modula "
        "la corriente entregada a los vehículos para no sobrepasar la potencia contratada de la industria."
    )

    add_h2("2.7 Dimensionado del Sistema Solar Fotovoltaico en Cubierta (ITC-BT-40)")
    doc.add_paragraph(
        "Para reducir el coste operativo de la panadería durante las horas diurnas de amasado y cocción, se proyecta una instalación "
        "solar fotovoltaica para autoconsumo sobre la cubierta sándwich de la nave compuesta por 56 módulos monocristalinos bifaciales "
        "de 450 Wp (25,2 kWp totales) acoplados a un inversor trifásico de red de 25 kW (400 V), dotado de sistema antivertido homologado "
        "y protecciones de interconexión exigidas por la ITC-BT-40."
    )

    doc.add_page_break()

    # ==========================================
    # 7. DOCUMENTO Nº 3: PLIEGO DE CONDICIONES Y MATERIALES
    # ==========================================
    add_h1("DOCUMENTO Nº 3: PLIEGO DE CONDICIONES Y SELECCIÓN DE MATERIALES")
    
    add_callout(
        "REQUISITO DEL ENUNCIADO (APARTADO 3.3)",
        "El enunciado especifica: '3.3 Selección de los materiales. Se indicarán de forma detallada los elementos seleccionados, las Normas que cumplirán, y un modelo de un fabricante de referencia de reconocido prestigio (se indicará modelo, calibre y poder de corte)'.",
        box_type="warning"
    )
    
    add_h2("3.1 Normativa Técnica y Reglamentación Aplicable")
    
    normas = [
        ("Reglamento Electrotécnico para Baja Tensión (REBT):", "Real Decreto 842/2002 y sus Instrucciones Técnicas Complementarias (ITC-BT-01 a ITC-BT-52)."),
        ("Norma UNE-HD 60364-5-52:", "Instalaciones eléctricas de baja tensión. Selección e instalación de equipos. Canalizaciones y corrientes admisibles."),
        ("Norma UNE-EN 60898-1 / UNE-EN 60947-2:", "Aparamenta de baja tensión. Interruptores automáticos para instalaciones domésticas e industriales."),
        ("Norma UNE-EN 61008-1 / UNE-EN 62423:", "Interruptores diferenciales tipo AC, A y B para instalaciones industriales."),
        ("Norma UNE-EN 50575 (Reglamento CPR):", "Cables de energía, control y comunicación. Requisitos de reacción al fuego (Clase Cca-s1b,d1,a1)."),
        ("Norma UNE-EN 60079-0 / UNE-EN 60079-31:", "Atmósferas explosivas. Material eléctrico y protección por envolvente contra polvo combustible (Ex tc).")
    ]
    for n_tit, n_desc in normas:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(2)
        r0 = p.add_run(n_tit + " ")
        r0.bold = True
        r0.font.size = Pt(8.5)
        r1 = p.add_run(n_desc)
        r1.font.size = Pt(8.5)

    add_h2("3.2 Especificación Justificada de Aparamenta y Materiales de Referencia")
    doc.add_paragraph(
        "En cumplimiento de los requisitos docentes, se seleccionan equipos de fabricantes líderes a nivel europeo "
        "(Schneider Electric, Prysmian Group, Philips y AISCAN), indicando sus características técnicas completas:"
    )
    
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
        ("Luminarias LED Industriales", "Signify (Philips)", "BN124C LED41S/840", "UNE-EN 60598-1", "36 W, 4.100 lm, 4000K, IP65, IK08"),
        ("Luminarias LED Oficinas", "Signify (Philips)", "DN065B LED20S/830", "UNE-EN 60598-1", "23 W, 2.000 lm, 3000K, IP20, empotrar"),
        ("Luminarias de Emergencia", "Legrand", "U21 LED 70 lm", "UNE-EN 60598-2-22", "Autonomía 1h, LED no permanente, IP42"),
        ("Canalizaciones Bajo Tubo", "AISCAN", "Tubo Rígido Blindado CR", "UNE-EN 61386-21", "Grado 5557, resistencia compresión > 4000 N")
    ]
    
    tbl_mat = doc.add_table(rows=len(mat_data)+1, cols=5)
    tbl_mat.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_mat, color="CBD5E0")
    for j, h in enumerate(mat_headers):
        c = tbl_mat.rows[0].cells[j]
        c.width = mat_widths[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, top=60, bottom=60, left=60, right=60)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    for i, data in enumerate(mat_data):
        row = tbl_mat.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]
            c.width = mat_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=60, right=60)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(7.5)
            if j == 0:
                r.bold = True

    add_h2("3.3 Prescripciones de Montaje y Tendido de Canalizaciones")
    doc.add_paragraph(
        "Las canalizaciones en obrador y zonas de trabajo se instalarán mediante bandejas perforadas de acero "
        "galvanizado en caliente suspendidas a 3,5 metros de altura o bajo tubo rígido blindado de alta resistencia "
        "mecánica (IK10). En las zonas ATEX (silos y tolvas) todos los pasos de cables dispondrán de prensaestopas metálicos "
        "con rosca métrica y certificado ATEX Ex d/e para evitar la propagación de atmósferas con polvo."
    )

    add_h2("3.4 Pruebas Reglamentarias y Ensayos de Puesta en Servicio (ITC-BT-05)")
    doc.add_paragraph(
        "Antes de la energización de la instalación se ejecutarán los ensayos oficiales previstos en la ITC-BT-05:\n"
        "1. Medida de continuidad de los conductores de protección y equipotenciales.\n"
        "2. Medida de la resistencia de aislamiento entre conductores activos y entre estos y tierra (R_aisl >= 1,0 MΩ a 500 Vcc).\n"
        "3. Medida de la resistencia del electrodo de puesta a tierra mediante telurómetro de 3 picas.\n"
        "4. Verificación del tiempo y corriente de disparo de todos los interruptores diferenciales con analizador de redes."
    )

    doc.add_page_break()

    # ==========================================
    # 8. DOCUMENTO Nº 4: PRESUPUESTO
    # ==========================================
    add_h1("DOCUMENTO Nº 4: PRESUPUESTO")
    
    add_callout(
        "ESTRUCTURA ECONÓMICA DE PROYECTO",
        "El presupuesto debe estructurarse en capítulos homogéneos y concluir con el Resumen General del Presupuesto, aplicando los porcentajes reglamentarios de Gastos Generales (13%), Beneficio Industrial (6%) e Impuesto sobre el Valor Añadido (IVA 21%).",
        box_type="tip"
    )
    
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
        c = tbl_cap.rows[0].cells[j]
        c.width = cap_widths[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        if j == 2:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
    for i, data in enumerate(cap_data):
        row = tbl_cap.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]
            c.width = cap_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=80, right=80)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if j == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r.bold = True

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
        c = tbl_res.rows[0].cells[j]
        c.width = res_widths[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        if j == 2:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
    for i, data in enumerate(res_data):
        row = tbl_res.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]
            c.width = res_widths[j]
            set_cell_margins(c, top=60, bottom=60, left=80, right=80)
            if "TOTAL" in data[0] or "SUBTOTAL" in data[0] or "PEM" in data[0]:
                set_cell_background(c, "EBF4FF")
            elif i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if "TOTAL" in data[0] or "SUBTOTAL" in data[0] or "PEM" in data[0]:
                r.bold = True
                if "TOTAL GENERAL" in data[0]:
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = RGBColor(27, 54, 93)
            if j == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p_letra = doc.add_paragraph(
        "Asciende el presente Presupuesto General a la cantidad de CIENTO TREINTA Y CUATRO MIL SEISCIENTOS UN EUROS CON OCHENTA Y CINCO CÉNTIMOS (134.601,85 €), I.V.A. incluido."
    )
    p_letra.runs[0].font.size = Pt(8.5)
    p_letra.runs[0].font.italic = True
    
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
        "Asegúrate de exportar tus planos de AutoCAD/Revit SIEMPRE con fondo blanco limpio.",
        box_type="warning"
    )
    
    tbl_planos = doc.add_table(rows=5, cols=4)
    tbl_planos.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_planos, color="CBD5E0")
    pl_widths = [Inches(1.0), Inches(2.5), Inches(1.5), Inches(1.5)]
    pl_headers = ["Nº Plano", "Título del Plano", "Formato Normalizado", "Escala de Representación"]
    for j, h in enumerate(pl_headers):
        c = tbl_planos.rows[0].cells[j]
        c.width = pl_widths[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    pl_data = [
        ("Plano Nº 1", "Situación, Emplazamiento y Acometida Eléctrica", "A3 (Fondo Blanco)", "1:500 / 1:1000"),
        ("Plano Nº 2", "Distribución en Planta, Maquinaria y Canalizaciones", "A3 (Fondo Blanco)", "1:100"),
        ("Plano Nº 3", "Esquema Unifilar General del CGMP y Subcuadros", "A3 (Fondo Blanco)", "Sin escala (Normalizado)"),
        ("Plano Nº 4", "Red de Puesta a Tierra y Red Equipotencial", "A3 (Fondo Blanco)", "1:100")
    ]
    for i, data in enumerate(pl_data):
        row = tbl_planos.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]
            c.width = pl_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=80, right=80)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if j == 0:
                r.bold = True

    doc.add_page_break()

    # ==========================================
    # 10. BIBLIOGRAFÍA Y REFERENCIAS NORMATIVAS (UPV)
    # ==========================================
    add_h1("BIBLIOGRAFÍA Y FUENTES DE INFORMACIÓN (FORMATO UPV / IEEE)")
    
    add_callout(
        "REQUISITO DE CITACIÓN UPV (COMENTARIOS)",
        "El profesor indica: 'Se ha de citar según normas de Biblioteca de UPV, esto se aprende en 10 minutos. Si no sabéis citar, preguntad. No cumplir este apartado genera un cero'. "
        "A continuación se incluyen las referencias completas y normalizadas de las normativas oficiales, libros de texto y manuales de fabricantes de referencia.",
        box_type="warning"
    )
    
    bib_items = [
        ("[1] MINISTERIO DE CIENCIA Y TECNOLOGÍA.", "Real Decreto 842/2002, de 2 de agosto, por el que se aprueba el Reglamento Electrotécnico para Baja Tensión (REBT) e Instrucciones Técnicas Complementarias (ITC) BT 01 a BT 52. Boletín Oficial del Estado (BOE), núm. 224, 18 de septiembre de 2002."),
        ("[2] MINISTERIO DE INDUSTRIA, ENERGÍA Y TURISMO.", "Guía Técnica de Aplicación del Reglamento Electrotécnico para Baja Tensión. Madrid: Dirección General de Industria y de la Pequeña y Mediana Empresa, 2019."),
        ("[3] ASOCIACIÓN ESPAÑOLA DE NORMALIZACIÓN (UNE).", "UNE-HD 60364-5-52:2014. Instalaciones eléctricas de baja tensión. Parte 5-52: Selección e instalación de equipos eléctricos. Canalizaciones. Madrid: AENOR, 2014."),
        ("[4] ASOCIACIÓN ESPAÑOLA DE NORMALIZACIÓN (UNE).", "UNE 157001:2014. Criterios generales para la elaboración formal de los documentos que constituyen un proyecto técnico. Madrid: AENOR, 2014."),
        ("[5] ASOCIACIÓN ESPAÑOLA DE NORMALIZACIÓN (UNE).", "UNE-EN 50575:2015/A1:2016. Cables de energía, control y comunicación. Requisitos de reacción al fuego (Reglamento de Productos de Construcción - CPR). Madrid: AENOR, 2016."),
        ("[6] SCHNEIDER ELECTRIC.", "Guía de diseño de instalaciones eléctricas según normas internacionales CEI. Rueil-Malmaison (Francia): Schneider Electric Industries SAS, 2024."),
        ("[7] PRYSMIAN GROUP.", "Catálogo General de Cables para Baja Tensión y Guía CPR. Vilanova i la Geltrú: Prysmian Cables Spain S.A., 2023."),
        ("[8] SIGNIFY IBERIA (PHILIPS LIGHTING).", "Manual Técnico de Iluminación Industrial y Catálogo de Luminarias LED CoreLine. Madrid: Signify, 2024."),
        ("[9] MITSUBISHI ELECTRIC IBERIA.", "Ficha Técnica y Manual de Instalación Climatización Cassettes Serie Mr. Slim (SPLZS-140YBA). Sant Cugat del Vallès: Mitsubishi Electric, 2023."),
        ("[10] GARCÍA TRASANCOS, J.", "Instalaciones Eléctricas de Enlace y Centros de Transformación. 6ª ed. Madrid: Ediciones Paraninfo, 2020.")
    ]
    for num, cit in bib_items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.first_line_indent = Inches(-0.4)
        p.paragraph_format.space_after = Pt(4)
        r0 = p.add_run(num + " ")
        r0.bold = True
        r0.font.size = Pt(8.5)
        r0.font.color.rgb = RGBColor(27, 54, 93)
        r1 = p.add_run(cit)
        r1.font.size = Pt(8.5)

    doc.add_page_break()

    # ==========================================
    # 11. ANEXO: CHECKLIST PARA OBTENER UN 10
    # ==========================================
    add_h1("ANEXO: CHECKLIST DE AUTOEVALUACIÓN PARA OBTENER UN 10")
    
    add_callout(
        "CRITERIOS DE CALIFICACIÓN DEL PROFESOR",
        "El profesor indica: 'La redacción correcta, sin errores importantes y presentación aceptable garantiza 5 puntos. El incremento hasta 10 puntos se corresponde con la presentación mejorada, complicación y exactitud del procedimiento de cálculo y soluciones aportadas más eficientes'. "
        "Utilizad esta lista de comprobación antes de generar el PDF definitivo para asegurar que cumplís todos los requisitos de excelencia.",
        box_type="tip"
    )
    
    chk_headers = ["Aspecto Evaluado", "Requisito Exigido", "Criterio de Excelencia para el 10", "Estado"]
    chk_widths = [Inches(1.5), Inches(1.8), Inches(2.4), Inches(0.8)]
    chk_data = [
        ("Potencia de la Nave", "Entre 120 y 190 kW", "Potencia justificada con coeficientes Ku y Ks y desglose por maquinaria real de obrador", "[  ] OK"),
        ("Superficie y Altillo", "Altillo <= 30% y >= 150 m²", "Distribución de superficies coherente con proceso de panadería y longitudes múltiplos de 0,5 m", "[  ] OK"),
        ("Aparcamiento y VE", "Entre 20 y 40 plazas", "Puntos de recarga Modo 3 con modulación SPL (ITC-BT-52) y diferencial Tipo B", "[  ] OK"),
        ("Cálculos y Fórmulas", "No pegar pantallazos de software", "Ecuaciones teóricas explicadas antes de las tablas; justificación de secciones por Iadm y caída de tensión", "[  ] OK"),
        ("Locales Especiales", "REBT general", "Clasificación rigurosa de Zona 22 ATEX por polvo de harina (ITC-BT-29) y locales mojados (ITC-BT-30)", "[  ] OK"),
        ("Contactos Indirectos", "Justificación de Rtierra", "Cálculo analítico del anillo y picas demostrando RA <= 24V / 0,3A = 80 Ω", "[  ] OK"),
        ("Sobretensiones", "Justificación transitoria y permanente", "Descargadores Tipo 1+2 coordinados con Tipo 2 y protección POP UNE-EN 50550", "[  ] OK"),
        ("Materiales y Marcas", "Indicar modelos y calibres", "Especificación completa con fabricantes de prestigio (Schneider, Prysmian, Philips)", "[  ] OK"),
        ("Planos A3 Normalizados", "Fondo claro obligatorio", "Fondo blanco pulcro, cajetín oficial UPV acotado y grafiado profesional (fondo negro = cero)", "[  ] OK"),
        ("Citas y Bibliografía", "Normas Biblioteca UPV", "Citas en formato formal IEEE/ISO con normas UNE, REBT y catálogos de fabricantes", "[  ] OK"),
        ("Firma Digital", "Documento único en PDF firmado", "Exportar a PDF y aplicar certificado digital oficial (DNIe / FNMT / ACCV)", "[  ] OK")
    ]
    
    tbl_chk = doc.add_table(rows=len(chk_data)+1, cols=4)
    tbl_chk.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_chk, color="CBD5E0")
    for j, h in enumerate(chk_headers):
        c = tbl_chk.rows[0].cells[j]
        c.width = chk_widths[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, top=60, bottom=60, left=70, right=70)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(8.0)
        r.font.color.rgb = RGBColor(255, 255, 255)
        if j == 3:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
    for i, data in enumerate(chk_data):
        row = tbl_chk.rows[i+1]
        for j, val in enumerate(data):
            c = row.cells[j]
            c.width = chk_widths[j]
            set_cell_margins(c, top=50, bottom=50, left=70, right=70)
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(7.5)
            if j == 0:
                r.bold = True
            if j == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r.bold = True
                r.font.color.rgb = RGBColor(13, 92, 117)

    output_filename = "Plantilla_Proyecto_Instalacion_Electrica_Panaderia.docx"
    doc.save(output_filename)
    print(f"Documento generado exitosamente como '{output_filename}'")

if __name__ == "__main__":
    create_document()
