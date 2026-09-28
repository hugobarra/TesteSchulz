import io
import zipfile
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors


def repair_xlsx_stream(file_bytes_or_path) -> io.BytesIO:
    """
    Sanitiza os metadados Strict OpenXML corrompidos diretamente da estrutura ZIP.
    """
    in_buf = file_bytes_or_path if isinstance(file_bytes_or_path, io.BytesIO) else open(file_bytes_or_path, "rb")
    out_buf = io.BytesIO()

    strict_ns = "http://purl.oclc.org/ooxml/spreadsheetml/main"
    trans_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    strict_rel = "http://purl.oclc.org/ooxml/officeDocument/relationships"
    trans_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    with zipfile.ZipFile(in_buf, "r") as z_in, zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as z_out:
        for item in z_in.infolist():
            data = z_in.read(item.filename)
            if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
                data_str = data.decode("utf-8", errors="ignore")
                data_str = data_str.replace(strict_ns, trans_ns).replace(strict_rel, trans_rel)
                data_str = data_str.replace('conformance="strict"', '')
                data = data_str.encode("utf-8")
            z_out.writestr(item, data)

    out_buf.seek(0)
    return out_buf


def process_and_validate_sheet(file_input, target_sheet_name="Registro"):
    """
    Extrai metadados dinamicamente, filtra rodapés indesejados e calcula os Limites Mín/Máx.
    """
    repaired_file = repair_xlsx_stream(file_input)

    try:
        excel_file = pd.ExcelFile(repaired_file, engine='calamine')
    except Exception:
        excel_file = pd.ExcelFile(repaired_file)

    matching_sheet = None
    for sheet in excel_file.sheet_names:
        if sheet.strip().lower() == target_sheet_name.lower():
            matching_sheet = sheet
            break

    if not matching_sheet:
        matching_sheet = excel_file.sheet_names[0]

    df_raw = pd.read_excel(excel_file, sheet_name=matching_sheet)

    # 1. EXTRAÇÃO DINÂMICA DE METADADOS
    keywords = {
        'cliente': ['cliente', 'solicitante'],
        'tecnico': ['técnico', 'tecnico', 'executante'],
        'codigo': ['código', 'codigo', 'identificação', 'identificacao'],
        'denominacao': ['denominação', 'denominacao', 'instrumento', 'equipamento'],
        'data': ['data']
    }

    meta_extracted = {
        'cliente': 'Não informado',
        'tecnico': 'Não informado',
        'codigo': 'S/N',
        'denominacao': 'Instrumento de medição',
        'data': '-'
    }

    for idx, row in df_raw.iterrows():
        for key, kw_list in keywords.items():
            if meta_extracted[key] in ['Não informado', 'S/N', 'Instrumento de medição', '-']:
                for col_i, cell in enumerate(row.values):
                    cell_s = str(cell).strip().lower()
                    if any(kw in cell_s for kw in kw_list):
                        for next_val in row.values[col_i + 1:]:
                            if pd.notna(next_val) and str(next_val).strip() != "":
                                meta_extracted[key] = str(next_val).strip()
                                break

    # 2. LOCALIZAÇÃO E FILTRAGEM DA TABELA DE MEDIÇÕES
    header_row_idx = None
    for idx, row in df_raw.iterrows():
        row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if any("parâmetro" in v or "parametro" in v for v in row_vals):
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise ValueError("A tabela de medições não foi encontrada na planilha.")

    df_med = df_raw.iloc[header_row_idx + 1:].copy()
    df_med.columns = [str(c).strip() for c in df_raw.iloc[header_row_idx].values]

    # Identificação de colunas essenciais
    col_param = next((c for c in df_med.columns if "parâmetro" in c.lower() or "parametro" in c.lower()), None)
    col_nom = next((c for c in df_med.columns if "nominal" in c.lower()), None)
    col_tol_plus = next((c for c in df_med.columns if "tol. +" in c.lower() or "tol +" in c.lower()), None)
    col_tol_minus = next((c for c in df_med.columns if "tol. -" in c.lower() or "tol -" in c.lower()), None)

    if not col_param or not col_nom:
        raise ValueError("Colunas obrigatórias de 'Parâmetro' e 'Valor nominal' não foram identificadas.")

    cols_medidas = [c for c in df_med.columns if "medida" in c.lower()]

    # FILTRAGEM RIGOROSA PARA REMOVER RODAPÉS (Como 2, 0, 2 da incerteza)
    df_med = df_med.dropna(subset=[col_param]).copy()
    df_med[col_param] = df_med[col_param].astype(str).str.strip()
    
    # Remove textos de rodapé/observações e valores espúrios no nome do parâmetro
    df_med = df_med[~df_med[col_param].isin(['', 'nan', 'None', '2', '0', '1', '3', '4'])].copy()
    df_med = df_med[~df_med[col_param].str.lower().str.contains("incerteza|observa|nota|folha", na=False)].copy()

    # Garante que Nominal e Tolerâncias são numéricos
    df_med[col_nom] = pd.to_numeric(df_med[col_nom], errors='coerce')
    df_med = df_med.dropna(subset=[col_nom]).reset_index(drop=True)

    df_med['Tol_Plus_Num'] = pd.to_numeric(df_med[col_tol_plus], errors='coerce').fillna(0.0) if col_tol_plus else 0.0
    df_med['Tol_Minus_Num'] = pd.to_numeric(df_med[col_tol_minus], errors='coerce').fillna(0.0) if col_tol_minus else 0.0

    # Converter e calcular a média das medições
    for m_col in cols_medidas:
        df_med[m_col] = pd.to_numeric(df_med[m_col], errors='coerce')

    if cols_medidas:
        df_med['Valor Obtido'] = df_med[cols_medidas].mean(axis=1)
    else:
        df_med['Valor Obtido'] = df_med[col_nom]

    # 3. CÁLCULO DOS LIMITES MÍNIMOS, MÁXIMOS E ERRO
    df_med['Limite Mín'] = df_med[col_nom] - df_med['Tol_Minus_Num'].abs()
    df_med['Limite Máx'] = df_med[col_nom] + df_med['Tol_Plus_Num'].abs()
    df_med['Erro'] = df_med['Valor Obtido'] - df_med[col_nom]

    # Validação com tolerância de precisão float
    df_med['Status_Bool'] = (df_med['Valor Obtido'] >= (df_med['Limite Mín'] - 1e-7)) & \
                            (df_med['Valor Obtido'] <= (df_med['Limite Máx'] + 1e-7))

    df_med['Status_Txt'] = df_med['Status_Bool'].map({True: 'APROVADO', False: 'REPROVADO'})

    status_geral = "APROVADO" if df_med['Status_Bool'].all() else "REPROVADO"
    meta_extracted['status_geral'] = status_geral

    df_med['Parâmetro'] = df_med[col_param]
    df_med['Nominal'] = df_med[col_nom]

    return meta_extracted, df_med


def gerar_pdf_bytes(info_dict: dict, df_medicoes: pd.DataFrame) -> io.BytesIO:
    """
    Gera o certificado PDF limpo sem linhas de rodapé espúrias.
    """
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=letter,
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=15,
        leading=18,
        alignment=1,
        textColor=colors.HexColor("#1A365D")
    )

    normal_style = styles['Normal']
    header_table_style = ParagraphStyle('HeaderTable', parent=normal_style, fontSize=7.5, leading=9, fontName="Helvetica-Bold", textColor=colors.white, alignment=1)
    cell_style = ParagraphStyle('CellTable', parent=normal_style, fontSize=7, leading=8.5, alignment=1)
    cell_param_style = ParagraphStyle('CellParam', parent=normal_style, fontSize=7, leading=8.5, alignment=0)

    status_aprov_style = ParagraphStyle('Aprov', parent=normal_style, fontSize=11, leading=13, fontName="Helvetica-Bold", textColor=colors.HexColor("#2F855A"), alignment=1)
    status_reprov_style = ParagraphStyle('Reprov', parent=normal_style, fontSize=11, leading=13, fontName="Helvetica-Bold", textColor=colors.HexColor("#C53030"), alignment=1)

    story = []

    # Título do Relatório
    story.append(Paragraph("CERTIFICADO DE CALIBRAÇÃO E VALIDAÇÃO", title_style))
    story.append(Spacer(1, 8))

    # Tabela de Metadados
    info_data = [
        [Paragraph(f"<b>Cliente:</b> {info_dict.get('cliente', '')}", normal_style),
         Paragraph(f"<b>Data:</b> {info_dict.get('data', '')}", normal_style)],
        [Paragraph(f"<b>Técnico Executante:</b> {info_dict.get('tecnico', '')}", normal_style),
         Paragraph(f"<b>Código do Instrumento:</b> {info_dict.get('codigo', '')}", normal_style)],
        [Paragraph(f"<b>Denominação:</b> {info_dict.get('denominacao', '')}", normal_style), ""]
    ]

    t_info = Table(info_data, colWidths=[320, 240])
    t_info.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F7FAFC")),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_info)
    story.append(Spacer(1, 8))

    # Banner do Status Geral
    status_geral = info_dict.get('status_geral', 'REPROVADO')
    banner_bg = colors.HexColor("#C6F6D5") if status_geral == "APROVADO" else colors.HexColor("#FED7D7")
    status_paragraph = Paragraph(f"PARECER FINAL: INSTRUMENTO {status_geral}", status_aprov_style if status_geral == "APROVADO" else status_reprov_style)

    t_status = Table([[status_paragraph]], colWidths=[560])
    t_status.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), banner_bg),
        ('BOX', (0, 0), (-1, -1), 1.5, colors.HexColor("#2F855A") if status_geral == "APROVADO" else colors.HexColor("#C53030")),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(t_status)
    story.append(Spacer(1, 10))

    # Tabela com Limites Mín, Máx, Erro e Status
    story.append(Paragraph("<b>Resultados Detalhados das Medições</b>", styles['Heading2']))
    story.append(Spacer(1, 4))

    col_headers = ["Parâmetro Avaliado", "Nominal", "Lím. Mín", "Lím. Máx", "Obtido", "Erro", "Status"]
    table_data = [[Paragraph(c, header_table_style) for c in col_headers]]

    for _, row in df_medicoes.iterrows():
        st_txt = row['Status_Txt']
        st_cell_style = ParagraphStyle('StCell', parent=cell_style, fontName="Helvetica-Bold",
                                       textColor=colors.HexColor("#2F855A") if st_txt == "APROVADO" else colors.HexColor("#C53030"))

        row_cells = [
            Paragraph(str(row['Parâmetro']), cell_param_style),
            Paragraph(f"{row['Nominal']:.4f}", cell_style),
            Paragraph(f"{row['Limite Mín']:.4f}", cell_style),
            Paragraph(f"{row['Limite Máx']:.4f}", cell_style),
            Paragraph(f"{row['Valor Obtido']:.4f}", cell_style),
            Paragraph(f"{row['Erro']:.4f}", cell_style),
            Paragraph(st_txt, st_cell_style)
        ]
        table_data.append(row_cells)

    col_widths = [170, 65, 65, 65, 65, 60, 70]
    t_medicoes = Table(table_data, colWidths=col_widths, repeatRows=1)
    t_medicoes.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))

    story.append(t_medicoes)

    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer