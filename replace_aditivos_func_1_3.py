from pathlib import Path
path = Path('gerar_relatorios_com_formulas_1.3.py')
text = path.read_text(encoding='utf-8')
start = text.index('def write_aditivos_distrato_sheet')
end = text.index('def write_client_sheet', start)
new_func = '''def write_aditivos_distrato_sheet(sheet, summary_normal, summary_distratado, price_data, bold_font, regular_font, currency_format, accounting_format, header_fill, bold_white_font=None, regular_white_font=None, client_normal=None, client_distratado=None):
    sheet.column_dimensions['A'].width = 12.29
    sheet.column_dimensions['B'].width = 43.37
    sheet.column_dimensions['C'].width = 10.14
    sheet.column_dimensions['D'].width = 5.00
    sheet.column_dimensions['E'].width = 13.91
    sheet.column_dimensions['F'].width = 13.20

    sheet.column_dimensions['G'].width = 2.00
    sheet.column_dimensions['H'].width = 12.29
    sheet.column_dimensions['I'].width = 43.37
    sheet.column_dimensions['J'].width = 10.14
    sheet.column_dimensions['K'].width = 5.00
    sheet.column_dimensions['L'].width = 13.91
    sheet.column_dimensions['M'].width = 13.20

    sheet.merge_cells('A1:F1')
    sheet['A1'] = 'ADITIVOS'
    sheet['A1'].font = bold_font
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['A1'].fill = header_fill

    sheet.merge_cells('H1:M1')
    sheet['H1'] = 'DISTRATO'
    sheet['H1'].font = bold_font
    sheet['H1'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['H1'].fill = PatternFill(start_color='31859c', end_color='31859c', fill_type='solid')

    headers = [('ID. Bloco/Torre', 1), ('Tipo R. Bassani', 2), ('Descrição', 3), ('Metragem', 4), ('Un', 5), ('Valor Total', 6)]
    for text, col in headers:
        cell = sheet.cell(row=2, column=col, value=text)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.fill = header_fill

    for text, col in headers:
        right_col = col + 7
        cell = sheet.cell(row=2, column=right_col, value=text)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.fill = PatternFill(start_color='31859c', end_color='31859c', fill_type='solid')

    def normalize_category(category):
        if category == 'Forro':
            return 'Forro'
        if category in ['Parede', 'Revestimento', 'Paredes_e_Revestimentos', 'Paredes e Revestimentos']:
            return 'Paredes e Revestimentos'
        return None

    def build_rows_by_category(client_data, category_name):
        rows = []
        if not client_data:
            return rows
        for client_id in sorted(client_data.keys(), key=natural_sort_key):
            items = [item for item in client_data[client_id].values() if normalize_category(item.get('Categoria')) == category_name]
            if not items:
                continue
            rows.append({'is_block_header': True, 'block_id': client_id})
            for item in sorted(items, key=lambda x: natural_sort_key(x.get('Tipo Code', ''))):
                rows.append({'is_block_header': False, 'block_id': client_id, 'item': item})
        return rows

    normal_forro_rows = build_rows_by_category(client_normal or {}, 'Forro')
    normal_paredes_rows = build_rows_by_category(client_normal or {}, 'Paredes e Revestimentos')
    distr_forro_rows = build_rows_by_category(client_distratado or {}, 'Forro')
    distr_paredes_rows = build_rows_by_category(client_distratado or {}, 'Paredes e Revestimentos')

    def write_section(left_rows, right_rows, start_row, section_title):
        sheet.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6)
        sheet.merge_cells(start_row=start_row, start_column=8, end_row=start_row, end_column=13)
        left_sec = sheet.cell(row=start_row, column=1, value=f'{section_title} - ADITIVOS')
        right_sec = sheet.cell(row=start_row, column=8, value=f'{section_title} - DISTRATO')
        for sec_cell, fill_color in [(left_sec, '6aa84f'), (right_sec, '31859c')]:
            sec_cell.font = bold_white_font or bold_font
            sec_cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
            sec_cell.alignment = Alignment(horizontal='center', vertical='center')

        current_row = start_row + 1
        max_rows = max(len(left_rows), len(right_rows))
        for i in range(max_rows):
            if i < len(left_rows):
                left = left_rows[i]
                if left['is_block_header']:
                    sheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
                    hcell = sheet.cell(row=current_row, column=1, value=f"Bloco/Torre: {left['block_id']}")
                    hcell.font = bold_font
                    hcell.fill = header_fill
                    hcell.alignment = Alignment(horizontal='left', vertical='center')
                else:
                    item = left['item']
                    tipo_code = item.get('Tipo Code', '')
                    descricao = item.get('Descricao', '')
                    metragem = round(item.get('BaseQuantity', 0), 2)
                    unit = price_data.get(tipo_code, {}).get('Un', '')
                    valor_unit = round(price_data.get(tipo_code, {}).get('Valor', 0), 2)

                    sheet.cell(row=current_row, column=1, value=left['block_id'])
                    sheet.cell(row=current_row, column=2, value=tipo_code)
                    sheet.cell(row=current_row, column=3, value=descricao)
                    sheet.cell(row=current_row, column=4, value=metragem).number_format = currency_format
                    sheet.cell(row=current_row, column=5, value=unit)
                    sheet.cell(row=current_row, column=6, value=f"=ROUND({metragem}*{valor_unit},2)").number_format = accounting_format

            if i < len(right_rows):
                right = right_rows[i]
                if right['is_block_header']:
                    sheet.merge_cells(start_row=current_row, start_column=8, end_row=current_row, end_column=13)
                    hcell = sheet.cell(row=current_row, column=8, value=f"Bloco/Torre: {right['block_id']}")
                    hcell.font = bold_font
                    hcell.fill = PatternFill(start_color='31859c', end_color='31859c', fill_type='solid')
                    hcell.alignment = Alignment(horizontal='left', vertical='center')
                else:
                    item = right['item']
                    tipo_code = item.get('Tipo Code', '')
                    descricao = item.get('Descricao', '')
                    metragem = round(item.get('BaseQuantity', 0), 2)
                    unit = price_data.get(tipo_code, {}).get('Un', '')
                    valor_unit = round(price_data.get(tipo_code, {}).get('Valor', 0), 2)

                    sheet.cell(row=current_row, column=8, value=right['block_id'])
                    sheet.cell(row=current_row, column=9, value=tipo_code)
                    sheet.cell(row=current_row, column=10, value=descricao)
                    sheet.cell(row=current_row, column=11, value=metragem).number_format = currency_format
                    sheet.cell(row=current_row, column=12, value=unit)
                    sheet.cell(row=current_row, column=13, value=f"=ROUND({metragem}*{valor_unit},2)").number_format = accounting_format

            current_row += 1

        end_row = current_row - 1
        return start_row + 1, end_row, current_row

    forro_data_start, forro_data_end, next_row = write_section(normal_forro_rows, distr_forro_rows, 3, 'FORROS')
    total_forro_row = next_row + 1
    sheet.cell(row=total_forro_row, column=1, value='TOTAL FORROS').font = bold_font
    sheet.merge_cells(start_row=total_forro_row, start_column=1, end_row=total_forro_row, end_column=4)
    if forro_data_end >= forro_data_start:
        sheet.cell(row=total_forro_row, column=6, value=f"=ROUND(SUM(F{forro_data_start}:F{forro_data_end}),2)").number_format = accounting_format
        sheet.cell(row=total_forro_row, column=13, value=f"=ROUND(SUM(M{forro_data_start}:M{forro_data_end}),2)").number_format = accounting_format
    else:
        sheet.cell(row=total_forro_row, column=6, value=0).number_format = accounting_format
        sheet.cell(row=total_forro_row, column=13, value=0).number_format = accounting_format

    next_row = total_forro_row + 2
    paredes_data_start, paredes_data_end, next_row2 = write_section(normal_paredes_rows, distr_paredes_rows, next_row, 'PAREDES e REVESTIMENTOS')
    total_paredes_row = next_row2 + 1
    sheet.cell(row=total_paredes_row, column=1, value='TOTAL PAREDES e REVESTIMENTOS').font = bold_font
    sheet.merge_cells(start_row=total_paredes_row, start_column=1, end_row=total_paredes_row, end_column=4)
    if paredes_data_end >= paredes_data_start:
        sheet.cell(row=total_paredes_row, column=6, value=f"=ROUND(SUM(F{paredes_data_start}:F{paredes_data_end}),2)").number_format = accounting_format
        sheet.cell(row=total_paredes_row, column=13, value=f"=ROUND(SUM(M{paredes_data_start}:M{paredes_data_end}),2)").number_format = accounting_format
    else:
        sheet.cell(row=total_paredes_row, column=6, value=0).number_format = accounting_format
        sheet.cell(row=total_paredes_row, column=13, value=0).number_format = accounting_format

    subtotal_row = total_paredes_row + 2
    sheet.cell(row=subtotal_row, column=1, value='SUBTOTAL FORROS + PAREDES e REVESTIMENTOS').font = bold_font
    sheet.merge_cells(start_row=subtotal_row, start_column=1, end_row=subtotal_row, end_column=4)
    sheet.cell(row=subtotal_row, column=6, value=f"=ROUND(F{total_forro_row}+F{total_paredes_row},2)").number_format = accounting_format
    sheet.cell(row=subtotal_row, column=13, value=f"=ROUND(M{total_forro_row}+M{total_paredes_row},2)").number_format = accounting_format

    sheet.row_dimensions[1].hidden = True
    sheet.sheet_view.view = 'pageBreakPreview'
'''
text = text[:start] + new_func + text[end:]
path.write_text(text, encoding='utf-8')
print('write_aditivos_distratado_sheet replaced in 1.3')
