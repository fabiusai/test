
import xlsxwriter
import pandas as pd

def genera_excel_format(df, output_stream):
    workbook = xlsxwriter.Workbook(output_stream, {'in_memory': True})
    worksheet = workbook.add_worksheet("Post Editoriali")

    blue = '#003DA5'
    light_gray = '#D3D3D3'

    # Definizione formati
    label_format = workbook.add_format({'bold': True, 'font_color': blue, 'align': 'center', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': blue})
    label_left = workbook.add_format({'bold': True, 'font_color': blue, 'align': 'left', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': blue})
    header_format = workbook.add_format({'bold': True, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': light_gray})
    cell_center = workbook.add_format({'font_color': 'black', 'align': 'center', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': light_gray})
    cell_left = workbook.add_format({'font_color': 'black', 'align': 'left', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': light_gray})
    bold_blue = workbook.add_format({'bold': True, 'font_color': blue, 'align': 'center', 'valign': 'vcenter', 'bottom': 1, 'bottom_color': light_gray})

    worksheet.set_column("A:A", 82.5)
    worksheet.set_column("B:B", 13.83)
    worksheet.set_column("C:G", 6.83)
    worksheet.set_column("H:H", 7.5)

    row_pos = 0
    for i, row in df.iterrows():
        values = ["" if pd.isna(v) or v == 0 else v for v in row.tolist()]
        is_label = isinstance(values[0], str) and values[0] == values[0].upper() and pd.isna(row[1]) if len(row) > 1 else False
        is_header = values[0] == "Argomento"

        for col, val in enumerate(values):
            if is_header:
                worksheet.write(row_pos, col, val, header_format)
            elif is_label:
                fmt = label_left if col == 0 else label_format
                worksheet.write_blank(row_pos, col, None, fmt) if val == "" else worksheet.write(row_pos, col, val, fmt)
            else:
                fmt = cell_left if col == 0 else (bold_blue if col == 1 else cell_center)
                worksheet.write(row_pos, col, val, fmt)
        row_pos += 1

    workbook.close()
