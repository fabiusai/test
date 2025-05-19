
import xlsxwriter
import pandas as pd

def genera_excel_format(df, output_stream):
    workbook = xlsxwriter.Workbook(output_stream, {'in_memory': True})
    worksheet = workbook.add_worksheet("Post Editoriali")

    blue = '#003DA5'
    light_gray = '#D3D3D3'

    label_format = workbook.add_format({
        'bold': True, 'font_color': blue, 'align': 'center', 'valign': 'vcenter', 
        'bottom': 1, 'bottom_color': blue
    })
    label_left = workbook.add_format({
        'bold': True, 'font_color': blue, 'align': 'left', 'valign': 'vcenter',
        'bottom': 1, 'bottom_color': blue
    })
    header_format = workbook.add_format({
        'bold': True, 'font_color': 'black', 'align': 'center', 'valign': 'vcenter',
        'bottom': 1, 'bottom_color': light_gray
    })
    cell_center = workbook.add_format({
        'font_color': 'black', 'align': 'center', 'valign': 'vcenter',
        'bottom': 1, 'bottom_color': light_gray
    })
    cell_left = workbook.add_format({
        'font_color': 'black', 'align': 'left', 'valign': 'vcenter',
        'bottom': 1, 'bottom_color': light_gray
    })
    bold_blue = workbook.add_format({
        'bold': True, 'font_color': blue, 'align': 'center', 'valign': 'vcenter',
        'bottom': 1, 'bottom_color': light_gray
    })

    worksheet.set_column("A:A", 82.5)
    worksheet.set_column("B:B", 13.83)
    worksheet.set_column("C:G", 6.83)
    worksheet.set_column("H:H", 7.5)

    row_pos = 0
    col_names = df.columns.tolist()
    worksheet.write_row(row_pos, 0, col_names, header_format)
    row_pos += 1

    current_label = None
    buffer_rows = []

    def write_label(label, rows):
        nonlocal row_pos
        if not rows:
            return
        temp_df = pd.DataFrame(rows, columns=col_names)
        numeric_cols = temp_df.select_dtypes(include='number').columns
        totals = temp_df[numeric_cols].sum().to_dict()
        total_row = [label, ""] + ["" if totals.get(col) == 0 else int(totals.get(col)) for col in col_names[2:]]
        for col_idx, val in enumerate(total_row):
            fmt = label_left if col_idx == 0 else label_format
            worksheet.write(row_pos, col_idx, val, fmt)
        row_pos += 1

        for row in rows:
            for col_idx, val in enumerate(row):
                if col_idx == 0:
                    fmt = cell_left
                elif col_idx == 1:
                    fmt = bold_blue
                else:
                    fmt = cell_center
                display_val = "" if val == 0 or pd.isna(val) else val
                worksheet.write(row_pos, col_idx, display_val, fmt)
            row_pos += 1

    for _, row in df.iterrows():
        if pd.isna(row[1]) and isinstance(row[0], str) and row[0] == row[0].upper():
            if current_label:
                write_label(current_label, buffer_rows)
                buffer_rows = []
            current_label = row[0]
        else:
            buffer_rows.append(row.tolist())

    if current_label:
        write_label(current_label, buffer_rows)

    workbook.close()
