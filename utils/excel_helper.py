import openpyxl


def get_barcode_data_from_xsl(file_name):
    workbook = openpyxl.load_workbook(file_name)
    worksheet = workbook.active

    for i in range(0, worksheet.max_row):
        data = []
        for col in worksheet.iter_cols(1, 2):
            if col[i].value is not None:
                data.append(col[i].value)
        if len(data) == 2:
            yield data