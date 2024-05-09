import functions as fn
import pypdf



file_list = ['10001-2023.1.18.xlsx', '10002-2023.1.18.xlsx', '10003-2023.1.18.xlsx']

for filename in file_list:
    excel_dict = fn.get_excel_file(filename)
    if excel_dict:
        invoice_file = generate_invoice(excel_dict)
