import functions as fn
import glob

file_paths = glob.glob("Excel_Files/*.xlsx")

for file_path in file_paths:
    invoice_dict = fn.get_excel_file(file_path)
    if invoice_dict:
        invoice_file = fn.generate_invoice(invoice_dict)
