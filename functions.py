import os
from pathlib import Path
import pandas as pd
from fpdf import FPDF

# To run this, I had to install pandas Excel dependency: openpyxl


def get_filepath(filename: str) -> str:
    """
    Return absolute path to specified filename (i.e., path from root folder to filename).

    :param filename: a file name or a relative path to a file name (relative to currently executing script)
    :return: a path-like object (i.e., absolute path to filename)
    """
    # Get directory of the currently executing script (e.g., main.py)
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # Use this directory as the base for our file paths
    filepath = os.path.join(script_directory, filename)

    return filepath


def get_excel_file(filepath: str) -> {}:
    """
    Use pandas to read an Excel file and return a Dictionary.

    :param filepath: relative path of the Excel file. (It is relative to the path to the currently
    executing script.) The first part of the filename is the Invoice Number.

    Note: filenames should be of the format: "invoiceNumber-invoiceDate.xlsx"

    :return: a dictionary as {'inv num': invoice_number, 'inv date': invoice_date, 'line items': line_items}
        line_items is a list of dictionaries, where each dict represents 1 line item and contains:
            'product_id': str
            'product_name': str
            'units_purchased': int
            'unit_price': int
            'total_price': int
    """
    # Method 1:
    # head, tail = os.path.split(filepath)
    # raw_filename = tail.removesuffix('.xlsx')

    # Method 2 (simpler):
    raw_filename = Path(filepath).stem

    inv_num, inv_date = raw_filename.split('-')

    excel_filepath = get_filepath(filepath)
    df = pd.read_excel(excel_filepath, sheet_name="Sheet 1", header=0, engine='openpyxl')

    line_items = []
    for _, record in df.iterrows():
        line_item_dict = {'product_id': record['product_id'],
                          'product_name': record['product_name'],
                          'units_purchased': record['amount_purchased'],
                          'unit_price': record['price_per_unit'],
                          'total_price': record['total_price']
                          }
        line_items.append(line_item_dict)

    invoice_dictionary = {'inv num': inv_num, 'inv date': inv_date, 'line items': line_items}

    return invoice_dictionary


def generate_invoice(inv_dict):
    """Generate PDF Invoice from an invoice dictionary. The dictionary was created by get_excel_file.

    :param inv_dict: see get_excel_file's docstring for this dictionary's format
    :return: filepath of newly created PDF file
    """
    pdf = FPDF(orientation='P', unit='mm', format='A4')

    """
    Product ID, Product Name, Amount, Price per Unit, Total Price - subhdr in box
    id, name, amt, price, total - in box
    grand total - in box

    The total due amount is 'grand total' dollars. - subhdr
    PythonHow, logo - subhdr

    """
    inv_num = inv_dict['inv num']
    inv_date = inv_dict['inv date']

    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=16)
    pdf.set_text_color(0, 0, 0)  # black (rgb)

    line1 = "Invoice nr. " + inv_num
    pdf.cell(w=0, h=8, txt=line1, align="L", ln=1, border=0)

    line2 = "Date " + inv_date
    pdf.cell(w=0, h=8, txt=line2, align="L", ln=1, border=0)

    pdf.set_font(family="Times", style="B", size=14)
    pdf.set_text_color(0, 0, 0)  # black (rgb)

    line_items = inv_dict['line items']

    # line3 =
    # pdf.cell(w=0, h=14, txt=str(line3), align="L", ln=1)

    pdf.set_font(family="Times", style="", size=12)
    pdf.set_text_color(80, 80, 80)  # grey

    for line_item in line_items:
        prod_id = line_item['product_id']
        prod_name = line_item['product_name']
        quantity = line_item['units_purchased']
        unit_price = line_item['unit_price']
        total_price = line_item['total_price']

        pdf.cell(w=30, h=10, txt=str(prod_id), align="L")
        pdf.cell(w=90, h=10, txt=prod_name, align="L")
        pdf.cell(w=20, h=10, txt=str(quantity), align="L")
        pdf.cell(w=30, h=10, txt=str(unit_price), align="L")
        pdf.cell(w=30, h=10, txt=str(total_price), align="L", ln=1)

    pdf.output(f"PDF_Invoices/{inv_num}-{inv_date}.pdf")

    return


if __name__ == '__main__':
    invoice_dict = get_excel_file('Excel_Files/10001-2023.1.18.xlsx')
    generate_invoice(invoice_dict)
