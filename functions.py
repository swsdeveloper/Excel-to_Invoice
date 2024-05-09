import pandas as pd
import pypdf
import os


def get_filepath(filename: str) -> str:
    # Get directory of the currently executing script (e.g., main.py)
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # Use this directory as the base for our file paths
    filepath = os.path.join(script_directory, filename)

    return filepath


def get_excel_file(excel_filename: str) -> dict:
    """
    Read an Excel file and convert its data into a dictionary. The Excel file represents a
    single invoice where each record is an invoice line item.

    :param excel_filename: name of the Excel file. The first part of that name is the Invoice Number.

    :return: a dictionary as {'invoice_number': line_item}
        where each line_item is a dictionary with this format:
            'product_id': str
            'product_name': str
            'amount_purchased': float
            'price_per_unit': float
            'total_price': float
    """
    excel_filepath = get_filepath(excel_filename)
    df = pd.read_excel(excel_filepath, header=0, engine='openpyxl')
    invoice_dict = {}
    for index, record in df.iterrows():
        line_item = {}
        product_id = line_item['product_id']  # str
        product_name = line_item['product_name']  # str
        amount_purchased = line_item['amount_purchased']  # float
        price_per_unit = line_item['price_per_unit']  # float
        total_price = line_item['total_price']  # float
        invoice_dict = {'invoice_number': line_item}
    return invoice_dict


def generate_invoice(invoice_dict: {}) -> str:
    """Generate PDF Invoice from invoice dictionary created by get_excel_file().
    See details, above.

    :param invoice_dict: a dictionary created by get_excel_file().

    :return: name of newly created PDF file. The first part of that name is the Invoice Number.
    """
