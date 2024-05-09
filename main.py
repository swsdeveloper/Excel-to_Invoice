import pandas as pd
import pypdf


def get_excel_file(filename: str) -> dict:
    """
    Read an Excel file and convert its data into a dictionary. The Excel file represents a
    single invoice where each record is an invoice line item.

    :param filename: Name of the Excel file. The first part of that name is the Invoice Number.

    :return: a dictionary as {'invoice_number': line_item}
        where each line_item is a dictionary with this format:
            'product_id': str
            'product_name': str
            'amount_purchased': float
            'price_per_unit': float
            'total_price': float
    """
    get filepath
    open filepath as excel_file

    invoice_dict = {}
    read excel_file
    for each record in excel_file:
        line_item = {}
        product_id = line_item['product_id']  # str
        product_name = line_item['product_name']  # str
        amount_purchased = line_item['amount_purchased']  # float
        price_per_unit = line_item['price_per_unit']  # float
        total_price = line_item['total_price']  # float
        invoice_dict = {'invoice_number': line_item}
    return invoice_dict


def generate_invoice(dictionary: {}) -> str:
    """Generate PDF Invoice from dictionary .

    :param dictionary: a dictionary of the format {'invoice_number': item_dict}
        where each item_dict has this format:
            'product_id': str
            'product_name': str
            'amount_purchased': float
            'price_per_unit': float
            'total_price': float
    :return: name of newly created PDF file
    """


file_list = ['10001-2023.1.18.xlsx', '10002-2023.1.18.xlsx', '10003-2023.1.18.xlsx']

for file in file_list:
    excel_dict = get_excel_file(file)
    if excel_dict:
        invoice_file = generate_invoice(excel_dict)

