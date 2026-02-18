import os
import sys

import pandas as pd
from openpyxl import load_workbook
from tqdm import tqdm
import win32com.client as win32

# File paths
kandela_path = 'Kandela_2026.xlsx'
output_folder = 'output'
os.makedirs(output_folder, exist_ok=True)

# Load source data
data = pd.read_excel(kandela_path)


def create_excel_app():
    """Create isolated Excel COM app and apply silent settings when possible."""
    excel_app = win32.DispatchEx('Excel.Application')
    for prop, value in (('Visible', False), ('DisplayAlerts', False), ('ScreenUpdating', False)):
        try:
            setattr(excel_app, prop, value)
        except AttributeError:
            # Some environments expose these COM properties as read-only.
            pass
    return excel_app


def parse_invoice_date(raw_date):
    """Parse invoice date with day-first to match dd/mm/yyyy source values."""
    return pd.to_datetime(raw_date, dayfirst=True, errors='coerce')


# Start Excel COM object once for performance
excel = create_excel_app()

try:
    # Process rows with terminal progress bar
    for index, row in tqdm(
        data.iterrows(),
        total=len(data),
        desc='PDF Fatura Olusturuluyor',
        file=sys.stdout,
        dynamic_ncols=True,
    ):

        # Basic validations
        if pd.isna(row.get('DATE')):
            print(f"Atlaniyor (satir {index + 1}): DATE eksik.")
            continue
        if pd.isna(row.get('Name')):
            print(f"Atlaniyor (satir {index + 1}): Name eksik.")
            continue
        parsed_date = parse_invoice_date(row.get('DATE'))
        if pd.isna(parsed_date):
            print(f"Atlaniyor (satir {index + 1}): DATE formati gecersiz.")
            continue

        # Pick template by currency
        if pd.notna(row.get('TRY')):
            invoice_template_path = 'invoice.xlsx'
            currency_value = row['TRY']
        elif pd.notna(row.get('Pound')):
            invoice_template_path = 'Invoice_Pound.xlsx'
            currency_value = row['Pound']
        elif pd.notna(row.get('Euro')):
            invoice_template_path = 'Invoice_Euro.xlsx'
            currency_value = row['Euro']
        elif pd.notna(row.get('usd')):
            invoice_template_path = 'Invoice_Dolar.xlsx'
            currency_value = row['usd']
        else:
            print(f"Atlaniyor (satir {index + 1}): Doviz degeri bulunamadi.")
            continue

        # Load template
        workbook = load_workbook(invoice_template_path)
        sheet = workbook.active

        # Write row data to template
        sheet['D4'].value = row['DATE']
        sheet['D7'].value = row['Inv No']
        sheet['A9'].value = row['Name']
        sheet['C15'].value = row.get('Hours', '')
        sheet['D15'].value = currency_value
        sheet['A15'].value = row.iloc[8]  # Column I from source sheet

        # Build output path
        date_str = parsed_date.strftime('%d.%m')
        month_folder = parsed_date.strftime('%m')
        subfolder_path = os.path.join(output_folder, month_folder)
        os.makedirs(subfolder_path, exist_ok=True)

        safe_name = ''.join(c for c in str(row['Name']) if c.isalnum() or c in ' _-').strip()
        filename = f'{safe_name} - {date_str}.pdf'
        output_path = os.path.join(subfolder_path, filename)

        # Save temporary xlsx for PDF conversion
        temp_invoice_path = 'temporary_invoice.xlsx'
        workbook.save(temp_invoice_path)
        workbook.close()

        # Convert to PDF
        try:
            wb = excel.Workbooks.Open(os.path.abspath(temp_invoice_path), ReadOnly=True)
            wb.ExportAsFixedFormat(0, os.path.abspath(output_path))  # 0 = PDF
            wb.Close(False)
        except Exception as e:
            print(f"Hata (satir {index + 1}, {safe_name}): {e}")
        finally:
            if os.path.exists(temp_invoice_path):
                os.remove(temp_invoice_path)

finally:
    # Always close Excel app
    excel.Quit()

print("PDF uretimi tamamlandi. Dosyalar 'output' klasorunde, aya gore siniflandirildi.")
