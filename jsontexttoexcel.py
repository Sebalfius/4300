import ast
import json
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

today = datetime.now().strftime('%Y-%m-%d')
json_file_path = f"C:\\Users\\sebastian.alfonso\\InformedeOperaciones_{today}.json"
with open(json_file_path, encoding='utf-16') as file:
    data = json.load(file)

df = pd.DataFrame(data)
df2 = pd.DataFrame()

def clean_currency(entry):
    if isinstance(entry, list) and len(entry) > 0 and isinstance(entry[0], dict):
        # Extract Moneda and Monto
        moneda = entry[0].get('Moneda', '')
        monto = entry[0].get('Monto', '')

        # Check if Monto is empty or 0
        if not monto or monto in ['0', '0.0', '']:
            return 0
        return f"{moneda} {monto}"
    return 0

def isthiscompra(row):
    if "Concurrencia Contado - Compra" in row:
        return 'Compra'
    elif "SENEBI Contado - Compra" in row:
        return 'Compra'
    elif "Futuros Financieros - Compra" in row:
        return 'Compra'
    elif "SENEBI Contado - Venta" in row:
        return 'Venta'
    elif "Concurrencia Contado - Venta" in row:
        return 'Venta'
    elif "Futuros Financieros - Venta" in row:
        return 'Venta'

def whattasa(row):
    if "ARS 24hs" in row:
        return '1'
    elif "USD 24hs" in row:
        return '1'
    if "ARS 24hs NG" in row:
        return '1'
    elif "USD 24hs NG" in row:
        return '1'
    elif 'ARS Inm' in row:
        return '0'
    elif 'USD Inm' in row:
        return '0'
    elif 'USD' in row:
        return '0'
    elif 'ARS' in row:
        return '0'

def update_mercado(df, df2):
    condition = df['tipoOperacion'].isin(["SENEBI Contado - Compra", "SENEBI Contado - Venta"])
    df2['Mercado'] = condition.map({True: "BYMA SENEBI", False: "BYMA"})
    return df2

temp_results = df['neto'].apply(clean_currency)
#arancelos = df['aranceles'].apply(clean_currency)

df['Neto'] = temp_results
df['Monedas'] = df['Neto'].apply(lambda x: x.split()[0] if isinstance(x, str) and len(x.split()) > 1 else '')

def pesodolar(row):
    if "USD" in row:
        return 'U$D'
    elif "ARS" in row:
        return '$'

df['concertacion'] = pd.to_datetime(df['concertacion'], errors='coerce')
df['liquidacion'] = pd.to_datetime(df['liquidacion'], errors='coerce')
                                   
df2['Fecha Concer.'] = df['concertacion'].dt.strftime('%d/%m/%Y')
df2['F. Liqui'] = df['liquidacion'].dt.strftime('%d/%m/%Y')
df2['FONDO'] = '32STGESTION VII'
df2['OPERACION'] = df['tipoOperacion'].apply(isthiscompra)
df2['Especie'] = df['instrumento'].str.extract(r'\] (\w+)')
df2['Contraparte'] = 'Cono Sur Inversiones S.A.'
df2['Plazo Liq.'] = df['condiciones'].apply(whattasa)
df2['TASA'] = ''
df2['Valor Nominal'] = df['cantidadTotal']
df2['Precio'] = df['precioPromedio'].abs()
df2['Importe'] = df['bruto']
df2['Precio'] = df2['Precio'].round(2)
df2['Importe'] = df2['Importe'].abs()

df2['Moneda'] = df['Monedas'].apply(pesodolar)
df2 = update_mercado(df, df2)
df2['GASTOS'] = '0'


excel_file_path1 = f'informeopetraducido_{today}.xlsx'
df.to_excel(excel_file_path1, index=False)

excel_file_path2 = f'informeprocesado_{today}.xlsx'
df2.to_excel(excel_file_path2, index=False)

print(f"DataFrames saved to {excel_file_path1} and {excel_file_path2}")

wb = load_workbook(excel_file_path2)
ws = wb.active
green_fill = PatternFill(start_color='92D050', end_color='92D050', fill_type='solid')
for cell in ws[1]:
    cell.fill = green_fill
wb.save(excel_file_path2)
print(f"Excel file saved with styled headers at {excel_file_path2}")
