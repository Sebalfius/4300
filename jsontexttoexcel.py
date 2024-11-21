import os
import sys
import json
import requests
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

if getattr(sys, 'frozen', False):
    # Cuando se ejecuta el programa como un .exe
    app_dir = os.path.dirname(sys.executable)
else:
    # Cuando se ejecuta el programa como un script .py
    app_dir = os.path.dirname(os.path.abspath(__file__))

credentials_path = os.path.join(app_dir, "credentialsforapi.json")
authkey_path = os.path.join(app_dir, "authkey.json")

file_path = f"C:\\Users\\sebastian.alfonso\\listofaccountstoread.txt"
headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJjb25vc3VyICxhcGlfc2ViYXN0aWFuICwyMDAuNjEuMTc4LjEyNiIsImV4cCI6MTczMjMwNjk2M30.VimqLXzZxe95f3CICVooQ0HiwUmS3hvDFF37X1Z_rfQysg3pt5fn_5xHsbRxc4Udh8DgFZATuvVt-UJM-QjD7A"
}
date_for_url = datetime.now().strftime("%d/%m/%Y")
with open(file_path, "r") as file:
    accounts = file.read().splitlines()
responses = []
for account in accounts:
    uri = f"https://conosur.aunesa.com/Irmo/api/operaciones/informes?cuenta={account}&fechaDesde={date_for_url}&fechaHasta={date_for_url}"    
    try:
        response = requests.get(uri, headers=headers)
        response.raise_for_status()  # Raise an exception for HTTP errors
        responses.append(response.json())
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data for account {account}: {e}")
final_json = json.dumps(responses, indent=4)
current_date = datetime.now().strftime("%Y-%m-%d")
file_name = f"InformedeOperaciones_{current_date}.json"
with open(file_name, "w", encoding="utf-8") as json_file:
    json_file.write(final_json)
print(f"Data saved to {file_name}")

'''
try:
    result = subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_script_path],
        capture_output=True,
        text=True
    )    
    print("Output:")
    print(result.stdout)
    
    if result.stderr:
        print("Errors:")
        print(result.stderr)
except Exception as e:
    print(f"An error occurred: {e}")
'''

today = datetime.now().strftime('%Y-%m-%d')
json_file_path = f"C:\\Users\\sebastian.alfonso\\InformedeOperaciones_{today}.json"
#json_file_path = f"C:\\Users\\sebastian.alfonso\\InformedeOperaciones_2024-11-20.json"
with open(json_file_path, encoding='utf-8') as file:
    raw_data = json.load(file)
data = raw_data[0]
df = pd.DataFrame(data)     # DF IS THE RAW JSON RECEIVED FROM AUNE'S API IN DATAFRAME FORM
print("Columns:", df.columns)
print(df.head())
df['instrumento'] = df['instrumento'].astype(str).str.strip()
dfb = df[df['instrumento'].str.match(r'^\[DLR\d{6}\]$', na=False)] # DFB IS JUST ROFEX OPERATIONS
df = df[~df['instrumento'].str.match(r'^\[DLR\d{6}\]$', na=False)]
df2 = pd.DataFrame()    # DF2 RESULTS IN THE PROCESSED LIST OF "BYMA" OPERATIONS.
df3 = pd.DataFrame()    # DF3 WILL PROCESSED PROCESSED ROFEX OPERATIONS

def clean_currency(entry):
    if isinstance(entry, list) and len(entry) > 0 and isinstance(entry[0], dict):
        moneda = entry[0].get('Moneda', '')
        monto = entry[0].get('Monto', '')
        
        if not monto or monto in ['0', '0.0', '']:
            return 0
        return f"{moneda} {monto}"
    return 0

def isthiscompra(row):
    if pd.isna(row):
        return None
    if "Concurrencia Contado - Compra" in row:
        return 'Compra'
    elif "SENEBI Contado - Compra" in row:
        return 'Compra'
    elif "Futuros Financieros - Compra" in row:
        return 'Compra'
    elif "Futuros Financieros - Venta" in row:
        return 'Venta'
    elif "SENEBI Contado - Venta" in row:
        return 'Venta'
    elif "Concurrencia Contado - Venta" in row:
        return 'Venta'
    return None

def whattasa(row):
    if "ARS 24hs" in row:
        return '1'
    elif "USD 24hs" in row:
        return '1'
    elif "ARS 24hs NG" in row:
        return '1'
    elif "USD 24hs NG" in row:
        return '1'
    elif "USDC 24hs NG" in row:
        return '1'
    elif 'ARS Inm' in row:
        return '0'
    elif 'USD Inm' in row:
        return '0'
    elif 'USD' in row:
        return '0'
    elif 'ARS' in row:
        return '0'
    
def whattasa2(row):
    if "Futuros Financieros - Venta" in row:
        return '1'
    elif "Futuros Financieros - Compra" in row:
        return '1'
    else:
        return '1'

def update_mercado(df, df2):
    condition = df['tipoOperacion'].isin(["SENEBI Contado - Compra", "SENEBI Contado - Venta"])
    df2['Mercado'] = condition.map({True: "BYMA SENEBI", False: "BYMA"})
    return df2

def extract_monto(value):
    if isinstance(value, list) and len(value) > 0 and isinstance(value[0], dict):
        monto = value[0].get('Monto', '0')
        if isinstance(monto, str):
            monto = monto.replace(',', '.')
        return float(monto)
    return 0

temp_results = df['neto'].apply(clean_currency)
df['Neto'] = temp_results
df['Monedas'] = df['Neto'].apply(lambda x: x.split()[0] if isinstance(x, str) and len(x.split()) > 1 else '')

def pesodolar(row):
    if "USD" in row:
        return 'U$D'
    elif "ARS" in row:
        return '$'

# PROCESSING BYMAS
if not df.empty:
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
    df2['Precio'] = df2['Precio'].round(2)
    df2['Importe'] = df['bruto']
    df2['Importe'] = df2['Importe'].abs()
    df2['Moneda'] = df['Monedas'].apply(pesodolar)
    df2 = update_mercado(df, df2)
    def adjust_importe(df2):
        df2.loc[df2['Mercado'] == "BYMA SENEBI", 'Importe'] *= 100
        return df2
    df2 = adjust_importe(df2)
    df2['GASTOS'] = '0'
    excel_file_path2 = f'operacionesbyma_{today}.xlsx' #BYMA
    df2.to_excel(excel_file_path2, index=False)
    print(df2)
    print(f"BYMA operationas saved to {excel_file_path2}.xlsxe")
    wb = load_workbook(excel_file_path2)
    ws = wb.active
    green_fill = PatternFill(start_color='92D050', end_color='92D050', fill_type='solid')
    for cell in ws[1]:
        cell.fill = green_fill
    wb.save(excel_file_path2)
    print(f"Excel file saved with styled headers at {excel_file_path2}")
    print('...')
else:
    print("No BYMA operations. Skipping processing.")

excel_file_path1 = f'informeopetraducido_{today}.xlsx' #JSON traducido
df.to_excel(excel_file_path1, index=False)

if not dfb.empty and len(dfb) > 0:
    temp_results_b = dfb['neto'].apply(clean_currency)
    print('neto')
    print(dfb['neto'])
    dfb['Neto'] = temp_results_b
    dfb['Monedas'] = dfb['Neto'].apply(lambda x: x.split()[0] if isinstance(x, str) and len(x.split()) > 1 else '')
    dfb['concertacion'] = pd.to_datetime(dfb['concertacion'], errors='coerce')
    dfb['liquidacion'] = pd.to_datetime(dfb['liquidacion'], errors='coerce') 
    df3['Fecha Concer.'] = dfb['concertacion'].dt.strftime('%d/%m/%Y')
    df3['F. Liqui'] = dfb['liquidacion'].dt.strftime('%d/%m/%Y')
    df3['FONDO'] = '32STGESTION VII'
    df3['OPERACION'] = dfb['tipoOperacion'].apply(isthiscompra)
    df3['Especie'] = dfb['instrumento']
    df3['Contraparte'] = 'Cono Sur Inversiones S.A.'
    df3['Plazo Liq.'] = dfb['tipoOperacion'].apply(whattasa2)
    df3['TASA'] = ''
    df3['Valor Nominal'] = dfb['cantidadTotal']
    df3['Precio'] = dfb['precioPromedio'].abs()
    df3['Precio'] = df3['Precio'].round(2)
    df3['Importe'] = dfb['neto'].apply(extract_monto)
    df3['Importe'] = pd.to_numeric(df3['Importe'], errors='coerce').fillna(0)
    df3['Importe'] = df3['Importe'].abs()
    df3['Moneda'] = dfb['Monedas'].apply(pesodolar)
    df3['Mercado'] = 'ROFEX'
    dfb['Gasto'] = dfb['gastos'].apply(extract_monto)
    dfb['Impuesto'] = dfb['impuestos'].apply(extract_monto)
    dfb['Gastos'] = pd.to_numeric(dfb['Gasto'], errors='coerce').fillna(0)
    dfb['Impuestos'] = pd.to_numeric(dfb['Impuesto'], errors='coerce').fillna(0)
    df3['GASTOS'] = dfb['Gastos'] + dfb['Impuestos']
    excel_file_path3 = f'operacionesrofex_{today}.xlsx' #ROFEX
    df3.to_excel(excel_file_path3, index=False)
    print(df3)
    print(f"ROFEX operationas saved to {excel_file_path3}")
    wb = load_workbook(excel_file_path3)
    ws = wb.active
    bourdeaux_fill = PatternFill(start_color='702222', end_color='702222', fill_type='solid')
    for cell in ws[1]:
        cell.fill = bourdeaux_fill
    wb.save(excel_file_path3)
    print(f"Excel file saved with styled headers at {excel_file_path3}")
else:
    print("No ROFEX operations. Skipping processing.")
