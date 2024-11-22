import os
import wx
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

class TextRedirect:
    """Redirect print statements to a wx.TextCtrl."""
    def __init__(self, text_ctrl):
        self.text_ctrl = text_ctrl

    def write(self, message):
        self.text_ctrl.AppendText(message)

    def flush(self):
        pass  # We don't need to do anything here.

class UpdateCredentialsDialog(wx.Dialog):
    """Dialog for updating the login credentials."""
    def __init__(self, parent):
        super().__init__(parent, title="Update Credentials", size=(350, 200))

        # Set up the dialog layout
        vbox = wx.BoxSizer(wx.VERTICAL)

        # Username field
        hbox_username = wx.BoxSizer(wx.HORIZONTAL)
        lbl_username = wx.StaticText(self, label="Username:")
        self.txt_username = wx.TextCtrl(self)
        hbox_username.Add(lbl_username, flag=wx.RIGHT, border=8)
        hbox_username.Add(self.txt_username, proportion=1)

        # Password field
        hbox_password = wx.BoxSizer(wx.HORIZONTAL)
        lbl_password = wx.StaticText(self, label="Password:")
        self.txt_password = wx.TextCtrl(self, style=wx.TE_PASSWORD)
        hbox_password.Add(lbl_password, flag=wx.RIGHT, border=8)
        hbox_password.Add(self.txt_password, proportion=1)

        # Load current values into the text fields
        self.load_credentials()

        # Buttons for Save and Cancel
        hbox_buttons = wx.BoxSizer(wx.HORIZONTAL)
        btn_save = wx.Button(self, label="Save")
        btn_cancel = wx.Button(self, label="Cancel")
        hbox_buttons.Add(btn_save)
        hbox_buttons.Add(btn_cancel, flag=wx.LEFT, border=10)

        # Bind button events
        btn_save.Bind(wx.EVT_BUTTON, self.on_save)
        btn_cancel.Bind(wx.EVT_BUTTON, self.on_cancel)

        # Add everything to the vbox
        vbox.Add(hbox_username, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(hbox_password, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(hbox_buttons, flag=wx.ALIGN_CENTER | wx.ALL, border=10)

        self.SetSizer(vbox)

    def load_credentials(self):
        """Load existing username and password from credentials.json file"""
        if os.path.exists(credentials_path):
            with open(credentials_path, 'r') as f:
                credentials = json.load(f)
                self.txt_username.SetValue(credentials.get("username", ""))
                self.txt_password.SetValue(credentials.get("password", ""))

    def on_save(self, event):
        """Save the new username and password to credentials.json file"""
        new_username = self.txt_username.GetValue()
        new_password = self.txt_password.GetValue()

        # Update JSON file
        with open(credentials_path, 'r+') as f:
            credentials = json.load(f)
            credentials["username"] = new_username
            credentials["password"] = new_password
            f.seek(0)
            json.dump(credentials, f, indent=4)
            f.truncate()  # Clear any remaining old data in the file after the new content

        wx.MessageBox("Credentials updated successfully.", "Success", wx.OK | wx.ICON_INFORMATION)
        self.EndModal(wx.ID_OK)

    def on_cancel(self, event):
        """Close the dialog without saving"""
        self.EndModal(wx.ID_CANCEL)

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, size=(1100, 650))
        
        self.panel = wx.Panel(self)
        self.sizer = wx.BoxSizer(wx.VERTICAL)

        # Create a horizontal sizer for the "Log In" button and the new button
        login_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.login_button = wx.Button(self.panel, label="Log In")
        self.ch_cred_button = wx.Button(self.panel, label="Change Credentials")

        # Add the login button and the new button to the horizontal sizer
        login_sizer.Add(self.login_button, flag=wx.RIGHT, border=10)
        login_sizer.Add(self.ch_cred_button, flag=wx.LEFT, border=10)

        # Add the horizontal sizer with the login buttons to the main vertical sizer
        self.sizer.Add(login_sizer, flag=wx.ALL, border=10)

        # Continue adding the other elements to the main vertical sizer                
        self.checkbox = wx.CheckBox(self.panel, label="Using saved accounts", pos=(50, 50))
        self.check_button = wx.Button(self.panel, label="Get Operations")              
        self.output_text_ctrl = wx.TextCtrl(self.panel, size=(300, 110), style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.run_button = wx.Button(self.panel, label="View Operations")

        self.sizer.Add(self.checkbox, flag=wx.ALL, border=10)
        self.sizer.Add(self.check_button, flag=wx.ALL, border=10)                
        self.sizer.Add(self.output_text_ctrl, flag=wx.ALL | wx.EXPAND, border=10)
        self.sizer.Add(self.run_button, flag=wx.ALL, border=10)

        self.panel.SetSizer(self.sizer)
        
        #self.check_button.Disable()
        self.checkbox.SetValue(True)

        # Bind events
        self.login_button.Bind(wx.EVT_BUTTON, self.on_update_token)
        self.ch_cred_button.Bind(wx.EVT_BUTTON, self.on_modify_cred_file)
        self.check_button.Bind(wx.EVT_BUTTON, self.fetch_operations)
        self.checkbox.Bind(wx.EVT_CHECKBOX, self.on_checkbox_toggled)
        self.run_button.Bind(wx.EVT_BUTTON, self.run_function)
        
        self.df = None  # Dataframe placeholder
        
        # Redirect print statements to output_text_ctrl
        sys.stdout = TextRedirect(self.output_text_ctrl)

        self.Show()
    
    def on_checkbox_toggled(self, event):
        self.check_button.Enable(self.checkbox.IsChecked())
    

    def on_modify_cred_file(self, event):
        """Allows to change login credentials"""
        dialog = UpdateCredentialsDialog(self)
        dialog.ShowModal()
        dialog.Destroy()
    
    def on_update_token(self, event):
        with open(credentials_path, 'r') as file:
            creds = json.load(file)
        username = creds.get('username')
        password = creds.get('password')
        #print(creds)

        """Run API login and update token"""
        base_login_url = 'https://conosur.aunesa.com/Irmo/api/login' 
        loginheaders = {
            "Content-Type": "application/json"
        }
        loginbody = {
                "clientId": "conosur",   
                "username": username,  
                "password": password
            }
        response = requests.post(base_login_url, json=loginbody, headers=loginheaders)

        if response.status_code == 200 or response.status_code == 201:
            with open(authkey_path, "w") as file:
                json.dump(response.json(), file)
            print("Nuevo token guardado en authkey.json. El mismo tendrá validez por las próximas 24hs")                         
        else:
            print(f"Request failed with status code {response.status_code}: {response.text}")

    def run_function(self, event):
        print('This will show grids for the files it created')

    def fetch_operations(self, event):
        file_path = os.path.join(app_dir, "listofaccountstoread.txt")
        with open(authkey_path, "r") as file:
            authkey = json.load(file)
        token = list(authkey.values())[1]
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}"
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
                print(f"Error recuperando operaciones para la cuenta {account}: {e}")
        final_json = json.dumps(responses, indent=4)
        current_date = datetime.now().strftime("%Y-%m-%d")
        file_name = f"InformedeOperaciones_{current_date}.json"
        with open(file_name, "w", encoding="utf-8") as json_file:
            json_file.write(final_json)
        print(f"Data de operaciones guardada en {file_name}")

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
        json_file_path = os.path.join(app_dir, f'InformedeOperaciones_{today}.json')   #f"C:\\Users\\sebastian.alfonso\\InformedeOperaciones_{today}.json"

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
        df3 = pd.DataFrame()    # DF3 WILL PROCESS ROFEX OPERATIONS

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
            print('Fetching BYMA operations...')
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
            print('Fetching ROFEX operations...')
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

if __name__ == "__main__":
    app = wx.App(False)
    frame = MyFrame(None, "[C.O.N.O.] Cargador Onírico de Neonumismáticos Onerosos")
    app.MainLoop()   
