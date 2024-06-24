import pandas as pd
import warnings
from tkinter import messagebox
from datetime import datetime
import os

warnings.filterwarnings('ignore')
user = os.getlogin()
link = rf'C:\Users\{user}\Emea\06_Qualitymanagement - Dokumente\01_QS\04_MPS\MPS.xlsx'

#Zeitstempel Letzte Änderung!
c_time = os.path.getmtime(link)
dt_c = datetime.fromtimestamp(c_time)

#Kopieren von MPS File zu Laufwerk
bestellungen = pd.read_excel(link)

bestellungen[['BELEGNR','PE14_MassProdRel', 'QTY','FIXPOSNR','BELEGART']] = bestellungen[['BELEGNR','PE14_MassProdRel', 'QTY','FIXPOSNR','BELEGART']].fillna(0).astype('int64')
bestellungen.to_csv(r'L:\Q\MPS.csv', sep=';', index=False)

#Log File
with open(rf'C:\Users\{user}\Emea\05_SupplyChainMgt - Dokumente\27_Datenuebertragung\EMEA\Kontrollabfragen\MPS.txt', 'a') as f:
    f.write(f'\nLast MPS copied at: {dt_c}')
    f.close()


#messagebox.showinfo('Update Erfolgreich!', f'Das MPS Update mit letzter Änderung vom {dt_c} wurde erfolgreich durchgeführt.')




