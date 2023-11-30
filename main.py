import pandas as pd
import warnings
from tkinter import messagebox
from datetime import datetime
import os

warnings.filterwarnings('ignore')

link = r'C:\Users\ChristophRazek\Emea\06_Qualitymanagement - Dokumente\01_QS\04_MPS\MPS.xlsx'

#Zeitstempel Letzte Änderung!
c_time = os.path.getmtime(link)
dt_c = datetime.fromtimestamp(c_time)

#Kopieren von MPS File zu Laufwerk
bestellungen = pd.read_excel(link)
bestellungen[['PE14_MassProdRel', 'QTY','FIXPOSNR','BELEGART']] = bestellungen[['PE14_MassProdRel', 'QTY','FIXPOSNR','BELEGART']].fillna(0).astype('int64')
bestellungen.to_csv(r'L:\Q\MPS.csv', sep=';', index=False)

#Log File
with open(r'S:\EMEA\Kontrollabfragen\MPS.txt', 'w') as f:
    f.write(f'Last MPS copied at: {dt_c}')
    f.close()


messagebox.showinfo('Update Erfolgreich!', f'Das MPS Update mit letzter Änderung vom {dt_c} wurde erfolgreich durchgeführt.')




