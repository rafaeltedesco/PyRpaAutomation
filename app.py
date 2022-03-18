from zipfile import ZipFile

from pyoutlookdispatcher import Mail, Outlook

import os
import pandas as pd

path = os.path.join(os.getcwd(), 'files')

save_path = os.path.join(os.getcwd(), 'results')
files_v1 = [os.path.join(path, file) for file in os.listdir(path)]

# modo for tradicional
# files_v2 = []
# for file in os.listdir(path):
#     files_v2.append(os.path.join(path, file))

df = pd.DataFrame()

for file in files_v1:
    df = pd.concat([df, pd.read_excel(file)], ignore_index=1)

df_grouped = df.groupby(['Categoria'])[['Valor']].sum(
).sort_values("Valor", ascending=False)

os.makedirs(save_path, exist_ok=True)

fullsavepath = os.path.join(save_path, 'acumulado.xlsx')

df_grouped.to_excel(fullsavepath)

# Criando arquivo zip

with ZipFile(os.path.join(save_path, 'results.zip'), 'w') as customZip:
    customZip.write(fullsavepath, arcname='acumulado.xlsx')

# Apagando arquivo acumulado temporário

os.unlink(os.path.join(save_path, 'acumulado.xlsx'))

# Disparando email

outlook = Outlook()

mail = Mail(
    Subject="Relatório Mensal Acumulado",
    To="example@example.com",
    HTMLBody="<h1>Aqui segue o relatório mensal</h1>",
    Attachments=os.path.join(save_path, 'results.zip')
)

outlook.send(mail)
