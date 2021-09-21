import pandas as pd
df = pd.read_csv('E:/Downloads/avail.cgi',
                 sep=',').replace('"', '', regex=True)
writer = pd.ExcelWriter('avail.xlsx')
df.to_excel(writer, index=False)
writer.save('E:\Downloads\Documents\Paladin_Backup_Files')
