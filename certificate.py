from docxtpl import DocxTemplate
import pandas as pd
import os
import time

j=0
ans=pd.read_csv('data.csv')
certificado=(len(ans))
def mkw(n):
    tpl=DocxTemplate("temp.docx")
    context={i+1:{"name":ans['name'][i],"tutor":ans['tutor'][i]}for i in range (n)}
    tpl.render(context[n])
    tpl.save("./gerados/" + "%s - "%str(n)+ans['name'][j]+".docx")
for i in range(1,certificado+1):
    mkw(i)
    j+=1

path = './gerados/'
 
l_files = os.listdir(path)
 
for file in l_files:
   
    file_path = f'{path}\\{file}'
 
    if os.path.isfile(file_path):
        try:
            os.startfile(file_path, 'print')
            print(f'Printing {file}')
            time.sleep(5)
        except:
            print(f'ALERT: {file} could not be printed! Please check\
            the associated softwares, or the file type.')
    else:
        print(f'ALERT: {file} is not a file, so can not be printed!')

