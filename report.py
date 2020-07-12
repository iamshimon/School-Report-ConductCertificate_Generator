import pandas as pd
import re
from docxtpl import DocxTemplate
import os

#df=pd.read_excel(r'./Students List 2019-20.xlsx')
xls = pd.ExcelFile(r'./Students List 2019-20.xlsx')
doc = DocxTemplate('./report_template.docx')

for i in range(0,32):
    df1=pd.read_excel(xls, i) #limit 31

    gradex=df1['Unnamed: 3'].dropna().values[0]
    gradex=gradex.replace('.', '')
    gradex=gradex.replace('-', '')
    gradex=gradex.strip()
    gradex=gradex.rstrip(' A')
    gradex=gradex.rstrip(' B')
    gradex=gradex.rstrip(' C')
    gradex=gradex.rstrip(' D')
    gradex=gradex.rstrip(' E')
    while '  ' in gradex:
        gradex = gradex.replace('  ', ' ')

    if gradex[0]=='K' and gradex[3]=='1':
        gradey ='KG 2'
    elif gradex[0]=='K' and gradex[3]=='2':
        gradey = 'GRADE 1'
    elif len(gradex) > 7:
        gradey ="".join(['GRADE ',str( int(gradex[6])*10 + int(gradex[7]) +1)])
    else:
        gradey = "".join(['GRADE ',str(int(gradex[6])+1)])

    print("GRADE IS", gradex)
    print("Promoted Grade is ", gradey)
    print('')

    if not os.path.exists('./'+gradex):
        os.makedirs('./'+gradex)
    for i in df1['Unnamed: 1'].dropna():
        if( not (i=='STUDENTS NAME') ):
            studentname= i
            cprno = df1.loc[df1['Unnamed: 1'] == studentname, 'Unnamed: 3'].iloc[0]
            dob = df1.loc[df1['Unnamed: 1'] == studentname, 'Unnamed: 5'].iloc[0]
            nationality = df1.loc[df1['Unnamed: 1'] == studentname, 'Unnamed: 4'].iloc[0]
            


            print("Student Name: ",studentname)
            print("CPR No: ", cprno)
            print("DOB: ", dob)
            print("Nationality: ",nationality)
            context = { 'studentname' : studentname, 'cprno' : cprno, "dob" : dob, "nationality" : nationality, "gradex" : gradex, "gradey" : gradey }
            doc.render(context)
            doc.save('./'+gradex+'/'+studentname+'.docx')

    print("")

print(" ALL REPORTS GENERATED SUCCESSFULLY")



    #print(df1)

