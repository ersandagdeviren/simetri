import pandas as pd
import xlsxwriter
import datetime
from datetime import date , time , datetime
today=date.today()
nowtime = datetime.now()
df=pd.read_excel(r'\\WIN-VCGK7KC2PST\Users\Administrator\Desktop\Uretim\source.xlsx')
df2=pd.read_excel(r'\\WIN-VCGK7KC2PST\Users\Administrator\Desktop\Uretim\input.xlsx')

formatted_date = "{}.{}.{}-{}{}{}".format(today.day, today.month, today.year, nowtime.hour, nowtime.minute, nowtime.second)
workbook = xlsxwriter.Workbook(r'\\WIN-VCGK7KC2PST\Users\Administrator\Desktop\Uretim\Excel\excel.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1',"DAT" )
worksheet.write('B1',str(today.day)+"."+str(today.month)+"."+str(today.year))
worksheet.write('C1',"DAT"+ str(today.day)+str(today.month)+str(today.year)+"-"+str(nowtime.hour)+str(nowtime.minute)+str(nowtime.second))
worksheet.write('D1',"D4" )
worksheet.write('E1',"D5")
worksheet.write('A2',"DAT")
worksheet.write('B2',"STOK ADI")
worksheet.write('C2',"BİRİM")
worksheet.write('D2',"MİKTAR")

workbook2=xlsxwriter.Workbook(r'\\WIN-VCGK7KC2PST\Users\Administrator\Desktop\Uretim\Formlar\Uretim_Formu_{}.xlsx'.format(formatted_date))

worksheet2 = workbook2.add_worksheet()

worksheet2.write('A1',"Ürün" )
worksheet2.write('B1',"Miktar")
worksheet2.write('C1',"Çip")
worksheet2.write('D1',"Kartuş")
worksheet2.write('E1',"Kafa")
worksheet2.write('F1',"Kutu")
worksheet2.write('G1',"Atık")
worksheet2.write('H1',"Toz")
worksheet2.write('I1',"Toz Gram")
worksheet2.write('J1',"Developer")
worksheet2.write('K1',"Developer Gram")
worksheet2.write('L1',"Çip")
worksheet2.write('M1',"Kartuş")
worksheet2.write('N1',"Kafa")
worksheet2.write('O1',"Kutu")
worksheet2.write('P1',"Atık")
worksheet2.write('Q1',"Toz")
worksheet2.write('R1',"Developer")


worksheet2.set_column('A:A', 10)
worksheet2.set_column('B:B', 10)
worksheet2.set_column('C:C', 10)
worksheet2.set_column('D:D', 10)
worksheet2.set_column('E:E', 10)
worksheet2.set_column('F:F', 10)
worksheet2.set_column('G:G', 10)
worksheet2.set_column('H:H', 10)
worksheet2.set_column('I:I', 10)
worksheet2.set_column('J:J', 10)
worksheet2.set_column('K:K', 10)
worksheet2.set_default_row(20)


counter=4
for i in range(len(df2)):
    for j in range(len(df)):
        if str(df2.iloc[i][0]).upper()== str(df.iloc[j][0]):
            worksheet2.write('A'+str(i+2),str(df2.iloc[i][0]).upper())
            worksheet2.write('B'+str(i+2),str(df2.iloc[i][1]))
            if str(df.iloc[j][5]) != "nan":
                if str(df2.iloc[i][2]).upper()=="Y": #chip
                    pass
                if str(df2.iloc[i][2])=="nan":
                    worksheet.write('A'+str(counter),str(df.iloc[j][5]))
                    worksheet.write('C'+str(counter),str(df.iloc[i][6]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1
                    worksheet2.write('C'+str(i+2),str(df.iloc[j][22]))
                else:
                    try:
                        index =df[df['Chip_short'] == str(df2.iloc[i][2]).upper()].index[0]
                        worksheet.write('A'+str(counter),str(df.iloc[index][5]))
                        worksheet.write('C'+str(counter),str(df.iloc[index][6]))
                        worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                        counter=counter+1
                        worksheet2.write('C'+str(i+2),str(df.iloc[index][22]))
                    except:
                        pass

                    
            if str(df.iloc[j][7]) != "nan":  
                if str(df2.iloc[i][4]).upper()=="Y":#empty
                    pass
                if str(df2.iloc[i][4]).upper()=="C" or str(df2.iloc[i][4]).upper()=="Ç" :
                    worksheet.write('A'+str(counter),str(df.iloc[j][7]).replace("80-","85-"))
                    worksheet.write('C'+str(counter),str(df.iloc[j][8]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1
                    
                    worksheet2.write('D'+str(i+2),str(df.iloc[j][23]))
                    
                if str(df2.iloc[i][4])=="nan":
                    
                    worksheet.write('A'+str(counter),str(df.iloc[j][7]))
                    worksheet.write('C'+str(counter),str(df.iloc[i][8]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1
                    
                    worksheet2.write('D'+str(i+2),str(df.iloc[j][23]))
                if str(df2.iloc[i][4])!="nan" and str(df2.iloc[i][4]).upper()!="Y" and str(df2.iloc[i][4]).upper()!="C" and str(df2.iloc[i][4]).upper()!="Ç":
                    
                    index=df[df['Empty_short'] == str(df2.iloc[i][4]).upper()].index[0]
                    worksheet.write('A'+str(counter),str(df.iloc[index][7]))
                    worksheet.write('C'+str(counter),str(df.iloc[index][8]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1

                    worksheet2.write('D'+str(i+2),str(df.iloc[index][23]))
                   
                    
            if str(df.iloc[j][9]) != "nan":
                if str(df2.iloc[i][6]).upper()=="Y":#head
                    pass
                if str(df2.iloc[i][6])=="nan":
                    worksheet.write('A'+str(counter),str(df.iloc[j][9]))
                    worksheet.write('C'+str(counter),str(df.iloc[i][10]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1
                    
                    worksheet2.write('E'+str(i+2),str(df.iloc[j][24]))
                if str(df2.iloc[i][6])!="nan" and str(df2.iloc[i][6]).upper()!="Y":
                    
                    index=df[df['Head_short'] == str(df2.iloc[i][6]).upper()].index[0]
                    worksheet.write('A'+str(counter),str(df.iloc[index][9]))
                    worksheet.write('C'+str(counter),str(df.iloc[index][10]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1

                    worksheet2.write('E'+str(i+2),str(df.iloc[index][24]))
                    
            if str(df.iloc[j][11]) != "nan": 
                if  str(df2.iloc[i][8]).upper()=="Y":#box
                    pass
                if  str(df2.iloc[i][8])=="nan":
                    worksheet.write('A'+str(counter),str(df.iloc[j][11]))
                    worksheet.write('C'+str(counter),str(df.iloc[i][12]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1
                    
                    worksheet2.write('F'+str(i+2),str(df.iloc[j][25]))
                if  str(df2.iloc[i][8])!="nan" and str(df2.iloc[i][8]).upper()!="Y":
                    
                    index=df[df['Box_short'] == str(df2.iloc[i][8]).upper()].index[0]
                    worksheet.write('A'+str(counter),str(df.iloc[index][11]))
                    worksheet.write('C'+str(counter),str(df.iloc[index][12]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1
                    worksheet2.write('F'+str(i+2),str(df.iloc[index][25]))
                   
                    
            if str(df.iloc[j][13]) != "nan": 
                if str(df2.iloc[i][10]).upper()=="Y":#waste
                    pass
                if str(df2.iloc[i][10])=="nan":
                    worksheet.write('A'+str(counter),str(df.iloc[j][13]))
                    worksheet.write('C'+str(counter),str(df.iloc[i][14]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1
                    
                    worksheet2.write('G'+str(i+2),str(df.iloc[j][26]))
                if str(df2.iloc[i][10])!="nan" and str(df2.iloc[i][10]).upper()!="Y":
                    if isinstance(df2.iloc[i][10], float):
                        index=df[df['Waste_short'] == str(str(int(df2.iloc[i][10])).upper())].index[0]
                    if isinstance(df2.iloc[i][10], str):
                        index=df[df['Waste_short'] == str(str(df2.iloc[i][10]).upper())].index[0]
                    worksheet.write('A'+str(counter),str(df.iloc[index][13]))
                    worksheet.write('C'+str(counter),str(df.iloc[index][14]))
                    worksheet.write('D'+str(counter),str(df2.iloc[i][1]).replace(".",","))
                    counter=counter+1
                

                    worksheet2.write('G'+str(i+2),str(df.iloc[index][26]))

                        
            if str(df.iloc[j][15]) != "nan":
                
                if str(df2.iloc[i][12]).upper()=="Y":#powder
                    pass
                if str(df2.iloc[i][12]).upper()!="NAN" and str(df2.iloc[i][14])=="nan":
                    
                   
                    index=df[df['Powder_short'] == str(df2.iloc[i][12]).upper()].index[0]
                    worksheet.write('A'+str(counter),str(df.iloc[index][15]))
                    worksheet.write('C'+str(counter),str(df.iloc[index][16]))
                    worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df.iloc[j][17])).replace(".",","))
                    counter=counter+1
                    worksheet2.write('H'+str(i+2),str(df.iloc[index][27]))
                    worksheet2.write('I'+str(i+2),str(float(df.iloc[j][17])).replace(".",","))

                    

                if str(df2.iloc[i][12])=="nan" and str(df2.iloc[i][14])!="nan":
                    worksheet.write('A'+str(counter),str(df.iloc[j][15]))
                    worksheet.write('C'+str(counter),str(df.iloc[j][16]))
                    if str(df.iloc[j][16])== "KG":
                        worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df2.iloc[i][14])/1000).replace(".",","))
                    else:
                        worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df2.iloc[i][14])).replace(".",","))
                    counter=counter+1
                    worksheet2.write('H'+str(i+2),str(df.iloc[j][27]))
                    if str(df.iloc[j][16])== "KG":
                        worksheet2.write('I'+str(i+2),str(float(df2.iloc[i][14])/1000).replace(".",","))
                    else:
                        worksheet2.write('I'+str(i+2),str(float(df2.iloc[i][14])).replace(".",","))


                if str(df2.iloc[i][12]).upper()!="NAN" and str(df2.iloc[i][14])!="nan":
                   
                        
                    index=df[df['Powder_short'] == str(df2.iloc[i][12]).upper()].index[0]
                    worksheet.write('A'+str(counter),str(df.iloc[index][15]))
                    worksheet.write('C'+str(counter),str(df.iloc[index][16]))
                    if str(df.iloc[j][16])== "KG":
                        worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df2.iloc[i][14])/1000).replace(".",","))
                    else:
                        worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df2.iloc[i][14])).replace(".",","))
                    counter=counter+1
                    worksheet2.write('H'+str(i+2),str(df.iloc[index][27]))
                    if str(df.iloc[j][16])== "KG":
                        worksheet2.write('I'+str(i+2),str(float(df2.iloc[i][14])/1000).replace(".",","))
                    else:
                        worksheet2.write('I'+str(i+2),str(float(df2.iloc[i][14])).replace(".",","))
                  


                if str(df2.iloc[i][12])=="nan" and str(df2.iloc[i][14])=="nan":
                    worksheet.write('A'+str(counter),str(df.iloc[j][15]))
                    worksheet.write('C'+str(counter),str(df.iloc[j][16]))
                    worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df.iloc[j][17])).replace(".",","))
                    counter=counter+1
                    worksheet2.write('H'+str(i+2),str(df.iloc[j][27]))
                    worksheet2.write('I'+str(i+2),str(float(df.iloc[j][17])).replace(".",","))
            if str(df.iloc[j][18]) != "nan":
                
                if str(df2.iloc[i][16]).upper()=="Y":#developer
                    pass 
                if str(df2.iloc[i][16]).upper()!="NAN" and str(df2.iloc[i][18])=="nan":
                  
                    index=df[df['Developer_short'] == str(df2.iloc[i][16]).upper()].index[0]
                    worksheet.write('A'+str(counter),str(df.iloc[index][18]))
                    worksheet.write('C'+str(counter),str(df.iloc[index][19]))
                    worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df.iloc[j][20])).replace(".",","))
                    counter=counter+1
                    worksheet2.write('J'+str(i+2),str(df.iloc[index][28]))
                    worksheet2.write('K'+str(i+2),str(float(df.iloc[j][20])).replace(".",","))

                  

                if str(df2.iloc[i][16])=="nan" and str(df2.iloc[i][18])!="nan":
                    worksheet.write('A'+str(counter),str(df.iloc[j][18]))
                    worksheet.write('C'+str(counter),str(df.iloc[j][19]))
                    worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df2.iloc[i][18])).replace(".",","))
                    counter=counter+1
                    worksheet2.write('J'+str(i+2),str(df.iloc[j][28]))
                    worksheet2.write('K'+str(i+2),str(float(df2.iloc[i][18])).replace(".",","))

                if str(df2.iloc[i][16]).upper()!="NAN" and str(df2.iloc[i][18])!="nan":
              
                
                    index=df[df['Developer_short'] == str(df2.iloc[i][16]).upper()].index[0]
                    worksheet.write('A'+str(counter),str(df.iloc[index][18]))
                    worksheet.write('C'+str(counter),str(df.iloc[index][19]))
                    worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df2.iloc[i][18])).replace(".",","))
                    counter=counter+1
                    worksheet2.write('J'+str(i+2),str(df.iloc[index][28]))
                    worksheet2.write('K'+str(i+2),str(float(df2.iloc[i][18])).replace(".",","))
                        

                  
                    
                if str(df2.iloc[i][16])=="nan" and str(df2.iloc[i][18])=="nan":
                
                    worksheet.write('A'+str(counter),str(df.iloc[j][18]))
                    worksheet.write('C'+str(counter),str(df.iloc[j][19]))
                    worksheet.write('D'+str(counter),str(int(df2.iloc[i][1])*float(df.iloc[j][20])).replace(".",","))
                    counter=counter+1
                    worksheet2.write('J'+str(i+2),str(df.iloc[j][28]))
                    worksheet2.write('K'+str(i+2),str(df.iloc[j][20]).replace(".",","))
            



workbook.close()
workbook2.close()
df_final = pd.read_excel(r'\\WIN-VCGK7KC2PST\Users\Administrator\Desktop\Uretim\Excel\excel.xlsx')

# Save the DataFrame as a CSV file.
df_final.to_csv(r'\\WIN-VCGK7KC2PST\Users\Administrator\Desktop\Uretim\Stok-hareket.csv', index=False,sep =';')
