# Program to access Excel data, pics and generate Report card of the students in PDF format 

# importing libraries
import pandas as pd
import numpy
from datetime import timezone
from datetime import datetime


df = pd.read_excel("Dummy Data.xlsx")
#print(df.head(6))
full_data = []
#df.columns


# list for round

round_df = df['Unnamed: 1']
round_dum = round_df.values.tolist()
l_round = []

for i in range(0,126,25):
    l_round.append(round_dum[i])
        
#print(l_round)
full_data.append(l_round)


# list for first name 

fir_name_df = df['Unnamed: 2']
firname_dum = fir_name_df.values.tolist()
fir_name = []

for i in range(0,126,25):
    fir_name.append(firname_dum[i])
         
#print(fir_name)
full_data.append(fir_name)


# list for last name 

last_name_df = df['Unnamed: 3']
lastname_dum = last_name_df.values.tolist()
last_name = []

for i in range(0,126,25):
    last_name.append(lastname_dum[i])
        
#print(last_name)
full_data.append(last_name)


# list for full name 

f_name_df = df['Unnamed: 4']
fname_dum = f_name_df.values.tolist()
f_name = []

for i in range(0,126,25):
    f_name.append(fname_dum[i])
        
#print(f_name)
full_data.append(f_name)


# list for registration number

reg_df = df['Unnamed: 5']
reg_dum = reg_df.values.tolist()
reg_name = []

for i in range(0,126,25):
    reg_name.append(reg_dum[i])
        
#print(reg_name)
full_data.append(reg_name)


# list for grade

grade_df = df['Unnamed: 6']
grade_dum = grade_df.values.tolist()
grade_name = []

for i in range(0,126,25):
    grade_name.append(grade_dum[i])
        
#print(grade_name)
full_data.append(grade_name)


# list for school

scl_df = df['Unnamed: 7']
scl_dum = scl_df.values.tolist()
scl_name = []

for i in range(0,126,25):
    scl_name.append(scl_dum[i])
        
#print(scl_name)
full_data.append(scl_name)


# list for gender

gen_df = df['Unnamed: 8']
gen_dum = gen_df.values.tolist()
gen_name = []

for i in range(0,126,25):
    gen_name.append(gen_dum[i])
        
#print(gen_name)
full_data.append(gen_name)


# list for DOB

dob_df = df['Unnamed: 9']
dob_dum = dob_df.values.tolist()
dob_name = []

for i in range(0,126,25):
    if(i==0):
        dob_name.append(dob_dum[i])
    else:
        dob = dob_dum[i]
        dt = dob.replace(tzinfo=timezone.utc).timestamp()
        d_obj = datetime.fromtimestamp(dt).strftime('%m/%d/%Y')
        dob_name.append(d_obj)
    
#print(dob_name)
full_data.append(dob_name)


# list for city

city_df = df['Unnamed: 10']
city_dum = city_df.values.tolist()
city_name = []

for i in range(0,126,25):
    city_name.append(city_dum[i])
        
#print(city_name)
full_data.append(city_name)


# list for DOT

dot_df = df['Unnamed: 11']
dot_dum = dot_df.values.tolist()
dot_name = []

for i in range(0,126,25):
    dot_name.append(dot_dum[i])
        
#print(dot_name)
full_data.append(dot_name)


# list for Country

cou_df = df['Unnamed: 12']
cou_dum = cou_df.values.tolist()
cou_name = []

for i in range(0,126,25):
    cou_name.append(cou_dum[i])
        
#print(cou_name)
full_data.append(cou_name)


# List for QNo.
qn_final = []
qn_df = df['Unnamed: 13']
qn_dum = qn_df.values.tolist()
qn_per = []

for i in range(1,126):
    qn_per.append(qn_dum[i])
    if i%25==0:
        qn_final.append(qn_per)
        qn_per = []
#print(qn_final)


# List for Marked


mark_df = df['Unnamed: 14']
df['Unnamed: 14'] = df['Unnamed: 14'].fillna('-')


mark_final = []
#mark_df = df['Unnamed: 14']
mark_dum = mark_df.values.tolist()
mark_per = []

for i in range(1,126):
    mark_per.append(mark_dum[i])
        
    if i%25==0:
        mark_final.append(mark_per)
        mark_per = []
#print(mark_final)


# List for crct
crct_final = []
crct_df = df['Unnamed: 15']
crct_dum = crct_df.values.tolist()
crct_per = []

for i in range(1,126):
    crct_per.append(crct_dum[i])
    if i%25==0:
        crct_final.append(crct_per)
        crct_per = []
#print(crct_final)


# List for otc
otc_final = []
otc_df = df['Unnamed: 16']
otc_dum = otc_df.values.tolist()
otc_per = []

for i in range(1,126):
    otc_per.append(otc_dum[i])
    if i%25==0:
        otc_final.append(otc_per)
        otc_per = []
        
#print(otc_final)


# List for scr
scr_final = []
scr_df = df['Unnamed: 17']
scr_dum = scr_df.values.tolist()
scr_per = []

for i in range(1,126):
    scr_per.append(scr_dum[i])
    if i%25==0:
        scr_final.append(scr_per)
        scr_per = []
#print(scr_final)


# List for ysc
ysc_final = []
ysc_df = df['Unnamed: 18']
ysc_dum = ysc_df.values.tolist()
ysc_per = []
ysc_total = []
total = 0

for i in range(1,126):
    ysc_per.append(ysc_dum[i])
    total += int(ysc_dum[i])
    if i%25==0:
        ysc_final.append(ysc_per)
        ysc_total.append(total)
        ysc_per = []
        total = 0
#print(ysc_final)
print(ysc_total)


# list for res

res_df = df['Unnamed: 19']
res_dum = res_df.values.tolist()
res_name = []

for i in range(0,126,25):
    res_name.append(res_dum[i])
        
#print(res_name)
full_data.append(res_name)


#print(full_data)


#!pip install reportlab


from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
from reportlab.lib.units import mm, cm


count = 0
for i in range(1,6):
    
    round_name = str(full_data[0][i])
    fir_name = full_data[1][i]
    las_name = full_data[2][i]
    full_name = full_data[3][i]
    reg_num = str(full_data[4][i])
    grd = str(full_data[5][i])
    scl_name = full_data[6][i]
    gen = full_data[7][i]
    dob = full_data[8][i]
    city = full_data[9][i]
    dot = full_data[10][i]
    coun = full_data[11][i]
    
    file_name = full_name
    stri = file_name+".pdf"
    fileName = stri
    pdf = canvas.Canvas(fileName)
    pdf.drawImage("logo.png", 70,750,4*cm,4*cm)
    # setting the title of the document
    #pdf.setTitle(documentTitle)
   
    pdf.setFont('Courier-Bold', 20)
    pdf.drawCentredString(300, 800, "SCORE CARD")
    pdf.setFont('Courier', 16)
    pdf.drawCentredString(300, 780, "PQR CHAMPIONSHIP (ROUND - "+round_name+")")
    pdf.line(600,770,0,770)
    #drawImage(image, x, y, width=None, height=None, mask=None, preserveAspectRatio=True, anchor='c')
    #pdf.drawImage("logo.png", 50,700, width=1, height=1, mask=None, preserveAspectRatio=True)
    
    pdf.drawString(20, 750, "First Name           - ")
    pdf.drawString(250, 750, fir_name)
    pdf.drawString(20, 730, "Last Name            - ")
    pdf.drawString(250, 730, las_name)
    pdf.drawString(20, 710, "Full Name            - ")
    pdf.drawString(250, 710, full_name)
    pdf.drawString(20, 690, "Registration Number  - ")
    pdf.drawString(250, 690, reg_num)
    pdf.drawString(20, 670, "Grade                - ")
    pdf.drawString(250, 670, grd)
    pdf.drawImage('Pics of students/'+ full_name +'.png', 420,580,5*cm,5*cm)
    pdf.drawString(20, 650, "School Name          - ")
    pdf.drawString(250, 650, scl_name)
    pdf.drawString(20, 630, "Gender               - ")
    pdf.drawString(250, 630, gen)
    pdf.drawString(20, 610, "Date of Birth        - ")
    pdf.drawString(250, 610, dob)
    pdf.drawString(20, 590, "City of Residence    - ")
    pdf.drawString(250, 590, city)
    pdf.drawString(20, 570, "Date and Time of Test- ")
    pdf.drawString(250, 570, dot)
    pdf.drawString(20, 550, "Country of Residence - ")
    pdf.drawString(250, 550, coun)
    pdf.line(600,530,0,530)
    
    pdf.setFont('Courier-Bold', 10)
    pdf.drawString(20,510, 'Q No.   What you marked?    Correct Answer    Outcome    Score if correct    Your Score')
    pdf.line(600,490,0,490)
    
    pdf.setFont('Courier', 10)
    for j in range(0,25):
        if len(qn_final[i-1][j]) == 2:
            if otc_final[i-1][j] == 'Correct':
                pdf.drawString(20, 470-15*j , str(qn_final[i-1][j])+"            "+str(mark_final[i-1][j])+"                 "+str(crct_final[i-1][j])+"             "+str(otc_final[i-1][j])+"            "+str(scr_final[i-1][j])+"             "+str(ysc_final[i-1][j]))
            elif otc_final[i-1][j] == 'Unattempted':
                pdf.drawString(20, 470-15*j , str(qn_final[i-1][j])+"            "+str(mark_final[i-1][j])+"                 "+str(crct_final[i-1][j])+"             "+str(otc_final[i-1][j])+"        "+str(scr_final[i-1][j])+"             "+str(ysc_final[i-1][j]))
            elif otc_final[i-1][j] == 'Incorrect':
                pdf.drawString(20, 470-15*j , str(qn_final[i-1][j])+"            "+str(mark_final[i-1][j])+"                 "+str(crct_final[i-1][j])+"             "+str(otc_final[i-1][j])+"          "+str(scr_final[i-1][j])+"             "+str(ysc_final[i-1][j]))

        elif len(qn_final[i-1][j]) == 3:
            if otc_final[i-1][j] == 'Correct':
                pdf.drawString(20, 470-15*j , str(qn_final[i-1][j])+"           "+str(mark_final[i-1][j])+"                 "+str(crct_final[i-1][j])+"             "+str(otc_final[i-1][j])+"            "+str(scr_final[i-1][j])+"             "+str(ysc_final[i-1][j]))
            elif otc_final[i-1][j] == 'Unattempted':
                pdf.drawString(20, 470-15*j , str(qn_final[i-1][j])+"           "+str(mark_final[i-1][j])+"                 "+str(crct_final[i-1][j])+"             "+str(otc_final[i-1][j])+"        "+str(scr_final[i-1][j])+"             "+str(ysc_final[i-1][j]))
            elif otc_final[i-1][j] == 'Incorrect':
                pdf.drawString(20, 470-15*j , str(qn_final[i-1][j])+"           "+str(mark_final[i-1][j])+"                 "+str(crct_final[i-1][j])+"             "+str(otc_final[i-1][j])+"          "+str(scr_final[i-1][j])+"             "+str(ysc_final[i-1][j]))

    pdf.line(600,90,0,90)
    
    pdf.setFont('Courier-Bold', 12)
    pdf.drawString(20, 80, "The Maximum Marks to Score in the Exam is 100. There are no negative marks.")
    pdf.setFont('Courier', 12)
    pdf.drawString(20, 60, "TOTAL SCORE   -  "+ str(ysc_total[i-1])+"/100")
    pdf.drawString(20, 40, "RESULT        -  "+ full_data[12][i])
    pdf.save()
    count+=1