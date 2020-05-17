import os
import xlrd
import numpy as np
import cv2

fname=input("Enter File name  (Excel file with extension .xlsx)= ")
path=os.getcwd().replace("\\","/")
loc=path+"/"+fname
wb = xlrd.open_workbook(loc)

x=int(input("Enter Sheet number= "))
sheet = wb.sheet_by_index(x)

rows=list()
for z in range(sheet.nrows):
    if sheet.row_values(x)[0] is not None:
        print("Columns present in file is : ")
        print(sheet.row_values(x))
        break

y=int(input("Enter number of extra columns in excel file (columns are in Stack so if you enter 2 , first 2 columns will be deleted) : "))
q=int(input("\nEnter start of serial Number (example- EF-0001 you have to type 1) "))-1

for i in range(sheet.nrows):
    if sheet.row_values(i)[0]=="":
        continue
    else:
        i=sheet.row_values(i)
        while y>0:
            del i[0]
            y-=1
        if q<10:
            sno="EF-000"+str(q)
        elif q<100:
            sno="EF-00"+str(q)
        elif q<1000:
            sno="EF-0"+str(q)
        i.append(sno)
        q+=1
        rows.append(i)


imag=input("Enter image name ( with extension (.jpg/.png) )=")
np=len(rows[0])
print("\nPosition sholud be in tuples i.e X and Y coordinates in pixels and in order like this : ")
print(rows[0])
print("\nEnter positons of blanks (example : x 'space' y (eg : 688 688) then press enter ) :")
position=list()
while np>0:
    t= tuple(map(int,input().split()))
    position.append(t)
    np-=1


for j in rows:
    image = cv2.imread(imag,cv2.IMREAD_UNCHANGED)
    for i in range(len(position)):
            cv2.putText(
                image, #numpy array on which text is written
                j[i], #text
                position[i], #position at which writing has to start
                cv2.FONT_HERSHEY_SIMPLEX, #font family
                1, #font size
                (0, 0, 0, 0), #font color
                3) #font stroke
    cv2.imwrite((j[5]+j[0]+'.png'), image)
