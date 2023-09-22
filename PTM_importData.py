﻿import io
import csv
from PyPDF2 import PdfReader

fileName = input("ป้อนชื่อไฟล์หนังสือแจ้งเตือน (.pdf) : ")
reader = PdfReader(fileName)
#reader = PdfReader("sampleData.pdf")

pageNum = len(reader.pages)

header = ['Ref1', 'Name', 'Address', 'Province', 'Postcode']
data = []

for page in reader.pages:    
    page_text = page.extract_text()
    
    ref1 = ""
    name = ""
    address = ""
    province = ""
    postcode = ""
    
    for line in io.StringIO(page_text):
            if line.find("Ref1") != -1:            
                ref1 = line.split(": ")
                ref1 = ref1[1].strip()
                print(ref1)
            elif line.find("นามสกุล :") != -1:
                name = line.split(": ")
                name = name[1].strip()
                print(name)
            elif line.find("ที่อยู่ :") != -1:
                address = line.split(": ")
                address = address[1].strip()
                print(address)
            elif line.find("อ.") != -1 and line.find("อ.") < 3:  #พบในตำแหน่งต้นๆ ของบันทัด          
                province = line.strip()
                print(province)
            elif line.find("ไปรษณีย์ :") != -1:
                #print(line)
                postcode = line.split(": ")
                postcode = postcode[1].strip()
                print(postcode) 

    row = [ref1, name, address, province, postcode]    
    data.append(row)
            
    print("-------------------------------------")

# print number of pages
print("รวมหนังสือแจ้ง จำนวน : ", pageNum)

fileName = input("\nตั้งชื่อไฟล์ข้อมูลที่จะบันทึก (.csv) : ")

with open(fileName, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(header)
    writer.writerows(data)

input("\nบันทึกเรียบร้อย กดปุ่ม Enter เพื่อจบการทำงาน...")