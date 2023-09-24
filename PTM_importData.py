import os, io, subprocess, csv
from PyPDF2 import PdfReader

if input("ต้องการเปิดโฟลเดอร์ข้อมูลหรือไม่ (1:ใช่) : ") == "1":
    subprocess.Popen(r'explorer /select,' + __file__)

while 1:
    fileName = input("ป้อนชื่อไฟล์หนังสือแจ้งเตือน (.pdf) : ")
    if fileName.find(".") == -1:
        fileName = fileName + ".pdf"
    elif fileName.split(".")[1] != "pdf":
        print("ผิดพลาด : ต้องเป็นไฟล์ .pdf เท่านั้น !")
        continue        
    try:
        reader = PdfReader(fileName)        
        break
    except FileNotFoundError:
        print("ผิดพลาด : ชื่อไฟล์ไม่ถูกต้อง !")
        
pageNum = len(reader.pages)
header = ['Ref1', 'Name', 'Address', 'Province', 'Postcode']
data = []
os.system('cls')
print("\033[3;37;42m   Ref1           ชื่อ - สกุล                        ที่อยู่                       อำเภอ/จังหวัด           รหัสไปรษณีย์ \033[0;37;40m")

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
                print("{0:<15}".format(ref1), end="")
            elif line.find("นามสกุล :") != -1:
                name = line.split(": ")
                name = name[1].strip()
                print("{0:<35}".format(name[0:33]), end="")
            elif line.find("ที่อยู่ :") != -1:
                address = line.split(": ")
                address = address[1].strip()
                print("{0:<30}".format(address[0:28]), end="")
            elif line.find("อ.") != -1 and line.find("อ.") < 3:  #พบในตำแหน่งต้นๆ ของบรรทัด          
                province = line.strip()
                print("{0:<30}".format(province[0:28]), end="")
            elif line.find("ไปรษณีย์ :") != -1:                
                postcode = line.split(": ")
                postcode = postcode[1].strip()
                print("{0:<9}".format(postcode)) 

    row = [ref1, name, address, province, postcode]    
    data.append(row)          

print("------------------------------------------------------------------------------------------------------------------------")
print("รวมหนังสือแจ้ง จำนวน : ", pageNum)
fileName = input("\nตั้งชื่อไฟล์ข้อมูลที่จะบันทึก (.csv) : ")

if fileName.find(".") == -1:
    fileName = fileName + ".csv"

with open(fileName, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(header)
    writer.writerows(data)
         
input("\nบันทึกเรียบร้อย กดปุ่ม \"Enter\" เพื่อปิดโปรแกรม..")