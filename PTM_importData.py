import io
from PyPDF2 import PdfReader

fileName = input("ป้อนชื่อไฟล์หนังสือแจ้งเตือน (.pdf) : ")
reader = PdfReader(fileName)
#reader = PdfReader("sampleData.pdf")

pageNum = len(reader.pages)

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
                print(ref1[1].strip())
            elif line.find("นามสกุล :") != -1:
                name = line.split(": ")
                print(name[1].strip())
            elif line.find("ที่อยู่ :") != -1:
                address = line.split(": ")
                print(address[1].strip())
            elif line.find("อ.") != -1 and line.find("อ.") < 3:  #พบในตำแหน่งต้นๆ ของบันทัด          
                province = line.strip()
                print(province)
            elif line.find("ไปรษณีย์ :") != -1:
                #print(line)
                postcode = line.split(": ")
                print(postcode[1].strip())        
            
    print("-------------------------------------")

# print number of pages
print("รวมหนังสือแจ้ง จำนวน : ", pageNum)