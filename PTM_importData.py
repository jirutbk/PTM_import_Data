import io
from PyPDF2 import PdfReader

reader = PdfReader("sampleData.pdf")

pageNum = len(reader.pages)

for page in reader.pages:    
    page_text = page.extract_text()
            
    i = 0
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
            elif line.find("อ.") != -1 and line.find("อ.") < 3:            
                province = line.strip()
                print(province)
            elif line.find("ไปรษณีย์ :") != -1:
                #print(line)
                postcode = line.split(": ")
                print(postcode[1].strip())        
            i = i + 1
    print("-------------------------------------")

# print number of pages
print("รวมจำนวนใบแจ้ง : ", pageNum)