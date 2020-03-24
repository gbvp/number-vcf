from openpyxl import Workbook, load_workbook
excel = load_workbook('temp.xlsx')
data = excel.active
file1 = open("temp.vcf","w",encoding="utf-8")
for i in range(1,len(data['A'])+1):	
	file1.write('BEGIN:VCARD')
	file1.write('\nVERSION:3.0')
	file1.write('\nN:'+str(data['B'+str(i)].value).strip()+' ;'+str(data['A'+str(i)].value).strip()+';;;')
	file1.write('\nFN:'+str(data['A'+str(i)].value).strip()+' '+str(data['B'+str(i)].value).strip())
	file1.write('\nTEL;TYPE=HOME:+'+str(data['C'+str(i)].value).strip()+str(data['D'+str(i)].value).strip()[-10:])
	file1.write('\nTEL;TYPE=HOME:+'+str(data['C'+str(i)].value).strip()+str(data['D'+str(i)].value).strip()[-10:])
	file1.write('\nEND:VCARD\n')
file1.close() 