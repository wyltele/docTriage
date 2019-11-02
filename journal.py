import os
import shutil
import glob
import sys
import codecs
from openpyxl import load_workbook

class journal:
	def run(self):
		filelist=glob.glob('C:\\Users\\hengyi\\Google Drive\\[ch]_*')
#		f=codecs.open('C:\\Users\\hengyi\\Google Drive\\Jingjing\\journal.csv','a',encoding="utf-8")
		wb=load_workbook('C:\\Users\\hengyi\\Google Drive\\python\\journal.xlsx')
		ws = wb.worksheets[0]
		
		if len(filelist)==0:
			print('nothing to process, exit.\n')
			sys.exit()
		for file in filelist:
			ent=receipt_entry(file)
			ent.move()
			if ent.valid:
				ws.append([ent.date,ent.amount,ent.merchant,ent.misc])
#				f.write(ent.date+','+ent.amount+','+ent.merchant+','+ent.misc+'\n')
#		f.close()
		wb.save('C:\\Users\\hengyi\\Google Drive\\python\\journal.xlsx')
		
class receipt_entry:
	def __init__(self, filename):
		self.fname=filename
		fields=os.path.splitext(os.path.basename(filename))[0].split("_")
		self.googledrv='C:\\Users\\hengyi\\Google Drive\\J BRANDS JOUNALS\\Journals 2018\\'
		self.localdrv='C:\\Users\\hengyi\\Downloads\\receipts\\'
		if len(fields)!=5:
			self.valid=False
		else:
			self.valid=True
			self.type=fields[0]
			self.date=fields[1]
			self.misc=fields[2]
			self.amount=fields[3]
			self.merchant=fields[4]
	def move(self):
		if self.valid:
			if self.type=='c':
				shutil.move(self.fname,self.googledrv)
			else:
				if self.misc=='mc':
					self.misc='买菜'
				elif self.misc=='yf':
					self.misc='衣服'
				elif self.misc=='qy':
					self.misc='汽油或车'
				elif self.misc=='jy':
					self.misc='教育'
				elif self.misc=='ry':
					self.misc='日用品'
				elif self.misc=='fd':
					self.misc='饭店'
				elif self.misc=='ly':
					self.misc='旅游'
				elif self.misc=='jj':
					self.misc='家居'
				elif self.misc=='yl':
					self.misc='医疗'
				elif self.misc=='dz':
					self.misc='电子产品'
				elif self.misc=='mp':
					self.misc='门票'
				else:
					pass
			
				target=self.localdrv+self.misc+'\\'
				if not os.path.exists(target):
					os.makedirs(target)
				shutil.move(self.fname,target)
			
def main():
	journal().run()
  
if __name__== "__main__":
	main()