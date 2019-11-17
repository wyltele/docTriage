import os
import shutil
import glob
import sys
import codecs
import json
from openpyxl import load_workbook

class journal:
	def __init__(self):
		cfgfile='C:\\Users\\hengyi\\Google Drive\\python\\journal.json'
		with codecs.open(cfgfile,'r','utf-8') as f:
			self.cfg=json.load(f)
		heads=''.join([i for i in self.cfg['triage']['base']])
		self.filelist=glob.glob(self.cfg["triage"]["search-location"]+'['+heads+']_*')
		self.wb=load_workbook(self.cfg['triage']['excel'])

	def run(self):
		if len(self.filelist)==0:
			print('nothing to process, exit.\n')
			sys.exit()
		for filename in self.filelist:
			ent=receipt_entry(filename,self.wb,self.cfg)
			ent.move_and_log()
		self.wb.save(self.cfg['triage']['excel'])

class receipt_entry:
	def __init__(self, filename,wb,cfg):
		self.wb=wb
		self.cfg=cfg
		self.fname=filename
		fields=os.path.splitext(os.path.basename(filename))[0].split("_")
		if len(fields)!=5:
			self.valid=False
		else:
			self.valid=True
			self.type=fields[0]
			self.date=fields[1]
			self.misc=fields[2]
			self.amount=fields[3]
			self.merchant=fields[4]
	def move_and_log(self):
		if self.valid:
			if self.misc in self.cfg["triage"]["categories"]:
				misc=self.cfg["triage"]["categories"][self.misc]
			to_dir=self.cfg["triage"]["base"][self.type]+misc
			if not os.path.exists(to_dir):
				os.makedirs(to_dir)
			to_dir=to_dir+"\\"+os.path.basename(self.fname)
			if os.path.exists(to_dir):
				os.remove(self.fname)
			else:
				shutil.move(self.fname,to_dir)
				ws=self.wb[self.cfg["triage"]["sheets"][self.type]]
				ws.append([self.date,self.amount,self.merchant,misc])

def main():
	journal().run()

if __name__== "__main__":
	main()
