import os
import pandas as pd
from tkinter import filedialog
import tkinter
import win32com.client

'''
# doi current working directory
folder='D:\Python\Pewfolder'
os.chdir(folder)

# go input folder
found=False
while not found:
	folder=input('Enter path to files: ')
	if not os.path.isdir(folder):
		print("Go sai path cmnr, nhap lai de.")
	else:
		print("Nhap link dung cmnr, de day lam tiep :)")
		found=True 
		os.chdir(folder)
'''

# select current directory bang graphic user interface (chuot)
root=tkinter.Tk()
root.destroy()

folder=filedialog.askdirectory()
os.chdir(folder)

# focus lai vao cmd, do phai bam chuot
shell=win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('%{TAB}')

# tao dataframe trong
all_data=pd.DataFrame()

# chay vong lap, ghep file, xong tim file co nhieu cot nhat, tao list theo column cua file do
n=len(pd.read_excel(os.listdir(folder)[0]).columns)
for x in os.listdir(folder):
	df=pd.read_excel(x)
	all_data=pd.concat([df,all_data],ignore_index=True)
	a=len(df.columns)
	if a-n>0:
		ds=list(df)

#tach phan theo list (main) va ko theo list (side), xong ghep lai
main=all_data.reindex(columns=ds)
side=all_data.drop(ds,axis=1)
all_data=pd.concat([main,side],axis=1)

# xuat file excel
all_data.to_excel(input('Chang hay tien sinh muon dat ten file la gi: ')+'.xlsx')
