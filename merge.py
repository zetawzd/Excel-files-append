import os
import pandas as pd
from tkinter import filedialog
import tkinter

# chon folder bang GUI
root=tkinter.Tk()
root.destroy()

folder=filedialog.askdirectory()
os.chdir(folder)

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
all_data.to_excel('master.xlsx')
