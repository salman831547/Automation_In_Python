import shutil
import os
import pandas as pd
from openpyxl import load_workbook
import schedule
import time

def job():

	#Get List of all processed files in the past
	processed_files = [file for file in os.listdir('Processed/') if file.endswith('.xlsx')]
	processed_path = os.path.join(os.getcwd(),'Processed/',''.join(processed_files))

	#Check if any new files appeared in drop folder
	dropped_files = [file for file in os.listdir('Drop/') if file.endswith('.xlsx')]
	drop_path = os.path.join(os.getcwd(),'Drop/',''.join(dropped_files))

	#if there is a new file lets load it to a dataframe and prepare it to write
	if dropped_files:
		df=pd.read_excel(drop_path, usecols= ['Name', 'Product', 'Production Run Time (Min)','Products Produced (Units)']) 
        #Find the current number of entries in the main file
		df_main=pd.read_excel(r'C:\Users\Sayyad Manzil\Desktop\Python_Data\shakil_123\Excel_automation\main.xlsx', usecols=  ['Name', 'Product', 'Production Run Time (Min)','Products Produced (Units)'])
		current_rows=df_main.shape[0]

		#Load the main workbook
		workbook_name = r'C:\Users\Sayyad Manzil\Desktop\Python_Data\shakil_123\Excel_automation\main.xlsx'
		wb = load_workbook(workbook_name)
		page=wb['Sheet1']
		#page = wb.active

		#Write new entries to the main workbook
		new_etries = df.values.tolist()
		for i in new_etries:
			page.append(i)
		wb.save(filename=workbook_name)
		df_main_new=pd.read_excel(r'C:\Users\Sayyad Manzil\Desktop\Python_Data\shakil_123\Excel_automation\main.xlsx', usecols= ['Name', 'Product', 'Production Run Time (Min)','Products Produced (Units)'])
		new_rows=df_main_new.shape[0]

		#Check to see if old rows+appended rows = total new rows in updated excel

		if new_rows == current_rows+df.shape[0]:
			shutil.move(drop_path, os.path.join(os.getcwd(),'Processed/'))
		print('All Files Processed. Completed')

	else:
		print('No New Files')

#schedule.every().day.at("05:53").do(job)
schedule.every().minute.at(":10").do(job)

while True:
    schedule.run_pending()
    time.sleep(1) # wait one minute


