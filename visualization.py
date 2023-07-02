# -*- coding: utf-8 -*-
"""
Created on Mon Feb  8 13:08:36 2021

@author: HP
"""
import os
import os.path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import xlsxwriter

# creating a new excel file
workbook = xlsxwriter.Workbook('module_wise_summary_visualization.xlsx', {'nan_inf_to_errors': True})    

# creating a function to convert data-type
def f(x):
    return np.float(x)

count = 0

# main method to perform visualization task
def plotCharts(sheetName, sheetName2):
    # slicing sheet name
    # sheetName2 = sheetName[0:25]+"..."
    # creating a new worksheet
    worksheet = workbook.add_worksheet(sheetName2)
    
    # reading excel data to add values in new excel file    
    data_info = pd.read_excel("module_wise_summary.xlsx",sheet_name=sheetName2)
    data_info = data_info.iloc[0:1]
    datarow0 = data_info.columns.values
    datarow1 = list(data_info.iloc[0])
    # writing data_info file in worksheet
    worksheet.write("A1", datarow0[0])
    worksheet.write("B1", datarow0[1])
    worksheet.write("A2", datarow1[0])
    worksheet.write("B2", datarow1[1])

    # reading data for visualization
    data=pd.read_excel("module_wise_summary.xlsx",sheet_name=sheetName2,skiprows=3)
    # data manipulation
    vis_data = data
    vis_data.rename(columns ={"Unnamed: 0":"Grade"}, inplace=True)
    vis_data.rename(columns ={"#":"Count"}, inplace=True)
    vis_data = vis_data.iloc[:6] 

    # preparing data for the workbook
    bold = workbook.add_format({'bold': 1}) 
    headings=vis_data.columns.values
    gradePer=[
    vis_data["Grade"].iloc[:6].values.astype(str),
    vis_data["Count"].iloc[:6].values.astype(str),
    vis_data["%"].iloc[:6].values.astype(str),
    vis_data["Avg Mark"].iloc[:6].values.astype(str)
    ]

    f2 = np.vectorize(f) # vectorizing data f
    
    # writing data file in worksheet
    worksheet.write_row('A4', headings, bold)  
    worksheet.write_column('A5', gradePer[0]) 
    worksheet.write_column('B5', f2(gradePer[1]))
    worksheet.write_column('C5', f2(gradePer[2]))
    worksheet.write_column('D5', f2(gradePer[3]))

    ''' visualization starts from here '''
    # grade count bar-graph visualization
    vis_grade=vis_data[["Grade","Count"]]
    # setting Seaborn style
    sns.set_style('darkgrid')    
    # constructing Seaborn bar-graph (Grade-Count)
    sns.barplot(x = "Grade", y = "Count", data = vis_grade)
    barName = sheetName+" Bar.png"
    plt.savefig(barName)
    plt.close()
    # adding Seaborn bar-graph in sheet
    worksheet.write('H2', 'Grade Count Bar Plot')
    worksheet.insert_image('H3', barName, {'x_scale': 0.8, 'y_scale': 0.8})

    # grade percentage pie-chart plot visualization    
    percentage=vis_data[["Grade","%"]]
    # setting plot figure size
    plt.subplots(figsize=[10,6])
    labels = percentage["Grade"]
    # constructing pie-chart using matplotlib
    percentage["%"].plot.pie(autopct="%.1f%%", labels=labels)
    pieName = sheetName+ " Pie.png"
    plt.savefig(pieName)
    plt.close()
    # adding matplotlib pie-chart in sheet
    worksheet.write('H21', 'Grade Percentage Pie Plot')
    worksheet.insert_image('H22', pieName, {'x_scale': 0.6, 'y_scale': 0.6})
    
# loop to fetch the each sheet names
base_dir = "data"
files = os.listdir(base_dir)
for file in files:
    file_path = os.path.join(base_dir,file)
    marks = pd.read_excel(file_path,skiprows=9)
    details = pd.read_excel(file_path,nrows=8,header=None,usecols="A,C")
    alist=list(details.iloc[[1]][2])
    sn1 = str(alist[0])
    sn2 = str(alist[0])[0:25]+ "..."
    plotCharts(sn1,sn2)

# calling
#plotCharts("Introduction to Information Systems","Introduction to Informati...")
#plotCharts("Security in Computing","Security in Computing...")
workbook.close()

#using listdir() method to list the files of the folder
test = os.listdir(os.getcwd())
#taking a loop to remove all the images
#using ".png" extension to remove only png images
#using os.remove() method to remove the files
for images in test:
    if images.endswith(".png"):
        os.remove(images)
