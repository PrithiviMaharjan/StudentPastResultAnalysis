# -*- coding: utf-8 -*-
"""
Created on Sat Jan  9 11:14:18 2021

@author: Shreyash
"""

import os 
import pandas as pd
import argparse
from visualization import init_codes
from grade_comparison import grade_comparison
from module_wise_summary import module_wise_summary
from visualization_graph import visualization

parser=argparse.ArgumentParser(description='Testing..')
parser.add_argument('-d','--dir',help='Enter directory path')
parser.add_argument('-f','--fname',help='Enter filename')
parser.add_argument('-m','--mode',help='Report generation mode, 1: Single Module Report,2: Module Wise Summary, 3: Comparison Report, 4: Plot charts')
args=parser.parse_args()



# base_dir = "data"
base_dir=args.dir
fname=args.fname
mode=args.mode

# print(base_dir)
files = os.listdir(base_dir)
report=fname+'_report_summary.xlsx'



# writer = pd.ExcelWriter('results_summary.xlsx', engine='xlsxwriter')


# for code,filename in data.items():
# if(code==fname):
# newFname=filename
def generate_report():
    global start_col, start_row, count, found, filecount
    start_col_ = 0
    start_row_ = 0
    count = 0
    found=False
    filecount=0
    for file in files:
        substring=file[0:6]
        # print(substring)
        if(fname==substring):
            count+=1
            file_path = os.path.join(base_dir,file)
            writer = pd.ExcelWriter(report, engine='xlsxwriter')
            marks = pd.read_excel(file_path,skiprows=9)
            details = pd.read_excel(file_path,nrows=8,header=None,usecols="A,C")
            
            # calculate grade frequencies and grade averages
            GRADES = ("A", "B", "C", "D", "E", "F")
            grade_frequencies = {}
            grade_avgs = {}
            for grade in GRADES:
                try:
                    d = marks[marks["Final Grade"] == grade]
                    grade_frequencies[grade] = d["Final Grade"].count()
                except:
                    d = marks[marks["Module Grade"] == grade]
                    grade_frequencies[grade] = d["Module Grade"].count()
                try:
                    grade_avgs[grade] = float("%.2f" % d["Final Marks"].mean())
                except:
                    grade_avgs[grade] = float("%.2f" % d["Module Mark"].mean())
            
            # calculate grade percents
            grade_percents = {}
            total = sum(grade_frequencies.values())
            for g in GRADES:
                grade_percents[g] = float("%.2f" % ((grade_frequencies[g] / total) * 100))
            
            # marks details
            marks_details = {"#": grade_frequencies,
                             "%": grade_percents,
                             "Avg Mark": grade_avgs}
            df = pd.DataFrame(marks_details)
            
            # other details
            other_details = {'#': {"Students Countable": df["#"].sum(),
                                   "Intermission": 0,
                                   "Withdraw": 0,
                                   "Course Transfer": 0},
                            '%': {},
                            'Avg Mark':{}}
            
            df2 = pd.DataFrame(other_details)
            df2.reindex(['Students Countable', 'Intermission', 'Withdraw', 'Course Transfer'])
            df3 = df.append(df2)
            df3.reindex(['A', 'B', 'C', 'D', 'E', 'F', 
                         'Students Countable', 'Intermission', 'Withdraw', 'Course Transfer'])
            
            #write to excel
            try:
                summary = {'#': {"Students Total": 0,
                                "Module Avg Mark": float("%.2f" % marks['Final Marks'].mean()),
                                "Pass %": float("%.2f" % (((df['#'].sum() - df.loc["F"]["#"]) / df['#'].sum()) * 100))},
                            '%': {},
                            'Avg Mark': {}}
            except:
                summary = {'#':{"Students Total": 0,
                                "Module Avg Mark": float("%.2f" % marks['Module Mark'].mean()),
                                "Pass %": float("%.2f" % (((df['#'].sum() - df.loc["F"]["#"]) / df['#'].sum()) * 100))},
                            '%': {},
                            'Avg Mark': {}}
                
            df4 = pd.DataFrame(summary)
            final_df = df3.append(df4)
            final_df = final_df.reindex(['A', 'B', 'C', 'D', 'E', 'F',
                                         'Students Countable', 'Intermission', 'Withdraw', 
                                         'Course Transfer', 'Students Total', 'Module Avg Mark', 'Pass %'])
            
            # write module details
            details.iloc[[0,1]][:].to_excel(writer, startrow=start_row_, startcol=start_col_, 
                                            header=False, index=False)
            
            # write summary
            final_df.to_excel(writer, startrow=start_row_+3, startcol=start_col_)
            
            count += 1
            if count % 2 == 0:
                start_col_ = 0 
                start_row_ += 19
            else:
               start_col_ += 5
               start_row_ += 0
            writer.save()
            print("Report has been generated.")
    
    if(count==0):
        print("Module not found")
        
if(mode=='1'):
    generate_report()
elif(mode=='2'):
    grade_comparison()
elif(mode=='3'):
    module_wise_summary()
elif(mode=='4'):
    visualization()