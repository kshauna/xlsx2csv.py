# xlsx2csv.py

import glob
import os
from os.path import exists
import sys
import numpy as np
import pandas as pd

f = glob.glob("*.xlsx") #grab all excel files
num_files=len(f) #count number of negatives by using total file count as proxy
for i in f:
    excel = pd.read_excel(i, 'Demultiplexing sheet', engine="openpyxl", keep_default_na=False) #xlrd not well supported hence use openpyxl
    excel['Specimen-code-number']=excel['Specimen-code-number'].map(str) #change zrc number from int to str (prevent loss of leading zeros)
    excel['Specimen-code-number'] = excel['Specimen-code-number'].apply(lambda x: x.zfill(7)) #add back leading zeros up to 7 characters total.
    excel.to_csv(i+'.csv', index=False) #change to csv

all_files = glob.glob("*.csv") #grab all csv

#read CSVs as string (prevent loss of leading zeros)
df= (pd.read_csv(f, dtype=str) for f in all_files) 

#concatenate all CSVs into one Dataframe
con_df = pd.concat(df,ignore_index=True) 

#generate the CSV for sorting sheet creation
if exists('Primertagreplacements.csv') == True: 
    con_df['ZRC'] = con_df['Specimen-code-prefix'] + con_df['Specimen-code-number']
    ss_df = con_df.filter(items=['ZRC', 'Plate-ID', 'F-primer-tag-ID', 'R-primer-tag-ID'])
    Fprimertag = ss_df['F-primer-tag-ID'].tolist()
    FHMtoU = []
    for i in Fprimertag:
        i = str(i)
        i = i.replace("HM", "U")
        FHMtoU.append(i)
    ss_df['New-F-primer-tag-ID'] = FHMtoU
    Rprimertag = ss_df['R-primer-tag-ID'].tolist()
    RHMtoU = []
    for i in Rprimertag:
        i = str(i)
        i = i.replace("HM", "U")
        RHMtoU.append(i)
    ss_df['New-R-primer-tag-ID'] = RHMtoU
    replacements = pd.read_csv("Primertagreplacements.csv")
    replacements = replacements.set_index('Original')
    rep_dict = replacements.to_dict()
    sorting_df = ss_df.replace({"New-R-primer-tag-ID": rep_dict['New']})
    nonans_ssdf = sorting_df.dropna()
    fourcolumns = nonans_ssdf.drop(columns=['F-primer-tag-ID', 'R-primer-tag-ID'])
    fourcolumns.to_csv("sortingsheetdemux.csv",index=False)

#Checks if collection dates are ok, merges them together, takes out nan
if 'Collection-Date' in con_df.columns:
    con_df['Collection Date'] = con_df['Collection Date'].map(str)
    con_df['Collection-Date'] = con_df['Collection-Date'].map(str)
    con_df['Combined Collection Date'] = con_df['Collection Date'] + con_df['Collection-Date']
    combinedcolldate_list = con_df['Combined Collection Date'].tolist()
    nonan_dates = []
    for i in combinedcolldate_list:
        i = str(i)
        i = i.replace("nan", "")
        nonan_dates.append(i)
    con_df['New Combined Collection Date'] = nonan_dates
    con_df['SpecimenID']= con_df['Researcher-name']+'_'+con_df['Plate-ID'].map(str)+'_'+con_df['Specimen-code-prefix']+con_df['Specimen-code-number'].map(str)+'_'+con_df['SampleID'].map(str)+'_'+con_df['Locality'].map(str)+'_'+con_df['New Combined Collection Date'].map(str)
else: 
    con_df['SpecimenID']= con_df['Researcher-name']+'_'+con_df['Plate-ID'].map(str)+'_'+con_df['Specimen-code-prefix']+con_df['Specimen-code-number'].map(str)+'_'+con_df['SampleID'].map(str)+'_'+con_df['Locality'].map(str)+'_'+con_df['Collection Date'].map(str)

#Checking for duplicates in Specimen-code-number which are a problem if 
# Plate-ID is different (same Plate-ID means a rerun plate)
#duplicated_ZRC = con_df[con_df.duplicated(['Specimen-code-number'], keep=False)]
#notrealduplicates = ['0000001', '1_rerun','1_repcr']
#duplicated_ZRC = duplicated_ZRC[duplicated_ZRC['Specimen-code-number'].isin(notrealduplicates) == False] 

#if duplicated_ZRC.empty == False:
#    duplicated_ZRC.to_csv("dupes.csv",index=False)

#if exists('dupes.csv') == True: 
#    print("You have duplicate ZRC codes, please check 'dupes.csv' file in your folder.")
#    sys.exit()

#makes new column
con_df['TagFsequence']=con_df['F-primer-tag']

#makes new column
con_df['TagRsequence']=con_df['R-primer-tag']

#Asks for user input (in case different primers are used)
PrimerF = input("Enter Primer F Sequence: ")
PrimerR = input("Enter Primer R Sequence: ")
con_df['PrimerF']= PrimerF
con_df['PrimerR']= PrimerR

#calculating the number of specimens
num_row = len(con_df.index)
num_spec = num_row - num_files
print("The total number of specimens is", num_spec)

#making final dataframe
finaldf= con_df.filter(items=['SpecimenID','TagFsequence','TagRsequence','PrimerF','PrimerR'])

#Changing dataframe to csv format
finaldf.to_csv("finalcsv.csv",index=False)
