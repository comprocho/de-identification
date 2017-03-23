import pandas as pd
import numpy as np
import re

df1 = pd.read_excel('Breast TMA test data2_IDENTIFIED.xlsx') # Imports Breast TMA test file
df2 = pd.read_excel('HTSR-Patient-Database-10-13-2016-xlsx.xlsx') #Imports HTSR-Patient-Database.xlsx

stems = pd.read_excel('Drug Suffix.xlsx')
tlist = pd.read_excel('treatment_dict.xlsx')
stems= stems['Stems']
tlist = tlist['dict']

#Creates column 'Index#' by importing Index# from HTSR-Patient-Database.xlsx
df1['Index#'] = pd.Series()
x1 = df1.set_index(['Medical_Record_Number'])['Index#']
x2 = df2.set_index(['MRN'])['Index#']
x1.update(x2)
df1['Index#'] = x1.values.astype(int)

#Creates new dataframe containing only date-related columns
date_column = ['Date_of_Last_Contact',
               'Date_of_Initial_Diagnosis',
               'Hx_Date_Recurrence',
               'Rx_Start_Date',
               'Date_RT_Started',
               'Date_RT_Ended']
date_only = df1[date_column]
do = date_only.copy()

#Exports column 'Date_of_Birth' to new series
dob = df1['Date_of_Birth']

#Parses all values in different form of datetime format to the unified datetime format
for d in do:
    do[d] = pd.to_datetime(do[d], errors='coerce')

#Creates a dictionary for a list of column headers in dataframe do
dict_do = {1: 'Date_of_Last_Contact',
           2: 'Date_of_Initial_Diagnosis',
           3: 'Hx_Date_Recurrence',
           4: 'Rx_Start_Date',
           5: 'Date_RT_Started',
           6: 'Date_RT_Ended'}
#Creates a dictionary for a list of column headers that will be inserted into dataframe do
dict_age = {1: 'Age_Last_Contact',
            2: 'Age_Initial_Diagnosis',
            3: 'Age_Hx_Recurrence',
            4: 'Age_Rx_Start',
            5: 'Age_RT_Started',
            6: 'Age_Rx_Ended'}
#Loops through dataframe columns, calculates difference between date and date of birth, and converts to number of years
for i in dict_do:
    td = (do[dict_do[i]].sub(dob))/365
    td = (td / np.timedelta64(1, 'D')).astype(float)
    do[dict_age[i]] = td.round(3)

#Drops columns containing sensitive data from dataframe do
do.drop(do.columns[0:6], axis=1, inplace=True)

#Drops columns containing sensitive data from dataframe df1
df1.drop(df1.columns[[0,1,2,3,5,9,34,35,39,40]], axis=1, inplace=True)

#Merges df1 and do
df1 = pd.concat([df1, do], axis=1)

#Converts the name of hospital center to code
location_code= ['A','B','C','D','E']
df1['Location_of_Radiation_Treatment_Desc'] = df1['Location_of_Radiation_Treatment_Desc'].map({
        'Hospital 1' : 'A',
        'Hospital 2' : 'B',
        'Hospital 3' : 'C',
        'Hospital 4' : 'D',
        'Hospital 5' : 'E',})
#All texts converted to XX if they are not in location_code
df1.loc[~df1["Location_of_Radiation_Treatment_Desc"].isin(location_code), "Location_of_Radiation_Treatment_Desc"] = "XX"

text_column =['Text_Ancillary_Therapy', 'Text_Chemotherapy', 'Text_Hormone_Therapy', 'Text_Immunotherapy', 
              'Text_Other_Radiation', 'Text_Other_Treatment','Text_Radiation_Therapy','Text_Treatment']
te = df1[text_column].copy()
te = te.apply(lambda x: x.astype(str).replace(np.nan, ' ', regex=True).str.lower())

df1.drop(df1.columns[[54,55,56,57,58,59,60,61]], axis=1, inplace=True)

# Functions to extract data
def date_extract(object):
    global te
    te['Date Ranges'] =te[object].str.extract(r'(\d+/\d+/\d+-\d+/\d+/\d+)', expand=True)
    te[object] = te[object].str.replace(r'(\d+/\d+/\d+-\d+/\d+/\d+)', '')

    te['Date Ranges Two'] = te[object].str.extract(r'(\d+\/\d{4} to \d+\/\d{4})', expand=True)
    te[object] = te[object].str.replace(r'(\d+\/\d{4} to \d+\/\d{4})', '')

    te['Single Date'] = te[object].str.extract(r'(\d+/\d+/\d+)', expand=True)
    te[object] = te[object].str.replace(r'(\d+/\d+/\d+)', '')

    te['Single Date Two'] = te[object].str.extract(r'(\d+\/\d{4})', expand=True)
    te[object] = te[object].str.replace(r'(\d+\/\d{4})', '')

    te['Date Range'] = pd.concat([te['Date Ranges'].dropna(), te['Date Ranges Two'].dropna()]).reindex_like(te)
    te['Date Range' +' (' + object + ')'] = te['Date Range'].str.replace(' to ', '-')

    te['Date'+' (' + object + ')'] = pd.concat([te['Single Date'].dropna(), te['Single Date Two'].dropna()]).reindex_like(te)
    
    te = te.drop(['Single Date', 'Single Date Two','Date Ranges', 'Date Ranges Two', 'Date Range'], axis=1)

    te[object] = te[object].str.replace(r'(\s){2,10}', ' ')
# Function to extract drug name
def drug_extract(object):
    pat = r'\b(\w*(?:{})\w*)\b'.format(stems.str.cat(sep='|'))
    try:
        te['Drugs'+' (' + object + ')'] = te[object].str.extractall(pat, flags=re.I).unstack().apply(lambda x:', '.join(x.dropna()), axis=1)
    except:
        te[object] = te[object].str.replace('nan', 'none')
    
    te[object] = te[object].str.replace(pat, ' ')
    te[object] = te[object].str.rjust(2, ' ')    
    te[object] = te[object].str.replace(',|\.|:|\+|(\s)(-)|(-)(\s)', ' ')
    te[object] = te[object].str.replace(r'(lt & rt)', ' bilateral ')
    te[object] = te[object].str.replace(r'(\s)(l)(\s)|(\s)(lt)(\s)|^(lt)(\s)', ' left ')
    te[object] = te[object].str.replace(r'(\s)(r)(\s)|(\s)(rt)(\s)|^(rt)(\s)', ' right ')
    te[object] = te[object].str.replace(r'(\s){2,10}', ' ')
# Function to extract treatment types
def treatment_extract(object):
    tpat = r'\b(\w*(?:{})\w*)\b'.format(tlist.str.cat(sep='|'))
    try:
        te['treatment' +' (' + object + ')'] = te[object].str.extractall(tpat, flags=re.I).unstack().apply(lambda x:', '.join(x.dropna()), axis=1)
    except:
        te[object] = te[object].str.replace('nan', 'none')
    
    te[object] = te[object].str.replace(tpat, '').str.replace(r'(\s){2,10}', ' ').str.strip()

for i in text_column:
    date_extract(i)
    drug_extract(i)
    treatment_extract(i)

df1 = pd.concat([df1, te], axis=1)

cols = df1.columns.tolist()
cols = cols[-42:] + cols[:-42]
df1 = df1[cols].set_index('Index#')

df1.to_excel('de_identified_document.xlsx') #Exports to xlsx type