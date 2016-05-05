# This module processes the files and creates a data table of the golf score data
from os import listdir
import pandas as pd



# to run this file in the console do this.
# execfile('/Users/kartiks/Documents/github/GolfScores/ExtractData.py')
# execfile('C:\docs\github\golfscores\ExtractData.py')
'''
path = './Data'
filelist = listdir(path)
filelist = [path + '/' + f for f in filelist]
'''

df = pd.DataFrame()
combined_df = pd.DataFrame()


#Opening using XL directly
srcFile = 'Data.xlsm'
xl = pd.ExcelFile(srcFile)
xlSheets = xl.sheet_names


for sheet in xlSheets:
    if 'Schedule' not in sheet: #TODO: deal with DS_Store file type in directory
        df = xl.parse(sheet)
        df['SHEETNAME'] = sheet
        df.columns = [x.upper() for x in df.columns]
        frames = [combined_df, df]
        combined_df = pd.concat(frames)

# Convert string to Numeric
combined_df = combined_df.replace({u'\u2010': '-'}, regex=True)


# remap series types
combined_df['DATE'] = combined_df['DATE'].astype('datetime64[ns]')
combined_df['LOW POINT'] = combined_df['LOW POINT'].astype('string')
combined_df['SHEETNAME'] = combined_df['SHEETNAME'].astype('string')
combined_df['SIDE'] = combined_df['SIDE'].astype('string')
combined_df['SIDE TOT.'] = combined_df['SIDE TOT.'].astype('string')
combined_df['CLUB'] = combined_df['CLUB'].astype('string')

#cast the rest of the columns to int 
all_columns= set (combined_df.columns.values)
col = {u'DATE',u'LOW POINT',u'SHEETNAME',u'SIDE',u'SIDE TOT.' ,u'CLUB'}
col = all_columns - col
for h in col:
    combined_df[h] = pd.to_numeric(combined_df[h], errors='coerce')


combined_df.to_excel('./output.xls')


#enable printing the entire dataframe
def print_full(x):
    pd.set_option('display.max_rows', len(x))
    print(x)
    pd.reset_option('display.max_rows')