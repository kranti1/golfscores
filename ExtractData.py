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
srcFile = 'Data2015.xlsm'
xl = pd.ExcelFile(srcFile)
xlSheets = xl.sheet_names


for sheet in xlSheets:
    if 'Schedule' not in sheet: #TODO: deal with DS_Store file type in directory
        df = xl.parse(sheet)
        df['SHEETNAME'] = sheet
        df.columns = [x.upper() for x in df.columns]
        frames = [combined_df, df]
        combined_df = pd.concat(frames)

#Final Combined File to a CSV
#combined_df.to_csv('./output.csv', encoding='ascii')
combined_df = combined_df.replace(u'\u2010\u2010\u2010', '')
combined_df = combined_df.replace(u'\u2010', '-')

combined_df.to_excel('./output.xls')

# Debug Helpers REMOVE when done
aa = combined_df['ATTACK ANG.'][:10]
sd = combined_df['SWING DIR.'][:10]

# replace "hyphen" with Unicode "minus"
aa = aa.replace(u'\u2010', u'\u2212')
sd = sd.replace(u'\u2010', u'\u2212')

# to convert a series to numeric try
#aa.convert_objects(convert_numeric=True)






#enable printing the entire dataframe
def print_full(x):
    pd.set_option('display.max_rows', len(x))
    print(x)
    pd.reset_option('display.max_rows')