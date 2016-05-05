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

combined_df['ATTACK ANG.'] = pd.to_numeric(combined_df['ATTACK ANG.'], errors='coerce')
combined_df['BALL SPEED'] = pd.to_numeric(combined_df['BALL SPEED'], errors='coerce')
combined_df['CARRY'] = pd.to_numeric(combined_df['CARRY'], errors='coerce')
combined_df['CLUB PATH'] = pd.to_numeric(combined_df['CLUB PATH'], errors='coerce')
combined_df['CLUB SPEED'] = pd.to_numeric(combined_df['CLUB SPEED'], errors='coerce')
combined_df['DYN. LOFT'] = pd.to_numeric(combined_df['DYN. LOFT'], errors='coerce')
combined_df['FACE ANG.'] = pd.to_numeric(combined_df['FACE ANG.'], errors='coerce')
combined_df['FACE TO PATH'] = pd.to_numeric(combined_df['FACE TO PATH'], errors='coerce')


#combined_df[''] = pd.to_numeric(combined_df[''], errors='coerce')

type_int = ('HEIGHT' , 'LAND. ANG.' ,'LAUNCH ANG.')
type_other = ('CLUB' , 'SIDE')
type_all = type_int + type_other


for h in type_int:
    combined_df[h] = pd.to_numeric(combined_df[h], errors='coerce')


'''

SMASH FAC.
SPIN AXIS
SPIN LOFT
SPIN RATE
STROKE NO
SWING DIR.
SWING PL.
TOTAL
'''


#combined_df = combined_df.convert_objects(convert_numeric=True)


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