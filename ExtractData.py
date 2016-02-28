# This module processes the files and creates a data table of the golf score data
from os import listdir
import pandas as pd

# to run this file in the console do this.
# execfile('/Users/kartiks/Documents/github/GolfScores/ExtractData.py')

path = './Data'
filelist = listdir(path)
filelist = [path + '/' + f for f in filelist]
df = pd.DataFrame()
combined_df = pd.DataFrame()


for fname in filelist:
    if 'Store' not in fname: #TODO: deal with DS_Store file type in directory
        df = pd.read_csv(fname)
        df['FILENAME'] = fname
        df.columns = [x.upper() for x in df.columns]
        frames = [combined_df, df]
        combined_df = pd.concat(frames)

#Final Combined File to a CSV
combined_df.to_csv('./output.csv')

#enable printing the entire dataframe
def print_full(x):
    pd.set_option('display.max_rows', len(x))
    print(x)
    pd.reset_option('display.max_rows')