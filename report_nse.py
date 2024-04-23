import pandas as pd
import os
import numpy as np

# Step 1: Define the folder containing the Excel files
folder_path = 'D:\\nsedata\\new\\DWN_DATA'

# Step 2: Get a list of all Excel files in the folder
excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]

# Step 3: Iterate through each Excel file, read the file, add filename as a column, and append to a list
dfs = []
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    # Read the Excel file
    df = pd.read_excel(file_path)
    df['close'] = df['close'].astype(int)
    sma10 = df.iloc[:10]['close'].mean()
    sma20 = df.iloc[:20]['close'].mean()
    sma50 = df.iloc[:50]['close'].mean()
    sma100 = df.iloc[:100]['close'].mean()
    sma200 = df.iloc[:200]['close'].mean()
    minlow5 = df.iloc[:5]['low'].min()
    maxhigh5 = df.iloc[:5]['high'].max()
    # Add filename as a column without the .xlsx extension
    filename = os.path.splitext(file)[0]
    df['Filename'] = filename
    df['sma10'] = sma10
    df['sma20'] = sma20
    df['sma50'] = sma50
    df['sma100'] = sma100
    df['sma200'] = sma200
    df['minlow5'] = minlow5
    df['maxhigh5'] = maxhigh5
    df['smacl'] = np.where((df['close'] > df['sma10']) & (df['sma10'] > df['sma20']) & (df['sma20'] > df['sma50']) & (df['sma50'] > df['sma100']) & (df['sma100'] > df['sma200']), 'Y', 'N')
    df['smacs'] = np.where((df['close'] < df['sma10']) & (df['sma10'] < df['sma20']) & (df['sma20'] < df['sma50']) & (df['sma50'] < df['sma100']) & (df['sma100'] < df['sma200']), 'Y', 'N')
    df['green'] = np.where((df['close'] > df['open']) & (df['close'] > df['prev_close']), 'Y', 'N')
    df['maxhigh5_close%'] = df['close'] / df['maxhigh5']*100
    df['minlow5_close%'] = df['close'] / df['minlow5']*100
    df['52whigh_close%'] = df['close'] / df['hi_52_wk']*100
    df['52wlow_close%'] = df['close'] / df['lo_52_wk']*100

    df1 = df.iloc[:1]
    
    
    # Append to the list
    dfs.append(df1)

# Step 4: Combine all DataFrames into a single DataFrame
cd = pd.concat(dfs, ignore_index=True)
cd1 = cd[(cd['smacl'] == 'Y') & (cd['green'] == 'Y')& (cd['maxhigh5_close%'] < 95)& (cd['minlow5_close%'] < 105)]
cd2 = cd[(cd['smacs'] == 'Y') & (cd['green'] == 'N')& (cd['maxhigh5_close%'] > 95)& (cd['minlow5_close%'] > 105)]
# Step 5: Write the combined data to a new Excel file
cd1.to_excel('long.xlsx')
cd2.to_excel('short.xlsx')
cd.to_excel('all.xlsx')

print("Combined report has been created successfully.")
