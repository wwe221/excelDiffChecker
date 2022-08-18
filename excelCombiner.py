import pandas as pd
import tkinter as tk
import tkinter.filedialog as fd

root = tk.Tk()
filez = fd.askopenfilenames(parent=root, title='Choose a file')
print(filez)
# filenames
excel_names = filez

# read them in
excels = [pd.ExcelFile(name) for name in excel_names]

# turn them into dataframes
frames = [x.parse(x.sheet_names[0], header=None, index_col=None)
          for x in excels]

# delete the first row for all frames except the first
# i.e. remove the header row -- assumes it's the first
frames[1:] = [df[1:] for df in frames[1:]]

# concatenate them..
combined = pd.concat(frames)
fileExtension = filez[0].split("/")[-1].split(".")[-1]
onlyName = filez[0].split("/")[-1].split("."+fileExtension)[0]
exportFileName = onlyName+"_Combined.xlsx"
# write it out
combined.to_excel(exportFileName, header=False, index=False)
