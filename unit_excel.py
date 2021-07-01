import pandas as pd
import glob
import xlrd
import openpyxl as xl

all_data = pd.DataFrame()

### Thư viện này hỗ trợ xử lý dữ liệu dạng Excel ###
### Hepl:
###
types = ('*.xls', '*.xlsx') # the tuple of file types
files_grabbed = []
for files in types:
    files_grabbed.extend(glob.glob(files))
error_files = []

for f in files_grabbed:
    try:
        print(f)
        df = pd.read_excel(f, skiprows=9, usecols='A:L', header=None)

        print(df.head(5))
        all_data = all_data.append(df)
    except Exception as e:
        error_files.append(f)

print('Danh sách file lỗi:')
print(error_files)
for file in error_files:
    wb = xl.load_workbook(file)
    wb.save("file.xls")
# now save the data frame
writer = pd.ExcelWriter('output.xlsx')
all_data.to_excel(writer,'sheet1')
writer.save()