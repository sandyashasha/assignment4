import xlsxwriter
d = {'Table Name': ['column 1', 'column 2', 'column 3', 'column 4', 'column 5', 'column 6']}

length_dict = {key: len(value) for key, value in d.items()}
length_key = length_dict['Table Name']
print(length_key)
alpha = ["B", "C", "E", "F", "H", "I", "K", "L", "N", "O", "Q", "R","T","U","W"]

workbook = xlsxwriter.Workbook('C:/Users/sandya.nandakumar/Desktop/sassignment2/working2.xlsx')
worksheet = workbook.add_worksheet()

for a in alpha:
    worksheet.set_column('A:Z', 14)
    format1 = workbook.add_format({'border': 2})
    print d.values()
    for i in range(2, length_key):
        for k, v in d.items():
            worksheet.write(a + str(i), v[i], format1)



cell_format = workbook.add_format()
cell_format.set_pattern(1)  # This is optional when using a solid fill.
cell_format.set_bg_color('white')
worksheet.write('A1:Z1', ' ', cell_format)


"""merge_format = workbook.add_format({
    'bold': 2,
    'border': 2,
    'align': 'center',
    'valign': 'vcenter'})


for key in d:
    worksheet.merge_range('B2:B' +str((length_key)-1),key,merge_format)"""
workbook.close()
