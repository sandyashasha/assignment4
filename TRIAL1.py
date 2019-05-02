import xlsxwriter
def collid(intv, cr, cc, er, ec, nc, nr):
    # pass
    list = []
    list.append(intv)
    list.append(cr)
    list.append(cc)
    list.append(er)
    list.append(ec)
    list.append(nc)
    list.append(nr)
    print(list)
workbook = xlsxwriter.Workbook("C:\Users\sandya.nandakumar\Desktop\sassignment2\working.xlsx")
worksheet = workbook.add_worksheet()
col = [['t_fnpl', 'npe', 'nt', 'npa', 'nt_prov']]
worksheet.set_column('A:Z', 14)
format2 = workbook.add_format({'border': 2})
format1 = workbook.add_format({'border': 1})
for i in range(0,len(col)):
    for j in range(0,len(col)):
        worksheet.write("B1",col[0][0],format2)
        worksheet.write("B2",col[0][1],format1)
        worksheet.write("B3",col[0][2],format1)
        worksheet.write("B4",col[0][3],format1)
        worksheet.write("B5",col[0][4],format1)
workbook.close()



c = 1
r = 1
cr = r
cc = c
col = [['t_fnpl', 'npe', 'nt', 'npa', 'nt_prov'], ['t_fpa', 'npe', 'npa', 'n', 'pa', 'np', 'pbea'],['npe','id','name','place']]
for i in range(0, len(col)):
    count = len(col[0][i])
    ec = cc
    er = count + 1 + r
    nc = cc + 2
    nr = 1
    collid(col[i][0], cc, cr, ec, er, nc, nr)
    cc = nc
    cr = nr
