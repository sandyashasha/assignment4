import xlsxwriter
workbook = xlsxwriter.Workbook("C:\Users\sandya.nandakumar\Desktop\sassignment2\important.xlsx")
worksheet = workbook.add_worksheet()




cellist = [['F', 1, 1, 1, 13, 3, 1], ['A', 3, 1, 3, 10, 5, 1], ['a', 5, 1, 5, 10, 7, 1], ['aa', 7, 1, 7, 10, 9, 1]
    , ['O', 9, 1, 9, 5, 11, 1], ['P', 11, 1, 11, 7, 13, 1]
    , ['hi', 13, 1, 13, 5, 15, 1], ['h', 15, 1, 15, 5, 17, 1]
    , ['hey', 17, 1, 17, 5, 19, 1]]

tablelist = [["F", "A", "B", "C"], ["A", "a", "b"], ["a", "aa", "bb"], ["aa", "O", "P"], ["P", "hi", "h"],
             ["h", "hey"], ["bb", "5", "10"], ["b", "x"], ["B", "a", "d"]]

columnlist = [["F", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10"],
              ["A", "A1", "A2", "A3", "A4", "A5", "A6", "A7"], ["a", "a1", "a2", "a3", "a4", "a5", "a6", "a7"],
              ["aa", "aa1", "aa2", "aa3", "aa4", "aa5", "aa6", "aa7"], ["O", "O1", "O2"], ["P", "P1", "P2", "P3", "P4"],
              ["hi", "hi1", "hi2"], ["h", "h1", "h2"], ["hey", "hey1", "hey2"]]

def collid(intv, cr, cc, er, ec, nc, nr):
    # pass
    list=[]
    list.append(intv)
    list.append(cr)
    list.append(cc)
    list.append(er)
    list.append(ec)
    list.append(nc)
    list.append(nr)
    """print(list)"""

c = 1
r = 1
cr = r
cc = c
for i in range(0, len(columnlist)):
    count = len(columnlist[i])
    ec = cc
    er = count + 1 + r
    nc = cc + 2
    nr = 1
    collid(columnlist[i][0], cc, cr, ec, er, nc, nr)
    cc = nc
    cr = nr



first = []
for i in range(0, len(tablelist)):
    for j in range(0, len(columnlist)):
        if columnlist[i][0] == tablelist[j][0]:
            first.append(columnlist[j])
print(first)

format2 = workbook.add_format({'border': 2})
format1 = workbook.add_format({'border': 1})

for k in range(0, len(first)):
    worksheet.write(0,0,first[0][0],format2)
    worksheet.write(1,0, first[0][1],format1)
    worksheet.write(2,0, first[0][2],format1)
    worksheet.write(3,0, first[0][3],format1)
    worksheet.write(4,0, first[0][4],format1)
    worksheet.write(5,0, first[0][5],format1)
    worksheet.write(6,0, first[0][6],format1)
    worksheet.write(7,0, first[0][7],format1)
    worksheet.write(8,0,first[0][8],format1)

    worksheet.write(0, 2, first[1][0],format2)
    worksheet.write(1, 2, first[1][1],format1)
    worksheet.write(2, 2, first[1][2],format1)
    worksheet.write(3, 2, first[1][3],format1)
    worksheet.write(4, 2, first[1][4],format1)
    worksheet.write(5, 2, first[1][5],format1)
    worksheet.write(6, 2, first[1][6],format1)
    worksheet.write(7, 2, first[1][7],format1)

    worksheet.write(0, 4, first[2][0],format2)
    worksheet.write(1, 4, first[2][1],format1)
    worksheet.write(2, 4, first[2][2],format1)
    worksheet.write(3, 4, first[2][3],format1)
    worksheet.write(4, 4, first[2][4],format1)
    worksheet.write(5, 4, first[2][5],format1)
    worksheet.write(6, 4, first[2][6],format1)
    worksheet.write(7, 4, first[2][7],format1)


    worksheet.write(0, 6, first[3][0],format2)
    worksheet.write(1, 6, first[3][1],format1)
    worksheet.write(2, 6, first[3][2],format1)
    worksheet.write(3, 6, first[3][3],format1)
    worksheet.write(4, 6, first[3][4],format1)
    worksheet.write(5, 6, first[3][5],format1)
    worksheet.write(6, 6, first[3][6],format1)
    worksheet.write(7, 6, first[3][7],format1)

    worksheet.write(0, 8, first[4][0],format2)
    worksheet.write(1, 8, first[4][1],format1)
    worksheet.write(2, 8, first[4][2],format1)

    worksheet.write(4, 8, first[5][0], format2)
    worksheet.write(5, 8, first[5][1], format1)
    worksheet.write(6, 8, first[5][2], format1)
    worksheet.write(7, 8, first[5][3], format1)
    worksheet.write(8, 8, first[5][4], format1)

    worksheet.write(4, 10, first[6][0], format2)
    worksheet.write(5, 10, first[6][1], format1)
    worksheet.write(6, 10, first[6][2], format1)









workbook.close()