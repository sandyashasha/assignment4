

columnlist=[["F","F1","F2","F3","F4","F5","F6","F7","F8","F9","F10"],["A","A1","A2","A3","A4","A5","A6","A7"]
    ,["a","a1","a2","a3","a4","a5","a6","a7"],["aa","aa1","aa2","aa3","aa4","aa5","aa6","aa7"]
    ,["P","P1","P2","P3","P4"],["h","h1","h2"],["hey","hey1","hey2"]]


tablelist=[["F","A","B","C"],["A","a","b"],["a","aa","bb"],["aa","O","P"],["P","hi","h"],["h","hey"],["bb","5","10"]
    ,["b","x"],["B","a","d"]]


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





