
# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import os

import xlrd
import xlwt
import cPickle as pickle



def excle2pickle():
    infos = []
    # excelpath = u"F:\\方成平\\群星外国语学校成绩数据库\\2017-2018第1学期成绩\\高中\\10月月考\\成绩\\高一\\卡\\高一化学成绩单.xls"
    excelpath = u"F:\方成平\群星外国语学校成绩数据库\\2017-2018第1学期成绩\高中\\10月月考\成绩\高一\\最新高一10月月考成绩汇总1.xls"

    # data = xlrd.open_workbook(excelpath.encode("utf8"))
    data = xlrd.open_workbook(excelpath)

    table = data.sheets()[0]
    result = open("results.txt","wb")
    book = xlwt.Workbook(encoding='utf8',style_compression=0)
    sheet = book.add_sheet('sheet1',cell_overwrite_ok=True)




    nrows = table.nrows
    for i in range(nrows):
        print str(table.row_values(i)).decode("unicode_escape").encode("utf8")
        infos.append(str(table.row_values(i)).decode("unicode_escape").encode("utf8"))

    student_name = u"\\蔡海波\\"


    for info in infos[2:]:

        if student_name.split("\\")[1] == info.split(" ")[1].split(",")[0].split("'")[1]:

            s =  info
            result.writelines(infos[0]+"\r\n")
            result.write(infos[1]+"\r\n")
            result.write(s)
            sheet.write(0,0,infos[0])
            line1 = [i.strip() for i in infos[1].strip().split(',')]
            line2 = [i.strip() for i in info.strip().split(',')]
            for i in range(len(infos[1].split(','))):
                sheet.write(1,i,line1[i])
                sheet.write(2,i,line2[i])
            print 1

    result.close()
    book.save("./results.xls")
    print nrows
    pickle.dump(infos,open("infos.pkl","wb"),protocol=2)

def check_pickle():
    pickle_path = "./infos.pkl"
    infos = pickle.load(open(pickle_path,"rb"))

    print 1

def multi_query():


    root = u"F:\\方成平\\群星外国语学校成绩数据库\\"
    student_name = u"\\赵梦洁\\"

    results = []

    book = xlwt.Workbook(encoding='utf8', style_compression=0)
    sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)


    xlsxes = []
    exts = ['.xls','.xlsx']
    for path, dirs, files in os.walk(root, followlinks=True):
        dirs.sort()
        files.sort()
        for fname in files:
            fpath = os.path.join(path, fname)
            suffix = os.path.splitext(fname)[1].lower()
            if os.path.isfile(fpath) and (suffix.encode("utf8") in exts):
                # print fpath
                xlsxes.append(fpath)
                # yield (i, os.path.relpath(fpath, root), cat[path])

    c = 0

    for excelpath in xlsxes:
        infos = []
        c +=1
        print c,"-----", len(xlsxes),excelpath
        if not os.path.isfile(excelpath):
            break
        try:
            data = xlrd.open_workbook(excelpath)
        except:
            print "error excel"
            continue
        table = data.sheets()[0]


        nrows = table.nrows
        for i in range(nrows):
            # print str(table.row_values(i)).decode("unicode_escape").encode("utf8")
            infos.append(str(table.row_values(i)).decode("unicode_escape").encode("utf8"))
        for info in infos[2:]:
            try:
                info.split(" ")[1].split(",")[0].split("'")[1]
            except:
                continue

            if student_name.split("\\")[1] == info.split(" ")[1].split(",")[0].split("'")[1]:

                # s = info
                print excelpath

                print infos[0]
                print infos[1]
                print info
                results.append(excelpath)
                results.append(infos[0])
                results.append(infos[1])
                results.append(info)
                # sheet.write(0, 0, infos[0])
                # line1 = [i.strip() for i in infos[1].strip().split(',')]
                # line2 = [i.strip() for i in info.strip().split(',')]
                # for i in range(len(infos[1].split(','))):
                #     sheet.write(1, i, line1[i])
                #     sheet.write(2, i, line2[i])
                # print 1

    print 'over'
    c = 0
    for result in results:

        line = [i.strip() for i in result.strip().split(',')]
        for i in range(len(line)):
            sheet.write(c, i, line[i])
        c += 1
    book.save("./results.xls")



if __name__ == "__main__":
    # excle2pickle()
    # check_pickle()
    multi_query()




