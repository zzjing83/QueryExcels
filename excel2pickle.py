
# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import os

import xlrd
import cPickle as pickle


infos = []
def excle2pickle():

    excelpath = u"C:/Users/lenovo/Desktop/2018年教材及教务/考务/高中/6月月考/成绩/高二6月月考成绩汇总.xls"

    data = xlrd.open_workbook(excelpath)

    table = data.sheets()[0]
    result = open("results.txt","wb")
    nrows = table.nrows
    for i in range(nrows):
        print str(table.row_values(i)).decode("unicode_escape").encode("utf8")
        infos.append(str(table.row_values(i)).decode("unicode_escape").encode("utf8"))

    student_name = u"\\黄禹涵\\"


    for info in infos[2:]:

        if student_name.split("\\")[1] == info.split(" ")[1].split(",")[0].split("'")[1]:

            s =  info
            result.writelines(infos[0]+"\r\n")
            result.write(infos[1]+"\r\n")
            # result.write("\n")
            result.write(s)



            print 1

    result.close()
    print nrows
    pickle.dump(infos,open("infos.pkl","wb"),protocol=2)


def check_pickle():
    pickle_path = "./infos.pkl"
    infos = pickle.load(open(pickle_path,"rb"))

    print 1



def multi_query():


    root = u"F:\\方成平\\月考类\\"
    exts = ['.xls','.xlsx']
    for path, dirs, files in os.walk(root, followlinks=True):
        dirs.sort()
        files.sort()
        for fname in files:
            fpath = os.path.join(path, fname)
            suffix = os.path.splitext(fname)[1].lower()
            if os.path.isfile(fpath) and (suffix.encode("utf8") in exts):
                print fpath
                # yield (i, os.path.relpath(fpath, root), cat[path])










if __name__ == "__main__":
    # excle2pickle()
    # check_pickle()
    multi_query()




