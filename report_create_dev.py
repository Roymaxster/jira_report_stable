# -*- coding: utf-8 -*-
from selenium import webdriver
import os
import xlwt


def files():
    wd = webdriver.Firefox()

    def general(folder, report, param1, ts, name):

        wd.get("file:///" + folder + "/" + report[param1])
        lb_add = wd.find_element_by_xpath("//table[1]/tbody/tr[2]/td/a").text
        name.append(lb_add)
        number_int = wd.find_element_by_xpath(
            "html/body/table[1]/tbody/tr[3]/td/strong[1]").text
        ts_cache = []
        i = 0
        while i < int(number_int):
            i = i + 1
            add = (wd.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[" + str(i) + "]/td[23]").text)
            #условие пропуска пустых значений
            if add == '':  
                pass
            else:
                ts_cache.append(int(add))

        ts.append(sum(ts_cache))

    def xls(title, summa, count):
        font0 = xlwt.Font()
        font0.name = 'Times New Roman'
        font0.bold = True
        font0.italic = True

        style0 = xlwt.easyxf()
        style0.font = font0

        style1 = xlwt.easyxf("", "0.00%")
        
        style2 = xlwt.easyxf("", "0.00")

        wb = xlwt.Workbook()
        ws = wb.add_sheet('Report')
        i = -1
        table = {1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6:
                 "F", 7: "G", 8: "H", 9: "I", 10: "J", 11: "K",
                 12: "L", 13: "M", 14: "N", 15: "O", 16: "P",
                 17: "Q", 18: "R", 19: "S", 20: "T", 21: "U",
                    22: "V", 23: "W", 24: "X", 25: "Y", 26: "Z",
                     27: "AA", 28: "AB", 29: "AC", 30: "AD"}
        ws.write(1, 0, "Total time", style0)
        ws.write(2, 0, "Percent", style0)
        while i < (count - 1):
            i += 1
            ws.col(i).width = 230 * 20
            task = title[i]
            task = task.replace('GameDevServer', '')
            task = task.replace('for', '')
            task = task.replace('(Globo-Tech)', '')
            ws.write(0, i + 1, task, style0)
            ws.write(1, i + 1, xlwt.Formula(str(str(summa[i]) + "/" + "3600")), style2)
            ws.write(2, i + 1, xlwt.Formula(
                str(str((summa[i])) + "/" + str(sum(summa)))), style1)

        ws.write(0, count + 1, "Total", style0)
        ws.write(1, count + 1, xlwt.Formula(
            str("SUM" + "(" + (table[2]) + "2" + ":" + str(table[count + 1] + "2" + ")"))))
        ws.write(2, count + 1, xlwt.Formula(
            str("SUM" + "(" + (table[2]) + "3" + ":" + str(table[count + 1] + "3" + ")"))), style1)
        
        wb.save('jira.xls')

    directory = os.getcwd()
    files = os.listdir(directory)
    reports = filter(lambda x: x.endswith('.html'), files)
    counter = len(reports)

    ts_test = []
    lable = []
    k = -1
    while k < (int(counter) - 1):
        k += 1
        general(directory, reports, k, ts_test, lable)
    pass
    xls(lable, ts_test, counter)
    wd.quit()

files()
