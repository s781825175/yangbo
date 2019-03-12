#coding=utf-8
import csv
import xlrd
import os
import time
import sys

def change(num):
    report = dirc + "\\WORKING\\{S1}".format(S1 = num[0])
    gene_name = ''
    title = num[1] + "-肠康计划-血451-"
    gene = '-阴性'
    for i in os.listdir(report):
        if 'Annokb' in i:
            with open(report + "\\" + i) as heat:
                point_reader = csv.reader(heat)
                heat_point = ['KRAS','NRAS','BRAF','EGFR','ERBB2','PIK3CA','POLE','POLD1','APC']
                for row in point_reader:
                    if row[1] in heat_point and row[10] == '1':
                        gene_name += '-' + row[1] + '-' + row[6].split('.')[-1]
                        gene = gene_name
            os.rename(report + '\\' + i, report + '\\' + title + (i.split('.',1)[1]).split('-')[0] + '.csv')
            os.system ("copy %s %s" % (report + '\\' + title + (i.split('.',1)[1]).split('-')[0] + '.csv', dirc + "\\WORKING"))

        
        if 'cnv' in i:
            with open(report + "\\" + i) as cnv_file:
                cnv_reader = csv.reader(cnv_file)
                cnv_name = '-无扩增-'
                cnv_count = 0
                for row in cnv_reader:
                    if cnv_reader.line_num == 1:
                        continue
                    if float(row[1]) >= 3:
                        cnv_count += 1
                        cnv_name = ''
                        cnv_name += row[0] + '-扩增-'
                    if float(row[1]) <= 1 and row[0] != 'ATRX':
                        cnv_count += 1
                        cnv_name = ''
                        cnv_name += row[0] + '-缺失-'
                if cnv_count >= 2:
                    cnv_name = cnv_name.split('-')[0] + '等' + str(cnv_count) + '扩增-'
                cnv = '-' + cnv_name
            print(report + '\\' + title + "cnv.csv")
            os.rename(report + '\\' + i, report + '\\' + title + "cnv.csv")
            os.system ("copy %s %s" % (report + '\\' + title + "cnv.csv", dirc + "\\WORKING"))
        
        if 'trans' in i:
            xlsfile = report + "\\" + i
            book = xlrd.open_workbook(xlsfile)
            sheet0 = book.sheet_by_index(0)
            sheet_name = book.sheet_names()[0]
            sheet1 = book.sheet_by_name(sheet_name)
            nrows = sheet0.nrows
            trans_name = ''
            dmmr = ['MSH2','MSH6','MLH3','PMS2']
            mmr = 'pmmr'
            for row in range(nrows):
                if '致病' in sheet0.cell_value(row, 7) and sheet0.cell_value(row, 0) in dmmr:
                    mmr = 'dmmr'
                if '致病' in sheet0.cell_value(row, 7):
                    trans_name += sheet0.cell_value(row, 0) + '-' + sheet0.cell_value(row, 7) + '-'
            if trans_name == '':
                trans = 'pmmr'
            else:
                trans = trans_name + mmr
            print(report + '\\' + title + i.split('trans_')[1])
            os.rename(report + '\\' + i, report + '\\' + title + i.split('trans_')[1])
            os.system ("copy %s %s" % (report + '\\' + title + i.split('trans_')[1], dirc + "\\WORKING"))
        
    info = gene + cnv + trans
    print(info)
    return info



dirc = sys.argv[0][:-18]
print(dirc)
report = []
for i in os.listdir(dirc + "\\WORKING"):
    if i.isdigit():
        report.append(i)
for i in os.listdir(dirc):
    if "上机信息" in i and '$' not in i:
        All_file = dirc + "\\" + i

book = xlrd.open_workbook(All_file)
sheet0 = book.sheet_by_index(0)
sheet_name = book.sheet_names()[0]
sheet1 = book.sheet_by_name(sheet_name)
nrows = sheet0.nrows
pair=[]
for row in range(nrows):
    for num in report:
        if ('_' + num) in sheet0.cell_value(row, 1):
            pair.append([num,sheet0.cell_value(row, 3)])

for i in pair:
    info = change(i)
    i.append(info)


working = os.listdir(dirc + "\\WORKING")
for i in working:
    for j in pair:
        if i.isdigit():
            continue
        if i.split('_')[0] in j and 'docx' in i:
            print(dirc + '\\' + i.split('.')[0] + j[2] + ".docx")
            os.rename(dirc + "\\WORKING\\" + i, dirc + "\\WORKING\\" + i.split('.')[0] + '-血451' + j[2] + ".docx")
