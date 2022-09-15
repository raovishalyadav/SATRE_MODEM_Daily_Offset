'''
Copyright (C) 2019 Vishal Yadav
'''
import easygui
import glob
import os
import re
import xlsxwriter
import numpy as np
import natsort as ns
from collections import Counter
from collections import OrderedDict

#To create Daily_Offset folder
if os.path.isdir(os.environ['USERPROFILE'] + '\Desktop\Daily_Offset') == False:
    os.mkdir(os.environ['USERPROFILE'] + '\Desktop\Daily_Offset')
os.chdir(os.environ['USERPROFILE'] + '\Desktop\Daily_Offset')

#To merge multiple text filetypes on the basis of time-slots and to filter out date ,time and offset data
filename = easygui.fileopenbox(title = 'Please locate the text files',default="./Desktop/*.txt", filetypes = '*.txt',multiple = True)
def nine_min(file,low,high,modem):
    date = []
    f1 = open(file+".txt","a+")
    for x in range(1,len(filename)+1):
        with open(filename[x-1], "r") as infile:
            lines = infile.readlines()
            for x in lines:
                if (low != 30) and (high != 40) and (low != 0) and (high != 10):
                    if ("%Rx1" in x) and (len(x) > 59) and (x[141:149] != "   0.000") and ((x[141:142] == " ") or (x[141:142] == "-")) and ((int(x[19:21]) in range (low,high))):
                        date.append((x[13:15]+x[10:12]))
                        f1.write(str(x[5:15])+" "+str(x[16:24])+" "+str(x[141:149])+"\n")

                if (low == 0) and (high == 10):
                    if ("%Rx1" in x) and (len(x) > 59) and (x[141:149] != "   0.000") and ((x[141:142] == " ") or (x[141:142] == "-")) and (x.split(';')[11] == modem):
                        date.append((x[13:15]+x[10:12]))
                        f1.write(str(x[5:15])+" "+str(x[16:24])+" "+str(x[141:149])+"\n")

                if (low == 30) and (high == 40):
                    if ("%Rx1" in x) and (len(x) > 59) and (x[141:149] != "   0.000") and ((x[141:142] == " ") or (x[141:142] == "-")) and (x.split(';')[11] == modem):
                        date.append((x[13:15]+x[10:12]))
                        f1.write(str(x[5:15])+" "+str(x[16:24])+" "+str(x[141:149])+"\n")

                if (low == 30) and (high == 40):
                    if ("%Rx1" in x) and (len(x) > 59) and (x[141:149] != "   0.000") and ((x[141:142] == " ") or (x[141:142] == "-")) and ((x.split(';')[11] == "422")and(x.split(';')[11] == "397")and(x.split(';')[11] == modem)):
                        date.append((x[13:15]+x[10:12]))
                        print((x.split(';')[11]))
                        f1.write(str(x[5:15])+" "+str(x[16:24])+" "+str(x[141:149])+"\n")

    date = Counter(date)
    s = [0]
    for k,v in date.items():
        v = v + s[-1]
        s.append(v)
        f1.writelines(str(k)+" "+str(v)+"/")
    f1.close()

nine_min('00-09(423)',0,10,'423')
nine_min('10-19',10,20,'')
nine_min('30-39(422)',30,40,'422')
nine_min('40-49',40,50,'')

#To create 12 noon to 12 noon files
def average_12_to_12(file1,file11):
    f2 = open(file1 + ".txt","r")
    lines1 = f2.readlines()
    dates = lines1[-1].count('/')
    sort = lines1[-1]
    lines1 = lines1[:-1]
    no = []
    so = []
    keep = []
    discard = []
    output = [re.split(r'\n', s)[0].strip() for s in lines1]
    for i in range(1, dates+1):
        for line_no, y in enumerate(lines1):
            if (int(y[5:7]) == int(sort.split('/')[i-1].split(' ')[0][2:4])) and (int(y[8:10]) == int(sort.split('/')[i-1].split(' ')[0][0:2])) and (int(y[11:13]) >= 12):
                so.append(int(sort.split('/')[i-1].split(' ')[0]))
                no.append(line_no+1)
                break

    no.append(len(lines1))
    no = sorted(set(no))
    no[-1] = int(no[-1])
    sort = list(sort.split('/'))[:-1]

    for first, second in zip(so, so[1:]):
        first = int(first)
        second = int(second)

        if (len(str(first)) == 4) and (len(str(second)) == 4):
            first = str(first)
            second = str(second)
            M1 = int(first[2:4])
            D1 = int(first[0:2])
            M2 = int(second[2:4])
            D2 = int(second[0:2])

        if (len(str(first)) == 3) and (len(str(second)) == 4):
            first = str(first)
            second = str(second)
            M1 = int(first[1:3])
            D1 = int(first[0:1])
            M2 = int(second[2:4])
            D2 = int(second[0:2])

        if (len(str(first)) == 4) and (len(str(second)) == 3):
            first = str(first)
            second = str(second)
            M1 = int(first[2:4])
            D1 = int(first[0:2])
            M2 = int(second[1:3])
            D2 = int(second[0:1])

        if (len(str(first)) == 3) and (len(str(second)) == 3):
            first = str(first)
            second = str(second)
            M1 = int(first[1:3])
            D1 = int(first[0:1])
            M2 = int(second[1:3])
            D2 = int(second[0:1])


        if ((D1==D2) or ((D2-D1 == 1) and (M1-M2 == 0)) or ((M1==1 and M2==2) and (D1==31 and D2==1)) or ((M1==2 and M2==3) and (D1==28 and D2==1)) or ((M1==2 and M2==3) and (D1==29 and D2==1)) or ((M1==3 and M2==4) and (D1==31 and D2==1)) or ((M1==4 and M2==5) and (D1==30 and D2==1)) or ((M1==5 and M2==6) and (D1==31 and D2==1)) or ((M1==6 and M2==7) and (D1==30 and D2==1)) or ((M1==7 and M2==8) and (D1==31 and D2==1)) or ((M1==8 and M2==9) and (D1==31 and D2==1)) or ((M1==9 and M2==10) and (D1==30 and D2==1)) or ((M1==10 and M2==11) and (D1==31 and D2==1)) or ((M1==11 and M2==12) and (D1==30 and D2==1)) or ((M1==12 and M2==1) and (D1==31 and D2==1))):
            first = int(first)
            second = int(second)
            keep.append(so.index(first))
            keep.append(so.index(second))
        else:
            first = int(first)
            second = int(second)
            discard.append(so.index(first))
            discard.append(so.index(second))

    keep = sorted( list(set(keep)))
    discard = sorted( list(set(discard)))

    p =[]
    for f, s in zip(keep, keep[1:]):
        if abs(f-s) == 1:
            p.append((no[f],no[s]-1))

    for c in range(len(discard)-1):
        p.append((no[discard[c]],int(sort[discard[c]].split(' ')[1])))
    p.append((no[-2],int(sort[-1].split(' ')[1])))
    p = sorted(p)
    l = []
    for f1, s1 in zip(p, p[1:]):
        f1 = str(f1).replace('(','').replace(')','').split(',')[0]
        s1 = str(s1).replace('(','').replace(')','').split(',')[0]
        if f1 == s1:
            l.append(f1)
    l1 = []
    for d in range(1,len(p)+1):
        for de in range(1,len(l)+1):
            if int(str(p[d-1]).replace('(','').replace(')','').split(',')[0]) == int(l[de-1]):
                l1.append(d-1)
    l1 = [i for a, i in enumerate(l1) if  a%2 == 0]
    p = [i for j, i in enumerate(p) if j not in l1]

    a1 = 1
    for write in range(1,len(p)+1):
        f3 = open('Two_Way_' + str(file11) + str(a1) + ".txt","w")
        start, stop = p[write-1][0], p[write-1][1]
        f3.write(''.join(lines1[start-1:stop]))
        a1 += 1
        f3.close()
    f2.close()

average_12_to_12('00-09(423)','A')
average_12_to_12('10-19','B')
average_12_to_12('30-39(422)','Cb')
average_12_to_12('40-49','D')

#To delete empty files
file_del = glob.glob('*Two_Way_*.txt')
for out in range(1,len(file_del)+1):
    if os.path.getsize(file_del[out-1]) < 10000:
        os.remove(file_del[out-1])

#To calculate average of each file
file2 = glob.glob('*Two_Way_*.txt')
for tmp in range(1,len(file2)+1):
    with open(file2[tmp-1],'r+') as f:
        Nos = [float(line.rstrip('\n').split(' ')[-1])
                   for line in f if not line.isspace()]
    with open(file2[tmp-1],'r+') as f4:
        lines = f4.readlines()
        if file2[tmp-1][8:9] == 'A':
            lines[0] = lines[0].replace('\n','') + " " + 'AVERAGE' + " = " + str(round(((sum(Nos) / len(Nos)) + 121.04385),3)) + " " + str(round(((sum(Nos) / len(Nos)) + 116.074),1)) +'\n'
        if file2[tmp-1][8:9] == 'B':
            lines[0] = lines[0].replace('\n','') + " " + 'AVERAGE' + " = " + str(round(((sum(Nos) / len(Nos)) - 88.42665),3)) + " " + str(round(((sum(Nos) / len(Nos)) - 93.397),1)) +'\n'
        if file2[tmp-1][8:10] == 'Cb':
            lines[0] = lines[0].replace('\n','') + " " + 'AVERAGE' + " = " + str(round(((sum(Nos) / len(Nos)) + 55.89185),3)) + " " + str(round(((sum(Nos) / len(Nos)) + 50.922),1)) +'\n'
        if file2[tmp-1][8:9] == 'D':
            lines[0] = lines[0].replace('\n','') + " " + 'AVERAGE' + " = " + str(round(((sum(Nos) / len(Nos)) - 89.58665),3)) + " " + str(round(((sum(Nos) / len(Nos)) - 94.557),1)) +'\n'
    with open(file2[tmp-1],'w+') as f4:
        f4.writelines(lines)

#To create daily offset file
f5 = open("daily offset.txt","a+")
file3 = glob.glob('*Two_Way_*.txt')
file3 =  ns.natsorted(file3, key=lambda x: (not x.isdigit(), x))
name = []
name1 = []
name2 = []
name3 = []
name4 = []
name5 = []
for tmp in range(1,len(file3)+1):
    with open(file3[tmp-1],'r+') as f6:
        lines = f6.readlines()
        lines[0] = lines[0].replace('\n','')
        name.append(lines[0].split(' ')[0])
        name.append(lines[-1].split(' ')[0])
        name.append(lines[0].split(' ')[-2])
        name.append(lines[0].split(' ')[-1])
        name.append(lines[0].split(' ')[1].split(':')[1])

    for type1,type2,type3,type4,type5 in zip(*[iter(name)]*5):
        if int(type5) in range(1,10):
            name1.append((str(type1) + ' to ' + str(type2) + ' ' + str(type3) + ' ' + str(type4)))
        if int(type5) in range(11,20):
            name2.append((str(type1) + ' to ' + str(type2) + ' ' + str(type3) + ' ' + str(type4)))
        if ((int(type5) in range(30,40)) and ('_Cb' in file3[tmp-1])):
            name4.append((str(type1) + ' to ' + str(type2) + ' ' + str(type3) + ' ' + str(type4)))
        if (int(type5) in range(40,50)):
            name5.append((str(type1) + ' to ' + str(type2) + ' ' + str(type3) + ' ' + str(type4)))

name1 = ns.natsorted(list(set(name1)), key=lambda x: (not x.isdigit(), x))
name2 = ns.natsorted(list(set(name2)), key=lambda x: (not x.isdigit(), x))
name4 = list(set(name4)-set(name3))
name4 = ns.natsorted(list(set(name4)), key=lambda x: (not x.isdigit(), x))
name5 = ns.natsorted(list(set(name5)), key=lambda x: (not x.isdigit(), x))
name1=np.array(name1)
name2=np.array(name2)
name4=np.array(name4)
name5=np.array(name5)

for a, b, d, e in zip(name1, name2,name4 ,name5):
    f5.write('{0} {1} {2} {3}  \n'.format(a, b, d, e))
f5.close()

#To convert daily offset text file into xlsx
def Txt2Xlsx(self, data, row = 0):
    for colNum, value in enumerate(data):
            self.write(row, colNum, value)
xlsxwriter.worksheet.Worksheet.addRow = Txt2Xlsx
workbook = xlsxwriter.Workbook("daily offset.xlsx")
worksheet = workbook.add_worksheet()
worksheet.set_column(1,1,1.7)
worksheet.set_column(6,6,1.7)
worksheet.set_column(11,11,1.7)
worksheet.set_column(16,16,1.7)
worksheet.set_column(21,21,1.7)
worksheet.set_column(0,0,10)
worksheet.set_column(2,2,10)
worksheet.set_column(5,5,10)
worksheet.set_column(7,7,10)
worksheet.set_column(10,10,10)
worksheet.set_column(12,12,10)
worksheet.set_column(15,15,10)
worksheet.set_column(17,17,10)
worksheet.set_column(20,20,10)
worksheet.set_column(22,22,10)
worksheet.set_column(3,3,7)
worksheet.set_column(8,8,7)
worksheet.set_column(13,13,7)
worksheet.set_column(18,18,7)
worksheet.set_column(23,23,7)
merge_format = workbook.add_format({'bold': 1,'align': 'center','valign': 'vcenter'})
worksheet.merge_range('A1:C1', 'DATE', merge_format)
worksheet.merge_range('D1:E1', '00-09', merge_format)
worksheet.merge_range('F1:H1', 'DATE', merge_format)
worksheet.merge_range('I1:J1', '10-19', merge_format)
worksheet.merge_range('K1:M1', 'DATE', merge_format)
worksheet.merge_range('N1:O1', '30-39(422)', merge_format)
worksheet.merge_range('P1:R1', 'DATE', merge_format)
worksheet.merge_range('S1:T1', '40-49', merge_format)

row = 1
with open('daily offset.txt', 'rt+') as f6:
    lines = f6.readlines()
    for line in lines:
        worksheet.addRow(data = line.split(" "), row = row)
        row += 1
workbook.close()
