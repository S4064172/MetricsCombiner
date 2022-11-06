# This is a sample Python script.
import argparse
import os

import pandas as pd

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    parser = argparse.ArgumentParser()
    parser.add_argument('--p', type=str, default="")
    parser.add_argument('--f1', type=str, default="")
    parser.add_argument('--f2', type=str, default="")
    parser.add_argument('--f3', type=str, default="")
    parser.add_argument('--s', type=str, default="")
    parser.add_argument('--c', type=str, default="")
    parser.add_argument('--d', type=str, default="")
    args = parser.parse_args()

    sheet = args.s  # ' LocatorDuplicationSameMethod '
    column = args.c  # 'Counter'
    descriminator = args.d  # 'isForm'
    descriminatorIsSet = descriminator != ''

    basePath = args.p
    fileName1 = args.f1  # 'H:\_TESI\_Fase7_EsecuzioneMetrice\MetricsReport\Report\Fase1SchoolManagement\SchoolManagement_manual.xlsx'
    fileName2 = args.f2  # 'H:\_TESI\_Fase7_EsecuzioneMetrice\MetricsReport\Report\Fase1SchoolManagement\SchoolManagement_newAssessor.xlsx'
    fileName3 = args.f3  # 'H:\_TESI\_Fase7_EsecuzioneMetrice\MetricsReport\Report\Fase1SchoolManagement\SchoolManagement_oldAssessor.xlsx'

    fileExist1 = fileName1 != ''
    fileExist2 = fileName2 != ''
    fileExist3 = fileName3 != ''
    index1 = 0
    index2 = 0
    index3 = 0

    outFile = os.path.basename(fileName1).split("_")[0] + "_" + sheet
    colums = []
    if (fileExist3):
        file3 = pd.ExcelFile(basePath + fileName3)
        df3 = pd.read_excel(file3, sheet)
        nameColum3 = os.path.basename(fileName3).split("_")[1].split(".")[0]
        colums.append(nameColum3)

    if (fileExist2):
        file2 = pd.ExcelFile(basePath + fileName2)
        df2 = pd.read_excel(file2, sheet)
        nameColum2 = os.path.basename(fileName2).split("_")[1].split(".")[0]
        colums.append(nameColum2)

    if (fileExist1):
        file1 = pd.ExcelFile(basePath + fileName1)
        df1 = pd.read_excel(file1, sheet)
        nameColum1 = os.path.basename(fileName1).split("_")[1].split(".")[0]
        colums.append(nameColum1)

    line = ';'.join(colums) + "\n"
    while (
            (fileExist1 and df1[column].size > index1) or
            (fileExist2 and df2[column].size > index2) or
            (fileExist3 and df3[column].size > index3)
    ):

        newLine = []
        if fileExist3:
            if df3[column].size > index3:
                if not descriminatorIsSet or not (descriminator in df3.columns) or df3[descriminator][index3] == 'N':
                    newLine.append(str(df3[column][index3]))
                index3 = index3 + 1
            else:
                newLine.append('')

        if fileExist2:
            if df2[column].size > index2:
                if not descriminatorIsSet or not (descriminator in df2.columns) or df2[descriminator][index2] == 'Ne':
                    newLine.append(str(df2[column][index2]))
                index2 = index2 + 1
            else:
                newLine.append('')

        if fileExist1:
            if df1[column].size > index1:
                if not descriminatorIsSet or not (descriminator in df1.columns) or df1[descriminator][index1] == 'Ne':
                    newLine.append(str(df1[column][index1]))
                index1 = index1 + 1
            else:
                newLine.append('')

        line = line + ';'.join(newLine) + "\n"
        #print(';'.join(newLine))


    try:
        with open(basePath + "\\" + outFile + ".csv", 'w') as f:
            f.write(line)
    except FileNotFoundError:
        print("The 'docs' directory does not exist")
