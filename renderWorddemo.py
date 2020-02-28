# -*- coding:utf-8 -*-

import re
from docxtpl import DocxTemplate
from glob import glob1
import os
import math
from win32com.client import gencache
from win32com.client import constants, gencache
def getAlltxt():
    return glob1("./","*.txt")


def readLines(filename="banji.txt"):
    lines=[]
    with open(filename) as f:
        lines=f.readlines()
    lines=list(filter(lambda x: x.strip(),lines))
    lines= list(map( lambda x:re.split("\s+",x.strip()),lines))
    dictlist=[]

    for line in lines:
        #print line
        item={}
        namelst=[]
        n=1
        item["score"]=[]
        total=0
        num=0
        for i,col in enumerate(line):
            if i==0:
                item["sno"]=col
                continue
            try:

                score=int(col)
                if n==3:
                    score=int(math.floor(score/45.0*100))
                elif n==4:
                    score = int(math.floor(score / 25.0 * 100))
                else:
                    score=int(math.floor((score+30)/45.0*100))
                item["score"].append(score)


                item["sname"]=" ".join(namelst)
                n+=1

            except:

                namelst.append(col)
        item["sum"]=sum(item["score"])
        total=item["sum"]
        num=+1
        item["avg"]=int(math.floor(item["sum"]/4.0))

        dictlist.append(item)


    savg=int(total/num)
    #lines=list(map(lambda x:[x[0],x[1]," ".join(x[2:])],lines))
    return dictlist,savg


def createPdf(wordPath, pdfPath):
    """
    word转pdf
    :param wordPath: word文件路径
    :param pdfPath:  生成pdf文件路径
    """
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(wordPath, ReadOnly=1)
    doc.ExportAsFixedFormat(pdfPath,
                            constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    word.Quit(constants.wdDoNotSaveChanges)


def main():
    txtfiles=getAlltxt()
    xlsfile="score.csv"
    teancher=["Mrs. Sudanee Kongkum","Mrs. Sineenat Tintakanonte",]
    with open(xlsfile,"w") as fw:
        for i,filename in enumerate(txtfiles) :
            print "[Info] Handling: "+filename
            lines,savg=readLines(filename)


            reportname=os.path.basename(filename)
            fname=os.path.splitext(reportname)[0]
            reportname=fname+".docx"
            grade="/".join(fname.split("."))
            fw.write(fname+",,总分平均分:,".decode("utf-8").encode("gbk")+str(savg)+",,平均分:,".decode("utf-8").encode("gbk")+str(int(savg/4.0))+",\n\n")
            lines1=lines
            lines=sorted(lines1,key=lambda s:s["avg"],reverse=True)
            for line in lines:
                fw.write(line["sno"]+","+line["sname"]+","+",".join(map(lambda x:str(x),line["score"]))+","+str(line["sum"])+","+str(line["avg"])+",\n")
            # fw.write(str(savg)+",\n")
            # doc =DocxTemplate('tpl3.docx')
            # doc.render({"datalist":lines1,"grade":grade,"hometeacher":teancher[i]})
            # docfile="./report/"+reportname
            # pdffile="./report/"+fname+".pdf"

            # doc.save(docfile)
            # createPdf(os.path.abspath(docfile),os.path.abspath(pdffile))
            print  "[Info] Successfull: ./report/"+reportname
main()