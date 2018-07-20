# -*- coding: utf-8 -*-
from win32com.client import Dispatch
import os
import pythoncom

from PIL import Image, ImageGrab

class PyExcelToImage:
    def __init__(self,excelname,rangdict,sheetname,vision=0):
        pythoncom.CoInitialize()
        self.filename = excelname
        self.rangearea = rangdict
        self.dir = os.path.join(os.getcwd(),'upload')
        self.sheetname = sheetname
        self.savedir = os.path.join(os.getcwd(),'jpg')
        self.vision = vision
        if os.path.exists(self.dir):
            pass
        else:
            os.makedirs(self.dir)
        if os.path.exists(self.savedir):
            pass
        else:
            os.makedirs(self.savedir)

    def __del__(self):
        pass

    def start_export(self):
        xlApp = Dispatch("Excel.Application")
        xlApp.Visible = self.vision
        try:
            if os.path.exists(os.path.join(self.dir, self.filename)):
                workbook = xlApp.Workbooks.Open(os.path.join(self.dir, self.filename), IgnoreReadOnlyRecommended=True, Editable=True)
            else:
                return {'code':-1, 'msg':'No file in the path'}
        except Exception as e:
            print "ExportExcelToImage Except:%s" %e
        item=1
        sheetname=[]
        while item<=workbook.Sheets.Count:
            sheetname=sheetname+[workbook.Sheets(item).name]
            item+=1
        if self.sheetname not in sheetname:
            xlApp.CutCopyMode = False
            workbook.Close(False)
            xlApp.quit()
            return {'code':-1,'msg':'the sheet not in the Workbooks' }
        sheet = workbook.Worksheets(self.sheetname)
        message={'code':0, 'image':{},'msg':''}
        #print self.rangearea
        for key in self.rangearea:
            #print "%s-%s" %(key, self.rangearea[key])
            sheet.Range(self.rangearea[key]).Copy()
            im = ImageGrab.grabclipboard()
            savefilename=self.filename[:-4]+"_"+key+".jpg"
            if isinstance(im, Image.Image):
                #print "Image:%s, size : %s, mode: %s" % (key, im.size, im.mode)
                im.save(os.path.join(self.savedir, savefilename), 'jpeg')
                message['image'][key]=savefilename
            else:
                print "clipboard is empty."
                message['code']=-1
        xlApp.CutCopyMode = False
        workbook.Close(False)
        xlApp.quit()
        return message

if __name__ == '__main__':
    filename = 'Sample.xls' #只支持xls文件
    rangedict = {'image1': 'I4:O23'}
    ett = PyExcelToImage(excelname=filename, rangdict=rangedict, sheetname=u'Sheet1', vision=1)
    print ett.start_export()