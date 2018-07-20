# -*- coding: utf-8 -*-
import requests
import os
import json
import time
import os


EXCEL_PATH = 'excel'
JPG_PATH = 'download_jpg'
if os.path.exists(EXCEL_PATH):
    pass
else:
    os.makedirs(EXCEL_PATH)

if os.path.exists(JPG_PATH):
    pass
else:
    os.makedirs(JPG_PATH)

todaydate = time.strftime("%Y%m%d")


filename_excel = 'Sample.xls'
rangedict = {'image_1': 'I4:O23'}
sheetname = u'Sheet1'
data={'rangedict':json.dumps(rangedict),'sheetname':sheetname}
files = {'file': open(os.path.join(EXCEL_PATH,filename_excel), 'rb')}
r = requests.post('http://127.0.0.1:5000/upload',data = data, files=files)

print r.text
image_url = json.loads(r.text)['image_url']
image_name = json.loads(r.text)['image']

html_image = ''
image_list = []
for key in image_url:
    html = requests.get(image_url[key])
    with open(os.path.join(JPG_PATH,image_name[key]), 'wb') as file:
        file.write(html.content)
    image_list.append(os.path.join(JPG_PATH,image_name[key]))

print image_list

