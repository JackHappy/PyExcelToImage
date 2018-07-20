# PyExcelToImage
 capture the chart and cells in the excel  to image
 
 本模块实现对excel中图表单元格截图输出为jpeg的图片片文件；
 
 本模块依赖包：
 win32com.client, pythoncom, PIL, flask
 
 只能在windows和windows server服务器上运行，运行的时候请确保服务器安装office；
 
 ExportExcelToImage.py： 为本模块主要工作类，用于打开excel文件，输出图表截图文件；
 upload 为打开excel的目录；
 jpg 为输出的截图目录；

http_server.py 将模块功能http接口化，通过http协议接受excel，返回截图文件URL；
upload 为接受excel文件保存目录；

http_client.py 实现上传excel文件，下载截图文件保存的功能
excel  为上传excel的文件目录
download_jpg 下载的截图文件保存目录

测试模块方法：

1，单独测试excel截图功能
运行 python ExportExcelToImage.py 会读取upload中Sample.xls文件，在jpg目录生成截图文件；

2，测试excel截图的http接口
运行 python http_server.py  会打开http的监听服务；
运行 python http_client.py  会测试excel上传和下载截图文件功能
