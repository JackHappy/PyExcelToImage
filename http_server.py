# -*- coding: utf-8 -*-
from flask import *
import logging
import os
import json

from logging import handlers
app = Flask(__name__, static_folder='jpg',static_url_path='/jpg')

from ExportExcelToImage import PyExcelToImage
@app.route('/')
def index():
    logger.info("test")
    return 'You have no permission to this page\n'

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            logger.debug('No file part')
            return jsonify({'code': -1, 'filename': '', 'msg': 'No file part'})
        if 'rangedict' not in request.form:
            logger.debug('No Range Area')
            return  jsonify({'code':-1,'msg':'No Range Area'})
        if 'sheetname' not in request.form:
            logger.debug('No SheetName')
            return  jsonify({'code':-1,'msg':'No SheetName'})
        file = request.files['file']
        sheetname = request.form['sheetname']
        rangedict = request.form['rangedict']
        # if user does not select file, browser also submit a empty part without filename
        if file.filename == '':
            logger.debug('No selected file')
            return jsonify({'code': -1, 'filename': '', 'msg': 'No selected file'})
        else:
            try:
                if file and allowed_file(file.filename):
                    origin_file_name = file.filename
                    logger.debug('filename is %s' % origin_file_name)
                    # filename = secure_filename(file.filename)
                    filename = origin_file_name

                    if os.path.exists(UPLOAD_PATH):
                        logger.debug('%s path exist' % UPLOAD_PATH)
                        pass
                    else:
                        logger.debug('%s path not exist, do make dir' % UPLOAD_PATH)
                        os.makedirs(UPLOAD_PATH)

                    file.save(os.path.join(UPLOAD_PATH, filename))
                    logger.debug('%s save successfully' % filename)
                    logger.debug("file:%s, sheetname:%s, rangearea:%s" % (filename, sheetname, rangedict))
                    ett = PyExcelToImage(excelname=filename, rangdict=json.loads(rangedict),sheetname=sheetname)
                    message = ett.start_export()

                    if message['code'] !=0:
                        return json.dumps(message)
                    logger.info("file:%s, sheetname:%s, rangearea:%s" %(filename, sheetname,rangedict))
                    image_url = {}
                    for key in message['image']:
                        image_url[key]=request.url_root+'jpg/'+message['image'][key]
                    message['image_url']=image_url
                    return json.dumps(message)
                else:
                    logger.debug('%s not allowed' % file.filename)
                    return jsonify({'code': -1, 'filename': '', 'msg': 'File not allowed'})
            except Exception as e:
                logger.debug('upload file exception: %s' % e)
                print 'upload file exception: %s' % e
                return jsonify({'code': -1, 'filename': '', 'msg': 'Error occurred'})
    else:
        return jsonify({'code': -1, 'filename': '', 'msg': 'Method not allowed'})

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
@app.route('/delete')
def delete_file():
    for file in os.listdir('jpg'):
        logger.info('delete jpg %s' % file)
        os.remove(os.path.join('jpg',file))
    for file in os.listdir('upload'):
        logger.info('delete upload %s' % file)
        os.remove(os.path.join('upload',file))
    return "delete success"


if __name__ == '__main__':

    ALLOWED_EXTENSIONS = ('xls')
    UPLOAD_PATH = 'upload'
    LOG_PATH = 'log'
    JPG_PATH = 'jpg'

    if os.path.exists(JPG_PATH):
        pass
    else:
        os.makedirs(JPG_PATH)

    if os.path.exists(UPLOAD_PATH):
        pass
    else:
        os.makedirs(UPLOAD_PATH)

    if os.path.exists(LOG_PATH):
        pass
    else:
        os.makedirs(LOG_PATH)

    fmt='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    logger = logging.getLogger()
    format_str = logging.Formatter(fmt)
    logger.setLevel(logging.DEBUG)
    sh = logging.StreamHandler()
    sh.setFormatter(format_str)
    th = handlers.TimedRotatingFileHandler(filename='log/log.txt', when='D',interval=1, backupCount=180 )
    th.setFormatter(format_str)  # 设置文件里写入的格式
    #logger.addHandler(sh)  # 把对象加到logger里
    logger.addHandler(th)
    app.run(host='0.0.0.0')

