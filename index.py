# coding: utf-8

from bottle import route, run, template, request, static_file, os
import datetime
import os
import glob
import subprocess
import shutil
import imghdr
import win32com
from win32com.client import *

root_dir = os.getcwd()

def all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)

def get_image_list(path):
    file_list = []
    for (root, dirs, files) in os.walk(path):
        for file in files:
            target = os.path.join(root,file).replace("\\", "/")
            fname, ext = os.path.splitext(target)
            if os.path.isfile(target):
                if imghdr.what(target) != None :
                    file_list.append(target)
                elif ext == '.emf': # include emf files too
                    file_list.append(target)
    return file_list

def doc2docx(doc_fullpath):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc_fullpath = doc_fullpath.replace("\\", "/")
    print(doc_fullpath)
    dirname = os.path.dirname(doc_fullpath)
    current_file = os.path.basename(doc_fullpath)
    fname, ext = os.path.splitext(current_file)
    doc = word.Documents.Open(doc_fullpath)
    doc.SaveAs(dirname + '/' + fname + '.docx', FileFormat = 16)
    doc.Close()
    word.Quit() # releases Word object from memory
    return dirname + '/' + fname + '.docx'

def ppt2pptx(ppt_fullpath):
    powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
    powerpoint.DisplayAlerts = 0
    ppt_fullpath = ppt_fullpath.replace("\\", "/")
    print(ppt_fullpath)
    dirname = os.path.dirname(ppt_fullpath)
    current_file = os.path.basename(ppt_fullpath)
    fname, ext = os.path.splitext(current_file)
    ppt = powerpoint.Presentations.Open(ppt_fullpath, False, False, False)
    ppt.SaveAs(dirname + '/' + fname + '.pptx')
    ppt.Close()
    powerpoint.Quit()
    return dirname + '/' + fname + '.pptx'

def xls2xlsx(xls_fullpath):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = 0
    xls_fullpath = xls_fullpath.replace("/", "\\")
    print(xls_fullpath)
    dirname = os.path.dirname(xls_fullpath)
    current_file = os.path.basename(xls_fullpath)
    fname, ext = os.path.splitext(current_file)
    xls = excel.Workbooks.Open(xls_fullpath)
    # xlsxに変換する
    xls.SaveAs(dirname + '/' + fname + '.xlsx', FileFormat = 51)
    xls.Close()
    excel.Quit()
    return dirname + '/' + fname + '.xlsx'

def extract_image(dirname, current_file, fname, ext):
    if ext == '.pptx' or ext == '.xlsx' or ext == '.docx':
        print(current_file)
        os.mkdir('tmp_' + fname)
        shutil.copyfile(dirname + '/' + current_file, dirname + '/' + 'tmp_' + fname + '/' + current_file + '.zip')
        os.chdir(dirname + '/' + 'tmp_' + fname)
        unzip_cmd = '7z x' + ' ' + '"' + current_file + '"' + '.zip'
        subprocess.call(unzip_cmd, shell=True, timeout=None)
        os.remove(current_file + '.zip')
        img_list = []
        img_list = get_image_list(dirname + '/' + 'tmp_' + fname)
        os.mkdir(dirname + '/' + fname)

        for images_fullpath in img_list:
            image_file = os.path.basename(images_fullpath)
            shutil.copyfile(images_fullpath, dirname + '/' + fname + '/' + image_file)
        
        os.chdir(dirname)
        shutil.rmtree(dirname + '/' + 'tmp_' + fname)

if __name__ == '__main__':
    @route('/image_file_extraction')
    def index():
        return template('''<!doctype html>
        <head>
            <meta charset="utf-8" />
            <script type="text/javascript">
              function getExtension(fileName) {
                var ret;
                if (!fileName) {
                  return ret;
                }
                var fileTypes = fileName.split(".");
                var len = fileTypes.length;
                if (len === 0) {
                  return ret;
                }
                ret = fileTypes[len - 1];
                return ret;
              }
          
              function IsTxtFile(fileName) {
                var Extension = getExtension(fileName);
                if (Extension.toLowerCase() === "zip") {
                  
                }
                else {
                  alert("It is not a zip file.");
                }
              }
            </script>
            <title>Extract image files</title>
        </head>

        <body>
            <h1>Extract image files</h1>
            <form action="/image_file_extraction/upload" method="post" enctype="multipart/form-data">
                <div class="form-group">
                    <label class="control-label" for="upload">Select a zip file:</label>
                    <input type="file" name="upload" id="elmFile" accept="application/zip" required>
                </div>
                <div class="form-group">
                    <input type="submit" id="btnUpload" value="upload" onclick="IsTxtFile(document.getElementById('elmFile').value)">
                </div>
            </form>
            </br>
            </br>
        </body>
        ''')

    @route('/image_file_extraction/static/<filepath:path>', name='static_file')
    def static(filepath):
        return static_file(filepath, root="./static")

    @route('/image_file_extraction/upload', method='POST')
    def do_upload():
        upload = request.files.get('upload')
        
        todaydetail = datetime.datetime.today()
        datetime_dir = todaydetail.strftime("%Y%m%d%H%M%S")
        p_dir = 'tmp/' + datetime_dir
        os.makedirs(p_dir)

        upload.save(p_dir, overwrite=True)
        os.chdir(p_dir)

        fileCheck = glob.glob('*')

        for f in fileCheck:
            f_name, f_ext = os.path.splitext(f)
            if f_name == 'zip':
                renamed_filename = 'tmp.zip'
                os.rename(f_name, renamed_filename)
                print('Renamed: ' + renamed_filename)
                break
            elif f_ext != '.zip':
                return 'Please upload a zip file.'
        
        zf = glob.glob('*.zip')
        
        for f in zf:
            z_name, z_ext = os.path.splitext(f)
            unzip = '7z x' + ' ' + z_name + '.zip'
            subprocess.call(unzip, shell=True, timeout=None)
            os.remove(f)

        target_folder = root_dir + '/' + p_dir

        temp_created_docx_pptx_xlsx = []
        for i in all_files(target_folder):
            dirname = os.path.dirname(i)
            current_file = os.path.basename(i)
            fname, ext = os.path.splitext(current_file)
            os.chdir(dirname)
            if ext == '.doc':
                converted_docx = doc2docx(dirname + '/' + current_file)
                temp_created_docx_pptx_xlsx.append(converted_docx)
            elif ext == '.ppt':
                converted_pptx = ppt2pptx(dirname + '/' + current_file)
                temp_created_docx_pptx_xlsx.append(converted_pptx)
            elif ext == '.xls':
                converted_xlsx = xls2xlsx(dirname + '/' + current_file)
                temp_created_docx_pptx_xlsx.append(converted_xlsx)

        print('Image file being extracted...')
        for i in all_files(target_folder):
            dirname = os.path.dirname(i)
            current_file = os.path.basename(i)
            fname, ext = os.path.splitext(current_file)
            os.chdir(dirname)
            extract_image(dirname, current_file, fname, ext)

        for temp_created_docx_pptx_xlsx_fullpath in temp_created_docx_pptx_xlsx:
            os.remove(temp_created_docx_pptx_xlsx_fullpath)

        archive_cmd = '7z a' + ' ' + target_folder + '.zip' + ' ' + target_folder
        subprocess.call(archive_cmd, shell=True, timeout=None)

        os.chdir(root_dir)
        result_zip_file = target_folder.replace('/', os.sep)
        result_zip_file = target_folder + '.zip'
        print(result_zip_file)
        cmd = 'python download_button_command.py ' + result_zip_file

        subprocess.call(cmd, shell=True)

        resultFile = 'result/' + datetime_dir + '.html'
        return template(resultFile)

    @route('/tmp/<file_path:path>')
    def static(file_path):
        return static_file(file_path, root='./tmp', download=True)

    run(host='localhost', port=8080, reloader=True)