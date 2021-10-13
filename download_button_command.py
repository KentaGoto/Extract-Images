# coding: utf-8

import os
from os.path import join, basename
import sys
import io
import codecs
from win32com.client import *


root_dir = os.getcwd()
result_dir = root_dir + r'\result'

if __name__ == '__main__':
    result_zip_file = sys.argv[1]

    current_file = os.path.basename(result_zip_file)
    fname, ext = os.path.splitext(current_file)

    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='cp932')

    result_file = result_dir + '\\' + fname + '.html'
    print(result_file)
    fout_html  = codecs.open(result_file, 'w', 'utf-8') # html

    html_header = '''<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"  xml:lang="ja" lang="ja">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Results</title>
<body>'''

    html_footer = '''</body>
</html>'''
        
    fout_html.write(html_header + '\n')

    fout_html.write('<p>Processing is completed.</br>Please Download.</p>')
    fout_html.write('<form id="export" calss="etable" method="GET" action="/tmp/' + fname + '.zip">' + '<input type="submit" name="export" value="Download"></form>')

    fout_html.write(html_footer)
    fout_html.close()

    print('\n' + 'Done!')
