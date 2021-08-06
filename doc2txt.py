from win32com import client as wc
import os
import fnmatch

def doc2txt(path):
    debug = 0
    if debug:
        print(path)
    # 该目录下所有文件的名字
    files = os.listdir(path)
    # 创建一个文本存入所有的word文件名/测试用
    fileNameSet = os.path.abspath(os.path.join(path, 'fileNames.txt'))
    o = open(fileNameSet, "w")
    try:
        for filename in files:
            if debug:
                print(filename)
            # 如果不是word文件：继续
            if not fnmatch.fnmatch(filename, '*.doc') and not fnmatch.fnmatch(filename, '*.docx'):
                continue;
            # 如果是word临时文件：继续
            if fnmatch.fnmatch(filename, '~$*'):
                continue;
            if debug:
                print(filename)
            docpath = os.path.abspath(os.path.join(path, filename))

            # 得到一个新的文件名,把原文件名的后缀改成txt
            new_txt_name = ''
            if fnmatch.fnmatch(filename, '*.doc'):
                new_txt_name = filename[:-4] + '.txt'
            else:
                new_txt_name = filename[:-5] + '.txt'
            if debug:
                print(new_txt_name)
            word_to_txt = os.path.join(os.path.join(path), new_txt_name)
            # print(word_to_txt)
            wordapp = wc.Dispatch('Word.Application')
            doc = wordapp.Documents.Open(docpath)
            # 为了让python可以在后续操作中r方式读取txt和不产生乱码，参数设置为4
            doc.SaveAs(word_to_txt, 4)
            doc.Close()
            o.write(word_to_txt + '\n')
    finally:
        wordapp.Quit()