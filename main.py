import sys
import os
import msvcrt
from doc2txt import doc2txt
from pdf2txt import pdf2txt
from vector import Vector
from sketch import Sketch
from io import open

e = 0
while (e < 10):

    # 筛选当前目录下的txt文档
    path = os.getcwd()
    names = os.listdir(path)
    print('当前目录下全部文档如下：')
    for name in names:
        if name.endswith(('.pdf','.docx','.txt')):
            print(name)
    print('调用函数进行格式转换......')
    doc2txt(path)
    path = os.getcwd()
    names = os.listdir(path)
    pdf2txt(path)
    path = os.getcwd()
    names = os.listdir(path)
    print('当前目录下的TXT文本文档如下：')
    for name in names:
        if name.endswith('.txt'):
            print(name)

        # 打开并量化参与测试的文档
    print('请输入要参与检测的文件名,输入‘end’结束输入')
    filenames = []
    for i in range(99):
        median = input()
        if (median == 'end'):
            break
        if (median not in names):
            print('输入的文件名有误，仅支持输入当前目录下的文件，请重新输入：')
            median = None
        else:
            filenames.append(median)

    k = 5  # k-grams.量化宽度
    d = 10000  # 文档摘要维度
    sketches = [0 for i in filenames]
    for i in range(len(filenames)):
        with open(filenames[i], 'r', encoding='UTF-8') as f:
            text = f.read()
            sketches[i] = Sketch(text, k, d)

    # 输出结果标题
    with open('clone.xls', 'w') as c:
        c.write('查重结果\n')
        c.write(' ' * 27)
        c.write('\t')
        print(' ' * 26, end='')
        for filename in filenames:
            print('{: <21}'.format(filename), end='  ')
            c.write(format(filename))
            c.write('\t')
        c.write('\n')
        print()

        # 输出结果比较明细
        for i in range(len(filenames)):
            print('{: <25}'.format(filenames[i]), end='')
            c.write(format(filenames[i]))
            c.write('\t')
            for j in range(len(filenames)):
                print('({: <19})'.format(sketches[i].similarTo(sketches[j])), end='  ')
                c.write(format(sketches[i].similarTo(sketches[j])))
                c.write('\t')
            print()
            c.write('\n')

        filenames.clear()
        e = e + 1
        print('按任意键继续....\n')
        q = input()