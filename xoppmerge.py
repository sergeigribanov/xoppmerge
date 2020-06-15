import re
import os
import gzip
import xlwt
import json
import collections
import argparse
import subprocess
from PyPDF2 import PdfFileReader
import xml.etree.ElementTree as et

def xopp_open(path):
    return et.parse(gzip.open(path, 'r')).getroot()

def xopps_open(path_list):
    iterables = []
    for path in path_list:
        iterables.append(xopp_open(path).iter('page'))

    return zip(*iterables)

def detect_scores(page, scores):
    mtemp = '[.*]?problem[\s+]?([0-9]+)[\s]?:[\s]?([+-]?(\d+((\.|\,)\d*)?|(\.|\,)\d+)([eE][+-]?\d+)?)[.*]?'
    for page_copy in page:
        for layer in page_copy.iter('layer'):
            for text in layer.iter('text'):
                match = re.match(mtemp, text.text.lower())
                if match:
                    assert(match.group(1) not in scores)
                    scores[match.group(1)] = float(match.group(2))
                else:
                    assert('problem' not in text.text.lower())

def xopp_background(bkg_type = 'solid', color = '#ffffffff', style='graph'):
    result = et.Element('background')
    result.attrib = {'type' : bkg_type, 'color' : color, 'style' : style}
    return result
                    
def xopp_rectangle(x1, y1, x2, y2,
                   tool = 'pen', ts = '0ll', fn = '',
                   color = '#000000ff', line_width = 2.26):
    result = et.Element('stroke')
    result.attrib = {'tool' : tool, 'ts' : ts, 'fn' : fn,
                   'color' : color, 'width' : str(line_width)}
    table_coord = [x1, y1, x1, y2, x2, y2, x2, y1, x1, y1]
    result.text = ' '.join([str(el) for el in table_coord])
    return result

def xopp_text(text, x, y, size=12, color='#000000ff',
              font='Sans', ts='0ll', fn=''):
    result = et.Element('text')
    result.attrib = {'font' : font, 'size' : str(size), 'x' : str(x), 'y' : str(y),
                     'color' : color, 'ts' : ts, 'fn' : fn}
    result.text = text
    return result

def xopp_line(x1, y1, x2, y2,
              tool = 'pen', ts = '0ll', fn = '',
              color = '#000000ff', line_width = 2.26):
    result = et.Element('stroke')
    result.attrib = {'tool' : tool, 'ts' : ts, 'fn' : fn,
                   'color' : color, 'width' : str(line_width)}
    result.text = '{x1} {y1} {x2} {y2}'.format(x1 = x1, y1 = y1, x2 = x2, y2 = y2)
    return result

def xopp_summ(x, y, width = 27.674):
    result = et.parse('xopp_summ.xml').getroot()
    result.attrib = {'text' : '\sum', 'left' : str(x), 'top' : str(y),
                     'right' : str(x + width),
                     'bottom' : str(y + 0.936 * width)}
    return result
                    
def xopp_score_table(root, width, tag, scores):
    scl = width / 725
    height = 200.0 * scl
    page_attrib = {'width' : str(width), 'height' : str(height)}
    page = et.Element('page')
    page.attrib = page_attrib
    layer = et.Element('layer')
    x1, y1, x2, y2 = 100 * scl, 80 * scl, 600 * scl, 150 * scl
    n = len(scores)
    h = (x2 - x1) / (n + 1)
    frame = xopp_rectangle(x1, y1, x2, y2, color='#3333ccff', line_width=2.26 * scl)
    title = xopp_text(tag, 100, 32, color='#3333ccff', size=int(24 * scl))
    layer.insert(1, title)
    layer.insert(2, frame)
    h0 = 0.33 * (y2 - y1)
    h1 = y2 - y1 - h0
    ly = y1 + h0
    layer.insert(3, xopp_line(x1, ly, x2, ly, color='#3333ccff', line_width=2.26 * scl))
    k = 4
    summ = 0
    for i, name in enumerate(scores):
        lx = x1 + (i + 1) * h
        layer.insert(k, xopp_line(lx, y1, lx, y2, color='#3333ccff', line_width=2.26 * scl))
        k += 1
        layer.insert(k, xopp_text(name, lx - 0.55 * h , y1 + 0.25 * h0,
                                  size = int(12 * scl)))
        k += 1
        layer.insert(k, xopp_text(str(scores[name]), lx - 0.6 * h , ly + 0.4 * h1,
                                  int(12 * scl)))
        k += 1
        summ += scores[name]

    layer.insert(k, xopp_summ(x1 + (n + 0.4) * h, y1 + 0.02 * h0,
                              width = 27.674 * scl))
    k += 1
    layer.insert(k, xopp_text(str(summ), x1 + (n + 0.4) * h, ly + 0.4 * h1,
                              int(12 * scl)))

    page.insert(1, xopp_background())
    page.insert(1, layer)
    root.insert(2, page)

def adjust_string_scale(string, scale):
    lst = string.replace('\n', ' ').split(' ')
    lst = list( filter(lambda x: x != '', lst))
    lst = [str(float(el) * scale) for el in lst]
    return ' '.join(lst)
    
def adjust_pen_scale(stroke, scale):
    stroke.attrib['width'] = adjust_string_scale(stroke.attrib['width'], scale)
    stroke.text = adjust_string_scale(stroke.text, scale)

def adjust_textimage_scale(textimage, scale):
    textimage.attrib['left'] = str(scale * float(textimage.attrib['left']))
    textimage.attrib['top'] = str(scale * float(textimage.attrib['top']))
    textimage.attrib['right'] = str(scale * float(textimage.attrib['right']))
    textimage.attrib['bottom'] = str(scale * float(textimage.attrib['bottom']))
    
def adjust_stroke_scale(stroke, scale):
    adjust_pen_scale(stroke, scale)

def adjust_text_scale(text, scale):
    text_size = float(text.attrib['size'])
    text_x = float(text.attrib['x'])
    text_y = float(text.attrib['y'])
    text.attrib['size'] = str(int(text_size * scale))
    text.attrib['x'] = str(text_x * scale)
    text.attrib['y'] = str(text_y * scale)

def page_size(page):
    return float(page.attrib['width']), float(page.attrib['height'])

def pdf_page_size(path):
    fl = open(path, 'rb')
    input1 = PdfFileReader(fl)
    npages = input1.getNumPages()
    result = []
    for pnum in range(npages):
        page = input1.getPage(pnum)
        box = page.mediaBox
        width = float(box.getWidth())
        height = float(box.getHeight())
        if page.get('/Rotate') == 90 or page.get('/Rotate') == 270:
            result.append((height, width))
        else:
            result.append((width, height))

    fl.close()
    return result

def adjust_scale(tag, page, tgt_size):
    src_size = page_size(page)
    scale = tgt_size[0] / src_size[0]
    page.attrib['width'] = str(src_size[0] * scale)
    page.attrib['height'] = str(src_size[1] * scale)
    if tag == 'Айдаков Е.Е. 17347':
        print('scale = {}'.format(scale))
        print(src_size)
        
    for layer in page:
        for text in layer.iter('text'):
            adjust_text_scale(text, scale)

        for stroke in layer.iter('stroke'):
            adjust_stroke_scale(stroke, scale)

        for textimage in layer.iter('textimage'):
            adjust_textimage_scale(textimage, scale)

def xopps_merge(tag, path_list, output_path, pdf_prefix, scoring = False):
    print(tag)
    tgt_sizes = None
    p0width = 725.0
    if pdf_prefix != None:
        path = os.path.join(pdf_prefix, '{}.pdf'.format(tag))
        assert(os.path.exists(path))
        tgt_sizes = pdf_page_size(path)
        p0width = tgt_sizes[0][0]
        
    docs = xopps_open(path_list)
    tree = et.parse('xopp_template.xml')
    root = tree.getroot()
    scores = dict()
    pnum = 0
    for page in docs:
        if pdf_prefix != None:
            for page_copy in page:
                adjust_scale(tag, page_copy, tgt_sizes[pnum])
        
        if scoring:
            detect_scores(page, scores)
             
        for page_copy in page[1:]:
            if tag == 'Айдаков Е.Е. 17347':
                print(page_copy.attrib['width'], page_copy.attrib['height'])
                
            for layer in page_copy.iter('layer'):
                page[0].append(layer)

        if pdf_prefix != None:
            path = os.path.join(pdf_prefix, '{}.pdf'.format(tag))
            page[0].find('background').set('filename', path)
            
        root.append(page[0])
        pnum +=1

    if scoring:
        scores = collections.OrderedDict(sorted(scores.items()))
        xopp_score_table(root, p0width, tag, scores)
        
    tree.write(gzip.open(output_path, 'wt'), encoding='unicode')
    return scores

def search_annotations(prefix):
    result = dict()
    mtemp = '(.*)_([0-9]+).pdf.xopp'
    for root, dirs, files in os.walk(prefix):
        for filename in files:
            match = re.match('(.*).pdf.xopp~', filename)
            if match:
                continue
            
            match = re.match(mtemp, filename)
            if not match:
                continue

            path = os.path.join(root, filename)
            tag = match.group(1)
            if tag not in result:
                result[tag] = []

            result[tag].append(path)

    for tag in result:
        result[tag] = sorted(result[tag], key = lambda x: int(re.match(mtemp, x).group(2)))

    return result

def pdf_export(xopp_path, pdf_path):
    subprocess.Popen(['/bin/bash', '-c',
                      'xournalpp {} -p {}'.format(
                          xopp_path.replace(' ', '\ '), pdf_path.replace(' ', '\ '))])

def colnum_string(m):
    n = m + 1
    string = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string
    
def export_excel(path, sheet, scores, n):
    book = xlwt.Workbook()
    sh = book.add_sheet(sheet)
    ntags = 1
    problems = None
    sh.write(0, 0, 'группа')
    sh.write(0, 1, 'ФИО')
    for i in range(n):
        sh.write(0, i + 2, i + 1)

    sh.write(0, n + 2, 'сумма')
        
    for i, tag in enumerate(scores):
        y = i + 1
        mres = re.match('(.*) ([0-9]+)', tag)
        assert(mres != None)
        sh.write(y, 0, mres.group(2))
        sh.write(y, 1, mres.group(1))
        book.save(path)
        s = ''
        for p in range(1, n + 1):
            problem = str(p)
            if problem in scores[tag]:
                sh.write(y, p + 1, float(scores[tag][problem]))
            else:
                sh.write(y, p + 1, 0)

            if s != '':
                s += '+'
                
            s += '{}{}'.format(colnum_string(p + 1), y + 1)

        sh.write(y, p + 2, xlwt.Formula(s))
                
    book.save(path)

def export_json(path, scores, n):
    result = dict()
    for stud in scores:
        result[stud] = [0 for i in range(n)]
        for prob in scores[stud]:
            index = int(prob) - 1
            result[stud][index] = scores[stud][prob]

    with open(path, 'w', encoding='utf-8') as fl:
        json.dump(result, fl, indent=4, ensure_ascii=False)
    
if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input-prefix', type=str, default=os.getcwd(),
                        help='Input prefix to .xopp annotations.')
    parser.add_argument('-o', '--output-prefix', type=str, default=os.getcwd(),
                        help='Output prefix to merged .xopp annotations.')
    parser.add_argument('-p', '--pdf-prefix', type=str, default=None,
                        help='Prefix to a local PDF files.')
    parser.add_argument('-e', '--pdf-export-prefix', type=str, default=None,
                        help='Prefix for a PDF export.')
    parser.add_argument('-s', '--scoring', action='store_true',
                        help='Score counting. Searching for '
                        'text "Problem <number>: <score>" (for example, "Problem 3: 10.5") '
                        'and summarizing scores.')
    args = parser.parse_args()
    annotations = search_annotations(args.input_prefix)
    scores = dict()
    for tag in annotations:
        xopp_path = os.path.join(args.output_prefix, '{}_final.pdf.xopp'.format(tag))
        scores[tag] = xopps_merge(tag, annotations[tag], xopp_path, args.pdf_prefix, args.scoring)
        if args.pdf_export_prefix != None:
            pdf_path = os.path.join(args.pdf_export_prefix, '{}_final.pdf'.format(tag))
            pdf_export(xopp_path, pdf_path)

    if args.scoring:
        export_excel('scores.xls', 'test-exam', scores, 5)
        export_json('scores.json', scores, 5)
