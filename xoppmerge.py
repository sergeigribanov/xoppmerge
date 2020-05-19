import re
import os
import gzip
import xlwt
import json
import collections
import argparse
import subprocess
import xml.etree.ElementTree as et

def xopp_open(path):
    return et.parse(gzip.open(path, 'r')).getroot()

def xopps_open(path_list):
    iterables = []
    for path in path_list:
        iterables.append(xopp_open(path).iter('page'))

    return zip(*iterables)

def detect_scores(page, scores):
    mtemp = 'Problem ([0-9]+): ([+-]?(\d+((\.|\,)\d*)?|(\.|\,)\d+)([eE][+-]?\d+)?)'
    for page_copy in page:
        for layer in page_copy.iter('layer'):
            for text in layer.iter('text'):
                match = re.match(mtemp, text.text)
                if match:
                    assert(match.group(1) not in scores)
                    scores[match.group(1)] = float(match.group(2))

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
                    
def xopp_score_table(root, tag, scores):
    page_attrib = {'width' : '725.76000000', 'height' : '200.0'}
    page = et.Element('page')
    page.attrib = page_attrib
    layer = et.Element('layer')
    x1, y1, x2, y2 = 100, 80, 600, 150
    n = len(scores)
    h = (x2 - x1) / (n + 1)
    frame = xopp_rectangle(x1, y1, x2, y2, color='#3333ccff')
    title = xopp_text(tag, 100, 32, color='#3333ccff', size=24)
    layer.insert(1, title)
    layer.insert(2, frame)
    h0 = 0.33 * (y2 - y1)
    h1 = y2 - y1 - h0
    ly = y1 + h0
    layer.insert(3, xopp_line(x1, ly, x2, ly, color='#3333ccff'))
    k = 4
    summ = 0
    for i, name in enumerate(scores):
        lx = x1 + (i + 1) * h
        layer.insert(k, xopp_line(lx, y1, lx, y2, color='#3333ccff'))
        k += 1
        layer.insert(k, xopp_text(name, lx - 0.55 * h , y1 + 0.25 * h0))
        k += 1
        layer.insert(k, xopp_text(str(scores[name]), lx - 0.6 * h , ly + 0.4 * h1))
        k += 1
        summ += scores[name]

    layer.insert(k, xopp_summ(x1 + (n + 0.4) * h, y1 + 0.02 * h0))
    k += 1
    layer.insert(k, xopp_text(str(summ), x1 + (n + 0.4) * h, ly + 0.4 * h1))

    page.insert(1, xopp_background())
    page.insert(1, layer)
    root.insert(2, page)
    
                    
def xopps_merge(tag, path_list, output_path, pdf_prefix, scoring = False):
    print(tag)
    docs = xopps_open(path_list)
    tree = et.parse('xopp_template.xml')
    root = tree.getroot()
    scores = dict()
    for page in docs:
        if (scoring):
            detect_scores(page, scores)
            
        for page_copy in page[1:]:
            for layer in page_copy.iter('layer'):
                page[0].append(layer)

        if pdf_prefix != None:
            path = os.path.join(pdf_prefix, '{}.pdf'.format(tag))
            page[0].find('background').set('filename', path)
            
        root.append(page[0])

    scores = collections.OrderedDict(sorted(scores.items()))
    xopp_score_table(root, tag, scores)
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
    
def export_excel(path, sheet, scores,
                 match_template = None,
                 tag_order = None):
    book = xlwt.Workbook()
    sh = book.add_sheet(sheet)
    ntags = 1
    problems = None
    for i, tag in enumerate(scores):
        y = i + 1
        n = 0
        if match_template:
            mres = re.match(match_template, tag)
            assert(mres != None)
            n = mres.lastindex
            ntags = mres.lastindex
            for j in range(0, n):
                sh.write(y, j, mres.group(tag_order[j]))

        else:
            n = 1
            sh.write(y, 0, tag)

        s = ''
        if problems == None:
            problems = [key for key in scores[tag]]
            
        for problem in scores[tag]:
            sh.write(y, n, float(scores[tag][problem]))
            if s != '':
                s += '+'

            s += '{}{}'.format(colnum_string(n), y)
            n += 1
            
        sh.write(y, n, xlwt.Formula(s))

    for i in range(ntags):
        sh.write(0, i, 'Tag {}'.format(i))

    for i, problem in enumerate(problems):
        sh.write(0, i + ntags, problem)

    sh.write(0, ntags + len(problems), 'sum')

    book.save(path)
    
    
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
    parser.add_argument('-c', '--scoring-conf', type=str, default=None,
                        help='Path to scoring JSON config (for example, scoring.json).')
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
        if args.scoring_conf:
            with open(args.scoring_conf, 'r') as fl:
                conf = json.load(fl)
                export_excel('scores.xls', 'sheet', scores,
                             match_template = conf['match_template'],
                             tag_order = conf['tag_order'])

        else:
            export_excel('scores.xls', 'sheet', scores)
