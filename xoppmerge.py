import re
import os
import gzip
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

def xopps_merge(tag, path_list, output_path, pdf_prefix):
    pages = xopps_open(path_list)
    tree = et.parse('xopp_template.xml')
    root = tree.getroot()
    for page in pages:
        for next_page in page[1:]:
            for layer in next_page.iter('layer'):
                page[0].append(layer)

        if pdf_prefix != None:
            path = os.path.join(pdf_prefix, '{}.pdf'.format(tag))
            page[0].find('background').set('filename', path)
            
        root.append(page[0])

    tree.write(gzip.open(output_path, 'wt'), encoding='unicode')

def search_annotations(prefix):
    result = dict()
    mtemp = '(.*)_([0-9]+).pdf.xopp'
    for root, dirs, files in os.walk(prefix):
        for filename in files:
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
    subprocess.Popen(['/bin/bash', '-c', 'xournalpp {} -p {}'.format(xopp_path, pdf_path)])
    
if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input-prefix', type=str, default=os.getcwd(), help='Input prefix')
    parser.add_argument('-o', '--output-prefix', type=str, default=os.getcwd(), help='Output prefix')
    parser.add_argument('-p', '--pdf-prefix', type=str, default=None, help='Prefix to a local PDF files')
    parser.add_argument('-e', '--pdf-export-prefix', type=str, default=None,
                        help='Prefix for a PDF export')
    args = parser.parse_args()
    annotations = search_annotations(args.input_prefix)
    for tag in annotations:
        xopp_path = os.path.join(args.output_prefix, '{}_final.pdf.xopp'.format(tag))
        xopps_merge(tag, annotations[tag], xopp_path, args.pdf_prefix)
        if args.pdf_export_prefix != None:
            pdf_path = os.path.join(args.pdf_export_prefix, '{}_final.pdf'.format(tag))
            pdf_export(xopp_path, pdf_path)
