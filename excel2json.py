# -*- coding: utf-8 -*-
from __future__ import unicode_literals, absolute_import

import argparse
import logging as log
from sys import stdout

from core.parser import parse

parser = argparse.ArgumentParser()
parser.add_argument('-f', '--file', help='Excel file path', type=str)
parser.add_argument('-l', '--lvl', help='start col', type=int, default=0)
parser.add_argument('-t', '--title', help='first useful line', type=int, default=2)
parser.add_argument('-n', '--nesting', help='number nested levels', type=int, default=2)
parser.add_argument('-v', action='store_true', help='show INFO log')
parser.add_argument('-O', '--output', help='save to file', type=str)
args = parser.parse_args()

if __name__ == '__main__':
    if args.v:
        log.getLogger().setLevel(log.INFO)

    if args.file is None:
        log.error('You need to set Excel file path:\npython excel2json.py -f file_name.xlsx')
        raise ValueError('file is None')

    res = parse(args.file, lvl=args.lvl, title_line=args.title, nesting=args.nesting)
    print(res, file=open('{}'.format(args.output), 'a') if args.output else stdout)
