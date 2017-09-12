# -*- coding: utf-8 -*-
from __future__ import unicode_literals, absolute_import

import json
import logging as log

from openpyxl import load_workbook

__all__ = ['parse', ]


def _get_scale(excel_list: list, title: str, lvl: int, title_line: int = 2) -> (int, int):
    """ Get slice for root category
    :param excel_list: excel data
    :param title: category title to search
    :param lvl: level to search
    :param title_line: first useful row
    :return: tuple with first and second row id
    """
    begin, end = None, None

    log.info('TRY TO FIND: {} on LVL {}'.format(title, lvl))

    for idx in range(title_line, len(excel_list)):
        if excel_list[idx][lvl] == title:
            log.info('OK: {} [id: {}; lvl: {}]'.format(title, idx, lvl))
            begin = idx
            if begin == len(excel_list) - 1:
                log.info('FIND LAZY SLICE FOR {}: [{}:{}]'.format(title, begin, begin))
                return begin, begin

            for x in range(idx + 1, len(excel_list)):
                if x == len(excel_list) - 1:
                    end = x
                    log.info('FIND END SLICE FOR {}: [{}:{}]'.format(title, begin, end))
                    return begin, end

                if (excel_list[x][lvl] is not None or (lvl != 0 and excel_list[x][lvl - 1] is not None)) \
                        and excel_list[x][lvl] != title:
                    end = x
                    log.info('FIND SLICE FOR {}: [{}:{}]'.format(title, begin, end))
                    return begin, end

    if not all([begin, end]):
        log.fatal('CAN NOT FIND: {}'.format(title))
        raise ValueError("Can't find {}".format(title))


def _get_ch(excel_list: list, title_slice: tuple, lvl, nesting):
    """ Recursively get a list of children
    :param excel_list: excel data
    :param title_slice: root category slice
    :param lvl: current level
    :param nesting: number nested levels
    :return: list with children with children with children lol ;D
    """
    a, b = title_slice
    r = []
    if a == b and excel_list[a][lvl] is not None:
        print('тут для {} {}'.format(a, b))
        r.append({
            'title': excel_list[a][lvl],
            'children': []
        })
    else:
        for x in range(a, b):
            if excel_list[x][lvl] is not None:
                r.append({
                    'title': excel_list[x][lvl],
                    'children': _get_ch(
                        excel_list,
                        _get_scale(
                            excel_list, excel_list[x][lvl], lvl), lvl + 1, nesting
                    ) if lvl + 1 <= nesting else []
                })

    return r


def parse(wb_path: str, lvl: int, title_line: int, nesting: int) -> list:
    """ Convert excel wb to MPTT ready list
    :param wb_path: path to Excel file
    :param lvl: root category level
    :param title_line: first useful row
    :param nesting: number nested levels

    :return: list of MPTT dict
    """
    excel_list = list(load_workbook(filename=wb_path).active.values)

    r = list()
    for i in range(title_line, len(excel_list)):
        if excel_list[i][lvl] is not None:
            log.info('ROOT CAT: {} [{}:{}]'.format(excel_list[i][lvl], i, lvl))
            r.append(
                {
                    'title': excel_list[i][lvl],
                    'children': _get_ch(
                        excel_list,
                        _get_scale(excel_list, excel_list[i][lvl], lvl),
                        lvl + 1, nesting)
                }
            )

    return json.dumps(r, sort_keys=False, indent=4, ensure_ascii=False, separators=(',', ': '))
