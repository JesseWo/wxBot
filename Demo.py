#!/usr/bin/env python
# coding: utf-8
#

from openpyxl import Workbook


class Demo(object):
    def __init__(self):

        wb = Workbook()

        # grab the active worksheet
        ws = wb.active

        # Data can be assigned directly to cells
        ws['A1'] = 'ID'

        # Rows can also be appended
        count = 1
        for unit in range(101):
            count = count + 1
            ws['A%d' % count] = unit
        # ws.append([1, 2, 3])

        # Save the file
        wb.save("outputs.xlsx")

        print 'outputs.xlsx 导出成功!'


def main():
    demo = Demo()


if __name__ == '__main__':
    main()
