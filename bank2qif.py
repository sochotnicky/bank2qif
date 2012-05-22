#!/usr/bin/env python
# -*- coding: utf-8 -*-
# mbank2qif.py - mBank CSV output convertor to QIF (quicken) file
# Copyright (C) 2011  Stanislav Ochotnicky <stanislav@ochotnicky.com>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see http://www.gnu.org/licenses.

import codecs
import csv
import argparse
import re
from datetime import date


def unicode_csv_reader(unicode_csv_data, dialect=csv.excel, **kwargs):
    # csv.py doesn't do Unicode; encode temporarily as UTF-8:
    csv_reader = csv.reader(utf_8_encoder(unicode_csv_data),
                            dialect=dialect, **kwargs)
    for row in csv_reader:
        # decode UTF-8 back to Unicode, cell by cell:
        yield [unicode(cell, 'utf-8') for cell in row]


def utf_8_encoder(unicode_csv_data):
    for line in unicode_csv_data:
        yield line.encode('utf-8')

def normalize_field(text):
    ret = re.sub(BankImporter.multispace_re, " ", text)
    return ret.replace('"', '').replace("'", "").strip()

def normalize_num(text):
    return text.replace(',', '.').replace(' ', '').strip()


class TransactionData(object):
    """Simple class to hold information about a transaction"""
    def __init__(self, date, amount, destination=None, message=None):
        self.date = date
        self.amount = amount
        self.destination = destination
        self.message = message


class BankImporter(object):
    """Base class for statement import"""

    multispace_re = re.compile('\s+')

    def __init__(self, infile):
        self.infile = open(infile, 'r')
        self.reader = codecs.getreader("utf-8")
        self.inputreader = self.reader(self.infile)
        self.transactions = []

    def bank_import(self):
        """Run import from file and return list of transactions

        bank_import() -> [TransactionData, ...]"""
        pass


class MBankImport(BankImporter):
    source = "mbank"

    def __init__(self, infile):
        BankImporter.__init__(self, infile)
        self.reader = codecs.getreader("cp1250")
        self.inputreader = self.reader(self.infile)

    def bank_import(self):
        items = False
        for row in unicode_csv_reader(self.inputreader.readlines(),
                                      delimiter=';'):
            if not items and len(row) > 0 and ( \
                    row[0] == u"#Datum uskutečnění transakce" or \
                    row[0] == u"#Dátum uskutočnenia transakcie" \
			):
                items = True
                continue
            if items:
                if len(row) == 0:
                    break
                d, m, y = row[1].split('-')
                tdate = date(int(y), int(m), int(d))
                tamount = float(normalize_num(row[9]))

                trans_type = normalize_field(row[2])
                trans_desc = normalize_field(row[3])
                trans_target = normalize_field(row[4])
                trans_acc = normalize_field(row[5])
                tmessage = u"M%s %s %s %s" % (trans_type,
                                              trans_desc,
                                              trans_target,
                                              trans_acc)
                self.transactions.append(TransactionData(tdate,
                                                         tamount,
                                                         message=tmessage))
        return self.transactions


class UnicreditImport(BankImporter):
    source = "unicredit"

    def bank_import(self):
        items = False
        for row in unicode_csv_reader(self.inputreader.readlines(),
                                      delimiter=';'):
            if not items and len(row) > 0 and row[0] == u"Účet":
                items = True
                continue
            if items:
                if len(row) == 0:
                    break
                y, m, d = row[3].split('-')
                tdate = date(int(y), int(m), int(d))
                tamount = float(normalize_num(row[1]))
                bank_no = row[5]
                bank_name = "%s %s" % (normalize_field(row[6]),
                                       normalize_field(row[7]))
                bank_name = bank_name.strip()
                account_number = normalize_field(row[8])
                account_name = normalize_field(row[9])
                tdest = None
                if account_number != "":
                    tdest = "%s: %s/%s %s" % (bank_name,
                                              account_number,
                                              bank_no,
                                              account_name)

                t_type = row[13].strip()
                if t_type == u"PLATBA PLATEBNÍ KARTOU" and \
                        tdest == None:
                    # when paid by card the description of place is in
                    # last of "transaction details"
                    for i in reversed(range(13,19)):
                        if normalize_field(row[i]) != "":
                            tdest = "%s" % (normalize_field(row[i]))
                            break

                tmessage = "%s %s %s %s %s %s" % (row[13],
                                                  row[14],
                                                  row[15],
                                                  row[16],
                                                  row[17],
                                                  row[18])
                tmessage = normalize_field(tmessage)
                self.transactions.append(TransactionData(tdate,
                                                         tamount,
                                                         message=tmessage,
                                                         destination=tdest))
        return self.transactions


def write_qif(outfile, transactions):
    with open(outfile, 'w') as output:
        writer = codecs.getwriter("utf-8")
        outputwriter = writer(output)
        outputwriter.write("!Type:Bank\n")
        for transaction in transactions:
            d, m, y = (transaction.date.day,
                      transaction.date.month,
                      transaction.date.year)
            outputwriter.write(u"D%s/%s/%s\n" % (m, d, y))
            outputwriter.write(u"T%s\n" % transaction.amount)
            if transaction.message:
                outputwriter.write(u"M%s\n" % transaction.message)
            if transaction.destination:
                outputwriter.write(u"P%s\n" % transaction.destination)
            outputwriter.write(u'^\n')


if __name__ == "__main__":
    importers = [MBankImport, UnicreditImport]
    sources = []
    for importer in importers:
        sources.append(importer.source)

    parser = argparse.ArgumentParser(description='Bank statement to QIF file converter')
    parser.add_argument('-i', '--input',
                        help='input file to process [default:stdin]',
                        default='/dev/stdin')
    parser.add_argument('-o', '--output',
                        help='output file [default:stdout]',
                        default='/dev/stdout')
    parser.add_argument('-t', '--type',
                        help='Type of input file [default:mbank]'
                             ' Possible values: %s' % ', '.join(sources),
                        default='mbank')
    args = parser.parse_args()
    transactions = []
    for importer in importers:
        if args.type == importer.source:
            inst = importer(args.input)
            transactions = inst.bank_import()

    write_qif(args.output, transactions)
