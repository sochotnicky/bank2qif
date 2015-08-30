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


class BadRecordTypeException(Exception):
    def __init__(self, line_no):
        self._line_no = line_no

    def __str__(self):
        return "Bad record type on line: %s" % (self._line_no,)


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


IMPORTERS = {}


def register_importer(source):
    def f(cls):
        assert source not in IMPORTERS, "More importers for the source?"
        IMPORTERS[source] = cls
        return cls
    return f


class TransactionData(object):
    """Simple class to hold information about a transaction"""
    def __init__(self, date, amount, destination=None, message=None,
            ident=None):
        self.date = date
        self.amount = amount
        self.destination = destination
        self.message = message
        self.ident = ident


class BankImporter(object):
    """Base class for statement import

    To work properly, you need to implement __iter__.
    Maybe the easyiest way is to call yield ever time
    you have new TransactionData."""

    multispace_re = re.compile('\s+')
    input_encoding = "utf-8"

    def __init__(self, infile):
        reader = codecs.getreader(self.input_encoding)
        self.inputreader = reader(open(infile, 'r'))

    def __iter__(self):
        pass


@register_importer("mbank")
class MBankImport(BankImporter):
    input_encoding = "cp1250"

    def __iter__(self):
        items = False
        for row in unicode_csv_reader(self.inputreader, delimiter=';'):
            if not items and len(row) > 0 and (
                    row[0] == u"#Datum uskutečnění transakce" or
                    row[0] == u"#Dátum uskutočnenia transakcie"
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

                # Extract date and the actual description from the "description" field
                matches = re.match(r"(.+)\s+DATUM PROVEDEN. TRANSAKCE: (\d{4})-(\d{2})-(\d{2})", trans_desc)
                if matches:
                    trans_desc = matches.group(1)
                    tdate = date(int(matches.group(2)), int(matches.group(3)), int(matches.group(4)))

                tmessage = u"%s %s %s %s" % (trans_type,
                                              trans_desc,
                                              trans_target,
                                              trans_acc)
                
                # print tmessage
                # print row

                yield TransactionData(tdate, tamount, message=tmessage)


@register_importer("kb")
class KBImport(BankImporter):
    input_encoding = "cp1250"

    def __iter__(self):
        items = False
        for row in unicode_csv_reader(self.inputreader, delimiter=';'):
            if not items and len(row) > 0 and (row[0] == u"Datum splatnosti"):
                items = True
                continue
            if items:
                if len(row) == 0:
                    break
                d, m, y = row[0].split('.')
                tdate = date(int(y), int(m), int(d))
                tamount = float(normalize_num(row[4]))

                #print ' :: '.join(row)
                trans_type = normalize_field(row[12]) # also 3
                trans_desc = normalize_field(row[15])
                trans_target = normalize_field(row[16])
                #trans_acc = normalize_field(row[3])
                trans_acc = ''
                tmessage = u"%s %s %s %s" % (trans_type,
                                              trans_desc,
                                              trans_target,
                                              trans_acc)
                #print tmessage
                #return
                yield TransactionData(tdate, tamount, message=tmessage)




@register_importer("unicredit")
class UnicreditImport(BankImporter):
    def __iter__(self):
        items = False
        for row in unicode_csv_reader(self.inputreader, delimiter=';'):
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
                        tdest is None:
                    # when paid by card the description of place is in
                    # last of "transaction details"
                    for i in reversed(range(13, 19)):
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
                yield TransactionData(tdate, tamount, message=tmessage,
                                      destination=tdest)


@register_importer("zuno")
class ZunoImport(BankImporter):
    def __iter__(self):
        items = False
        for row in unicode_csv_reader(self.inputreader, delimiter=';'):
            if not items and len(row) > 0 and row[0] == u"Dátum transakcie:":
                items = True
                continue
            if items:
                if len(row) <= 1:
                    break

                d, m, y = row[0].split('.')
                tdate = date(int(y), int(m), int(d))
                tamount = float(normalize_num(row[6]))

                account_number = normalize_field(row[3])
                bank_code = normalize_field(row[4])
                tdest = None

                if account_number != "":
                    tdest = "%s/%s" % (account_number, bank_code)
                    tdest = tdest.strip()

                tmessage = normalize_field(row[5])
                yield TransactionData(tdate, tamount, message=tmessage,
                                      destination=tdest)


@register_importer("fio")
class FioImport(BankImporter):
    input_encoding = "cp1250"

    def __iter__(self):
        # For GPC format documentation see here:
        # http://www.fio.cz/docs/cz/struktura-gpc.pdf
        line_no = 1

        # The first line contains info about account
        line = self.inputreader.readline()
        record_type = line[0:3]
        if record_type != '074':
            raise BadRecordTypeException(line_no)

        # The following lines contain transactions
        for line in self.inputreader:
            line_no += 1
            # Record type must be '075' (transaction)
            record_type = line[0:3]
            if record_type != '075':
                raise BadRecordTypeException(line_no)

            # Transaction type; 1 = debet, 2 = credit, 4 = storno of debet,
            # 5 = storno of credit
            ttype = int(line[60])

            # Transaction amount
            tamount = float(line[48:60]) / 100.0
            if ttype in (1, 5):
                tamount = -tamount

            # Transaction date (DDMMYY)
            d = int(line[122:124])
            m = int(line[124:126])
            y = 2000 + int(line[126:128])
            tdate = date(y, m, d)

            # Destination account
            tbankcode = line[71:81].lstrip('0')[0:4]
            tdestacc = line[19:35].lstrip('0')
            tdest = tdestacc and ('%s/%s' % (tdestacc, tbankcode)) or None

            # Message
            tmessage = line[97:117].strip()

            # Transaction identifier
            tident = line[35:48]

            # Append transaction to the list
            yield TransactionData(tdate, tamount, destination=tdest,
                                  message=tmessage, ident=tident)


@register_importer("slsp")
class SlSpImport(BankImporter):
    input_encoding = "cp1250"

    def __iter__(self):
        for row in unicode_csv_reader(self.inputreader, delimiter=';'):
            if len(row) <= 1:
                break

            d, m, y = row[0].split('.')
            tdate = date(int(y), int(m), int(d))
            tamount = float(normalize_num(row[7]))

            account_number = normalize_field(row[4])
            prepend_number = normalize_field(row[3])
            bank_code = normalize_field(row[5])
            tdest = None

            if account_number != "":
                    tdest = "%s/%s" % (account_number, bank_code)
                    tdest = tdest.strip()
                    if prepend_number != "":
                        tdest = "%s-%s" % (prepend_number, tdest)
                        tdest = tdest.strip()

            account_name = normalize_field(row[6])
            name = normalize_field(row[11])
            information = normalize_field(row[16])
            extend_information = normalize_field(row[18]) + " " + \
                normalize_field(row[22])

            tmessage = "%s %s %s" % (name, information, extend_information)
            tmessage = tmessage.strip()
            if account_name != "":
                tmessage += ", " + account_name

            yield TransactionData(tdate, tamount, message=tmessage,
                                  destination=tdest)


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
            if transaction.ident:
                outputwriter.write(u"#%s\n" % transaction.ident)
            if transaction.message:
                outputwriter.write(u"M%s\n" % transaction.message)
            if transaction.destination:
                outputwriter.write(u"P%s\n" % transaction.destination)
            outputwriter.write(u'^\n')


if __name__ == "__main__":
    sources = sorted(IMPORTERS.keys())

    parser = argparse.ArgumentParser(
        description='Bank statement to QIF file converter')
    parser.add_argument('-i', '--input',
                        help='input file to process [default:stdin]',
                        default='/dev/stdin')
    parser.add_argument('-o', '--output',
                        help='output file [default:stdout]',
                        default='/dev/stdout')
    parser.add_argument('-t', '--type',
                        help='Type of input file [default:mbank]',
                        choices=sources,
                        default='mbank')
    args = parser.parse_args()
    importer_class = IMPORTERS.get(args.type)
    importer = importer_class(args.input)
    write_qif(args.output, importer)
