from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import Statement, StatementLine
from datetime import datetime
from xlrd import open_workbook,xldate_as_datetime,XL_CELL_DATE,XL_CELL_TEXT
from re import match,search,VERBOSE,IGNORECASE


class ZabaPlugin(Plugin):
    """Croatian Zagrebacka banka plugin XLS
    """

    def get_parser(self, filename):
        return ZabaParser(filename)


class ZabaParser(StatementParser):
    """ Parses Zaba's xls file
    """

    bank_id = 'ZABAHR2X'
    header_date_format = "%d.%m.%Y."

    statement = None
    in_header = True
    row_nr = 0
    datemode = None
    mappings = {
            "date": 0,
            "refnum": 1,
            "memo": 2,
            "debit": 3,
            "credit": 4,
            "balance": 5,
            "currency": 6,
            }

    def __init__(self, filename):
        self.statement = Statement()
        self.filename = filename

    def parse(self):
        """Main entry point for parsers

        super() implementation will call to split_records and parse_record to
        process the file.
        """
        with open_workbook(self.filename) as book:
            self.sh = book.sheet_by_index(0)
            self.datemode = book.datemode
            self.statement.bank_id = self.bank_id
            return super(ZabaParser, self).parse()

    def split_records(self):
        """ Return generator for iterating through each row
        """
        return self.sh.get_rows()

    def parse_record(self, row):
        """Parse given table row and return StatementLine object
        """

        # if we're still in header, no transactions yet, just collect account data
        if self.in_header:
            m = match(r"""(Prometi\ za\ razdoblje\ od\ (?P<start>[0-9.]+)\ do\ (?P<end>[0-9.]+))
                         |(Račun:\ (?P<acct>\w+))   # account           ^- start&end dates
                         |(Valuta:\ (?P<curr>\w+))  # currency
                         |(?P<eoh>Datum)            # end of headers""",
                row[0].value, VERBOSE)
            if m:
                # start/end
                if m['start']:
                    self.statement.start_date = datetime.strptime(m['start'], self.header_date_format)
                    self.statement.end_date = datetime.strptime(m['end'], self.header_date_format)
                # account number
                if m['acct']:
                    self.statement.account_id = m['acct']
                # currency
                elif m['curr']:
                    self.statement.currency = m['curr']
                # end of headers
                elif m['eoh']:
                    self.in_header = False
            self.row_nr += 1
            return None

        # main body of transactions after headers
        stmt_line = StatementLine(amount=0)

        for field, col in self.mappings.items():
            if col >= len(row):
                raise ValueError("Cannot find column %s in a row of %s items "
                                 % (col, len(row)))
            cell = row[col]
            value = self.parse_value(cell)
            # calculate debits and credits to amount
            if field == 'debit' and value != 0:
                stmt_line.amount += value
                stmt_line.trntype = 'DEBIT'
            elif field == 'credit' and value != 0:
                stmt_line.amount -= value
                stmt_line.trntype = 'CREDIT'
            else:
                setattr(stmt_line, field, value)

        # apply generated transaction id
        stmt_line.id = self.gen_id(stmt_line)

        # next row
        self.row_nr += 1
        return stmt_line

    def parse_value(self, cell):
        """ Returns value of the xlrd.sheet.cell with rendering of the date """
        if cell.ctype == XL_CELL_DATE:
            return xldate_as_datetime(cell.value, self.datemode)
        elif cell.ctype == XL_CELL_TEXT:
            return cell.value.strip()
        else:
            return cell.value

    def gen_id(self, stmtln):
        """ generate transaction id
        """
        return self.statement.account_id + stmtln.refnum
