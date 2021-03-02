import itertools
import math
import re
import sys
import threading
import time
from queue import Queue

from PySide2.QtWidgets import QFileDialog, QMainWindow, QApplication
from pip._vendor.distlib.compat import raw_input
import xlrd

class Main(QMainWindow):
    def __init__(self,parent):
        super(Main, self).__init__()
        self._running = False
        self.file_ = None
        self.get_file(1)

        if self.file_:
            self.data()
            # self.file2('01/02/2017 09') #remove input msg for month/day/year mm
        else:
            quit()

    def get_file(self,path):
        file_name, _ = QFileDialog.getOpenFileName(self, caption="Data Parse", filter="Excel(*.xlsx)")
        if file_name != '':
            self.file_ = file_name
            self.path = True
            return file_name
        else:
            if path == 1:
                return False
            else:
                self.msg_file()


    def data(self):
        month = raw_input("Enter Month (1-12):")
        m = re.match('[0-9]+', month)
        if m:
            if int(month) > 12:
                print('enter 1-12 for month')
                return self.data()
            else:
                if int(month) in [1,2,3,4,5,6,7,8,9]:
                    self.day(f'0{int(month)}')
                else:
                    self.day(month)
        else:
            print('enter 1-12 for month')
            return self.data()

    def day(self,month):
        days = raw_input("Enter Day (1-31):")
        d = re.match('[0-9]+', days)
        if d:
            if int(days) > 31:
                print('enter 1-31 for day')
                return self.day(month)
            else:
                if int(days) in [1,2,3,4,5,6,7,8,9]:
                    self.year(month,f'0{int(days)}')
                else:
                    self.year(month, days)
        else:
            print('enter 1-31 for day')
            return self.day(month)


    def year(self,m,day):
        years = raw_input("Enter Year :")
        y = re.match('[0-9]+', years)
        if y:
            d = f'{m}/{day}/{years}'
            self.hour(d)

        else:
            print('enter a valid year')
            return self.year(self,m,day)
    def hour(self,_):
        hours = raw_input("Enter hours (0-24):")
        h = re.match('[0-9]+', hours)
        if h:
            if int(hours) > 24:
                print('enter 0-24 hours')
                return self.hour(_)
            else:
                if int(hours) in [1,2,3,4,5,6,7,8,9]:
                    date_filter = f'{_} 0{hours}'
                else:
                    date_filter = f'{_} {hours}'

                t = threading.Thread(target=self.loading)
                t.start()


                que = Queue()
                t1 = threading.Thread(target=lambda q, arg1: q.put(self.file2(arg1)), args=(que, date_filter))
                t1.start()
        else:
            print('enter 0-24 hours')
            return self.hour(_)


    # def file(self,date_):
    #     import pandas as pd
    #     from pandas import ExcelWriter
    #     from pandas import ExcelFile
    #     file = self.file_
    #     df = pd.read_excel(file, sheet_name=None)
    #     # df = pd.read_excel("file/PMO Pier 1 Jan-Dec 2017.xlsx", sheet_name=None)
    #     # df = pd.read_excel("file/Book1.xlsx", sheet_name=None)
    #
    #     # print("Column headings:")
    #     # print(DataF.columns)
    #     # print(df)\
    #     i = 0
    #     o = 0
    #     for d in df:
    #         sheet = pd.read_excel(file, sheet_name=d, header=[0])
    #         # sheet = pd.read_excel("file/PMO Pier 1 Jan-Dec 2017.xlsx", sheet_name=d, header=[0])
    #         # sheet = pd.read_excel("file/Book1.xlsx", sheet_name=d, header=[0])
    #
    #         h = []
    #         for k in sheet.to_dict('dict'):
    #             h.append(k)
    #
    #         berth_dt_key = []
    #         if 'BERTH ARR DATE & TIME' in h:
    #             for k in sheet.to_dict('dict')['BERTH ARR DATE & TIME']:
    #                 try:
    #                     berth_dt = sheet.to_dict('dict')['BERTH ARR DATE & TIME'][k].strftime("%m/%d/%Y")
    #                 except:
    #                     # print(k)
    #                     # print(sheet.to_dict('dict')['BERTH ARR DATE & TIME'][k])
    #                     # print("Please check your excel file empty column found")
    #                     # exit()
    #                     pass
    #
    #                 if berth_dt == date_:
    #                     berth_dt_key.append(k)
    #
    #         for b in berth_dt_key:
    #             in_ = sheet.to_dict('dict')['PASSENGER IN'][b]
    #             out_ = sheet.to_dict('dict')['PASSENGER OUT'][b]
    #
    #             i += float(in_)
    #             o += float(out_)
    #
    #     self.terminate()
    #     print('')
    #     print('==================================')
    #     print(f'Date : {date_}')
    #     print(f'Total PASSENGER IN : {i}')
    #     print(f'Total PASSENGER OUT : {o}')
    #     print('==================================')
    #     print('')

    def file2(self,search):
        import pandas as pd
        file = self.file_
        # df = pd.read_excel("file/PMO Pier 1 Jan-Dec 2017.xlsx", None)
        # df = pd.read_excel("file/Book1.xlsx",sheet_name='Sheet1')
        # df = pd.read_excel("file/Book1.xlsx",None)
        # df = pd.read_excel("file/PMO Pier 1 Jan-Dec 2017.xlsx",None)
        df = pd.read_excel(file,None)

        i = 0
        o = 0
        for d in df.keys():
            print()
            print('Sheet Name : ',d)
            print('Passenger in : ', i)
            print('Passenger out : ', o)
            print()
            record = pd.read_excel(file,sheet_name=d)
            header = record.to_dict('dict').keys()
            if 'BERTH ARR DATE & TIME' in header:
                dataframe = pd.DataFrame(record, columns=header)
                try:
                    date_ = dataframe.loc[dataframe['BERTH ARR DATE & TIME'].dt.strftime('%m/%d/%Y %H') == search]
                except:
                    print(f'Please check BERTH ARR DATE & TIME column date format in {d} sheet')
                    exit()
                i += float(sum(date_['PASSENGER IN'].dropna()))
                o += float(sum(date_['PASSENGER OUT'].dropna()))

            # print(d,search, i)

        # record = pd.read_excel("file/PMO3 Pier 3 Jan-Dec 2017.xlsx", sheet_name='PMOPier3 Jan2017')
        # header = record.to_dict('dict').keys()
        #
        # dataframe = pd.DataFrame(record, columns=header)
        # date_ = dataframe.loc[dataframe['BERTH ARR DATE & TIME'].dt.strftime('%m/%d/%Y %H') == search]

        # print(date_['PASSENGER IN'])
        # print(date_['PASSENGER IN'].astype(float))
        # print(type(date_['PASSENGER IN'].astype(str)))

        # print(date_['PASSENGER IN'].isnull)

        # if date_['PASSENGER IN'].notnull:
        #     i += float(sum(date_['PASSENGER IN']))
        #
        # if date_['PASSENGER OUT'].notnull:
        #     o += float(sum(date_['PASSENGER OUT']))

        self.terminate()
        print('')
        print('==================================')
        print(f'File Path : {self.file_}')
        print(f'Date : {search}')
        print(f'Total PASSENGER IN : {i}')
        print(f'Total PASSENGER OUT : {o}')
        print('==================================')
        print('')


    def msg_(self):
        msg = raw_input('Do you want to search again (Y/N)?:')
        if msg.upper() not in ['Y', 'N', 'YES', 'NO']:
            return self.msg_()
        elif msg.upper() in ['Y', 'YES']:
            self.msg_file()
        else:
            exit()

    def msg_file(self):
        msg = raw_input('New file (Y/N)?:')
        if msg.upper() not in ['Y', 'N', 'YES', 'NO']:
            return self.msg_file()
        elif msg.upper() in ['Y', 'YES']:
            return self.get_file(2)
        else:
            self.data()
    def terminate(self):
        self._running = True

    def loading(self):
        for c in itertools.cycle(['|', '/', '-', '\\']):
            if self._running:
                break
            sys.stdout.write('\rSearching ' + c)
            sys.stdout.flush()
            time.sleep(0.1)
        self._running = False
        self.msg_()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = Main(app)
    sys.exit(app.exec_())
