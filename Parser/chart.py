import itertools
import re
import sys
import threading
import time
from queue import Queue

from pip._vendor.distlib.compat import raw_input
import xlrd
from PySide2.QtWidgets import QFileDialog, QMainWindow, QApplication

class Main(QMainWindow):
    def __init__(self,parent):
        super(Main, self).__init__()
        self._running = False

        passenger_in = []

        # passenger_in.append(self.get_passenger_in())

        # self.month()
        self.total_pax_in = []
        self.file_ = None
        self.get_file()
        if self.file_:
            self.year_chart()
            # self.file2('01/02/2017 09') #remove input msg for month/day/year mm
        else:
            quit()

    def get_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, caption="Data Parse", filter="Excel(*.xlsx)")
        if file_name != '':
            self.file_ = file_name
            return file_name
        else:
            return False

    def year_chart(self):
        years = raw_input("Enter Year :")
        y = re.match('[0-9]+', years)
        if y:

            t = threading.Thread(target=self.loading)
            t.start()

            d = f'/{years}'
            que = Queue()
            t1 = threading.Thread(target=lambda q, arg1: q.put(self.get_year_passenger_in(arg1)), args=(que, d))
            t1.start()

        else:
            print('enter a valid year')
            return self.year_chart(self)

    def get_year_passenger_in(self,search):
        import pandas as pd
        # df = pd.read_excel("file/PMO Pier 1 Jan-Dec 2017.xlsx", None)
        # df = pd.read_excel("file/Book1.xlsx",sheet_name='Sheet1')
        # df = pd.read_excel("file/Book1.xlsx",None)
        # df = pd.read_excel("file/PMO Pier 1 Jan-Dec 2017.xlsx",None)
        # df = pd.read_excel("file/PMO3 Pier 3 Jan-Dec 2017.xlsx",None)
        df = pd.read_excel(self.file_,None)


        self.total_pax_in = []
        in_month = {}
        for month in range(1, 13):
            i = 0
            o = 0

            if int(month) in [1, 2, 3, 4, 5, 6, 7, 8, 9]:
                month_filter = f'0{int(month)}'
            else:
                month_filter = month

            for d in df.keys():

                print()
                print('Sheet Name : ', d)
                print('Month : ', month)
                print()

                record = pd.read_excel(self.file_,sheet_name=d)
                # record = pd.read_excel("file/PMO3 Pier 3 Jan-Dec 2017.xlsx",sheet_name=d)
                header = record.to_dict('dict').keys()
                if 'BERTH ARR DATE & TIME' in header:
                    dataframe = pd.DataFrame(record, columns=header)

                    _ = f'{month_filter}{search}'

                    date_ = dataframe.loc[dataframe['BERTH ARR DATE & TIME'].dt.strftime('%m/%Y') == _]
                    i += float(sum(date_['PASSENGER IN'].dropna()))
                    # o += float(sum(date_['PASSENGER OUT']))
                    in_month[month] = i

            self.total_pax_in.append(i)

        self.terminate()
        # print('')
        # print('==================================')
        # print(f'Date : {search}')
        # print(f'Total PASSENGER IN : {i}')
        # print(f'Total PASSENGER OUT : {o}')
        # print('==================================')
        # print('')

        # return i

    def terminate(self):
        self._running = True

    def loading(self):
        for c in itertools.cycle(['|', '/', '-', '\\']):
            if self._running:
                break
            sys.stdout.write('\rSearching ' + c)
            sys.stdout.flush()
            time.sleep(0.1)
        sys.stdout.write('\rGenarate chart ')
        self._running = False
        self.year_generate_chart()


    def msg_(self):
        msg = raw_input('Do you want to search again (Y/N)?:')
        if msg.upper() not in ['Y', 'N', 'YES', 'NO']:
            return self.msg_()
        elif msg.upper() in ['Y', 'YES']:
            return self.year_chart()
        else:
            exit()


    def year_generate_chart(self):
        import pandas as pd
        import matplotlib.pyplot as plt

        data = {'months': ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE',
                  'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'],
                'passenger_in': self.total_pax_in}

        print(data)
        df = pd.DataFrame(data, columns=['months', 'passenger_in'])
        df.plot(x='months', y='passenger_in', kind='bar')
        plt.show()

        self.msg_()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = Main(app)
    sys.exit(app.exec_())
