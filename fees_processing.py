import os, xlrd, re
import pandas as pd

os.chdir('/home/qf/fees_rgkkh')

class fee_file:
    
    def __init__(self, fname):

        self.fname = fname
        self.proc_status=[]

        self.wb = xlrd.open_workbook(fname)
        self.ws = self.wb.sheet_by_name('Sheet 1')

        #Get account number
        try:
            self.ac_no = re.findall(r'\d{13,15}', self.ws.cell(14,4).value)[0]
        except:
            self.proc_status.append("Err - AcNo")
        
        #Get account type - One-time, lodging, tuition fee
        if self.ac_no == '50100103043898':
            self.ac_type = 'Lodging Boarding'
        elif self.ac_no == '50100103044174':
            self.ac_type = 'Tuition'
        elif self.ac_no == '50100103044190':
            self.ac_type = 'One time'
        else:
            self.ac_type = 'Error'
            self.proc_status.append("Err - A/c Type")

        #Get date range
        try:
            self.date_range = self.ws.cell(15,0).value
            self.date_from = re.findall('\d{2}/\d{2}/\d{4}', self.date_range)[0]
            self.date_to = re.findall('\d{2}/\d{2}/\d{4}', self.date_range)[1]
        except:
            self.proc_status.append("Err - Date Range")

        #Get number of rows
        self.last_row = self.ws.col_values(0).index('STATEMENT SUMMARY  :-') - 5
        self.count_rows = self.last_row - 21

        self.data = self.get_data()
        print(self.data.head())


    def get_data(self):
        df = pd.read_excel(self.fname, skiprows=21, nrows=self.count_rows)
        df.columns = ["Dt", "Narr", "Ref", "Dt_post", "Dr", "Cr", "Close",]
        
    
    def __str__(self):
        ret = '\n'.join([self.fname, self.ac_no, self.ac_type, self.date_from + ' - ' + self.date_to \
            , str(self.last_row), str(self.count_rows) \
            ] +  self.proc_status)
        return ret + '\n\n'


if __name__ == '__main__':
    f1 = fee_file("62302760_1569307498937.xls")
    f2 = fee_file("62302760_1569307542076.xls")
    f3 = fee_file("62302760_1569307578042.xls")

    print(f1, f2, f3)