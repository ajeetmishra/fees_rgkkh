import os, xlrd, re
import numpy as np
import pandas as pd

os.chdir('/home/qf/Documents/QfProjects/fees_rgkkh/')

class fee_file:
    
    def __init__(self, fname):

        self.fname = fname
        self.proc_status=[]

        self.wb = xlrd.open_workbook(self.fname)
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
            self.ac_type = 'Unknown'
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




    def get_data(self):
        df = pd.read_excel(self.fname, skiprows=21, nrows=self.count_rows)
        df.columns = ["TxnDate", "Narr", "Ref", "Dt_post", "Dr", "Cr", "Close",]
        
        #Remove debit entries
        df = df[np.isnan(df.Dr)]
        
        #Drop columns not needed
        df = df.drop(columns=['Dt_post', 'Dr', 'Close'])

        #Remove credit entries of Internal Transfer
        df = df[df.Narr.str.contains('INTERNAL TRANSFER') == False]

        #Remove NEFT returns
        df = df[df.Narr.str.contains('NEFT RETURN') == False]

        #Remove credit interest capitalized
        df = df[df.Narr.str.contains('CREDIT INTEREST CAPITALISED') == False]

        # Get payment mode
        df['Mode'] = df.Narr.apply(self.get_paymode)

        # Convert string dates to actual dates
        df.TxnDate = pd.to_datetime(df.TxnDate, format="%d/%m/%y")

        # Create yymm
        df['YYMM'] = str(pd.DatetimeIndex(df['TxnDate']).year) + "-" + str(pd.DatetimeIndex(df['TxnDate']).month)
        df['YYMM'] = pd.to_datetime(df['TxnDate']).dt.to_period('M')
        
        # Capture account type as column
        df['Acc Type'] = self.ac_type
        
        #Guess student name
        df['Student'] = df.Narr.apply(self.guess_student)
        
        #Sort by student name
        df = df.sort_values('Student')

        return df



    def get_paymode(self, text):
        if text.startswith('IMPS'):
            return 'IMPS'
        elif text.startswith('NEFT'):
            return 'NEFT'
        elif '-TPT-' in text:
            return 'TPT'
        elif text.startswith('UPI'):
            return 'UPI'
        elif 'NET BANKING SI' in text:
            return 'SI'
        elif 'CHQ DEP' in text or 'CHEQUE DEPOSIT' in text:
            return 'CHQ'
        elif 'CASH DEP' in text:
            return 'CASH'
        else:
            return ''

        
    def guess_student(self, text):
        pattern_m = {}
        pattern_m['XX - Sumit/Vidula Naikode'] = 'SHAILESH DAGADU'
        pattern_m['XX - Kali/Shri Shukla'] = 'KALI SHRI'
        pattern_m['XX - Shreya/Rishi'] = 'SANDEEP MADHUKAR'
        pattern_m['XX - Riya/Siya'] = 'BHUPENDRA'
        pattern_m['XX - Vinayak/Shivraj'] = 'GAJANAN SHANKAR SHINDE'

        pattern_m['PP - Aarohi Mishra'] = 'AAROHI ISHI'

        pattern_m['03 - Dev Madkaikar'] = 'MAST DEV'
        pattern_m['03 - Tashvi Vartak'] = 'TASHVI'
        
        pattern_m['04 - Joy Tanna'] = 'JOY'
        pattern_m['04 - Agrima Mishra1'] = 'AGRIMA'
        pattern_m['04 - Agrima Mishra2'] = 'MAHADEV'
        
        pattern_m['05 - Gaurav Chopra'] = 'RAHUL VINOD CHOPRA'
        pattern_m['05 - Rohit Nayak'] = 'RASHMI V NAYAK'
        
        pattern_m['06 - Isha Bapat'] = 'ISHA BAPAT'
        
        pattern_m['07 - Yash Gajra'] = 'YASH GAJRA'
        pattern_m['07 - Harsh Ambesange'] = 'RAKHEE SUDHANSHU AMBESANGE'
        pattern_m['07 - Jaivik Vyas'] = 'RENUKADARSHANVYAS'
        pattern_m['07 - Kaustubh Abnave'] = 'KAUSTUBH'
        pattern_m['07 - Saee Gujare'] = 'PRASHANT DATTU GUJAR'
        pattern_m['07 - Sarthak Sawant'] = 'SARTHAK'

        pattern_m['08 - Aditya Mahajan'] = 'ANIL N MAHAJAN'
        pattern_m['08 - Arnav Sawant1'] = 'AARYAAMITSAWANT'
        pattern_m['08 - Arnav Sawant2'] = 'RAVINDRA R SAWANT'
        pattern_m['08 - Diya Pamnani'] = 'PALLAVI HIRANAND'
        pattern_m['08 - Rishi Agarwal1'] = 'MINU AGARWAL'
        pattern_m['08 - Rishi Agarwal2'] = 'RITU AGARWAL'
        pattern_m['08 - Yuvaraj Mudaliyar1'] = 'VIJAY RAMSWAMY MUDALIYAR'
        pattern_m['08 - Yuvaraj Mudaliyar2'] = 'VIJAY RAMSWAMY MUDAL'
        pattern_m['09 - Aditya Bhise'] = 'NITIN CHHAGAN BHISE'
        
        pattern_m['09 - Moulya'] = 'MOULYA'
        pattern_m['09 - Krushna Borhade'] = 'BORHADE'
        
        pattern_m['10 - Arya Gawas'] = 'PRADIP GOVIND GAWAS'
        pattern_m['10 - Bhavya'] = 'BHAVYA PATEL'
        pattern_m['10 - Parth Agarwal'] = 'GANESH MAHAVIR PRASAD AGARWAL'        
        pattern_m['10 - Rucha Shelar'] = 'SANJAY S SHELAR'
        pattern_m['10 - Shubh Shah1'] = 'NAYAN VASANT SHAH'
        pattern_m['10 - Shubh Shah2'] = 'SHUBH TRADING'
        pattern_m['10 - Rasika Desai'] = 'CHAITRA DESAI'
        pattern_m['10 - Riya Patil'] = 'SHAILESH S PATIL'
        pattern_m['10 - Sai Bagade'] = 'SANTOSH SHANTARAM'
        pattern_m['10 - Sakshi Tekawade'] = 'OCEAN'
        pattern_m['10 - Samruddhi Pawar1'] = 'AMOL JANAR'
        pattern_m['10 - Samruddhi Pawar2'] = 'AMOL PAWAR'
        pattern_m['10 - Siddhant Kamble'] = 'CHOPSTIX'
        pattern_m['10 - Soham Gawali'] = 'SURYAKANT JAGANNATH'
        pattern_m['10 - Ved Ahire1'] = 'VED AHIRE'
        pattern_m['10 - Ved Ahire2'] = '28021000006912'

        pattern_m['Past - Kukreja'] = 'KUKREJA'
        pattern_m['Past - Hariharan'] = 'HARIHARAN'
        pattern_m['Past - Sanavi'] = 'VISHWRAJ SHIVRAJ'
        pattern_m['Past - Utsav'] = 'ARVIND KUMAR GUPTA'

        for key, val in pattern_m.items():
            if val in text:
                return key
        return None


    def get_minmaxdate(self, df):
        # Get max and min date from real data to validate data-importing quality
        return str(min(df.TxnDate)), str(max(df.TxnDate))


    def __str__(self):

        ret = '\n'.join([self.fname, self.ac_no, self.ac_type, self.date_from + ' - ' + self.date_to \
            , str(self.last_row), str(self.count_rows) \
            ] +  self.proc_status)
        return ret + '\n\n'


if __name__ == '__main__':
    f1 = fee_file("62302760_1569307498937.xls")
    df1 = f1.get_data()
    f2 = fee_file("62302760_1569307542076.xls")
    df2 = f2.get_data()
    f3 = fee_file("62302760_1569307578042.xls")
    df3 = f3.get_data()

    print(f1, df1, f2, df2, f3, df3)

    df_all = pd.concat([df1, df2, df3])

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    # Also set the default datetime and date formats.
    df_all.to_excel("output.xlsx", engine='xlsxwriter', index=False)

    # Pivot the data
    # print(pd.pivot_table(df1, index=["YYMM"], values=["Cr"]))