import numpy
from pyxlsb import open_workbook as open_xlsb
import pandas as pd
# import progressbar
import datetime
import json
import os

import xlsxwriter

from tkinter import *
from tkinter import ttk, messagebox


root = Tk()
def beginning():

    def createFolder(path):
        try:
            if not os.path.exists(path):
                os.makedirs(path)
        except OSError:
            messagebox.showinfo('ERROR', 'Creating directory. ' + path)


    doc_training = {}
    doc_no_training = {}
    verified_trainings_list = ["TRAINING-TAS", "TRAINING-125-TAS", "TRAINING-150-TAS"]

    # with open('excel_variables.json', 'r') as f:
    #     config = json.load(f)
    #--- takes the json file excel_variables.json and puts them to work ----
    user_xls = 'given_reports/user_search.xls'
    ts_order_xlsb = "given_reports/TS Order Detail.xlsb"



    global_list_saved_ts=[]
    def get_spreadSheets():
        '''gets the excel sheets named in the json file and opens them'''
        pd.set_option('display.max_colwidth', 10000)
        try:
            users = pd.read_excel(user_xls)
        except FileNotFoundError:
            messagebox.showinfo('ERROR', 'No such file or directory: "given_reports/user_search.xls"')
            createFolder('./given_reports/')
        ts_r_order = []
        try:
            with open_xlsb(ts_order_xlsb) as wb:
                with wb.get_sheet(2) as sheet:
                    for row in sheet.rows():
                        ts_r_order.append([item.v for item in row])

                ts_r_order = pd.DataFrame(ts_r_order[1:], columns=ts_r_order[0])
        except FileNotFoundError:
            messagebox.showinfo('ERROR', 'No such file or directory: "given_reports/TS Order Detail.xlsb"')
            createFolder('./given_reports/')

        '''removes the first row which does not have the headings and makes the
        second row with the headings the first row'''
        new_header = ts_r_order.iloc[0]
        ts_r_order = ts_r_order[1:]
        ts_r_order.columns = new_header

        compare_for_ts_trainings(ts_r_order, users)

    def save_no_trainings(df):
        '''Formats all relevant info into a json file format'''
        order = df.get('Order #').__str__()
        shipname = df.get("Shipto Name").__str__()
        shipto = df.get('Shipto Address').__str__()
        rep = df.get("Rep Name").__str__()
        serial = df.get("Serial").__str__()
        product = df.get("Product").__str__()
        shipdate = df.get("Ship Date").__str__()
        del_type = df.get("Del Type").__str__()
        notes = df.get("Notes").__str__()

        doc_no_training.update({order:{'Order #': order,
            'ShipTo Name': shipname,
            'ShipTo Address': shipto,
            'Rep Name': rep,
            'Serial': serial,
            'Product': product,
            'Ship Date': shipdate,
            'Del Type': del_type,
            "Notes": notes,}})

    def save_trainings(df, qty):

        '''Formats all relevant info into a json file format'''
        '''
        Move all info from users_excel form to TS_order excel form. 
        This is only for adding the quantity of TS training purchased
        Also add address, city, state, zip, Order header shipping instructions
        '''
        order = df.get('Order #').__str__()
        customer = df.get("Shipto Name").__str__()
        shipto = df.get('Shipto Address').__str__()
        rep = df.get("Rep Name").__str__()
        serial = df.get("Serial").__str__()
        product = df.get("Product").__str__()
        shipdate = df.get("Ship Date").__str__()
        ts_purchased = qty.get('Qty').__str__()
        del_type = df.get("Del Type").__str__()
        notes = df.get("Notes").__str__()

        doc_training.update({order:{'Order #': order,
            'ShipTo Name': customer,
            'ShipTo Address': shipto,
            'Rep Name': rep,
            'Serial': serial,
            'Product': product,
            'Qty': ts_purchased,
            'Ship Date': shipdate,
            'Del Type': del_type,
            "Notes": notes}})


    def tas_writer():
        '''saves the dict of order numbers that had TAS training that was delivered today'''
        createFolder("./tas_trainings/")

        x= datetime.datetime.now()
        with open('tas_trainings/' + str(x.strftime('%m'+'_'+'%d'+'_'+'%Y')) + '.json', 'w+') as fp:
            json.dump(doc_training, fp, indent=4)

    def no_tas_writer():
        '''saves the dict of order numbers that had no training that was delivered today'''
        createFolder("./no_tas_trainings/")

        x= datetime.datetime.now()
        with open('no_tas_trainings/' + str(x.strftime('%m'+'_'+'%d'+'_'+'%Y')) + '.json', 'w+') as fp:
            json.dump(doc_no_training, fp, indent=4)


    def compare_for_ts_trainings(ts_r_order, users):
        order_num = ts_r_order["Item"].isin(verified_trainings_list) #national list comparing TS... paid trainings True/False
        users_combined_comparison = users["Order #"].isin(ts_r_order[order_num]["Order No"]) #local trainings file as a padas.Series
        # pbar = progressbar.ProgressBar(widgets=[progressbar.Percentage(), ' ', progressbar.ETA()])
        # for n in pbar(len(order_num)):
        x=0
        for i in users_combined_comparison:
            if i == True:
                qty = newcompare(ts_r_order,users)
                save_trainings(users.ix[x], qty) # sends all column info that is in users excel file in row x
                x+=1
            else:
                save_no_trainings(users.ix[x])
                x+=1


    def error():
        messagebox.showinfo('Data type has changed in excel file, please contact Admin')
        # print('Data type has changed in excel file, please contact Admin')


    def newcompare(ts_r_order, users):
        ts_r_order_trainings_df = ts_r_order.loc[ts_r_order["Item"].isin(verified_trainings_list)] #creats a df of all rows with training
        users_order_list = users['Order #'].tolist()  #creates a list of all order # in users file

        row_num = 0
        for x in ts_r_order_trainings_df['Order No']:
            for order in users_order_list:
                if int(x) == int(order) and x not in global_list_saved_ts:
                    global_list_saved_ts.append(order) # created list to not duplicate order #'s
                    return ts_r_order_trainings_df.iloc[row_num] #returns the row for Qty trainings sold field
            row_num+=1
        error()

    def main():
        get_spreadSheets()
        tas_writer()
        no_tas_writer()
        messagebox.showinfo('completed', 'Part 1 is done')


    # ----------------------------------------print trainings start------------------------------

    def printer():
        x = datetime.datetime.now()
        order_list = []
        '''need to ask or grab file name'''
        try:
            with open('tas_trainings/' + str(x.strftime('%m' + '_' + '%d' + '_' + '%Y')) + '.json', 'r') as f:
                db = json.load(f)
        except FileNotFoundError:
            messagebox.showinfo('ERROR', 'No TAS trainings found from todays search. Try again with a new "TS Order Detail.xlsb"'
                                         'or "user_search.xls".')
            tkinter_run()


        def writer(order):
            createFolder("./delivery_docs/")

            workbook = xlsxwriter.Workbook("./delivery_docs/trainings" + db[order]['Order #'] + ".xlsx")
            title = 'Order Delivery Report'
            order_num = 'Ricoh Order Number '
            num_order_num = db[order]['Order #']

            worksheet_data = workbook.add_worksheet('TAS')
            header = workbook.add_format({'bold': True, 'font_size': 25})
            bold = workbook.add_format({'border': 1, 'bold': True, 'font_size': 12, 'align': 'left'})
            bold_no_border = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'top'})

            worksheet_data.set_column('A:A', 21)
            worksheet_data.set_column('B:B', 68)
            worksheet_data.set_default_row(25)
            worksheet_data.set_row(13, 75)

            merge_format = workbook.add_format({'border': 1, 'align': 'left',
                                                'text_wrap': True})
            order_num_format = workbook.add_format({'align': 'right', 'bold': True, 'font_size': 12})
            num_order_num_format = workbook.add_format({'align': 'right', 'bold': True, 'font_size': 12, 'valign': 'top'})
            comments_format = workbook.add_format({'align': 'left', 'valign': 'top', 'border': 1, 'bold': True})

            worksheet_data.write('A1:A1', title, header)
            worksheet_data.write('B2:B2', order_num, order_num_format)
            worksheet_data.write('B3:B3', num_order_num, num_order_num_format)
            worksheet_data.write_row('A5', ["Order #"], bold)
            worksheet_data.write_row('A6', ["Customer Name"], bold)
            worksheet_data.write_row('A7', ["Address"], bold)
            worksheet_data.write_row('A8', ["Rep Name"], bold)
            worksheet_data.write_row('A9', ["Serial"], bold)
            worksheet_data.write_row('A10', ["Product"], bold)
            worksheet_data.write_row('A11', ["Number of trainings"], bold)
            worksheet_data.write_row('A12', ["Ship Date"], bold)
            worksheet_data.write_row('A13', ["Del Type"], bold)
            worksheet_data.write_row('A14', ["Notes"], bold)
            worksheet_data.write_row('B19', ["Customer signature: ___________________________  Date:_________"],
                                     bold_no_border)

            worksheet_data.merge_range('B20:B23', 'Comments', comments_format)
            # worksheet_data.write_row('A20', ["Comments:"],)

            worksheet_data.write_row('B5', [db[order]["Order #"]], merge_format),
            worksheet_data.write_row('B6', [db[order]["ShipTo Name"]], merge_format),
            worksheet_data.write_row('B7', [db[order]["ShipTo Address"]], merge_format),
            worksheet_data.write_row('B8', [db[order]["Rep Name"]], merge_format),
            worksheet_data.write_row('B9', [db[order]["Serial"]], merge_format),
            worksheet_data.write_row('B10', [db[order]["Product"]], merge_format),
            worksheet_data.write_row('B11', [db[order]["Qty"]], merge_format),
            worksheet_data.write_row('B12', [db[order]["Ship Date"]], merge_format),
            worksheet_data.write_row('B13', [db[order]["Del Type"]], merge_format),
            worksheet_data.write_row('B14', [db[order]["Notes"]], merge_format),

            workbook.close()


        if db == {}:
            messagebox.showinfo('Trainings', 'There are no trainings within searchable files')
        else:
            for k in db:
                order_list.append(k)

        for order in order_list:
            writer(order)

    # ---------------------tkinter starts-----------------------------------------
    def tkinter_run():
        root.title("Ricoh training generator")

        mainframe = ttk.Frame(root, padding="3 3 12 12")
        mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        mainframe.columnconfigure(0, weight=1)
        mainframe.rowconfigure(0, weight=1)


        # ttk.Label(mainframe, textvariable=meters).grid(column=2, row=2, sticky=(W, E))
        ttk.Button(mainframe, text="Execute First", command=main).grid(column=1, row=4, sticky=W)
        ttk.Button(mainframe, text="Execute Second", command=printer).grid(column=2, row=4, sticky=E)

        ttk.Label(mainframe, text='First verify a folder called “Downloaded Reports” is located in the same directory as '
                                  'this program and has the "TS Order Detail" and "user_search" per directions ', wraplength=200,)\
                                    .grid(column=1, row=1, sticky=W)
        ttk.Label(mainframe, text='Once first button has completed there should be 2 folder "no_tas_trainings" and '
                                  '"tas_trainings". Once they have appeared if there is a file in "tas_trainings" '
                                  'with todays date on it, run the "Execute Second.', wraplength=200,).grid(column=2, row=1, sticky=E)

        for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

        root.bind('<Button-1>', main)
        root.bind('<Button-1>', printer)
    tkinter_run()
beginning()
if __name__==('__main__'):
    root.mainloop()


