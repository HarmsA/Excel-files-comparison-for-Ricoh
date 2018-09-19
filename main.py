import numpy
from pyxlsb import open_workbook as open_xlsb
import pandas as pd
# import progressbar
import xlsxwriter
import datetime
import json
import os

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
    # verified_trainings_list = ["TRAINING-TAS", "TRAINING-125-TAS", "TRAINING-150-TAS"]

    #--- takes the json file excel_variables.json and puts them to work ----
    user_xls = 'given_reports/user_search.xls'
    ts_order_xlsb = "given_reports/TS Order Detail.xlsb"



    global_list_saved_ts=[]
    def get_spreadSheets():
        '''gets the excel sheets and opens them'''
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

        compare_for_ts_trainings(ts_r_order, users) # both ts_r_order and users are type Dataframe

# ---------------------save no training as json file------------------------------------

    def save_no_trainings(df):
        '''Formats all relevant info into a json file format'''
        order = df.get('Order #').__str__()
        customer = df.get("Customer Name").__str__()
        shipto = df.get('Address').__str__()
        shiptoC = df.get('City').__str__()
        shiptoS = df.get('State').__str__()
        shiptoZ = df.get('Zip').__str__()
        rep = df.get("Rep Name").__str__()
        serial = df.get("Serial").__str__()
        product = df.get("Model").__str__()
        shipdate = df.get("Ship Date").__str__()
        # del_type = df.get("Del Type").__str__()
        notes = df.get("Notes").__str__()

        doc_no_training.update({order:{'Order #': order,
            'ShipTo Name': customer,
            'ShipTo Address': shipto,
            'ShipTo City': shiptoC,
            'ShipTo State': shiptoS,
            'ShipTo Zip': shiptoZ,
            'Rep Name': rep,
            'Serial': serial,
            'Product': product,
            'Ship Date': shipdate,
            # 'Del Type': del_type,
            "Notes": notes,}})

# ---------------------Saved trainings to Json file-----------------------------

    def save_trainings(df):
        '''Formats all relevant info into a json file format'''
        '''
        Move all info from users_excel form to TS_order excel form. 
        This is only for adding the quantity of TS training purchased
        Also add address, city, state, zip, Order header shipping instructions
        '''

        order = df.get('Order #').__str__()
        customer = df.get("Customer Name").__str__()
        shipto = df.get('Address').__str__()
        shiptoC = df.get('City').__str__()
        shiptoS = df.get('State').__str__()
        shiptoZ = df.get('Zip').__str__()
        rep = df.get("Rep Name").__str__()
        serial = df.get("Serial").__str__()
        product = df.get("Model").__str__()
        shipdate = df.get("Ship Date").__str__()
        ts_purchased = df.get('Number of trainings').__str__()
        # del_type = df.get("Del Type").__str__()
        notes = df.get("Notes").__str__()


        doc_training.update({order:{'Order #': order,
            'ShipTo Name': customer,
            'ShipTo Address': shipto,
            'ShipTo City': shiptoC,
            'ShipTo State': shiptoS,
            'ShipTo Zip': shiptoZ,
            'Rep Name': rep,
            'Serial': serial,
            'Product': product,
            'Number of trainings': ts_purchased,
            'Ship Date': shipdate,
            # 'Del Type': del_type,
            "Notes": notes,}})

# ------------------------Create and open training file and json----------------------
    def tas_writer():
        '''saves the dict of order numbers that had TAS training that was delivered today'''
        createFolder("./tas_trainings/")

        x= datetime.datetime.now()
        with open('tas_trainings/' + str(x.strftime('%m'+'_'+'%d'+'_'+'%Y')) + '.json', 'w+') as fp:
            json.dump(doc_training, fp, indent=4)


# ------------------------Create and open no-training file and json----------------------

    def no_tas_writer():
        '''saves the dict of order numbers that had no training that was delivered today'''
        createFolder("./no_tas_trainings/")

        x= datetime.datetime.now()
        with open('no_tas_trainings/' + str(x.strftime('%m'+'_'+'%d'+'_'+'%Y')) + '.json', 'w+') as fp:
            json.dump(doc_no_training, fp, indent=4)

#----------------------

    def compare_for_ts_trainings(ts_r_order, users):
        """Compares the list of verified TAS trainings arguments to the TS Order Details excel sheet,
        then compares the user_search order# to the yes/no order_num of TS Order Details [Order No]"""

        users_ordernum = set()
        for order_num in users["Order #"]:
            if order_num not in users_ordernum:
                users_ordernum.update([order_num])

        for item in users_ordernum:
            dict1 = {}
            models_list = []
            serial_list = []
            trainings_list = []
            count = 0
            for order in ts_r_order['Order No']:
                count += 1
                if str(order) == str(item):
                    if ts_r_order.ix[count]['Order No'] != None:
                        dict1.update({'Order #':ts_r_order.ix[count]['Order No']})
                    if ts_r_order.ix[count]['Customer Name'] != None:
                        dict1.update({'Customer Name': ts_r_order.ix[count]['Customer Name']})
                    if ts_r_order.ix[count]['Shipto Addr'] != None:
                        dict1.update({'Address': ts_r_order.ix[count]['Shipto Addr']})
                    if ts_r_order.ix[count]['Shipto City'] != None:
                        dict1.update({'City': ts_r_order.ix[count]['Shipto City']})
                    if ts_r_order.ix[count]['Shipto State'] != None:
                        dict1.update({'State': ts_r_order.ix[count]['Shipto State']})
                    if ts_r_order.ix[count]['Shipto Zip'] != None:
                        dict1.update({'Zip': ts_r_order.ix[count]['Shipto Zip']})
                    if ts_r_order.ix[count]['Ship Date'] != None:
                        dict1.update({'Ship Date': ts_r_order.ix[count]['Ship Date']})

                    if ts_r_order.ix[count]['Salesrep Name'] != None:
                        dict1.update({'Rep Name': ts_r_order.ix[count]['Salesrep Name']})
                    if ts_r_order.ix[count]['Order Line Shipping Instructions'] != None \
                            and ts_r_order.ix[count]['Order Line Shipping Instructions'] != '.':
                        dict1.update({'Notes': ts_r_order.ix[count]['Order Line Shipping Instructions']})


                    if ts_r_order.ix[count]['Item'] == "TRAINING-TAS" or ts_r_order.ix[count]['Item'] == "TRAINING-150-TAS" \
                        or ts_r_order.ix[count]['Item'] == "TRAINING-125-TAS":
                        trainings_list.append(ts_r_order.ix[count]['Qty'])
                    if ts_r_order.ix[count]['Model'] != None:
                        models_list.append(ts_r_order.ix[count]['Model'])
                    if ts_r_order.ix[count]['Serial'] != None:
                        serial_list.append(ts_r_order.ix[count]['Serial'])

            dict1['Serial'] = serial_list
            dict1['Model'] = models_list
            dict1['Number of trainings'] = trainings_list
            try:
                if dict1['Number of trainings'] != []:
                    save_trainings(dict1)
                else:
                    save_no_trainings(dict1)
            except KeyError:
                print(KeyError)


    def error():
        messagebox.showinfo('Data type has changed in excel file, please contact Admin')

    def main():
        messagebox.showinfo('Running Main Script', 'Please be patient the program is running. Select OK to continue.')
        get_spreadSheets()
        tas_writer()
        no_tas_writer()
        messagebox.showinfo('completed', 'Part 1 is done')


    # ----------------------------------------print trainings start------------------------------

    def printerTrainings():
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

            workbook = xlsxwriter.Workbook("./delivery_docs/"+ db[order]["ShipTo Name"]+"_" + db[order]['Order #'] + ".xlsx")
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
            worksheet_data.set_row(15, 75)
            worksheet_data.set_row(11, 50)
            worksheet_data.set_row(12, 50)

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
            worksheet_data.write_row('A8', ["City"], bold)
            worksheet_data.write_row('A9', ["State"], bold)
            worksheet_data.write_row('A10', ["Zip"], bold)
            worksheet_data.write_row('A11', ["Rep Name"], bold)
            worksheet_data.write_row('A12', ["Serial"], bold)
            worksheet_data.write_row('A13', ["Product"], bold)
            worksheet_data.write_row('A14', ["Number of trainings"], bold)
            worksheet_data.write_row('A15', ["Ship Date"], bold)
            # worksheet_data.write_row('A16', ["Del Type"], bold)
            worksheet_data.write_row('A16', ["Notes"], bold)
            worksheet_data.write_row('B19', ["Customer signature: ___________________________  Date:_________"],
                                     bold_no_border)

            worksheet_data.merge_range('B21:B24', 'Comments', comments_format)

            worksheet_data.write_row('B5', [db[order]["Order #"]], merge_format),
            worksheet_data.write_row('B6', [db[order]["ShipTo Name"]], merge_format),
            worksheet_data.write_row('B7', [db[order]["ShipTo Address"]], merge_format),
            worksheet_data.write_row('B8', [db[order]["ShipTo City"]], merge_format),
            worksheet_data.write_row('B9', [db[order]["ShipTo State"]], merge_format),
            worksheet_data.write_row('B10', [db[order]["ShipTo Zip"]], merge_format),
            worksheet_data.write_row('B11', [db[order]["Rep Name"]], merge_format),
            worksheet_data.write_row('B12', [db[order]["Serial"]], merge_format),
            worksheet_data.write_row('B13', [db[order]["Product"]], merge_format),
            worksheet_data.write_row('B14', [db[order]["Number of trainings"]], merge_format),
            worksheet_data.write_row('B15', [db[order]["Ship Date"]], merge_format),
            # worksheet_data.write_row('B16', [db[order]["Del Type"]], merge_format),
            worksheet_data.write_row('B16', [db[order]["Notes"]], merge_format),

            workbook.close()

        if db == {}:
            messagebox.showinfo('Trainings', 'There are no trainings within searchable files')
        else:
            for k in db:
                order_list.append(k)

        for order in order_list:
            writer(order)
    # ---------------------No training prints--------------------------------------

    def printerNoTrainings():
        x = datetime.datetime.now()
        order_list = []
        '''need to ask or grab file name'''
        try:
            with open('no_tas_trainings/' + str(x.strftime('%m' + '_' + '%d' + '_' + '%Y')) + '.json', 'r') as f:
                db = json.load(f)
        except FileNotFoundError:
            messagebox.showinfo('ERROR', 'No TAS file found from todays search. Try again with a new "TS Order Detail.xlsb"'
                                         'or "user_search.xls".')
            tkinter_run()


        def writer(order):
            createFolder("./no_training_delivery_docs/")

            workbook = xlsxwriter.Workbook("./no_training_delivery_docs/"+ db[order]["ShipTo Name"]+"_"+ db[order]['Order #'] + ".xlsx")
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
            worksheet_data.set_row(15, 75)
            worksheet_data.set_row(11, 50)
            worksheet_data.set_row(12, 50)

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
            worksheet_data.write_row('A8', ["City"], bold)
            worksheet_data.write_row('A9', ["State"], bold)
            worksheet_data.write_row('A10', ["Zip"], bold)
            worksheet_data.write_row('A11', ["Rep Name"], bold)
            worksheet_data.write_row('A12', ["Serial"], bold)
            worksheet_data.write_row('A13', ["Product"], bold)
            worksheet_data.write_row('A14', ["Number of trainings"], bold)
            worksheet_data.write_row('A15', ["Ship Date"], bold)
            # worksheet_data.write_row('A16', ["Del Type"], bold)
            worksheet_data.write_row('A16', ["Notes"], bold)
            worksheet_data.write_row('B19', ["Customer signature: ___________________________  Date:_________"],
                                     bold_no_border)

            worksheet_data.merge_range('B20:B23', 'Comments', comments_format)

            worksheet_data.write_row('B5', [db[order]["Order #"]], merge_format),
            worksheet_data.write_row('B6', [db[order]["ShipTo Name"]], merge_format),
            worksheet_data.write_row('B7', [db[order]["ShipTo Address"]], merge_format),
            worksheet_data.write_row('B8', [db[order]["ShipTo City"]], merge_format),
            worksheet_data.write_row('B9', [db[order]["ShipTo State"]], merge_format),
            worksheet_data.write_row('B10', [db[order]["ShipTo Zip"]], merge_format),
            worksheet_data.write_row('B11', [db[order]["Rep Name"]], merge_format),
            worksheet_data.write_row('B12', [db[order]["Serial"]], merge_format),
            worksheet_data.write_row('B13', [db[order]["Product"]], merge_format),
            # worksheet_data.write_row('B14', [db[order]["Qty"]], merge_format),
            worksheet_data.write_row('B15', [db[order]["Ship Date"]], merge_format),
            # worksheet_data.write_row('B16', [db[order]["Del Type"]], merge_format),
            worksheet_data.write_row('B16', [db[order]["Notes"]], merge_format),

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
        ttk.Button(mainframe, text="TAS Trainings", command=printerTrainings).grid(column=2, row=4, sticky=E)
        ttk.Button(mainframe, text="No TAS Trainings", command=printerNoTrainings).grid(column=3, row=4, sticky=W)

        ttk.Label(mainframe, text='First verify a folder called “Downloaded Reports” is located in the same directory as '
                                  'this program and has the "TS Order Detail" and "user_search" per directions ', wraplength=200,)\
                                    .grid(column=1, row=1, sticky=W)
        ttk.Label(mainframe, text='Once first button has completed there should be 2 folder "no_tas_trainings" and '
                                  '"tas_trainings". Once they have appeared if there is a file in "tas_trainings" '
                                  'with todays date on it, run the "Execute Second.', wraplength=200,).grid(column=2, row=1, sticky=E)

        for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

        root.bind('<Button-1>', main)
        root.bind('<Button-1>', printerTrainings)
        root.bind('<Button-1>', printerNoTrainings)
    tkinter_run()
beginning()
if __name__==('__main__'):
    root.mainloop()


