import pathlib
import re

import pygubu
import tkinter as tk
import tkinter.ttk as ttk
import requests
import pandas as pd
from tkinter import *
from tkinter import messagebox
import json
from tkinter import ttk, filedialog
from tkinter.filedialog import askopenfile

PROJECT_PATH = pathlib.Path(__file__).parent
PROJECT_UI = PROJECT_PATH / "ti_query.ui"


class TiQueryApp:
    def __init__(self, master=None):
        # build ui
        self.toplevel1 = tk.Tk() if master is None else tk.Toplevel(master)
        self.menu2 = tk.Menu(self.toplevel1)
        self.submenu1 = tk.Menu(self.menu2)
        self.menu2.add(tk.CASCADE, menu=self.submenu1, hidemargin='false', label='File')
        self.mi_command5 = 1
        self.submenu1.add('command', label='command5')
        self.mi_separator2 = 2
        self.submenu1.add('separator')
        self.mi_command7 = 3
        self.submenu1.add('command', label='command7')
        self.mi_command6 = 4
        self.submenu1.add('command', label='command6')
        self.toplevel1.configure(menu=self.menu2)
        self.labelframe1 = ttk.Labelframe(self.toplevel1)
        self.btn_login = ttk.Button(self.labelframe1)
        self.btn_login.configure(text='Login')
        self.btn_login.pack(side='top')
        self.btn_login.configure(command=self.API_login)
        self.separator1 = ttk.Separator(self.labelframe1)
        self.separator1.configure(orient='horizontal')
        self.separator1.pack(side='top')
        self.label1 = ttk.Label(self.labelframe1)
        self.label1.configure(text='Access Token : ')
        self.label1.pack(side='left')
        self.txt_token = tk.Text(self.labelframe1)
        self.txt_token.configure(height='1', width='50')
        self.txt_token.pack(side='top')
        self.labelframe1.configure(height='200', relief='groove', text='Login', width='200')
        self.labelframe1.grid(column='0', row='1', sticky='nw')
        self.labelframe2 = ttk.Labelframe(self.toplevel1)
        self.label3 = ttk.Label(self.labelframe2)
        self.label3.configure(text='GPN')
        self.label3.place(anchor='nw', height='25', x='0', y='0')
        self.et_text = ttk.Entry(self.labelframe2)
        self.et_text.configure(width='25')
        self.et_text.place(height='25', relx='0.0', x='50', y='0')
        self.btn_submit = ttk.Button(self.labelframe2)
        self.btn_submit.configure(text='Submit')
        self.btn_submit.place(bordermode='inside', height='25', relheight='0.0', relwidth='0.0', relx='0.0', rely='0.0', width='100', x='225', y='0')
        self.btn_submit.configure(command=self.productQuery)
        self.tr_view = ttk.Treeview(self.labelframe2)
        self.tr_view_cols = ['column1', 'column2', 'column3', 'column4', 'column5', 'column6', 'column7']
        self.tr_view_dcols = ['column1', 'column2', 'column3', 'column4', 'column5', 'column6', 'column7']
        self.tr_view.configure(columns=self.tr_view_cols, displaycolumns=self.tr_view_dcols)
        self.tr_view.column('column1', anchor='w',stretch='true',width='100',minwidth='10')
        self.tr_view.column('column2', anchor='w',stretch='true',width='50',minwidth='20')
        self.tr_view.column('column3', anchor='w',stretch='true',width='50',minwidth='20')
        self.tr_view.column('column4', anchor='w',stretch='true',width='50',minwidth='20')
        self.tr_view.column('column5', anchor='w',stretch='true',width='50',minwidth='20')
        self.tr_view.column('column6', anchor='w',stretch='true',width='50',minwidth='20')
        self.tr_view.column('column7', anchor='w',stretch='true',width='50',minwidth='20')
        self.tr_view.heading('column1', anchor='w',text='OPN')
        self.tr_view.heading('column2', anchor='w',text='LifeCycle Status')
        self.tr_view.heading('column3', anchor='w',text='Package Type')
        self.tr_view.heading('column4', anchor='w',text='Price')
        self.tr_view.heading('column5', anchor='w',text='MOQ')
        self.tr_view.heading('column6', anchor='w',text='SPQ')
        self.tr_view.heading('column7', anchor='w',text='Lead Time')
        self.tr_view.place(width='950', x='25', y='50')
        self.button4 = ttk.Button(self.labelframe2)
        self.button4.configure(text='Export')
        self.button4.place(anchor='center', relx='0.0', x='60', y='290')
        self.button4.configure(command=self.export_ti)
        self.labelframe2.configure(height='330', text='Ordereable Query', width='1000')
        self.labelframe2.grid(column='0', row='2', sticky='w')
        self.toplevel1.rowconfigure('2', weight='0')
        self.labelframe3 = ttk.Labelframe(self.toplevel1)
        self.label2 = ttk.Label(self.labelframe3)
        self.label2.configure(text='GPN')
        self.label2.place(anchor='nw', height='25', x='0', y='0')
        self.et_text1 = ttk.Entry(self.labelframe3)
        self.et_text1.configure(width='25')
        self.et_text1.place(height='25', relx='0.0', x='50', y='0')
        self.btn_submit1 = ttk.Button(self.labelframe3)
        self.btn_submit1.configure(text='Submit')
        self.btn_submit1.place(bordermode='inside', height='25', relheight='0.0', relwidth='0.0', relx='0.0', rely='0.0', width='100', x='225', y='0')
        self.btn_submit1.configure(command=self.storeQuery)
        self.tr_view1 = ttk.Treeview(self.labelframe3)
        self.tr_view1_cols = ['column8', 'column9', 'column10', 'column12', 'column13', 'column15']
        self.tr_view1_dcols = ['column8', 'column9', 'column10', 'column12', 'column13', 'column15']
        self.tr_view1.configure(columns=self.tr_view1_cols, displaycolumns=self.tr_view1_dcols)
        self.tr_view1.column('column8', anchor='w',stretch='true',width='100',minwidth='10')
        self.tr_view1.column('column9', anchor='w',stretch='true',width='50',minwidth='20')
        self.tr_view1.column('column10', anchor='w',stretch='true',width='400',minwidth='20')
        self.tr_view1.column('column12', anchor='w',stretch='true',width='100',minwidth='20')
        self.tr_view1.column('column13', anchor='w',stretch='true',width='50',minwidth='20')
        self.tr_view1.column('column15', anchor='w',stretch='true',width='200',minwidth='20')
        self.tr_view1.heading('column8', anchor='w',text='OPN')
        self.tr_view1.heading('column9', anchor='w',text='Inventory')
        self.tr_view1.heading('column10', anchor='w',text='Buy Now URL')
        self.tr_view1.heading('column12', anchor='w',text='Package Type')
        self.tr_view1.heading('column13', anchor='w',text='MOQ')
        self.tr_view1.heading('column15', anchor='w',text='SPQ')
        self.tr_view1.place(width='950', x='25', y='50')
        self.export1 = ttk.Button(self.labelframe3)
        self.export1.configure(text='Export')
        self.export1.place(anchor='center', relx='0.0', x='60', y='290')
        self.export1.configure(command=self.export_store)
        self.btn_load1 = ttk.Button(self.labelframe3)
        self.btn_load1.configure(text='Load')
        self.btn_load1.place(bordermode='inside', height='25', relheight='0.0', relwidth='0.0', relx='0.2', rely='0.0', width='100', x='225', y='0')
        self.btn_load1.configure(command=self.storeOPN_Query)
        self.labelframe3.configure(height='330', text='TI Store Query', width='1000')
        self.labelframe3.grid(column='0', row='3', sticky='sw')
        self.labelframe5 = ttk.Labelframe(self.toplevel1)
        self.tr_view2 = ttk.Treeview(self.labelframe5)
        self.tr_view2_cols = ['Details']
        self.tr_view2_dcols = ['Details']
        self.tr_view2.configure(columns=self.tr_view2_cols, displaycolumns=self.tr_view2_dcols)
        self.tr_view2.column('Details', anchor='w',stretch='true',width='160',minwidth='20')
        self.tr_view2.heading('Details', anchor='w',text='Details')
        self.tr_view2.grid(column='0', row='0')
        self.labelframe5.configure(height='0', text='Details Tab', width='0')
        self.labelframe5.grid(column='0', ipadx='70', ipady='0', padx='1010', pady='0', row='2', sticky='n')
        self.labelframe6 = ttk.Labelframe(self.toplevel1)
        self.tr_view4 = ttk.Treeview(self.labelframe6)
        self.tr_view4_cols = ['column17']
        self.tr_view4_dcols = ['column17']
        self.tr_view4.configure(columns=self.tr_view4_cols, displaycolumns=self.tr_view4_dcols)
        self.tr_view4.column('column17', anchor='w',stretch='true',width='160',minwidth='20')
        self.tr_view4.heading('column17', anchor='w',text='Details')
        self.tr_view4.grid(column='0', row='0')
        self.labelframe6.configure(height='0', text='Details Tab', width='0')
        self.labelframe6.grid(column='0', ipadx='70', ipady='0', padx='1010', pady='0', row='3', sticky='n')
        self.toplevel1.geometry('1400x1050')
        self.toplevel1.title('TI Information Tool')

        # Main widget
        self.mainwindow = self.toplevel1

        # self.toplevel1.attributes('-fullscreen', True)
        w, h = self.toplevel1.winfo_screenwidth(), self.toplevel1.winfo_screenheight()
        self.toplevel1.geometry("%dx%d+0+0" % (w, h))
        self.tr_view.column('#0', width=0)
        self.tr_view1.column('#0', width=0)
        # self.tr_view2.column('#0', width=0)
        self.tr_view.bind('<ButtonRelease-1>', self.selectItem)
        self.tr_view1.bind('<ButtonRelease-1>', self.selectItem1)

    def run(self):
        self.mainwindow.mainloop()


    def API_login(self):
        global access_token, pre_access_token
        url = 'https://transact.ti.com/v1/oauth'
        # Opening JSON file
        f = open('config.json', )
        myobj = json.load(f)
        x = requests.post(url, data=myobj)
        data = x.json()
        access_token = data['access_token']
        text1 = "Bearer " + data['access_token']
        self.txt_token.delete('1.0', END)
        self.txt_token.insert('insert', access_token)
        myobj1 = {'token_type': 'bearer', 'grant_type': 'client_credentials',
                  'client_id': "Z3iVoLjdiVZuhAKZ4NLuW35ADKo4grbc", 'client_secret': 'uzAk78Eijo62D6JJ',
                  'access_token': data['access_token'], 'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'}
        # print(access_token)

        url = 'https://transact-pre.ti.com/v1/oauth'
        # Opening JSON file
        f = open('config.json', )
        myobj = json.load(f)
        x = requests.post(url, data=myobj)

        data = x.json()
        pre_access_token = data['access_token']
        text1 = "Bearer " + data['access_token']
        self.txt_token.delete('1.0', END)
        self.txt_token.insert('insert', access_token)
        myobj1 = {'token_type': 'bearer', 'grant_type': 'client_credentials',
                  'client_id': "Z3iVoLjdiVZuhAKZ4NLuW35ADKo4grbc", 'client_secret': 'uzAk78Eijo62D6JJ',
                  'access_token': data['access_token'], 'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'}
        # print(pre_access_token)

    def productQuery(self):
        global access_token
        token = str(self.txt_token.get("1.0", END))
        token = token.strip()
        header = {"Authorization": 'Bearer ' + token, 'Content-Type': 'application/json',
                  'User-Agent': 'Mozilla/5.0'}
        gpn = str(self.et_text.get()).upper()
        gpn = gpn.strip()
        param = {'GenericProductIdentifier': gpn, 'Page': '0', 'Size': '25'}
        r = requests.get('https://transact.ti.com/v1/products/', headers=header, params=param)
        resp = r.json()
        df = pd.json_normalize(resp['Content'])
        cols = list(df.columns)
        self.df_ti_data = df
        df = df[["GenericProductIdentifier", "Identifier", "LifeCycleStatus", "IndustryPackageType", "Price.Value",
                 "MinOrderQty", "StandardPackQty", "LeadTimeWeeks"]]
        df = df.set_index('GenericProductIdentifier')
        df = df.sort_values(by=['Identifier'])
        self.df_ti = df

        for i in self.tr_view.get_children():
            self.tr_view.delete(i)
        for index, row in df.iterrows():
            self.tr_view.insert("", 0, text=index, values=list(row))

    def storeQuery(self):
        global access_token
        token = str(self.txt_token.get("1.0", END))
        token = token.strip()
        header = {"Authorization": 'Bearer ' + token, 'Content-Type': 'application/json',
                  'User-Agent': 'Mozilla/5.0'}
        gpn = str(self.et_text1.get()).upper()
        gpn = gpn.strip()
        param = {'gpn': gpn, 'page': '0', 'size': '25'}
        r = requests.get('https://transact.ti.com/v1/store/products/', headers=header, params=param)
        resp = r.json()
        df = pd.json_normalize(resp['Content'])
        cols = list(df.columns)
        self.df_store_data = df
        # cols = list(df.columns)
        df = df[["GenericProductIdentifier", "ProductIdentifier", "Quantity", "BuyNowURL", "PackageType",
                 "MinimumOrderQuantity", "StandardPackQuantity"]]
        df = df.set_index('GenericProductIdentifier')
        df = df.sort_values(by=['ProductIdentifier'])
        self.df_store = df
        for i in self.tr_view1.get_children():
            self.tr_view1.delete(i)
        for index, row in df.iterrows():
            self.tr_view1.insert("", 0, text=index, values=list(row))

    def storeOPN_Query(self):
        global access_token, pre_access_token
        pd_response = pd.DataFrame()
        token = str(self.txt_token.get("1.0", END))
        token = token.strip()
        pre_token = pre_access_token.strip()
        header = {"Authorization": 'Bearer ' + token, 'Content-Type': 'application/json',
                  'User-Agent': 'Mozilla/5.0'}
        pre_header = {"Authorization": 'Bearer ' + pre_token, 'Content-Type': 'application/json',
                  'User-Agent': 'Mozilla/5.0'}
        import_file_path = filedialog.askopenfilename()
        df1 = pd.read_excel(import_file_path)
        for i in range(len(df1)):
            # print(str(df1.loc[i, "OPN"]))
            weblink = "https://transact.ti.com/v1/store/products/" + str(df1.loc[i, "OPN"])
            r = requests.get(weblink, headers=header)
            resp = r.json()
            pd_resp = pd.DataFrame.from_dict(resp)
            pd_response = pd_response.append(pd_resp)
        #not working as i am not a customer as per Mohammed
        # for i in range(len(df1)):
        #     weblink = "https://transact-pre.ti.com/v1/products/availability/"
        #     param = {
        #         "CustomerPartnerIdentifier": "1",
        #         "AvailabilityRequest": [
        #             {
        #             "CustomerItemReferenceNumber": str(i),
        #             "SupplierProductIdentifier": str(df1.loc[i, "OPN"]),
        #             "CustomerRequestedQuantity": 1000
        #         }
        #         ]
        #     }
        #     print(param)
        #     r = requests.post(weblink, headers=pre_header,data=param)
        #     resp = r.json()
        #     print(resp)
        self.df_store_data = pd_response
        pd_response = pd_response.reset_index()
        pd_response = pd_response[
            ["GenericProductIdentifier", "ProductIdentifier", "Quantity", "BuyNowURL", "PackageType",
             "MinimumOrderQuantity", "StandardPackQuantity"]]
        pd_response = pd_response.set_index('GenericProductIdentifier')
        pd_response = pd_response.sort_values(by=['ProductIdentifier'])
        self.df_store = pd_response
        for i in self.tr_view1.get_children():
            self.tr_view1.delete(i)
        for index, row in pd_response.iterrows():
            self.tr_view1.insert("", 0, text=index, values=list(row))

    def export_ti(self):
        gpn = str(self.et_text.get())
        gpn = gpn.strip()
        f_name = "TI_" + gpn + ".xlsx"
        writer = pd.ExcelWriter(f_name)  # Creates this excel file
        self.df_ti.to_excel(writer, 'sheet1')  # Writes the dataframe to excel file
        writer.save()  # Saves the file
        messagebox.showinfo("Alert", "File Exported Successfully " + f_name)

    def export_store(self):
        gpn = str(self.et_text1.get())
        gpn = gpn.strip()
        if gpn == "":
            f_name = "STORE_" + "OPN" + ".xlsx"
        else:
            f_name = "STORE_" + gpn + ".xlsx"
        writer = pd.ExcelWriter(f_name)  # Creates this excel file
        self.df_store.to_excel(writer, 'sheet1')  # Writes the dataframe to excel file
        writer.save()  # Saves the file
        messagebox.showinfo("Alert", "File Exported Successfully " + f_name)

    def selectItem(self,a):
        curItem = self.tr_view.focus()
        temp=self.tr_view.item(curItem)
        small_dict = {k: v for k, v in temp.items() if k.startswith('values')}.values()
        value_iterator = iter(small_dict)
        first_value = next(value_iterator)
        select_opn=pd.DataFrame(self.df_ti_data.loc[self.df_ti_data['Identifier'] == first_value[0]])
        select_opn= select_opn.transpose()
        select_opn.to_clipboard(sep='\t')
        select_opn = select_opn[::-1]

        # print(self.df_ti_data)
        for i in self.tr_view2.get_children():
            self.tr_view2.delete(i)
        for index, row in select_opn.iterrows():
            self.tr_view2.insert("", 0, text=index, values=list(row))

    def selectItem1(self,a):
        curItem = self.tr_view1.focus()
        temp=self.tr_view1.item(curItem)
        small_dict = {k: v for k, v in temp.items() if k.startswith('values')}.values()
        value_iterator = iter(small_dict)
        first_value = next(value_iterator)
        select_opn=pd.DataFrame(self.df_store_data.loc[self.df_store_data['ProductIdentifier'] == first_value[0]])
        select_opn= select_opn.transpose()
        select_opn.to_clipboard(sep='\t')
        select_opn = select_opn[::-1]

        # print(self.df_store_data)
        for i in self.tr_view4.get_children():
            self.tr_view4.delete(i)
        for index, row in select_opn.iterrows():
            self.tr_view4.insert("", 0, text=index, values=list(row))


if __name__ == '__main__':
    app = TiQueryApp()
    app.run()
