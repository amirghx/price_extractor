import tkinter as tk
import tkinter.ttk as ttk
from urllib.request import urlopen
from xml.etree.ElementTree import parse
import pandas as pd


class AppBase:

    def __init__(self):
        self.mywin = tk.Tk()
        self.mywin.title("قيمت ارز")
        self.frame1 = tk.Frame(self.mywin)
        self.frame1.pack()

        lb_header = ['تغييرات (درصد)', 'قيمت (ريال)', 'نام ارز', 'شماره']
        var_url = urlopen('http://parsijoo.ir/api?serviceType=price-API&query=Currency')
        xmldoc = parse(var_url)
        lb_list = []
        name_tag = []
        price_tag = []
        i = 0;
        for item in xmldoc.iterfind('sadana-services/price-service/item'):
            name = item.findtext('name')
            price = item.findtext('price')
            change = item.findtext('change')
            percent = item.findtext('percent')
            i = i + 1;
            name_tag.append(name)
            price_tag.append(price)
            lb_list += [(percent + change, price, name, i)]
        data = {'نام ارز': name_tag,
                'قیمت': price_tag}
        df = pd.DataFrame(data, columns=['نام ارز', 'قیمت'])
        df.to_excel('Pricelist.xlsx')
        style = ttk.Style()
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 12, 'bold'))  # Modify the font of the headings
        self.tree = ttk.Treeview(columns=lb_header, show="headings", height=len(lb_list), style='mystyle.Treeview')
        self.tree.column('نام ارز', anchor="center", width=100)
        self.tree.column('قيمت (ريال)', anchor="center", width=100)
        self.tree.column('تغييرات (درصد)', anchor="center", width=100)
        self.tree.column('شماره', anchor="center", width=100)
        self.tree.grid(in_=self.frame1)

        for col in lb_header:
            self.tree.heading(col, text=col.title())
        for item in lb_list:
            self.tree.insert('', 'end', values=item)

    def start(self):
        self.mywin.mainloop()


app = AppBase()
app.start()
