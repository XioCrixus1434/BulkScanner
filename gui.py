import os
import csv
import tkinter
from tkinter import *
from tkinter import filedialog
import PyPDF2
import pdfplumber

class GUI(tkinter.Tk):

    def __init__(self):
        super().__init__()
        self.title("BulkScan")
        self.geometry("800x400+100+50")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.path = None
        self.chosen_dir = None
        self.rules = []

        self.fieldnames = []
        self.files = []
        self.opfiles = []
        self.sorted = []
        self.rows = []

        self.blacklist_entry = None
        self.blacklist = None
        self.blacklist_button = None
        self.blacklist_delete = None
        self.blacklist_checks = []

        self.select_label = Label(text="Select the file location", font=("Segoe UI", 20))
        self.select_label.place(x=110, y=100, w=600, h=75)

        self.browse_button = Button(text="Browse", font=("Segoe UI", 10))
        self.browse_button.bind("<Button-1>", self.open_file_input)
        self.browse_button.place(x=362.5, y=175, w=75, h=25)

        self.dir = Label(text="", font=("Segoe UI", 10), wraplength=300)
        self.dir.place(x=105, y=225, w=600, h=25)

        self.select_button = Button(text="Select", font=("Segoe UI", 20), command=self.on_select)
        self.select_button["state"] = "disabled"
        self.select_button.place(x=349, y=300, w=100, h=50)

        self.rule_list = None
        self.plus = None
        self.munis = None
        self.temp = None
        self.cat = None
        self.fil = None
        self.sort_button = None

    def open_file_input(self, event):
        tkinter.Tk().withdraw()
        folder_path = filedialog.askdirectory()
        if folder_path != "":
            self.path = folder_path
            self.chosen_dir = self.get_scope()
            self.select_button["state"] = "normal"
            self.dir.config(text=folder_path)

    def get_scope(self):
        path = self.path
        dir = ""
        while dir == "":
            try:
                path = path[path.index("/")+1:]
            except ValueError:
                dir = path
        return dir

    def on_select(self):
        self.select_button.place_forget()
        self.dir.place_forget()
        self.browse_button.place_forget()
        self.select_label.place_forget()

        self.rule_list = Listbox()
        self.rule_list.place(x=550, y=100, w=200, h=200)

        self.blacklist = Listbox()
        self.blacklist.place(x=50, y=100, w=150, h=100)

        self.blacklist_entry = Entry()
        self.blacklist_entry.place(x=50, y=215, w=150, h=25)

        self.blacklist_button = Button(text="Add", font=("Segoe UI", 10), command=self.on_bad)
        self.blacklist_button.place(x=50, y=250, w=75, h=25)

        self.blacklist_button = Button(text="Delete", font=("Segoe UI", 10), command=self.on_del)
        self.blacklist_button.place(x=125, y=250, w=75, h=25)

        self.plus = Button(text="+", font=("Segoe UI", 10), command=self.add)
        self.plus.place(x=525, y=100, w=25, h=25)

        self.minus = Button(text="-", font=("Segoe UI", 10), command=self.on_subtraction)
        self.minus.place(x=525, y=125, w=25, h=25)

        self.sort_button = Button(text="Sort", font=("Segoe UI", 20), command=self.on_sort)
        self.sort_button["state"] = "disabled"
        self.sort_button.place(x=349, y=300, w=100, h=50)

    def on_bad(self):
        blacklisting = self.blacklist_entry.get()
        self.blacklist_entry.delete(0, END)
        self.blacklist_checks.append(blacklisting)
        self.blacklist.insert(self.blacklist_checks.index(blacklisting), blacklisting)

    def on_del(self):
        i = self.blacklist.curselection()
        selection = self.blacklist.get(i)
        self.blacklist_checks.remove(selection)
        self.blacklist.delete(i)

    def add(self):
        self.plus["state"] = "disabled"
        self.minus["state"] = "disabled"
        self.temp = tkinter.Toplevel()
        self.temp.title("Add Rule")
        self.temp.geometry("400x250+200+100")
        self.temp.resizable(False, False)

        self.cat = Entry(self.temp)
        self.cat.place(x=175, y=75, w=150, h=25)

        catlab = Label(self.temp, text="Category:", font=("Segoe UI", 10))
        catlab.place(x=100, y=73, w=55, h=25)

        self.fil = Entry(self.temp)
        self.fil.place(x=175, y=125, w=150, h=25)

        fillab = Label(self.temp, text="Filter:", font=("Segoe UI", 10))
        fillab.place(x=100, y=123, w=55, h=25)

        ok_button = Button(self.temp, text="Ok", font=("Segoe UI", 10), command=self.on_addition)
        ok_button.place(x=132.5, y=200, w=50, h=25)

    def temp_close(self):
        self.temp.destroy()

    def on_addition(self):
        category = self.cat.get()
        filter = self.fil.get()
        if category != "" and filter != "":
            rule = self.cat.get() + " - " + self.fil.get()
            self.rules.append(rule)
            self.rule_list.insert(self.rules.index(rule), rule)
            self.temp_close()
            self.plus["state"] = "normal"
            self.minus["state"] = "normal"
            self.sort_button["state"] = "normal"

    def on_subtraction(self):
        i = self.rule_list.curselection()
        selection = self.rule_list.get(i)
        self.rules.remove(selection)
        self.rule_list.delete(i)
        if len(self.rules) == 0:
            self.sort_button["state"] = "disabled"

    def on_sort(self):
        i = 0
        while i < len(self.rules):
            rule_name = self.rules[i]
            end = self.rules[i].index("-")-1
            self.fieldnames.append(rule_name[0:end])
            self.sorted.append([])
            i += 1
        self.fieldnames.append("Unknown")
        self.sorted.append([])
        self.sort("")
        self.create_export()

    def sort(self, ext):
        dir = self.path + ext
        try:
            file_list = os.listdir(dir)
        except PermissionError:
            file_list = []
        except NotADirectoryError:
            file_list = []
        except FileNotFoundError:
            file_list = []
        dir += "/"
        i = 0
        while i < len(file_list):
            blacklisted = False
            file = file_list[i]
            for blacklisting in self.blacklist_checks:
                if blacklisting in file:
                    blacklisted = True
            if not blacklisted:
                try:
                    report = self.scan_in_pdf(dir, file)
                except TypeError:
                    report = None
                except PermissionError:
                    print(dir + file)
                    report = None
                except UnicodeDecodeError:
                    report = None
                except Exception:
                    report = None
                if report is None:
                    self.sort("/" + file)
                elif report is not None:
                    self.files.append(file)
                    self.opfiles.append(report)
            i += 1


    def create_export(self):
        user = os.environ.get("USERNAME")
        exp_path = "C:/Users/" + user + "/Desktop"

        i = 0
        while i < len(self.opfiles):
            j = 0
            found = False
            while j < len(self.rules):
                filter = self.rules[j][self.rules[j].index("-")+2:]

                if filter.lower() in self.opfiles[i].lower():
                    found = True
                    self.sorted[j].append(self.files[i])
                j += 1
            if not found:
                self.sorted[len(self.sorted) - 1].append(self.files[i])
            i += 1

        i = 1
        j = 0
        while i < len(self.sorted):
            if len(self.sorted[i]) > len(self.sorted[j]):
                j = i
            i += 1

        i = 0
        while i < len(self.sorted[j]):
            row = []
            k = 0
            while k < len(self.sorted):
                try:
                    row.append(self.sorted[k][i])
                except IndexError:
                    row.append("")
                k += 1
            self.rows.append(row)
            i += 1

        with open(exp_path + "/output", "w") as output:
            writer = csv.writer(output, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

            writer.writerow(self.fieldnames)
            for row in self.rows:
                writer.writerow(row)





    def scan_in_pdf(self, dir, file):
        r = ""
        pdf = open(dir + file, "rb")
        pdf_reader = PyPDF2.PdfFileReader(pdf)
        pages = len(pdf_reader.pages)
        pdf.close()
        i = 0
        with pdfplumber.open(dir + file) as temp:
            while i < pages:
                page = temp.pages[i]
                r += page.extract_text()
                i += 1
        temp.close()
        return r

    def on_close(self):
        self.quit()


root = GUI()
root.mainloop()
