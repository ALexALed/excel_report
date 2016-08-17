import tkinter as tk
from tkinter.filedialog import askopenfilename
import xlrd


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()
        self.source = ''
        self.target = ''

    def create_widgets(self):
        self.choice_source = tk.Button(self)
        self.choice_source["text"] = "Выбор банка"
        self.choice_source["command"] = self.select_source
        self.choice_source.pack(side="top")

        self.choice_target = tk.Button(self)
        self.choice_target["text"] = "Выбор отчета по счетам"
        self.choice_target["command"] = self.select_target
        self.choice_target.pack(side="top")

        self.start_report = tk.Button(self)
        self.start_report["text"] = "Старт отчета"
        self.start_report["command"] = self.run_report
        self.start_report.pack(side="top")

        self.report_text = tk.Text(self, height=200, width=300)
        self.report_text.pack(side="bottom")

        self.quit = tk.Button(self, text="Выход", fg="red",
                              command=root.destroy)
        self.quit.pack(side="bottom")

    def select_source(self):
        self.source = askopenfilename()

    def select_target(self):
        self.target = askopenfilename()

    def run_report(self):
        rb = xlrd.open_workbook(self.source, formatting_info=True)
        sheet = rb.sheet_by_index(0)
        # G
        vals = {rownum: {'order': sheet.cell(rownum, 0).value,
                         'value': sheet.cell(rownum, 6).value} for rownum in range(sheet.nrows)}
        target = {}
        for k, v in vals.items():
            if type(v['order']) == str and "Рахунок на оплату" in v['order']:
                target[k] = v

        rb_target = xlrd.open_workbook(self.target)
        sheet = rb_target.sheet_by_index(0)
        # Задолженность
        vals = {rownum: {'order': sheet.cell(rownum, 0).value,
                         'value': sheet.cell(rownum, 8).value} for rownum in range(sheet.nrows)}

        result = {}
        for k, v in vals.items():
            if type(v['order']) == str and "Рахунок на оплату" in v['order']:
                result[k] = v

        wb = xlrd.open_workbook(self.target, formatting_info=True)

        new_values = {}
        for v in target.values():
            for k, r in result.items():
                if r['order'] == v['order'] and v['value'] and r['value'] != v['value']:
                    new_values[k] = v['value']

        not_equals = []
        for n, v in new_values.items():
            not_equals.append(
                "Для строки " + str(n) + " новое значение в банке " + str(v)
                )

        self.report_text.insert(tk.END, "Несовпадения сумм:\n")
        self.report_text.insert(tk.END, '\n'.join(not_equals))
        result_values_set = set([value['order'] for value in result.values()])
        target_values_set = set([value['order'] for value in target.values()])

        self.report_text.insert(tk.END, "\nЕсть в банке, но нет в отчете:\n")
        self.report_text.insert(tk.END, '\n'.join(
            target_values_set-result_values_set)
        )


root = tk.Tk()
app = Application(master=root)
app.mainloop()
