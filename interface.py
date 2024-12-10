import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkinter import simpledialog
import tkinter.font
import pandas as pd
import os
import sys
#os.chdir(os.path.dirname(sys.executable))
os.chdir(os.path.dirname(os.path.realpath(__file__)))
import Scheduler
from datetime import datetime, timedelta, time


class AddDataDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("데이터 추가")
        
        self.entry_widgets = {}
        for column in self.parent.dataframe.columns:
            frame = tk.Frame(self)
            frame.pack(fill=tk.X, padx=5, pady=5)
            label = tk.Label(frame, text=column)
            label.pack(side=tk.LEFT)
            entry = tk.Entry(frame)
            entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
            self.entry_widgets[column] = entry
        
        submit_button = tk.Button(self, text="확인", command=self.submit)
        submit_button.pack(pady=5)
    
    def submit(self):
        data = {col: entry.get() for col, entry in self.entry_widgets.items()}
        self.parent.add_row(data)
        self.destroy()

class DateSelectionDialog(simpledialog.Dialog):
    def body(self, master):
        self.title("날짜 및 시간 선택")

        today = datetime.now()  
        current_year = today.year
        current_month = today.strftime("%m")
        current_day = today.strftime("%d")
        default_hour = "08"
        default_minute = "30"

        style = ttk.Style()
        style.configure('TLabel', font=('Helvetica', 12))
        style.configure('TCombobox', font=('Helvetica', 12))

        ttk.Label(master, text="년:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.year_var = tk.StringVar(master, value=current_year)
        self.year_combobox = ttk.Combobox(master, textvariable=self.year_var, values=[str(year) for year in range(current_year - 5, current_year + 20)], width=5, style='TCombobox')
        self.year_combobox.grid(row=0, column=1, padx=5, pady=5, sticky='w')

        ttk.Label(master, text="월:", style='TLabel').grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.month_var = tk.StringVar(master, value=current_month)
        self.month_combobox = ttk.Combobox(master, textvariable=self.month_var, values=[str(i).zfill(2) for i in range(1, 13)], width=5, style='TCombobox')
        self.month_combobox.grid(row=1, column=1, padx=5, pady=5, sticky='w')

        ttk.Label(master, text="일:", style='TLabel').grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.day_var = tk.StringVar(master, value=current_day)
        self.day_combobox = ttk.Combobox(master, textvariable=self.day_var, values=[str(i).zfill(2) for i in range(1, 32)], width=5, style='TCombobox')
        self.day_combobox.grid(row=2, column=1, padx=5, pady=5, sticky='w')

        ttk.Label(master, text="시:", style='TLabel').grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.hour_var = tk.StringVar(master, value=default_hour)
        self.hour_combobox = ttk.Combobox(master, textvariable=self.hour_var, values=[str(i).zfill(2) for i in range(24)], width=5, style='TCombobox')
        self.hour_combobox.grid(row=3, column=1, padx=5, pady=5, sticky='w')

        ttk.Label(master, text="분:", style='TLabel').grid(row=4, column=0, padx=5, pady=5, sticky='e')
        self.minute_var = tk.StringVar(master, value=default_minute)
        self.minute_combobox = ttk.Combobox(master, textvariable=self.minute_var, values=[str(i).zfill(2) for i in range(60)], width=5, style='TCombobox')
        self.minute_combobox.grid(row=4, column=1, padx=5, pady=5, sticky='w')

        return self.year_combobox

    def apply(self):
        year = self.year_var.get()
        month = self.month_var.get()
        day = self.day_var.get()
        hour = self.hour_var.get()
        minute = self.minute_var.get()
        self.result = [f"{year}",f"{month}-{day}",f"{hour}:{minute}"]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("스케줄러")
        self.geometry("1400x850")
        self.frames = {}

        
        self.schedule = Scheduler.productionScheduler()
        
        Frame_Page = (StartPage, PageOne, PageThree)
        for Page in Frame_Page:
            frame = Page(self)
            self.frames[Page] = frame
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.show_frame(StartPage)
        
        self.file_path= {}
        
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()
        
    def quit_app(self):
        self.destroy()
    
    def inicial(self):
        self.schedule = Scheduler.productionScheduler()
    
    def load_file_path(self, file_name):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsm")])
        if path:
            self.file_path[file_name] = [path, '목록']
    
    def load_data(self, p_filepath, d_filepath):
        try:
            self.schedule.load_data(p_filepath, d_filepath)
        except Exception as e:
            self.show_error(f"데이터 로딩 중 오류 발생: {e}")
    
    def data_structures(self):
        try:
            self.schedule.data_structures()
        except Exception as e:
            self.show_error(f"데이터 구조 설정 중 오류 발생: {e}")
    
    def setup_constraints(self):
        try:
            self.schedule.setup_constraints()
        except Exception as e:
            self.show_error(f"제약 조건 설정 중 오류 발생: {e}")
    
    
    def solve(self):
        dialog = DateSelectionDialog(self)
        scenario_number = simpledialog.askfloat("입력", "작동시간을 입력하세요(단위 : 분)", parent=self)
        self.schedule.running_time = scenario_number*60
        if dialog.result:  
            self.schedule.senario = dialog.result  
            try:
                self.schedule.solve()
            except Exception as e:
                self.show_error(f"문제 해결 중 오류 발생: {e}")
            try:
                self.schedule.output_results()
                self.schedule.rest_combine()
            except Exception as e:  
                self.show_error(f"결과 처리 중 오류 발생: {e}")

    def output_excel(self, Working_order):
        try:
            self.schedule.output_excel(Working_order)
        except Exception as e:  
            self.show_error(f"액샐 저장 중 오류 발생: {e}")
    
    def gantt_chart(self,Working_order):
        try:
            self.schedule.gantt_chart(Working_order)
        except Exception as e:  
            self.show_error(f"간트차트 생성 중 오류 발생: {e}")
    
    def show_error(self, error_message):
        messagebox.showerror("오류 발생", error_message)
    
class StartPage(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        # 페이지 배경색 지정
        self.configure(bg = '#EFF5FB')
        
        #labe_font = ("에스코어 드림 9 Black", 40)
        labe_font = ('현대하모니 L', 40, 'bold')
        
        
        label = tk.Label(self, text="스케줄러",font=labe_font, borderwidth=0, relief="ridge", fg = '#9CCDDB')
        label.place(x=40, y=20)
        label.configure(bg = "#064469")
        
        current_path = os.path.dirname(os.path.realpath(__file__))
        
        #button_font = ("나눔스퀘어_ac ExtraBold", 30)  
        button_font = ('현대하모니 L', 30, 'bold')
        button_width = 15
        button_height = 1
        
        button_borderwidth = 2
        button_relief = "ridge"
        button_bd = 5
        button_fg = 'white'
        button_bg = '#9CCDDB'

        button1 = tk.Button(self, text="스케쥴링 하기", command=lambda: parent.show_frame(PageOne), font=button_font, width=button_width, height=button_height, anchor="center", padx=10, pady=20, borderwidth=button_borderwidth, relief=button_relief, bd = button_bd, bg = button_bg, fg = button_fg)
        button1.place(x=50, y=150)

        
        button3 = tk.Button(self, text="생산정보 조회", command=lambda: parent.show_frame(PageThree), font=button_font, width=button_width, height=button_height, anchor="center", padx=10, pady=20, borderwidth=button_borderwidth, relief=button_relief, bd = button_bd, bg = button_bg, fg = button_fg)
        button3.place(x=50, y=300)
        
        end_button = tk.Button(self, text="종료", command=lambda: parent.quit_app(), font=button_font, width=button_width, height=button_height, anchor="center", padx=10, pady=20, borderwidth=button_borderwidth, relief=button_relief, bd = button_bd, bg = button_bg, fg = button_fg)
        end_button.place(x=50, y=450)

class PageOne(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)

        style = ttk.Style()
        style.configure('TNotebook.Tab', font=('', 15))  
        
        button_style = ttk.Style()
        button_style.configure('Large.TButton', font=('', 15))

        notebook = ttk.Notebook(self)

        tab1 = self.PageOne_TabOne(notebook, parent)
        tab2 = self.PageOne_TabTwo(notebook, parent)
        tab3 = self.PageOne_TabThree(notebook, parent)
        
        notebook.add(tab1, text='스케쥴러 실행')
        notebook.add(tab2, text='생산수량 확인')
        notebook.add(tab3, text='생산제품정보 확인')

        notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        button_font = ("나눔스퀘어_ac ExtraBold", 20)  
        button_width = 15  
        button_height = 1  
        button_borderwidth = 0
        button_relief = "solid"
        
        main_button = tk.Button(self, text="메인화면", command=lambda: parent.show_frame(StartPage),
                                        font=button_font, width=button_width, height=button_height,
                                        anchor="center", padx=10, pady=20, 
                                        borderwidth=button_borderwidth, relief=button_relief)
        main_button.place(x=1100, y=720)
        main_button.configure(bg = "#73C6D9")

    class PageOne_TabOne(tk.Frame):
        def __init__(self, notebook, parent):
            tk.Frame.__init__(self, notebook)
            
            button_font = ("나눔스퀘어_ac ExtraBold", 20)  
            button_width = 15  
            button_height = 1  
            button_borderwidth = 0
            button_relief = "solid"
            
            y_place = 20
            
            inicializimi_button = tk.Button(self, text="초기화", command=lambda: parent.inicial(),
                                            font=button_font, width=button_width, height=button_height,
                                            anchor="center", padx=10, pady=20, 
                                            borderwidth=0, relief=button_relief)
            inicializimi_button.place(x=20, y=675)
            inicializimi_button.configure(bg = "#73C6D9")
            
            

            read_data_button = tk.Button(self, text="제품정보\n파일경로 지정", command=lambda: parent.load_file_path('p_file_path'),
                                         font=button_font, width=button_width, height=button_height,
                                         anchor="center", padx=10, pady=20, 
                                         borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            read_data_button.place(x=20, y=10+y_place)
            read_data_button.configure(bg = "#072D44")
            

            read_data_button2 = tk.Button(self, text="생산정보\n파일경로 지정", command=lambda: parent.load_file_path("d_file_path"),
                                          font=button_font, width=button_width, height=button_height,
                                          anchor="center", padx=10, pady=20, 
                                          borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            read_data_button2.place(x=300, y=10+y_place)
            read_data_button2.configure(bg = "#072D44")
            
            load_data_button = tk.Button(self, text="데이터 불러오기",
                                         command=lambda: parent.load_data(parent.file_path['p_file_path'], parent.file_path['d_file_path']),
                                         font=button_font, width=button_width, height=button_height,
                                         anchor="center", padx=10, pady=20, 
                                         borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            load_data_button.place(x=20, y=150+y_place)
            load_data_button.configure(bg = "#064469")
            
            data_structure_button = tk.Button(self, text="데이터 구조화", command=lambda: parent.data_structures(),
                                              font=button_font, width=button_width, height=button_height,
                                              anchor="center", padx=10, pady=20, 
                                              borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            data_structure_button.place(x=300, y=150+y_place)
            data_structure_button.configure(bg = "#064469")
            
            
            setup_constraints_button = tk.Button(self, text="제약식 설정", command=lambda: parent.setup_constraints(),
                                                 font=button_font, width=button_width, height=button_height,
                                                 anchor="center", padx=10, pady=20, 
                                                 borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            setup_constraints_button.place(x=20, y=290+y_place)
            setup_constraints_button.configure(bg = "#5790AB")
            
            run_scheduler_button = tk.Button(self, text="스케쥴 시작", command=lambda: parent.solve(),
                                             font=button_font, width=button_width, height=button_height,
                                             anchor="center", padx=10, pady=20, 
                                             borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            run_scheduler_button.place(x=300, y=290+y_place)
            run_scheduler_button.configure(bg = "#5790AB")
            
            
            out_excel_button1 = tk.Button(self, text="엑셀불러오기", command=lambda: parent.output_excel(parent.schedule.output_excel_Working_Order),
                                          font=button_font, width=button_width, height=button_height,
                                          anchor="center", padx=10, pady=20,
                                          borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            out_excel_button1.place(x=20, y=430+y_place)
            out_excel_button1.configure(bg = "#9CCDDB")
            
            
            out_excel_button2 = tk.Button(self, text="엑셀불러오기\n(휴식시간 추가)", command=lambda: parent.output_excel(parent.schedule.output_excel_rest_Working_Order),
                                          font=button_font, width=button_width, height=button_height,
                                          anchor="center", padx=10, pady=20,
                                          borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            out_excel_button2.place(x=300, y=430+y_place)
            out_excel_button2.configure(bg = "#9CCDDB")
            
            out_ganttchart_button1 = tk.Button(self, text="간트차트 불러오기", command=lambda: parent.gantt_chart(parent.schedule.Working_order),
                                               font=button_font, width=button_width, height=button_height,
                                               anchor="center", padx=10, pady=20,
                                               borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            out_ganttchart_button1.place(x=580, y=430+y_place)
            out_ganttchart_button1.configure(bg = "#9CCDDB")
            
            out_ganttchart_button2 = tk.Button(self, text="간트차트 불러오기\n(휴식시간 추가)", command=lambda: parent.gantt_chart(parent.schedule.rest_Working_order),
                                               font=button_font, width=button_width, height=button_height,
                                               anchor="center", padx=10, pady=20,
                                               borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            out_ganttchart_button2.place(x=860, y=430+y_place)
            out_ganttchart_button2.configure(bg = "#9CCDDB")
            

    class PageOne_TabTwo(tk.Frame):
        def __init__(self, notebook, parent):
            tk.Frame.__init__(self, notebook)
            self.parent = parent  
    
            self.table = None
    
            self.display_dataframe()
    
            button_font = ("나눔스퀘어_ac ExtraBold", 20)  
            button_width = 15  
            button_height = 1  
            button_borderwidth = 0
            button_relief = "solid"
            
            delete_row_button = tk.Button(self, text="선택한 행 제거", command=self.delete_row,
                                          font=button_font, width=button_width, height=button_height,
                                          anchor="center", padx=10, pady=20,
                                          borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            delete_row_button.place(x=50, y=30)  
            delete_row_button.configure(bg = "#064469")
            
            add_data_button = tk.Button(self, text="데이터 추가", command=self.open_add_data_dialog,
                                        font=button_font, width=button_width, height=button_height,
                                        anchor="center", padx=10, pady=20,
                                        borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            add_data_button.place(x=550, y=30)  
            add_data_button.configure(bg = "#064469")
            
            refresh_data_button = tk.Button(self, text="데이터 갱신", command=self.refresh_dataframe,
                                            font=button_font, width=button_width, height=button_height,
                                            anchor="center", padx=10, pady=20,
                                            borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            refresh_data_button.place(x=1040, y=30)  
            refresh_data_button.configure(bg = "#064469")
        
        def display_dataframe(self):
            self.dataframe = self.parent.schedule.all_production
        
            if self.table:
                self.table.destroy()
        
            self.table = ttk.Treeview(self)
            self.table["columns"] = list(self.dataframe.columns)
            self.table["show"] = "headings"
        
            for column in self.table["columns"]:
                self.table.heading(column, text=column)
                self.table.column(column, width=100)
        
            for index, row in self.dataframe.iterrows():
                self.table.insert("", "end", iid=str(index), values=list(row))
        
        
            self.table.place(x=30, y=140, width=1300, height=520)
        
        def delete_row(self):
            selected_items = self.table.selection()
            for item_id in selected_items:
                index_to_drop = int(item_id)
                self.dataframe = self.dataframe.drop(index_to_drop)
                self.table.delete(item_id)
        
            self.dataframe.reset_index(drop=True, inplace=True)
            self.parent.schedule.production = self.dataframe
        
            self.display_dataframe()
    
        def refresh_dataframe(self):
            self.display_dataframe()
    
        def open_add_data_dialog(self):
            newWindow = AddDataDialog(self)
            self.wait_window(newWindow)  
    
        def add_row(self, data):
            new_row_df = pd.DataFrame([data])
            self.dataframe = pd.concat([self.dataframe, new_row_df], ignore_index=True)
            self.parent.schedule.production = self.dataframe
            self.display_dataframe()
        
    class PageOne_TabThree(tk.Frame):
        def __init__(self, notebook, parent):
            tk.Frame.__init__(self, notebook)
            self.parent = parent  
    
            self.table = None
    
            self.display_dataframe()
            
            
            button_font = ("나눔스퀘어_ac ExtraBold", 20)  
            button_width = 15  
            button_height = 1  
            button_borderwidth = 0
            button_relief = "solid"
            
            refresh_data_button = tk.Button(self, text="데이터 갱신", command=self.refresh_dataframe,
                                            font=button_font, width=button_width, height=button_height,
                                            anchor="center", padx=10, pady=20,
                                            borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
            refresh_data_button.place(x=1040, y=30)
            refresh_data_button.configure(bg = "#064469")
            
        def display_dataframe(self):
            self.dataframe = self.parent.schedule.product_production_information
        
            if self.table:
                self.table.destroy()
        
            self.table = ttk.Treeview(self)
            self.table["columns"] = list(self.dataframe.columns)
            self.table["show"] = "headings"
        
            for column in self.table["columns"]:
                self.table.heading(column, text=column)
                self.table.column(column, width=100)
        
            for index, row in self.dataframe.iterrows():
                self.table.insert("", "end", iid=str(index), values=list(row))
        
        
            self.table.place(x=30, y=140, width=1300, height=520)
    
        def refresh_dataframe(self):
            self.display_dataframe()
            

class PageThree(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent  

        self.selected_key = tk.StringVar()
        self.keys_dropdown = ttk.Combobox(self, textvariable=self.selected_key, 
                                          values=list(self.parent.schedule.Working_order.keys()), 
                                          state="readonly")
        self.keys_dropdown.place(x=630, y=10)
        self.keys_dropdown.bind("<<ComboboxSelected>>", self.update_display)

        self.table = None

        if self.parent.schedule.Working_order.keys():
            first_key = list(self.parent.schedule.Working_order.keys())[0]
            self.selected_key.set(first_key)
            self.update_display(first_key)  
            
        
        button_font = ("나눔스퀘어_ac ExtraBold", 20)  
        button_width = 15  
        button_height = 1  
        button_borderwidth = 0
        button_relief = "solid"
        
        update_button = tk.Button(self, text="데이터 갱신", command=self.update_keys_dropdown,
                                        font=button_font, width=button_width, height=button_height,
                                        anchor="center", padx=10, pady=20,
                                        borderwidth=button_borderwidth, relief=button_relief, fg = 'white')
        update_button.place(x=580, y=40)  
        update_button.configure(bg = "#064469")
        
        main_button = tk.Button(self, text="메인화면", command=lambda: parent.show_frame(StartPage),
                                        font=button_font, width=button_width, height=button_height,
                                        anchor="center", padx=10, pady=20,
                                        borderwidth=button_borderwidth, relief=button_relief)
        main_button.place(x=1100, y=750)  
        main_button.configure(bg = "#73C6D9")
        

    def update_display(self, event=None):
        key = self.selected_key.get()
        
        if self.table:
            self.table.destroy()
        
        self.table = ttk.Treeview(self)
        self.table["columns"] = list(self.parent.schedule.Working_order[key].columns)
        self.table["show"] = "headings"
        
        for column in self.table["columns"]:
            self.table.heading(column, text=column)
            self.table.column(column, width=100)
            
        for index, row in self.parent.schedule.Working_order[key].iterrows():
            self.table.insert("", "end", values=list(row))
        
        self.table.place(x=30, y=140, width=1350, height=600)
        
    def update_keys_dropdown(self):
        new_values = list(self.parent.schedule.Working_order.keys())
        self.keys_dropdown['values'] = new_values
        
        if self.selected_key.get() not in new_values and new_values:
            self.selected_key.set(new_values[0])
            self.update_display()
    
    
app = App()
app.mainloop()









