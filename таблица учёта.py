import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import copy
import os
import sys

class ScrollableTable(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.header_frame = tk.Frame(self)
        self.header_frame.pack(side=tk.TOP, fill=tk.X)
        
        self.h_scroll_header = tk.Scrollbar(self.header_frame, orient=tk.HORIZONTAL)
        self.h_scroll_header.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.header_canvas = tk.Canvas(self.header_frame, height=50,
                                      xscrollcommand=self.h_scroll_header.set)
        self.header_canvas.pack(side=tk.TOP, fill=tk.X, expand=True)
        self.h_scroll_header.config(command=self.header_canvas.xview)
        
        self.header_inner = tk.Frame(self.header_canvas)
        self.header_window = self.header_canvas.create_window((0,0), window=self.header_inner, anchor='nw')
        self.header_inner.bind("<Configure>", lambda e: self.header_canvas.configure(scrollregion=self.header_canvas.bbox("all")))
        
        self.table_frame = tk.Frame(self)
        self.table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        self.v_scroll = tk.Scrollbar(self.table_frame, orient=tk.VERTICAL)
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.h_scroll = tk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.table_canvas = tk.Canvas(self.table_frame,
                                     yscrollcommand=self.v_scroll.set,
                                     xscrollcommand=self.h_scroll.set)
        self.table_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.v_scroll.config(command=self.table_canvas.yview)
        self.h_scroll.config(command=self.table_canvas.xview)
        
        self.table_inner = tk.Frame(self.table_canvas)
        self.table_window = self.table_canvas.create_window((0,0), window=self.table_inner, anchor='nw')
        
        self.table_inner.bind("<Configure>", self.on_table_configure)
        self.table_canvas.bind("<Configure>", self.on_canvas_configure)
        
        self.header_canvas.config(xscrollcommand=self.sync_scroll_x)
        self.table_canvas.config(xscrollcommand=self.sync_scroll_x)
        self.h_scroll.config(command=self.sync_scrollbar)
        
        self._scroll_x = 0
    
    def sync_scrollbar(self, *args):
        self.header_canvas.xview(*args)
        self.table_canvas.xview(*args)
        self.h_scroll.set(*args)
        self._scroll_x = self.header_canvas.xview()[0]
    
    def sync_scroll_x(self, *args):
        self.header_canvas.xview_moveto(args[0])
        self.table_canvas.xview_moveto(args[0])
        self.h_scroll.set(*args)
        self._scroll_x = float(args[0])
    
    def on_table_configure(self, event):
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))
        canvas_width = self.table_canvas.winfo_width()
        self.table_canvas.itemconfig(self.table_window, width=canvas_width)
    
    def on_canvas_configure(self, event):
        self.table_canvas.itemconfig(self.table_window, width=event.width)
        self.header_canvas.itemconfig(self.header_window, width=event.width)

class JournalApp(tk.Tk):
    def __init__(self, init_file=None):
        super().__init__()
        self.title("Таблица учёта продукции на уголковой линии")
        self.geometry("450x220")

        # Попытка загрузить иконку из упакованного exe (сделано для Windows + pyinstaller)
        try:
            self.iconbitmap(sys.executable)
        except Exception:
            pass
        
        self.data_rows = 1
        self.columns = []
        
        self.undo_stack = []
        self.redo_stack = []
        self.max_undo = 50
        
        self.current_file = None
        self.unsaved_changes = False
        
        # Меню
        menubar = tk.Menu(self)
        
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Сохранить как...", command=self.save_to_file)
        filemenu.add_command(label="Открыть файл", command=self.load_from_file)
        menubar.add_cascade(label="Файл", menu=filemenu)
        
        editmenu = tk.Menu(menubar, tearoff=0)
        editmenu.add_command(label="Отменить", command=self.undo_action)
        editmenu.add_command(label="Вернуть", command=self.redo_action)
        menubar.add_cascade(label="Правка", menu=editmenu)
        
        name_menu = tk.Menu(menubar, tearoff=0)
        name_menu.add_command(label="Импортировать из директории", command=self.import_columns_from_directory)
        menubar.add_cascade(label="Маркировки", menu=name_menu)
        
        # Новое меню "О программе"
        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label="Сведения", command=self.show_about_info)
        menubar.add_cascade(label="О программе", menu=about_menu)
        
        self.config(menu=menubar)
        
        top_frame = ttk.Frame(self)
        top_frame.pack(side=tk.TOP, fill=tk.X)
        
        self.add_col_btn = ttk.Button(top_frame, text="Добавить столбец", command=self.add_column)
        self.add_col_btn.pack(side=tk.LEFT, padx=5, pady=6)
        
        self.quick_save_btn = ttk.Button(top_frame, text="Быстрое сохранение", command=self.quick_save)
        self.quick_save_btn.pack(side=tk.LEFT, padx=5, pady=6)
        
        self.scrollable_table = ScrollableTable(self)
        self.scrollable_table.pack(fill=tk.BOTH, expand=True)
        
        self.del_row_buttons = []
        self.add_row_btn = None
        
        for _ in range(3):
            self.add_column(init=True)
        self.refresh_delete_row_buttons()
        
        self.bind_all("<Control-z>", lambda e: self.undo_action())
        self.bind_all("<Control-y>", lambda e: self.redo_action())
        
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        if init_file:
            if os.path.isfile(init_file):
                try:
                    with open(init_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    self.restore_state(data)
                    self.undo_stack.clear()
                    self.redo_stack.clear()
                    self.save_state()
                    self.current_file = init_file
                    self.unsaved_changes = False
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось загрузить файл при запуске:\n{e}")
        
        self.save_state()
    
    def show_about_info(self):
        about_text = (
            "Имя программы: Таблица учёта продукции\n"
            "Автор: Павел Чуликов\n"
            "Версия: 3.0.22\n"
            "Дата выхода версии: 10.08.2025\n"
            "Дата выхода первой версии: 10.07.2025\n"
            "Контакты: P30534437@gmail.com\n"
        )
        messagebox.showinfo("Сведения о программе", about_text)
    
    def mark_changes(self, event=None):
        self.unsaved_changes = True
    
    def add_column(self, init=False):
        col_index = len(self.columns)
        col_entries = []
        
        del_col_btn = ttk.Button(self.scrollable_table.header_inner, text="Удалить", width=12,
                                 command=lambda c=col_index: self.delete_column(c))
        del_col_btn.grid(row=0, column=col_index, padx=3, pady=3)
        
        header = ttk.Entry(self.scrollable_table.header_inner, justify='center', width=13)
        header.insert(0, f"Столбец {col_index+1}")
        header.grid(row=1, column=col_index, padx=3, pady=3)
        header.bind("<KeyRelease>", lambda e: [self.mark_changes(), self.save_state()])
        self._bind_ctrl_v(header)
        
        col_entries.append(del_col_btn)
        col_entries.append(header)
        
        for row in range(self.data_rows):
            entry = ttk.Entry(self.scrollable_table.table_inner, justify='center', width=13)
            entry.grid(row=row, column=col_index, padx=3, pady=1)
            entry.bind("<KeyRelease>", lambda e, c=col_index: [self.update_sum(c), self.mark_changes(), self.save_state()])
            self._bind_ctrl_v(entry)
            col_entries.append(entry)
        
        sum_entry = ttk.Entry(self.scrollable_table.table_inner, justify='center', width=13, state='readonly')
        sum_entry.grid(row=self.data_rows, column=col_index, padx=3, pady=1)
        col_entries.append(sum_entry)
        
        self.columns.append(col_entries)
        
        if not init:
            self.refresh_delete_row_buttons()
            self.update_all_sums()
            self.save_state()
    
    def add_row(self):
        self.data_rows += 1
        row_index = self.data_rows - 1
        
        for col_index, col_entries in enumerate(self.columns):
            sum_entry = col_entries[-1]
            sum_entry.grid_forget()
            sum_entry.grid(row=self.data_rows, column=col_index, padx=3, pady=1)
            
            new_entry = ttk.Entry(self.scrollable_table.table_inner, justify='center', width=13)
            new_entry.grid(row=row_index, column=col_index, padx=3, pady=1)
            new_entry.bind("<KeyRelease>", lambda e, c=col_index: [self.update_sum(c), self.mark_changes(), self.save_state()])
            self._bind_ctrl_v(new_entry)
            col_entries.insert(-1, new_entry)
        
        self.refresh_delete_row_buttons()
        self.update_all_sums()
        self.save_state()
    
    def delete_column(self, col_index):
        if len(self.columns) <= 1:
            messagebox.showwarning("Предупреждение", "Нельзя удалить последний столбец!")
            return
        
        col_entries = self.columns.pop(col_index)
        for widget in col_entries:
            widget.grid_forget()
            widget.destroy()
        
        for c in range(col_index, len(self.columns)):
            col_entries = self.columns[c]
            col_entries[0].config(command=lambda c=c: self.delete_column(c))
            for r, widget in enumerate(col_entries):
                widget.grid_configure(column=c)
        
        self.refresh_delete_row_buttons()
        self.update_all_sums()
        self.mark_changes()
        self.save_state()
    
    def delete_row(self, row_index):
        if self.data_rows <= 1:
            messagebox.showwarning("Предупреждение", "Нельзя удалить последнюю строку!")
            return
        
        btn_to_remove = self.del_row_buttons.pop(row_index)
        btn_to_remove.grid_forget()
        btn_to_remove.destroy()
        
        for col_entries in self.columns:
            entry_to_remove = col_entries[2 + row_index]
            entry_to_remove.grid_forget()
            entry_to_remove.destroy()
            col_entries.pop(2 + row_index)
        
        self.data_rows -= 1
        
        for col_index, col_entries in enumerate(self.columns):
            for i in range(self.data_rows):
                col_entries[2 + i].grid_configure(row=i)
            col_entries[-1].grid_configure(row=self.data_rows)
        
        col_for_buttons = len(self.columns)
        for i, btn in enumerate(self.del_row_buttons):
            btn.config(command=lambda r=i: self.delete_row(r))
            btn.grid(row=i, column=col_for_buttons, padx=3, pady=1)
        
        if self.add_row_btn:
            self.add_row_btn.grid_forget()
            self.add_row_btn.grid(row=self.data_rows, column=col_for_buttons, padx=3, pady=6)
        
        self.update_all_sums()
        self.mark_changes()
        self.save_state()
    
    def refresh_delete_row_buttons(self):
        for btn in getattr(self, 'del_row_buttons', []):
            btn.grid_forget()
            btn.destroy()
        if self.add_row_btn:
            self.add_row_btn.grid_forget()
            self.add_row_btn.destroy()
            self.add_row_btn = None
        
        self.del_row_buttons = []
        
        col_for_buttons = len(self.columns)
        
        for row in range(self.data_rows):
            btn = ttk.Button(self.scrollable_table.table_inner, text="Удалить строку", width=16)
            btn.grid(row=row, column=col_for_buttons, padx=3, pady=1)
            btn.config(command=lambda r=row: self.delete_row(r))
            self.del_row_buttons.append(btn)
        
        self.add_row_btn = ttk.Button(self.scrollable_table.table_inner, text="Добавить строку", width=16, command=self.add_row)
        self.add_row_btn.grid(row=self.data_rows, column=col_for_buttons, padx=3, pady=6)
    
    def update_sum(self, col_index):
        col_entries = self.columns[col_index]
        total = 0.0
        for entry in col_entries[2:-1]:
            val = entry.get().strip()
            if val:
                try:
                    total += float(val)
                except ValueError:
                    pass
        sum_entry = col_entries[-1]
        sum_entry.config(state='normal')
        sum_entry.delete(0, tk.END)
        sum_entry.insert(0, f"{total:.2f}")
        sum_entry.config(state='readonly')
    
    def update_all_sums(self):
        for i in range(len(self.columns)):
            self.update_sum(i)
    
    def save_state(self):
        state = {
            "data_rows": self.data_rows,
            "columns": []
        }
        for col_entries in self.columns:
            col_data = {
                "header": col_entries[1].get(),
                "values": [entry.get() for entry in col_entries[2:-1]]
            }
            state["columns"].append(col_data)
        
        if self.undo_stack and state == self.undo_stack[-1]:
            return
        
        self.undo_stack.append(copy.deepcopy(state))
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)
        self.redo_stack.clear()
    
    def restore_state(self, state):
        for col_entries in self.columns:
            for widget in col_entries:
                widget.grid_forget()
                widget.destroy()
        self.columns.clear()
        
        self.data_rows = state.get("data_rows", 1)
        
        for col_data in state.get("columns", []):
            self.add_column()
            col_index = len(self.columns) - 1
            self.columns[col_index][1].delete(0, tk.END)
            self.columns[col_index][1].insert(0, col_data.get("header", f"Столбец {col_index+1}"))
            
            values = col_data.get("values", [])
            while len(values) < self.data_rows:
                values.append("")
            while len(values) > self.data_rows:
                self.add_row()
            
            for row_i, val in enumerate(values):
                entry = self.columns[col_index][2 + row_i]
                entry.delete(0, tk.END)
                entry.insert(0, val)
        
        self.refresh_delete_row_buttons()
        self.update_all_sums()
    
    def undo_action(self):
        if len(self.undo_stack) < 2:
            return
        current_state = self.undo_stack.pop()
        self.redo_stack.append(current_state)
        prev_state = self.undo_stack[-1]
        self.restore_state(copy.deepcopy(prev_state))
    
    def redo_action(self):
        if not self.redo_stack:
            return
        state = self.redo_stack.pop()
        self.undo_stack.append(state)
        self.restore_state(copy.deepcopy(state))
    
    def prepare_save_data(self):
        data = {
            "data_rows": self.data_rows,
            "columns": []
        }
        for col_entries in self.columns:
            col_data = {
                "header": col_entries[1].get(),
                "values": [entry.get() for entry in col_entries[2:-1]]
            }
            data["columns"].append(col_data)
        return data
    
    def save_to_file(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".json",
                                                filetypes=[("JSON файлы", "*.json"), ("Все файлы", "*.*")])
        if not filepath:
            return
        try:
            data = self.prepare_save_data()
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            self.current_file = filepath
            self.unsaved_changes = False
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
    
    def quick_save(self):
        if not self.current_file:
            self.save_to_file()
            return
        try:
            data = self.prepare_save_data()
            with open(self.current_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            self.unsaved_changes = False
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
    
    def load_from_file(self):
        filepath = filedialog.askopenfilename(defaultextension=".json",
                                              filetypes=[("JSON файлы", "*.json"), ("Все файлы", "*.*")])
        if not filepath:
            return
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.restore_state(data)
            self.undo_stack.clear()
            self.redo_stack.clear()
            self.save_state()
            self.current_file = filepath
            self.unsaved_changes = False
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")
    
    def import_columns_from_directory(self):
        directory = filedialog.askdirectory()
        if not directory:
            return
        try:
            files = os.listdir(directory)
            files = [f for f in files if os.path.isfile(os.path.join(directory, f))]
            if not files:
                messagebox.showinfo("Импорт из директории", "В выбранной директории нет файлов.")
                return
            
            for file_name in files:
                self.add_column()
                col_index = len(self.columns) - 1
                self.columns[col_index][1].delete(0, tk.END)
                self.columns[col_index][1].insert(0, file_name)
            
            self.mark_changes()
            self.save_state()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать из директории:\n{e}")
    
    def on_close(self):
        if self.unsaved_changes:
            answer = messagebox.askyesnocancel("Сохранение",
                        "У вас есть несохранённые изменения. Хотите сохранить перед выходом?")
            if answer is True:
                self.quick_save()
                self.destroy()
            elif answer is False:
                self.destroy()
        else:
            self.destroy()
    
    def _bind_ctrl_v(self, entry_widget):
        def on_paste(event):
            event.widget.event_generate("<<Paste>>")
            return "break"
        entry_widget.bind("<Control-v>", on_paste)
        entry_widget.bind("<Control-V>", on_paste)

if __name__ == "__main__":
    init_file = None
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if os.path.isfile(arg) and arg.lower().endswith(".json"):
            init_file = arg
    app = JournalApp(init_file=init_file)
    app.mainloop()
