import tkinter as tk
from tkinter import filedialog, messagebox
import json
import copy

class ScrollableTable(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Фрейм заголовков (фиксирован по вертикали, прокручивается по горизонтали)
        self.header_frame = tk.Frame(self)
        self.header_frame.pack(side=tk.TOP, fill=tk.X)
        
        # Горизонтальный скроллбар для заголовков
        self.h_scroll_header = tk.Scrollbar(self.header_frame, orient=tk.HORIZONTAL)
        self.h_scroll_header.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Канвас для заголовков
        self.header_canvas = tk.Canvas(self.header_frame, height=50,
                                       xscrollcommand=self.h_scroll_header.set)
        self.header_canvas.pack(side=tk.TOP, fill=tk.X, expand=True)
        self.h_scroll_header.config(command=self.header_canvas.xview)
        
        self.header_inner = tk.Frame(self.header_canvas)
        self.header_window = self.header_canvas.create_window((0,0), window=self.header_inner, anchor='nw')
        
        self.header_inner.bind("<Configure>", lambda e: self.header_canvas.configure(scrollregion=self.header_canvas.bbox("all")))
        
        # Фрейм с таблицей и скроллами
        self.table_frame = tk.Frame(self)
        self.table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Вертикальный и горизонтальный скроллбары для таблицы
        self.v_scroll = tk.Scrollbar(self.table_frame, orient=tk.VERTICAL)
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.h_scroll = tk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Канвас для таблицы
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
        
        # Синхронизируем горизонтальный скроллбар заголовков и таблицы
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
    def __init__(self):
        super().__init__()
        self.title("Таблица учёта продукции на уголковой линии")
        self.geometry("360x600")

        try:
            self.iconbitmap('favicon.ico')
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")
        
        self.data_rows = 1
        self.columns = []
        
        self.undo_stack = []
        self.redo_stack = []
        self.max_undo = 50
        self.current_file = None
        
        # Меню
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Сохранить как...", command=self.save_to_file)
        filemenu.add_command(label="Загрузить", command=self.load_from_file)
        menubar.add_cascade(label="Файл", menu=filemenu)
        
        editmenu = tk.Menu(menubar, tearoff=0)
        editmenu.add_command(label="Отменить (Ctrl+Z)", command=self.undo_action)
        editmenu.add_command(label="Вернуть (Ctrl+Y)", command=self.redo_action)
        menubar.add_cascade(label="Правка", menu=editmenu)
        
        self.config(menu=menubar)
        
        # Верхняя панель с кнопками (отдельно, вне прокрутки)
        self.top_buttons_frame = tk.Frame(self)
        self.top_buttons_frame.pack(side=tk.TOP, fill=tk.X)
        
        self.add_col_btn = tk.Button(self.top_buttons_frame, text="Добавить столбец", command=self.add_column)
        self.add_col_btn.pack(side=tk.LEFT, padx=5, pady=2)
        
        self.quick_save_btn = tk.Button(self.top_buttons_frame, text="Быстрое сохранение", command=self.quick_save)
        self.quick_save_btn.pack(side=tk.LEFT, padx=5, pady=2)
        
        # Прокручиваемая таблица
        self.scrollable_table = ScrollableTable(self)
        self.scrollable_table.pack(fill=tk.BOTH, expand=True)
        
        # Кнопка "Добавить строку" снизу (вне прокрутки)
        self.add_row_btn = tk.Button(self, text="Добавить строку", command=self.add_row)
        self.add_row_btn.pack(side=tk.BOTTOM, pady=5)
        
        self.del_row_buttons = []
        
        for _ in range(3):
            self.add_column()
        self.refresh_delete_row_buttons()
        
        self.bind_all("<Control-z>", lambda e: self.undo_action())
        self.bind_all("<Control-y>", lambda e: self.redo_action())
        
        self.save_state()
    
    # --- Методы работы с таблицей ---
    def add_column(self):
        col_index = len(self.columns)
        col_entries = []
        
        del_col_btn = tk.Button(self.scrollable_table.header_inner, text="Удалить", fg="red",
                                command=lambda c=col_index: self.delete_column(c))
        del_col_btn.grid(row=0, column=col_index, padx=5, pady=2)
        
        header = tk.Entry(self.scrollable_table.header_inner, width=12, justify='center')
        header.insert(0, f"Столбец {col_index+1}")
        header.grid(row=1, column=col_index, padx=5, pady=2)
        header.bind("<KeyRelease>", lambda e: self.save_state())
        
        col_entries.append(del_col_btn)
        col_entries.append(header)
        
        for row in range(self.data_rows):
            entry = tk.Entry(self.scrollable_table.table_inner, width=12, justify='center')
            entry.grid(row=row, column=col_index, padx=5, pady=2)
            entry.bind("<KeyRelease>", lambda e, c=col_index: [self.update_sum(c), self.save_state()])
            col_entries.append(entry)
        
        sum_entry = tk.Entry(self.scrollable_table.table_inner, width=12, justify='center', state='readonly', readonlybackground='lightgray')
        sum_entry.grid(row=self.data_rows, column=col_index, padx=5, pady=2)
        col_entries.append(sum_entry)
        
        self.columns.append(col_entries)
        self.update_sum(col_index)
        self.refresh_delete_row_buttons()
        self.save_state()
    
    def add_row(self):
        self.data_rows += 1
        row_index = self.data_rows - 1
        
        for col_index, col_entries in enumerate(self.columns):
            sum_entry = col_entries[-1]
            sum_entry.grid_forget()
            sum_entry.grid(row=self.data_rows, column=col_index, padx=5, pady=2)
            
            new_entry = tk.Entry(self.scrollable_table.table_inner, width=12, justify='center')
            new_entry.grid(row=row_index, column=col_index, padx=5, pady=2)
            new_entry.bind("<KeyRelease>", lambda e, c=col_index: [self.update_sum(c), self.save_state()])
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
        self.save_state()
    
    def delete_row(self, row_index):
        if self.data_rows <= 1:
            messagebox.showwarning("Предупреждение", "Нельзя удалить последнюю строку!")
            return
        
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
        
        self.refresh_delete_row_buttons()
        self.update_all_sums()
        self.save_state()
    
    def refresh_delete_row_buttons(self):
        for btn in self.del_row_buttons:
            btn.grid_forget()
            btn.destroy()
        self.del_row_buttons.clear()
        
        col_for_buttons = len(self.columns)
        for row in range(self.data_rows):
            btn = tk.Button(self.scrollable_table.table_inner, text="Удалить", fg="red",
                            command=lambda r=row: self.delete_row(r))
            btn.grid(row=row, column=col_for_buttons, padx=5, pady=2)
            self.del_row_buttons.append(btn)
    
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
    
    # --- Undo/Redo ---
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
            messagebox.showinfo("Отмена", "Отменять нечего.")
            return
        current_state = self.undo_stack.pop()
        self.redo_stack.append(current_state)
        prev_state = self.undo_stack[-1]
        self.restore_state(copy.deepcopy(prev_state))
    
    def redo_action(self):
        if not self.redo_stack:
            messagebox.showinfo("Повтор", "Повторять нечего.")
            return
        state = self.redo_stack.pop()
        self.undo_stack.append(state)
        self.restore_state(copy.deepcopy(state))
    
    # --- Сохранение/загрузка ---
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
            messagebox.showinfo("Сохранение", "Данные успешно сохранены.")
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
            messagebox.showinfo("Быстрое сохранение", f"Изменения сохранены в:\n{self.current_file}")
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
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")
            return
        
        self.restore_state(data)
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.save_state()
        self.current_file = filepath
        messagebox.showinfo("Загрузка", "Данные успешно загружены.")

if __name__ == "__main__":
    app = JournalApp()
    app.mainloop()
