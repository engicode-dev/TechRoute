import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from .mk_logic import prepare_operations_data
from .mk_excel_writer import write_to_excel
from .mk_editor import OperationEditor
from modules.ok_module.ok_gui import OKWindow
from utils import convert_to_pdf

class ProjectWindow:
    def __init__(self, root, user_settings):
        self.window = tk.Toplevel(root)
        self.window.title("Проект: Разработка детали")
        self.window.geometry("950x750")
        self.window.configure(padx=20, pady=20)

        self.saved_times = {}
        self.user_settings = user_settings
        self.save_path = user_settings.get('save_path', '')

        if not self.load_all_data():
            return
        self.create_widgets()

    def load_all_data(self):
        base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        data_dir = os.path.join(base_dir, 'data')

        try:
            with open(os.path.join(data_dir, 'operations.json'), 'r', encoding='utf-8') as f:
                self.ops_db = json.load(f)["operations"]
            with open(os.path.join(data_dir, 'transitions.json'), 'r', encoding='utf-8') as f:
                self.std_transitions = json.load(f)["transitions"]
            with open(os.path.join(data_dir, 'stock_shapes.json'), 'r', encoding='utf-8') as f:
                self.stock_db = json.load(f)["stock_shapes"]
            with open(os.path.join(data_dir, 'materials_grades.json'), 'r', encoding='utf-8') as f:
                self.grades_db = json.load(f)["materials_grades"]
            with open(os.path.join(data_dir, 'tools.json'), 'r', encoding='utf-8') as f:
                self.tools_db = json.load(f)["tools"]
            with open(os.path.join(data_dir, 'professions.json'), 'r', encoding='utf-8') as f:
                self.prof_db = json.load(f)["professions"]
            with open(os.path.join(data_dir, 'ranks_coefficients.json'), 'r', encoding='utf-8') as f:
                self.rank_coefs = json.load(f)["rank_coefficients"]

            self.ops_display = {f"{k}. {v['name']}": k for k, v in self.ops_db.items()}
            return True
        except Exception as e:
            messagebox.showerror("ОШИБКА ДАННЫХ", f"Не удалось загрузить JSON:\n{e}")
            self.window.destroy()
            return False

    def create_widgets(self):
        # --- ВЕРХНЯЯ ЧАСТЬ ---
        top_frame = tk.Frame(self.window)
        top_frame.pack(fill="x", pady=(0, 15))

        tk.Label(top_frame, text="Название детали:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.entry_name = tk.Entry(top_frame, width=25)
        self.entry_name.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(top_frame, text="Обозначение:", font=("Arial", 10, "bold")).grid(row=0, column=2, sticky="w")
        self.entry_code = tk.Entry(top_frame, width=15)
        self.entry_code.grid(row=0, column=3, padx=5, pady=5)

        # Заготовка
        tk.Label(top_frame, text="Тип заготовки:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w")
        self.combo_stock_type = ttk.Combobox(top_frame, values=[v['name'] for v in self.stock_db.values()], state="readonly", width=20)
        self.combo_stock_type.grid(row=1, column=1, padx=5, pady=5)
        self.combo_stock_type.bind("<<ComboboxSelected>>", self.on_stock_type_change)

        tk.Label(top_frame, text="Размер:", font=("Arial", 10, "bold")).grid(row=1, column=2, sticky="w")
        self.combo_stock_size = ttk.Combobox(top_frame, state="readonly", width=15)
        self.combo_stock_size.grid(row=1, column=3, padx=5, pady=5)

        # Длина и Ширина
        tk.Label(top_frame, text="Длина загот. (мм):", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w")
        self.entry_length = tk.Entry(top_frame, width=10)
        self.entry_length.insert(0, "300")
        self.entry_length.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        self.lbl_width = tk.Label(top_frame, text="Ширина загот. (мм):", font=("Arial", 10, "bold"))
        self.lbl_width.grid(row=2, column=2, sticky="w")
        self.entry_width = tk.Entry(top_frame, width=10)
        self.entry_width.insert(0, "100")
        self.entry_width.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        # Материал
        tk.Label(top_frame, text="Группа материала:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky="w")
        self.combo_grade_group = ttk.Combobox(top_frame, values=[v['name'] for v in self.grades_db.values()], state="readonly", width=20)
        self.combo_grade_group.grid(row=3, column=1, padx=5, pady=5)
        self.combo_grade_group.bind("<<ComboboxSelected>>", self.on_grade_group_change)

        tk.Label(top_frame, text="Марка:", font=("Arial", 10, "bold")).grid(row=3, column=2, sticky="w")
        self.combo_grade_name = ttk.Combobox(top_frame, state="readonly", width=15)
        self.combo_grade_name.grid(row=3, column=3, padx=5, pady=5)

        if self.stock_db:
            self.combo_stock_type.current(0)
            self.on_stock_type_change()
        if self.grades_db:
            self.combo_grade_group.current(0)
            self.on_grade_group_change()

        # --- ЦЕНТРАЛЬНАЯ ЧАСТЬ ---
        center_frame = tk.Frame(self.window)
        center_frame.pack(fill="both", expand=True)

        left_frame = tk.Frame(center_frame)
        left_frame.pack(side="left", fill="both", expand=True)
        tk.Label(left_frame, text="Доступные операции", font=("Arial", 10, "bold")).pack()
        self.listbox_avail = tk.Listbox(left_frame)
        self.listbox_avail.pack(fill="both", expand=True, pady=5)

        btn_frame = tk.Frame(center_frame)
        btn_frame.pack(side="left", padx=10, fill="y")
        tk.Button(btn_frame, text="Добавить >>", command=self.add_op).pack(pady=20)
        tk.Button(btn_frame, text="<< Удалить", command=self.remove_op).pack(pady=5)

        right_frame = tk.Frame(center_frame)
        right_frame.pack(side="right", fill="both", expand=True)
        tk.Label(right_frame, text="Маршрут изготовления", font=("Arial", 10, "bold")).pack()
        self.listbox_selected = tk.Listbox(right_frame)
        self.listbox_selected.pack(fill="both", expand=True, pady=5)

        # --- НИЖНЯЯ ЧАСТЬ ---
        bottom_frame = tk.Frame(self.window)
        bottom_frame.pack(pady=20, fill="x", side="bottom")

        tk.Button(bottom_frame, text="Сгенерировать МК", bg="#4CAF50", fg="white",
                  font=("Arial", 12, "bold"), command=self.generate_mk).pack(side="left", expand=True, fill="x", padx=5)

        tk.Button(bottom_frame, text="Сгенерировать ОК", bg="#2196F3", fg="white",
                  font=("Arial", 12, "bold"), command=self.generate_ok).pack(side="right", expand=True, fill="x", padx=5)

        self.window.update_idletasks()
        for name in self.ops_display.keys():
            self.listbox_avail.insert(tk.END, name)

    def on_stock_type_change(self, event=None):
        selected_type_name = self.combo_stock_type.get()
        if selected_type_name == "Лист":
            self.lbl_width.grid()
            self.entry_width.grid()
        else:
            self.lbl_width.grid_remove()
            self.entry_width.grid_remove()

        for key, data in self.stock_db.items():
            if data['name'] == selected_type_name:
                sizes = [f"{item['size']} ({item['gost']})" for item in data['items']]
                self.combo_stock_size['values'] = sizes
                if sizes: self.combo_stock_size.current(0)
                break

    def on_grade_group_change(self, event=None):
        selected_group_name = self.combo_grade_group.get()
        for key, data in self.grades_db.items():
            if data['name'] == selected_group_name:
                names = [f"{item['name']} ({item['gost']})" for item in data['items']]
                self.combo_grade_name['values'] = names
                if names: self.combo_grade_name.current(0)
                break

    def add_op(self):
        sel = self.listbox_avail.curselection()
        if sel: self.listbox_selected.insert(tk.END, self.listbox_avail.get(sel))

    def remove_op(self):
        sel = self.listbox_selected.curselection()
        if sel: self.listbox_selected.delete(sel)

    def _get_material_info(self):
        grade_display = self.combo_grade_name.get()
        grade_clean = grade_display.split(' (')[0]
        for group in self.grades_db.values():
            for item in group['items']:
                if item['name'] == grade_clean:
                    return item
        return {"name": grade_clean, "gost": "", "hardness_hb": 120}

    def calculate_weight(self):
        stock_type = self.combo_stock_type.get()
        stock_size = self.combo_stock_size.get()
        mat_info = self._get_material_info()
        density = mat_info.get('density', 7850)

        try:
            length = float(self.entry_length.get()) / 1000
            if self.combo_stock_type.get() == "Лист":
                width = float(self.entry_width.get()) / 1000
            else:
                width = 0.0
        except ValueError:
            return 0.0

        for group in self.stock_db.values():
            for item in group['items']:
                if item['size'] in stock_size:
                    geom_type = item['type']
                    if geom_type == "plate":
                        thickness = item['thickness'] / 1000
                        vol = length * width * thickness
                        return vol * density
                    elif geom_type == "round":
                        d = item['diameter'] / 1000
                        vol = 3.14159 * (d ** 2 / 4) * length
                        return vol * density
                    elif geom_type == "tube":
                        d_out = item['diameter'] / 1000
                        wall = item['wall'] / 1000
                        d_in = d_out - 2 * wall
                        vol = 3.14159 * ((d_out ** 2 - d_in ** 2) / 4) * length
                        return vol * density
        return 0.0

    def generate_mk(self):
        part_name = self.entry_name.get().strip()
        part_code = self.entry_code.get().strip()

        if not part_name or not part_code or self.listbox_selected.size() == 0:
            messagebox.showwarning("Внимание", "Заполните название, обозначение и выберите операции!")
            return

        stock_type = self.combo_stock_type.get()
        stock_size = self.combo_stock_size.get()
        mat_info = self._get_material_info()

        full_mat_string = f"{stock_type} {stock_size} / {mat_info['name']} {mat_info['gost']}"
        mass_zag = round(self.calculate_weight(), 2)

        self.header_data = {
            'name': part_name,
            'code': part_code,
            'material': full_mat_string,
            'mass': mass_zag,
            'developer': self.user_settings.get('developer_name', 'Инженер'),
            'organization': self.user_settings.get('organization', 'Организация')
        }

        raw_ops = []
        for i in range(self.listbox_selected.size()):
            op_key = self.ops_display[self.listbox_selected.get(i)]
            raw_ops.append(self.ops_db[op_key].copy())

        prepared_ops = prepare_operations_data(raw_ops)

        OperationEditor(
            self.window, prepared_ops, self.std_transitions, self.tools_db,
            self.prof_db, self.rank_coefs, mat_info['hardness_hb'], self.saved_times, self.finalize_mk
        )

    def finalize_mk(self, edited_ops):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "template_mk.xlsx")

        if self.save_path:
            output_filename = os.path.join(self.save_path, f"МК_{self.header_data['name']}.xlsx")
        else:
            output_filename = f"МК_{self.header_data['name']}.xlsx"

        try:
            write_to_excel(template_path, output_filename, self.header_data, edited_ops)

            if messagebox.askyesno("Успех", f"Файл МК создан!\nСделать PDF-версию?"):
                convert_to_pdf(output_filename)
            else:
                messagebox.showinfo("Успех", "Файл Excel сохранен.")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении: {e}")

    def generate_ok(self):
        sel = self.listbox_selected.curselection()
        if not sel:
            messagebox.showwarning("Внимание", "Выберите операцию в маршруте для создания ОК!")
            return

        op_display_name = self.listbox_selected.get(sel)
        op_key = self.ops_display[op_display_name]
        op_data = self.ops_db[op_key]

        mat_info = self._get_material_info()

        header_data = {
            'name': self.entry_name.get().strip(),
            'code': self.entry_code.get().strip(),
            'material': mat_info['name'],
            'hardness_hb': mat_info['hardness_hb'],
            'op_name': op_data['name'],
            'op_equipment': op_data['equipment'],
            'developer': self.user_settings.get('developer_name', 'Инженер'),
            'organization': self.user_settings.get('organization', 'Организация')
        }

        OKWindow(self.window, header_data, self.std_transitions, self.tools_db, self.save_path, self.save_ok_time)

    def save_ok_time(self, op_name, time_value):
        self.saved_times[op_name] = time_value
        messagebox.showinfo("Связка", f"Время для операции '{op_name}' ({time_value} мин) сохранено! МК его подхватит.")