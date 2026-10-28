import tkinter as tk
from tkinter import ttk, messagebox
import os
import math
from .ok_excel_writer import write_to_ok


class OKWindow:
    def __init__(self, root, header_data, std_transitions, tools_db, save_path, on_save_callback):
        self.window = tk.Toplevel(root)
        self.window.title(f"Операционная карта: {header_data.get('op_name', '')}")
        self.window.geometry("1000x750")
        self.window.configure(padx=20, pady=20)

        self.header_data = header_data
        self.std_transitions = std_transitions
        self.tools_db = tools_db
        self.save_path = save_path  # <-- Путь для сохранения
        self.on_save = on_save_callback
        self.transitions = []
        self.total_main_time = 0.0
        self.aux_time = 2.0

        self.op_name = header_data.get('op_name', '').lower()
        self.is_welding = "свар" in self.op_name

        if "токарн" in self.op_name:
            self.op_type = "turning"
        elif "сверл" in self.op_name:
            self.op_type = "drilling"
        elif "свар" in self.op_name:
            self.op_type = "welding"
        elif "слесар" in self.op_name:
            self.op_type = "bench"
        else:
            self.op_type = "any"

        self.create_widgets()

    def create_widgets(self):
        # 1. Список переходов
        tk.Label(self.window, text="Добавленные переходы:", font=("Arial", 10, "bold")).pack(anchor="w")
        list_frame = tk.Frame(self.window)
        list_frame.pack(fill="both", expand=True, pady=5)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        # 2. Форма добавления
        input_frame = tk.LabelFrame(self.window, text="Добавить переход", padx=10, pady=10)
        input_frame.pack(fill="x", pady=10)

        # Переходы
        tk.Label(input_frame, text="Выбрать шаблон перехода:").grid(row=0, column=0, sticky="w")
        avail_trans = self.std_transitions.get(self.op_type, [])
        trans_texts = [t["text"] for t in avail_trans]
        self.combo_trans = ttk.Combobox(input_frame, values=trans_texts, width=70)
        self.combo_trans.grid(row=1, column=0, columnspan=4, sticky="we", pady=2, padx=5)

        tk.Label(input_frame, text="Текст перехода:").grid(row=2, column=0, sticky="w")
        self.entry_text = tk.Entry(input_frame, width=70)
        self.entry_text.grid(row=3, column=0, columnspan=4, sticky="we", pady=2, padx=5)

        def on_trans_select(event):
            self.entry_text.delete(0, tk.END)
            self.entry_text.insert(0, self.combo_trans.get())

        self.combo_trans.bind("<<ComboboxSelected>>", on_trans_select)

        # Инструмент
        tk.Label(input_frame, text="Инструмент:").grid(row=4, column=0, sticky="w")
        self.combo_tool = ttk.Combobox(input_frame, width=70)
        self.combo_tool.grid(row=5, column=0, columnspan=4, sticky="we", pady=2, padx=5)

        # Режимы / Параметры
        self.params_frame = tk.Frame(input_frame)
        self.params_frame.grid(row=6, column=0, columnspan=4, sticky="we", pady=10)

        self.entries_params = {}

        if self.is_welding:
            params = [("D", "Катет шва (мм)"), ("L", "Длина шва (мм)"), ("V", "Ток (А)")]
        else:
            params = [("D", "Диаметр (мм)"), ("L", "Длина (мм)"), ("t", "Глубина (мм)"),
                      ("i", "Проходов"), ("S", "Подача (мм/об)"), ("n", "Обороты (об/мин)"), ("V", "Скорость (м/мин)")]

        for i, (key, label) in enumerate(params):
            tk.Label(self.params_frame, text=label).grid(row=0, column=i, sticky="w")
            entry = tk.Entry(self.params_frame, width=10)
            entry.grid(row=1, column=i, padx=2, pady=2)
            self.entries_params[key] = entry

        # Кнопки
        btn_frame = tk.Frame(input_frame)
        btn_frame.grid(row=7, column=0, columnspan=4, pady=10, sticky="we")

        tk.Button(btn_frame, text="Добавить переход", bg="#2196F3", fg="white", font=("Arial", 10, "bold"),
                  command=self.add_transition).pack(side="left", fill="x", expand=True, padx=5)

        if self.is_welding:
            tk.Button(btn_frame, text="⚡ Рассчитать сварку", bg="#FF9800", fg="white", font=("Arial", 10, "bold"),
                      command=self.auto_welding).pack(side="right", fill="x", expand=True, padx=5)
        else:
            tk.Button(btn_frame, text="⚡ Автоподбор режимов", bg="#FF9800", fg="white", font=("Arial", 10, "bold"),
                      command=self.auto_select_modes).pack(side="right", fill="x", expand=True, padx=5)

        # 3. Время и сохранение
        time_frame = tk.Frame(self.window)
        time_frame.pack(fill="x", pady=10)

        tk.Label(time_frame, text="Вспомогательное время (Тв), мин:").pack(side="left")
        self.entry_aux = tk.Entry(time_frame, width=10)
        self.entry_aux.insert(0, "2.0")
        self.entry_aux.pack(side="left", padx=5)

        tk.Label(time_frame, text="Итоговое Тшт:").pack(side="left", padx=(20, 0))
        self.lbl_total_time = tk.Label(time_frame, text="0.00", fg="blue", font=("Arial", 10, "bold"))
        self.lbl_total_time.pack(side="left", padx=5)

        tk.Button(self.window, text="Сгенерировать ОК в Excel", bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
                  command=self.generate).pack(side="bottom", pady=10, fill="x")

        self.update_tools()

    def update_tools(self):
        valid_tools = []
        for tool_data in self.tools_db.values():
            if tool_data["type"] == self.op_type:
                valid_tools.append(tool_data["name"])

        self.combo_tool['values'] = valid_tools
        if valid_tools:
            self.combo_tool.set(valid_tools[0])

    def auto_welding(self):
        try:
            k = float(self.entries_params["D"].get())
            L = float(self.entries_params["L"].get())
        except ValueError:
            messagebox.showwarning("Ошибка", "Введите Катет и Длину шва!")
            return

        area = (k * k) / 2
        mass = (area * L) * 0.00785
        I = 150

        main_time = (mass / (10 * I)) * 60

        self.entries_params["V"].delete(0, tk.END)
        self.entries_params["V"].insert(0, str(I))

        messagebox.showinfo("Расчет", f"Масса: {mass:.1f} г\nВремя (То): {main_time:.2f} мин\nТок: {I} А")
        self._calc_main_time = main_time

    def auto_select_modes(self):
        d_str = self.entries_params["D"].get().strip()
        if not d_str:
            messagebox.showwarning("Ошибка", "Введите Диаметр (D)!")
            return
        try:
            D = float(d_str)
        except ValueError:
            return

        mat_hb = self.header_data.get('hardness_hb', 120)

        if "токарн" in self.op_name:
            if mat_hb < 150:
                V = 140;
                S = 0.4;
                t = 2.0
            elif mat_hb < 220:
                V = 100;
                S = 0.3;
                t = 1.5
            else:
                V = 70;
                S = 0.2;
                t = 1.0

            n = int((1000 * V) / (math.pi * D))
            self.entries_params["V"].delete(0, tk.END);
            self.entries_params["V"].insert(0, str(V))
            self.entries_params["S"].delete(0, tk.END);
            self.entries_params["S"].insert(0, str(S))
            self.entries_params["n"].delete(0, tk.END);
            self.entries_params["n"].insert(0, str(n))
            self.entries_params["t"].delete(0, tk.END);
            self.entries_params["t"].insert(0, str(t))

        elif "сверл" in self.op_name:
            V = 30 if mat_hb < 150 else 15
            S = 0.2 if mat_hb < 150 else 0.1
            n = int((1000 * V) / (math.pi * D))
            self.entries_params["V"].delete(0, tk.END);
            self.entries_params["V"].insert(0, str(V))
            self.entries_params["S"].delete(0, tk.END);
            self.entries_params["S"].insert(0, str(S))
            self.entries_params["n"].delete(0, tk.END);
            self.entries_params["n"].insert(0, str(n))

        messagebox.showinfo("Успех", "Режимы подобраны!")

    def add_transition(self):
        text = self.entry_text.get().strip()
        if not text:
            messagebox.showwarning("Внимание", "Введите текст перехода!")
            return

        params = {k: v.get().strip() for k, v in self.entries_params.items()}

        main_time = 0.0
        if self.is_welding and hasattr(self, '_calc_main_time'):
            main_time = self._calc_main_time
            del self._calc_main_time
        else:
            try:
                L = float(params.get("L", 0))
                i = float(params.get("i", 1))
                S = float(params.get("S", 0))
                n = float(params.get("n", 0))
                if n > 0 and S > 0:
                    main_time = (L * i) / (n * S)
            except ValueError:
                pass

        self.total_main_time += main_time
        self.update_total_time()

        self.transitions.append({"text": text, "tool": self.combo_tool.get(), "params": params, "main_time": main_time})

        display_text = f"{len(self.transitions):02d}. {text} (То: {main_time:.2f} мин)"
        self.listbox.insert(tk.END, display_text)

        self.entry_text.delete(0, tk.END)
        for entry in self.entries_params.values():
            entry.delete(0, tk.END)

    def update_total_time(self):
        try:
            self.aux_time = float(self.entry_aux.get())
        except ValueError:
            self.aux_time = 0.0
        total_time = self.total_main_time + self.aux_time
        self.lbl_total_time.config(text=f"{total_time:.2f}")

    def generate(self):
        if not self.transitions:
            messagebox.showwarning("Ошибка", "Добавьте хотя бы один переход!")
            return

        self.update_total_time()
        final_time = float(self.lbl_total_time.cget("text"))

        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "template_ok.xlsx")

        # Используем сохраненный путь или папку модуля
        if self.save_path:
            output_path = os.path.join(self.save_path, f"ОК_{self.header_data.get('name', 'Деталь')}.xlsx")
        else:
            output_path = os.path.join(base_dir, f"ОК_{self.header_data.get('name', 'Деталь')}.xlsx")

        try:
            saved_path = write_to_ok(template_path, output_path, self.header_data, self.transitions,
                                     self.total_main_time, self.aux_time, final_time)
            messagebox.showinfo("Успех", f"Файл создан!\n\nПуть:\n{saved_path}")

            # Передаем время обратно в МК
            self.on_save(self.header_data.get('op_name', ''), final_time)

            self.window.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении:\n{e}")