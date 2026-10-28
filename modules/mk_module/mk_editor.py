import tkinter as tk
from tkinter import ttk


class OperationEditor:
    def __init__(self, root, operations_list, standard_transitions, tools_db, prof_db, rank_coefs, hardness_hb,
                 saved_times, on_save_callback):
        self.window = tk.Toplevel(root)
        self.window.title("Настройка переходов и оснастки")

        self.window.transient(root)
        self.window.grab_set()

        window_width = 750
        window_height = 650
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.window.configure(padx=10, pady=10)

        self.ops = operations_list
        self.std_transitions = standard_transitions
        self.tools_db = tools_db
        self.prof_db = prof_db
        self.rank_coefs = rank_coefs
        self.hardness_hb = hardness_hb
        self.saved_times = saved_times  # <-- Времена из ОК
        self.on_save = on_save_callback
        self.entries = []

        self.create_widgets()

    def create_widgets(self):
        btn_save = tk.Button(self.window, text="Применить и сохранить в Excel",
                             bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
                             command=self.save_and_close)
        btn_save.pack(side="bottom", fill="x", pady=10)

        main_frame = tk.Frame(self.window)
        main_frame.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for i, op in enumerate(self.ops):
            frame = tk.LabelFrame(scrollable_frame, text=f"Операция: {op['name']}", padx=10, pady=10)
            frame.pack(fill="x", pady=5, padx=5)

            op_name = op.get("name", "").lower()

            if "токарн" in op_name: op_type = "turning"
            elif "сверл" in op_name: op_type = "drilling"
            elif "свар" in op_name: op_type = "welding"
            elif "слесар" in op_name: op_type = "bench"
            elif "заготов" in op_name: op_type = "cutting"
            elif "контрол" in op_name: op_type = "control"
            else: op_type = "any"

            # 1. Переходы
            available_transitions_data = self.std_transitions.get(op_type, [])
            transition_texts = [t["text"] for t in available_transitions_data]

            tk.Label(frame, text="Переходы:").pack(anchor="w")
            combo_transitions = ttk.Combobox(frame, values=transition_texts, width=60)
            combo_transitions.pack(fill="x", pady=2)
            if transition_texts:
                combo_transitions.set(transition_texts[0])

            entry_desc = tk.Text(frame, height=2, width=60)
            entry_desc.insert("1.0", combo_transitions.get())
            entry_desc.pack(fill="x", pady=2)

            # 2. Инструменты
            tk.Label(frame, text="Инструмент:").pack(anchor="w")
            combo_tools = ttk.Combobox(frame, width=60)
            combo_tools.pack(fill="x", pady=2)

            # 3. Блок нормирования (Профессия, Разряд, Время)
            prof_frame = tk.Frame(frame)
            prof_frame.pack(fill="x", pady=5)

            tk.Label(prof_frame, text="Профессия:").pack(side="left")
            combo_prof = ttk.Combobox(prof_frame, values=list(self.prof_db.keys()), width=15, state="readonly")
            combo_prof.pack(side="left", padx=5)

            default_prof = op.get("profession", "Слесарь")
            if default_prof in self.prof_db:
                combo_prof.set(default_prof)
            else:
                combo_prof.current(0)

            tk.Label(prof_frame, text="Разряд:").pack(side="left")
            combo_rank = ttk.Combobox(prof_frame, width=5, state="readonly")
            combo_rank.pack(side="left", padx=5)

            # Поле для Тшт (то, чего у тебя не хватало)
            tk.Label(prof_frame, text="Тшт (мин):").pack(side="left", padx=(10, 0))
            entry_time = tk.Entry(prof_frame, width=8)
            entry_time.pack(side="left", padx=5)

            # --- АВТОМАТИЧЕСКАЯ ПОДСТАНОВКА ВРЕМЕНИ ИЗ ОК ---
            if op['name'] in self.saved_times:
                # Если время для этой операции уже посчитано в ОК — берем его
                entry_time.insert(0, str(self.saved_times[op['name']]))
            else:
                # Если нет — ставим базовое (например, 5.0)
                entry_time.insert(0, "5.0")

            # Функция обновления разрядов
            def update_ranks(event=None, prof_combo=combo_prof, rank_combo=combo_rank):
                prof = prof_combo.get()
                ranks = [str(r) for r in self.prof_db.get(prof, [])]
                rank_combo['values'] = ranks
                if ranks: rank_combo.current(0)

            combo_prof.bind("<<ComboboxSelected>>", update_ranks)
            update_ranks()

            # Функция инструментов
            def update_tools(text_widget=entry_desc, tool_combo=combo_tools, base_type=op_type):
                text_lower = text_widget.get("1.0", tk.END).lower()
                target_tool_type = base_type
                if "сверл" in text_lower:
                    target_tool_type = "drilling"
                elif "точить" in text_lower or "подрезать" in text_lower:
                    target_tool_type = "turning"
                elif "сварить" in text_lower or "прихватить" in text_lower:
                    target_tool_type = "welding"

                valid_tools = []
                for tool_data in self.tools_db.values():
                    if tool_data["type"] != target_tool_type: continue
                    # Сравниваем твердость детали и максимальную твердость инструмента
                    if self.hardness_hb > tool_data.get("max_hardness", 999): continue
                    valid_tools.append(tool_data["name"])

                tool_combo['values'] = valid_tools
                if valid_tools and not tool_combo.get():
                    tool_combo.set(valid_tools[0])

            combo_transitions.bind("<<ComboboxSelected>>", lambda e: (entry_desc.delete("1.0", tk.END),
                                                                      entry_desc.insert("1.0", combo_transitions.get()),
                                                                      update_tools()))
            entry_desc.bind("<KeyRelease>", update_tools)
            update_tools()

            # Сохраняем все виджеты для сбора данных
            self.entries.append((entry_desc, combo_tools, combo_prof, combo_rank, entry_time))

    def save_and_close(self):
        for i, op in enumerate(self.ops):
            entry_desc, combo_tools, combo_prof, combo_rank, entry_time = self.entries[i]
            op["description"] = entry_desc.get("1.0", tk.END).strip()
            op["tool"] = combo_tools.get().strip()
            op["profession"] = combo_prof.get().strip()
            op["rank"] = combo_rank.get().strip()
            op["time_Tsht"] = entry_time.get().strip() # Забираем время из поля ввода

        self.window.destroy()
        self.on_save(self.ops)