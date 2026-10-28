import tkinter as tk
from tkinter import ttk, messagebox
import os
from .spec_excel_writer import write_spec_excel


# Импорт convert_to_pdf убран, так как мы не делаем PDF для спецификации

class SpecWindow:
    def __init__(self, root, header_data, user_settings, on_save_callback):
        self.window = tk.Toplevel(root)
        self.window.title("Генератор Спецификации")
        self.window.geometry("900x700")
        self.window.configure(padx=20, pady=20)

        self.header_data = header_data
        self.user_settings = user_settings
        self.on_save = on_save_callback
        self.items = []

        self.create_widgets()

    def create_widgets(self):
        # --- ВЕРХНЯЯ ЧАСТЬ ---
        top_frame = tk.Frame(self.window)
        top_frame.pack(fill="x", pady=(0, 15))

        tk.Label(top_frame, text="Название изделия:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.entry_name = tk.Entry(top_frame, width=30)
        self.entry_name.insert(0, self.header_data.get('name', ''))
        self.entry_name.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(top_frame, text="Обозначение:", font=("Arial", 10, "bold")).grid(row=0, column=2, sticky="w")
        self.entry_code = tk.Entry(top_frame, width=20)
        self.entry_code.insert(0, self.header_data.get('code', ''))
        self.entry_code.grid(row=0, column=3, padx=10, pady=5)

        # --- ЦЕНТРАЛЬНАЯ ЧАСТЬ ---
        tk.Label(self.window, text="Состав изделия:", font=("Arial", 10, "bold")).pack(anchor="w")
        list_frame = tk.Frame(self.window)
        list_frame.pack(fill="both", expand=True, pady=5)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        # --- ФОРМА ДОБАВЛЕНИЯ ---
        input_frame = tk.LabelFrame(self.window, text="Добавить позицию", padx=10, pady=10)
        input_frame.pack(fill="x", pady=10)

        tk.Label(input_frame, text="Категория:").grid(row=0, column=0, sticky="w")
        self.combo_category = ttk.Combobox(input_frame, values=[
            "Документация", "Детали", "Сборочные единицы", "Стандартные изделия", "Материалы"
        ], state="readonly", width=20)
        self.combo_category.grid(row=1, column=0, padx=5, pady=2)
        self.combo_category.current(1)

        tk.Label(input_frame, text="Формат:").grid(row=0, column=1, sticky="w")
        self.entry_format = tk.Entry(input_frame, width=8)
        self.entry_format.insert(0, "А4")
        self.entry_format.grid(row=1, column=1, padx=5, pady=2)

        tk.Label(input_frame, text="Зона:").grid(row=0, column=2, sticky="w")
        self.entry_zone = tk.Entry(input_frame, width=8)
        self.entry_zone.grid(row=1, column=2, padx=5, pady=2)

        tk.Label(input_frame, text="Обозначение:").grid(row=0, column=3, sticky="w")
        self.entry_code_item = tk.Entry(input_frame, width=20)
        self.entry_code_item.grid(row=1, column=3, padx=5, pady=2)

        tk.Label(input_frame, text="Наименование:").grid(row=0, column=4, sticky="w")
        self.entry_name_item = tk.Entry(input_frame, width=35)
        self.entry_name_item.grid(row=1, column=4, padx=5, pady=2)

        tk.Label(input_frame, text="Кол-во:").grid(row=0, column=5, sticky="w")
        self.entry_count = tk.Entry(input_frame, width=5)
        self.entry_count.insert(0, "1")
        self.entry_count.grid(row=1, column=5, padx=5, pady=2)

        tk.Button(input_frame, text="Добавить позицию", bg="#2196F3", fg="white", font=("Arial", 10, "bold"),
                  command=self.add_item).grid(row=2, column=0, columnspan=6, pady=10, sticky="we")

        # --- НИЖНЯЯ ЧАСТЬ ---
        tk.Button(self.window, text="Сгенерировать спецификацию в Excel", bg="#4CAF50", fg="white",
                  font=("Arial", 12, "bold"), command=self.generate).pack(side="bottom", pady=10, fill="x")

    def add_item(self):
        category = self.combo_category.get()
        fmt = self.entry_format.get().strip()
        zone = self.entry_zone.get().strip()
        code = self.entry_code_item.get().strip()
        name = self.entry_name_item.get().strip()
        count = self.entry_count.get().strip()

        if not name:
            messagebox.showwarning("Ошибка", "Введите наименование!")
            return

        self.items.append({
            "category": category,
            "format": fmt,
            "zone": zone,
            "code": code,
            "name": name,
            "count": count
        })

        self.listbox.insert(tk.END, f"[{category}] {code} - {name} ({count} шт.)")

        self.entry_code_item.delete(0, tk.END)
        self.entry_name_item.delete(0, tk.END)
        self.entry_count.delete(0, tk.END)
        self.entry_count.insert(0, "1")

    def generate(self):
        part_name = self.entry_name.get().strip()
        part_code = self.entry_code.get().strip()

        if not part_name or not part_code:
            messagebox.showwarning("Ошибка", "Введите название и обозначение изделия!")
            return

        if not self.items:
            messagebox.showwarning("Ошибка", "Добавьте хотя бы одну позицию!")
            return

        # Добавляем данные пользователя
        self.header_data['name'] = part_name
        self.header_data['code'] = part_code
        self.header_data['developer'] = self.user_settings.get('developer_name', 'Инженер')
        self.header_data['organization'] = self.user_settings.get('organization', 'Организация')

        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "template_spec.xlsx")

        # Путь сохранения
        save_path = self.user_settings.get('save_path', '')
        if save_path:
            output_path = os.path.join(save_path, f"Спецификация_{part_name}.xlsx")
        else:
            output_path = os.path.join(base_dir, f"Спецификация_{part_name}.xlsx")

        try:
            write_spec_excel(template_path, output_path, self.header_data, self.items)

            # Показываем сообщение без предложения PDF
            messagebox.showinfo(
                "Успех",
                f"Спецификация сохранена в Excel!\n\n"
                f"Путь: {output_path}\n\n"
                f"(Для создания PDF: откройте файл, выделите все строки,\n"
                f"поставьте высоту 0.1 см и экспортируйте в PDF)"
            )
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении:\n{e}")