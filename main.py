import tkinter as tk
from tkinter import messagebox, filedialog
import json
import os
from modules.mk_module.mk_gui import ProjectWindow
from modules.spec_module.spec_gui import SpecWindow


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("TechRoute - Инженерный Помощник")
        self.root.geometry("450x500")
        self.root.configure(padx=20, pady=20)

        self.settings_path = os.path.join(os.path.dirname(__file__), 'data', 'user_settings.json')
        self.user_settings = {}
        self.save_path = ""  # Путь сохранения

        if not self.load_settings():
            self.show_auth_window()
        else:
            self.create_widgets()

    def load_settings(self):
        if os.path.exists(self.settings_path):
            try:
                with open(self.settings_path, 'r', encoding='utf-8') as f:
                    self.user_settings = json.load(f)
                self.save_path = self.user_settings.get('save_path', '')
                return True
            except:
                return False
        return False

    def save_settings(self, name, org):
        self.user_settings = {"developer_name": name, "organization": org, "save_path": self.save_path}
        os.makedirs(os.path.dirname(self.settings_path), exist_ok=True)
        with open(self.settings_path, 'w', encoding='utf-8') as f:
            json.dump(self.user_settings, f, ensure_ascii=False, indent=2)

    def show_auth_window(self):
        self.auth_win = tk.Toplevel(self.root)
        self.auth_win.title("Добро пожаловать!")
        self.auth_win.geometry("400x300")
        self.auth_win.configure(padx=20, pady=20)
        self.auth_win.grab_set()

        tk.Label(self.auth_win, text="Давайте познакомимся!", font=("Arial", 14, "bold")).pack(pady=10)

        tk.Label(self.auth_win, text="Ваша фамилия и инициалы:").pack(anchor="w")
        entry_name = tk.Entry(self.auth_win, width=30)
        entry_name.pack(fill="x", pady=5)

        tk.Label(self.auth_win, text="Организация:").pack(anchor="w")
        entry_org = tk.Entry(self.auth_win, width=30)
        entry_org.pack(fill="x", pady=5)

        def confirm():
            name = entry_name.get().strip()
            org = entry_org.get().strip()
            if not name or not org:
                messagebox.showwarning("Ошибка", "Заполните все поля!")
                return
            self.save_settings(name, org)
            self.auth_win.destroy()
            self.create_widgets()

        tk.Button(self.auth_win, text="Начать работу", bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
                  command=confirm).pack(pady=20, fill="x")

    def create_widgets(self):
        title = tk.Label(self.root, text="TechRoute", font=("Arial", 24, "bold"))
        title.pack(pady=(10, 10))

        welcome_text = f"Пользователь: {self.user_settings.get('developer_name')}\n{self.user_settings.get('organization')}"
        tk.Label(self.root, text=welcome_text, font=("Arial", 9), fg="gray").pack(pady=(0, 10))

        # Кнопка выбора папки
        self.lbl_path = tk.Label(self.root, text=f"Папка: {self.save_path or 'Не выбрана'}", font=("Arial", 8),
                                 fg="blue", wraplength=400)
        self.lbl_path.pack(pady=5)

        btn_path = tk.Button(self.root, text="Выбрать папку для сохранения", font=("Arial", 9),
                             command=self.choose_folder)
        btn_path.pack(fill="x", pady=5)

        # Кнопка создания проекта
        btn_project = tk.Button(self.root, text="Создать проект (Деталь)",
                                font=("Arial", 12), bg="#4CAF50", fg="white", command=self.open_project)
        btn_project.pack(fill="x", pady=5)

        # Кнопка спецификации
        btn_spec = tk.Button(self.root, text="Создать спецификацию",
                             font=("Arial", 12), bg="#FF9800", fg="white", command=self.open_spec)
        btn_spec.pack(fill="x", pady=5)

        # Кнопка смены пользователя
        btn_change_user = tk.Button(self.root, text="Сменить пользователя",
                                    font=("Arial", 10), bg="#e0e0e0", command=self.change_user)
        btn_change_user.pack(fill="x", pady=20)

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку для сохранения файлов")
        if folder:
            self.save_path = folder
            self.user_settings['save_path'] = folder
            # Сохраняем в файл
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump(self.user_settings, f, ensure_ascii=False, indent=2)
            self.lbl_path.config(text=f"Папка: {self.save_path}")

    def change_user(self):
        if messagebox.askyesno("Смена пользователя", "Вы уверены, что хотите сменить данные?"):
            if os.path.exists(self.settings_path):
                os.remove(self.settings_path)
            for widget in self.root.winfo_children():
                widget.destroy()
            self.show_auth_window()

    def open_project(self):
        ProjectWindow(self.root, self.user_settings)

    def open_spec(self):
        header_data = {'name': '', 'code': ''}
        SpecWindow(self.root, header_data, self.user_settings, self.finalize_spec)

    def finalize_spec(self):
        pass


if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()