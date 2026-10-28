import subprocess
import os
from tkinter import messagebox

def convert_to_pdf(xlsx_path):
    """Конвертирует Excel в PDF через LibreOffice"""
    if not os.path.exists(xlsx_path):
        messagebox.showerror("Ошибка", "Файл Excel не найден!")
        return None

    try:
        # Вызываем LibreOffice в фоновом режиме
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf',
            '--outdir', os.path.dirname(xlsx_path),
            xlsx_path
        ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        pdf_path = xlsx_path.replace('.xlsx', '.pdf')
        messagebox.showinfo("Успех", f"PDF создан!\n\nПуть: {pdf_path}")
        return pdf_path
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось создать PDF:\n{e}")
        return None