import pandas as pd
import re
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime


class PhoneBookApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PhoneBook Pro")
        self.root.geometry("1000x600")

        try:
            # Загрузка данных с правильными параметрами
            self.df_fio = pd.read_excel("Данные (ФИО).xlsx")
            self.df_phone = pd.read_excel("Данные (ID, phone).xlsx")

            # Файл сообщений без заголовка
            self.df_messages = pd.read_excel("Текстовые сообщения (1).xlsx", header=None)
            self.df_messages.columns = ['Сообщение']

            self.process_data()
            self.create_widgets()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка загрузки данных: {str(e)}")
            self.root.destroy()

    def process_data(self):
        # Объединение данных
        self.df = pd.merge(self.df_fio, self.df_phone, on="ID", how="left")

        # Обработка сообщений
        inactive_phones = self.extract_inactive_phones()
        self.df["Активен"] = ~self.df["Номер телефона"].astype(str).apply(
            lambda x: re.sub(r'[^\d]', '', x)[-10:] if pd.notna(x) else ''
        ).isin(inactive_phones)

        # Валидация номеров
        self.df["Корректный номер"] = self.df["Номер телефона"].apply(self.validate_phone)

        # Определение типа аккаунта и проверка пола
        self.df["Тип аккаунта"] = self.df.apply(self.detect_account_type, axis=1)

        # Расчет возраста
        current_year = datetime.now().year
        birth_years = pd.to_datetime(self.df["Дата рождения"]).dt.year
        self.df["Возраст"] = current_year - birth_years

    def extract_inactive_phones(self):
        phones = set()
        for msg in self.df_messages['Сообщение'].astype(str):
            # Ищем все возможные форматы номеров
            matches = re.findall(
                r'(?:\+?7|8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}',
                msg
            )
            # Нормализуем номера (оставляем последние 10 цифр)
            for phone in matches:
                cleaned = re.sub(r'[^\d]', '', phone)[-10:]
                if len(cleaned) == 10:
                    phones.add('7' + cleaned)
        return phones

    def validate_phone(self, phone):
        if pd.isna(phone):
            return False
        cleaned = re.sub(r'[^\d]', '', str(phone))
        return len(cleaned) == 11 and cleaned.startswith(('7', '8'))

    def detect_account_type(self, row):
        name = str(row['ФИО'])
        gender = str(row['Пол']).strip().upper()

        # Проверка на семейный аккаунт
        if ' и ' in name:
            return 'Семейный'

        # Проверка пола с учетом возможных вариантов написания
        if gender in ['М', 'МУЖ', 'M', 'МУЖСКОЙ']:
            # Дополнительная проверка по имени для женщин
            if any(ending in name.lower() for ending in ['вна', 'на', 'та', 'ия']):
                return 'Женщина (исправлено)'
            return 'Мужчина'

        if gender in ['Ж', 'ЖЕН', 'F', 'ЖЕНСКИЙ']:
            # Дополнительная проверка по имени для мужчин
            if any(ending in name.lower() for ending in ['вич', 'влы', 'ер', 'ий']):
                return 'Мужчина (исправлено)'
            return 'Женщина'

        # Автоматическое определение по ФИО если пол не указан
        if any(ending in name.lower() for ending in ['вна', 'на', 'та', 'ия']):
            return 'Женщина (автоопределение)'
        if any(ending in name.lower() for ending in ['вич', 'влы', 'ер', 'ий']):
            return 'Мужчина (автоопределение)'

        return 'Не определен'

    def create_widgets(self):
        # Панель поиска
        search_frame = ttk.Frame(self.root)
        search_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT)
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        ttk.Button(search_frame, text="Найти", command=self.search_data).pack(side=tk.LEFT)
        ttk.Button(search_frame, text="Сброс", command=self.reset_search).pack(side=tk.LEFT, padx=5)

        # Таблица данных
        columns = [
            "ID", "ФИО", "Пол", "Дата рождения", "Возраст",
            "Номер телефона", "Корректный номер", "Активен", "Тип аккаунта"
        ]

        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.root, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.W)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Панель статистики
        stats_frame = ttk.Frame(self.root)
        stats_frame.pack(pady=10, fill=tk.X)

        ttk.Button(stats_frame, text="Статистика по возрастам",
                   command=self.show_age_stats).pack(side=tk.LEFT, padx=5)
        ttk.Button(stats_frame, text="Статистика по полу",
                   command=self.show_gender_stats).pack(side=tk.LEFT, padx=5)

        self.update_table()

    def update_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in self.df.iterrows():
            values = [
                row["ID"],
                row["ФИО"],
                row["Пол"],
                row["Дата рождения"].strftime("%d.%m.%Y") if pd.notna(row["Дата рождения"]) else "",
                row["Возраст"],
                row["Номер телефона"],
                "Да" if row["Корректный номер"] else "Нет",
                "Да" if row["Активен"] else "Нет",
                row["Тип аккаунта"]
            ]
            self.tree.insert("", tk.END, values=values)

    def search_data(self):
        query = self.search_entry.get().lower()
        if not query:
            return

        mask = (
                self.df["ФИО"].str.lower().str.contains(query) |
                self.df["Номер телефона"].astype(str).str.contains(query) |
                self.df["Тип аккаунта"].str.lower().str.contains(query)
        )

        self.df = self.df[mask]
        self.update_table()

    def reset_search(self):
        self.process_data()  # Перезагружаем данные
        self.search_entry.delete(0, tk.END)
        self.update_table()

    def show_age_stats(self):
        age_groups = pd.cut(
            self.df["Возраст"],
            bins=[0, 18, 30, 45, 60, 100],
            labels=["0-18", "19-30", "31-45", "46-60", "60+"]
        )

        stats = self.df.groupby(age_groups).size().reset_index(name="Количество")
        stats_window = tk.Toplevel(self.root)
        stats_window.title("Статистика по возрастам")

        tree = ttk.Treeview(stats_window, columns=("Возрастная группа", "Количество"), show="headings")
        tree.heading("Возрастная группа", text="Возрастная группа")
        tree.heading("Количество", text="Количество")

        for _, row in stats.iterrows():
            tree.insert("", tk.END, values=(row[0], row["Количество"]))

        tree.pack(fill=tk.BOTH, expand=True)

    def show_gender_stats(self):
        stats = self.df["Тип аккаунта"].value_counts().reset_index()
        stats.columns = ["Тип аккаунта", "Количество"]

        stats_window = tk.Toplevel(self.root)
        stats_window.title("Статистика по полу")

        tree = ttk.Treeview(stats_window, columns=("Тип аккаунта", "Количество"), show="headings")
        tree.heading("Тип аккаунта", text="Тип аккаунта")
        tree.heading("Количество", text="Количество")

        for _, row in stats.iterrows():
            tree.insert("", tk.END, values=(row["Тип аккаунта"], row["Количество"]))

        tree.pack(fill=tk.BOTH, expand=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = PhoneBookApp(root)
    root.mainloop()