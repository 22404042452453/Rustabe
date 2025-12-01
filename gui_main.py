#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
from datetime import datetime
from analysis_core import MultiScenarioAnalyzer
import threading
import queue

class RastrAnalysisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализ устойчивости энергосистемы - RastrWin")
        self.root.geometry("1000x700")

        # Переменные для хранения данных
        self.generators = []  # Список генераторов
        self.selected_parameters = {
            'delta': tk.BooleanVar(value=True),  # Угол ротора
            'voltage': tk.BooleanVar(value=False),  # Напряжение
            'power_p': tk.BooleanVar(value=True),  # Активная мощность
            'power_q': tk.BooleanVar(value=False),  # Реактивная мощность
            'current': tk.BooleanVar(value=False)  # Ток
        }
        self.scenario_folder = tk.StringVar()
        self.calculation_method = tk.StringVar(value="simple")
        self.net_file = tk.StringVar(value="Рем. СГРЭС-Компр.rst")

        # Очередь для обновления GUI из потока анализа
        self.progress_queue = queue.Queue()

        self.create_widgets()
        self.setup_layout()

    def create_widgets(self):
        """Создание всех виджетов интерфейса"""

        # === Настройки генераторов ===
        self.generators_frame = ttk.LabelFrame(self.root, text="Конфигурация генераторов", padding=10)
        self.create_generators_section()

        # === Параметры мониторинга ===
        self.parameters_frame = ttk.LabelFrame(self.root, text="Параметры для мониторинга", padding=10)
        self.create_parameters_section()

        # === Настройки сценариев ===
        self.scenarios_frame = ttk.LabelFrame(self.root, text="Сценарии анализа", padding=10)
        self.create_scenarios_section()

        # === Настройки расчета ===
        self.calculation_frame = ttk.LabelFrame(self.root, text="Настройки расчета", padding=10)
        self.create_calculation_section()

        # === Управление ===
        self.control_frame = ttk.Frame(self.root, padding=10)
        self.create_control_section()

        # === Прогресс и результаты ===
        self.results_frame = ttk.LabelFrame(self.root, text="Результаты анализа", padding=10)
        self.create_results_section()

    def create_generators_section(self):
        """Создание секции конфигурации генераторов"""

        # Таблица генераторов
        columns = ("ID", "Имя", "Мин P", "Макс P")
        self.generators_tree = ttk.Treeview(self.generators_frame, columns=columns, show="headings", height=5)

        for col in columns:
            self.generators_tree.heading(col, text=col)
            self.generators_tree.column(col, width=100)

        # Scrollbar для таблицы
        scrollbar = ttk.Scrollbar(self.generators_frame, orient=tk.VERTICAL, command=self.generators_tree.yview)
        self.generators_tree.configure(yscroll=scrollbar.set)

        # Кнопки управления генераторами
        btn_frame = ttk.Frame(self.generators_frame)
        ttk.Button(btn_frame, text="Добавить генератор", command=self.add_generator).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Удалить выбранного", command=self.remove_generator).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Очистить все", command=self.clear_generators).pack(side=tk.LEFT, padx=5)

        # Размещение
        self.generators_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        # Добавляем тестовые данные
        self.add_generator_defaults()

    def create_parameters_section(self):
        """Создание секции выбора параметров"""

        ttk.Checkbutton(self.parameters_frame, text="Угол ротора (Delta)",
                        variable=self.selected_parameters['delta']).pack(anchor=tk.W)
        ttk.Checkbutton(self.parameters_frame, text="Напряжение (U)",
                        variable=self.selected_parameters['voltage']).pack(anchor=tk.W)
        ttk.Checkbutton(self.parameters_frame, text="Активная мощность (P)",
                        variable=self.selected_parameters['power_p']).pack(anchor=tk.W)
        ttk.Checkbutton(self.parameters_frame, text="Реактивная мощность (Q)",
                        variable=self.selected_parameters['power_q']).pack(anchor=tk.W)
        ttk.Checkbutton(self.parameters_frame, text="Ток (I)",
                        variable=self.selected_parameters['current']).pack(anchor=tk.W)

    def create_scenarios_section(self):
        """Создание секции выбора сценариев"""

        # Поле для пути к папке
        path_frame = ttk.Frame(self.scenarios_frame)
        ttk.Entry(path_frame, textvariable=self.scenario_folder, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="Выбрать папку", command=self.select_scenario_folder).pack(side=tk.RIGHT, padx=(5, 0))

        # Список найденных сценариев
        self.scenarios_listbox = tk.Listbox(self.scenarios_frame, height=4)
        scenarios_scrollbar = ttk.Scrollbar(self.scenarios_frame, orient=tk.VERTICAL, command=self.scenarios_listbox.yview)
        self.scenarios_listbox.configure(yscroll=scenarios_scrollbar.set)

        # Файл сети
        net_frame = ttk.Frame(self.scenarios_frame)
        ttk.Label(net_frame, text="Файл сети:").pack(side=tk.LEFT)
        ttk.Entry(net_frame, textvariable=self.net_file, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(net_frame, text="Выбрать файл", command=self.select_net_file).pack(side=tk.RIGHT, padx=(5, 0))

        # Размещение
        path_frame.pack(fill=tk.X, pady=(0, 5))
        self.scenarios_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=(0, 5))
        scenarios_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        net_frame.pack(fill=tk.X)

    def create_calculation_section(self):
        """Создание секции настроек расчета"""

        ttk.Radiobutton(self.calculation_frame, text="Простой расчет",
                        variable=self.calculation_method, value="simple").pack(anchor=tk.W)
        ttk.Radiobutton(self.calculation_frame, text="EMS режим",
                        variable=self.calculation_method, value="ems").pack(anchor=tk.W)

    def create_control_section(self):
        """Создание секции управления"""

        ttk.Button(self.control_frame, text="Запустить анализ",
                  command=self.start_analysis, style="Accent.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(self.control_frame, text="Остановить", command=self.stop_analysis).pack(side=tk.LEFT)
        ttk.Button(self.control_frame, text="Экспорт результатов", command=self.export_results).pack(side=tk.RIGHT)

    def create_results_section(self):
        """Создание секции результатов"""

        # Прогресс бар
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.results_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        # Статус
        self.status_var = tk.StringVar(value="Готов к работе")
        ttk.Label(self.results_frame, textvariable=self.status_var).pack(anchor=tk.W)

        # Область результатов
        results_frame = ttk.Frame(self.results_frame)
        self.results_text = tk.Text(results_frame, height=15, wrap=tk.WORD)

        # Scrollbar для текста
        text_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscroll=text_scrollbar.set)

        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        results_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

    def setup_layout(self):
        """Настройка размещения виджетов"""

        # Левая колонка
        left_frame = ttk.Frame(self.root)
        self.generators_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.parameters_frame.pack(fill=tk.X, pady=(0, 10))
        self.scenarios_frame.pack(fill=tk.BOTH, expand=True)

        # Правая колонка
        right_frame = ttk.Frame(self.root)
        self.calculation_frame.pack(fill=tk.X, pady=(0, 10))
        self.control_frame.pack(fill=tk.X, pady=(0, 10))
        self.results_frame.pack(fill=tk.BOTH, expand=True)

        # Разделение на колонки
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 5), pady=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 10), pady=10)

    def add_generator_defaults(self):
        """Добавление генераторов по умолчанию для демонстрации"""
        defaults = [
            {"id": "51804050", "name": "Ген-1", "p_min": 80, "p_max": 210},
            {"id": "51804051", "name": "Ген-2", "p_min": 100, "p_max": 250}
        ]

        for gen in defaults:
            self.generators_tree.insert("", tk.END, values=(gen["id"], gen["name"], gen["p_min"], gen["p_max"]))
            self.generators.append(gen)

    def add_generator(self):
        """Добавление нового генератора"""
        # Диалог для ввода данных генератора
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить генератор")
        dialog.geometry("300x280")
        dialog.resizable(False, False)

        # Поля ввода
        ttk.Label(dialog, text="ID генератора:").pack(pady=(10, 0))
        id_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=id_var).pack(fill=tk.X, padx=20)

        ttk.Label(dialog, text="Имя:").pack(pady=(10, 0))
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var).pack(fill=tk.X, padx=20)

        ttk.Label(dialog, text="Мин. мощность (МВт):").pack(pady=(10, 0))
        p_min_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=p_min_var).pack(fill=tk.X, padx=20)

        ttk.Label(dialog, text="Макс. мощность (МВт):").pack(pady=(10, 0))
        p_max_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=p_max_var).pack(fill=tk.X, padx=20)

        def save_generator():
            try:
                gen_id = id_var.get().strip()
                name = name_var.get().strip() or f"Ген-{gen_id}"
                p_min = float(p_min_var.get())
                p_max = float(p_max_var.get())

                if not gen_id:
                    messagebox.showerror("Ошибка", "Введите ID генератора")
                    return

                if p_min >= p_max:
                    messagebox.showerror("Ошибка", "Мин. мощность должна быть меньше макс.")
                    return

                # Добавляем в таблицу и список
                self.generators_tree.insert("", tk.END, values=(gen_id, name, p_min, p_max))
                self.generators.append({
                    "id": gen_id,
                    "name": name,
                    "p_min": p_min,
                    "p_max": p_max
                })

                dialog.destroy()

            except ValueError:
                messagebox.showerror("Ошибка", "Введите корректные числовые значения")

        ttk.Button(dialog, text="Добавить", command=save_generator).pack(pady=10)

    def remove_generator(self):
        """Удаление выбранного генератора"""
        selection = self.generators_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите генератор для удаления")
            return

        for item in selection:
            values = self.generators_tree.item(item, "values")
            gen_id = values[0]

            # Удаляем из дерева и списка
            self.generators_tree.delete(item)
            self.generators = [g for g in self.generators if g["id"] != gen_id]

    def clear_generators(self):
        """Очистка всех генераторов"""
        self.generators_tree.delete(*self.generators_tree.get_children())
        self.generators.clear()

    def select_scenario_folder(self):
        """Выбор папки со сценариями"""
        folder = filedialog.askdirectory(title="Выберите папку со сценариями")
        if folder:
            self.scenario_folder.set(folder)
            self.scan_scenarios(folder)

    def select_net_file(self):
        """Выбор файла сети"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл сети",
            filetypes=[("RastrWin файлы", "*.rst"), ("Все файлы", "*.*")]
        )
        if file_path:
            self.net_file.set(file_path)

    def scan_scenarios(self, folder):
        """Сканирование папки на наличие .scn файлов"""
        self.scenarios_listbox.delete(0, tk.END)

        if not os.path.exists(folder):
            return

        scn_files = [f for f in os.listdir(folder) if f.endswith('.scn')]
        scn_files.sort()

        for file in scn_files:
            self.scenarios_listbox.insert(tk.END, file)

    def start_analysis(self):
        """Запуск анализа"""
        if not self.generators:
            messagebox.showerror("Ошибка", "Добавьте хотя бы один генератор")
            return

        if not self.scenario_folder.get():
            messagebox.showerror("Ошибка", "Выберите папку со сценариями")
            return

        # Получаем список сценариев
        scenarios = list(self.scenarios_listbox.get(0, tk.END))
        if not scenarios:
            messagebox.showerror("Ошибка", "В выбранной папке нет .scn файлов")
            return

        # Получаем выбранные параметры
        selected_params = [param for param, var in self.selected_parameters.items() if var.get()]

        if not selected_params:
            messagebox.showerror("Ошибка", "Выберите хотя бы один параметр для мониторинга")
            return

        # Очищаем результаты
        self.results_text.delete(1.0, tk.END)
        self.progress_var.set(0)
        self.status_var.set("Запуск анализа...")

        # Инициализируем анализатор если нужно
        if not hasattr(self, 'analyzer') or self.analyzer is None:
            try:
                self.analyzer = MultiScenarioAnalyzer()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось инициализировать анализатор: {e}")
                return

        # Запускаем анализ в отдельном потоке
        analysis_thread = threading.Thread(
            target=self.run_analysis,
            args=(scenarios, selected_params),
            daemon=True
        )
        analysis_thread.start()

        # Запускаем обновление GUI
        self.root.after(100, self.check_progress)

    def run_analysis(self, scenarios, selected_params):
        """Выполнение анализа в отдельном потоке"""
        try:
            # Конфигурация генераторов в формате анализатора
            gen_config = []
            for gen in self.generators:
                gen_config.append({
                    "table": "Generator",
                    "column": "P",
                    "key": f"Num = {gen['id']}",
                    "name": gen["name"]
                })

            power_ranges = [(gen["p_min"], gen["p_max"]) for gen in self.generators]

            # Запуск анализа
            results = self.analyzer.run_analysis(
                net_file=self.net_file.get(),
                scenario_files=scenarios,
                scenario_folder=self.scenario_folder.get(),
                generators_config=gen_config,
                power_ranges=power_ranges,
                calculate_func=self.calculation_method.get(),
                selected_params=selected_params,
                progress_callback=self.update_progress
            )

            # Отправляем результаты в GUI
            self.progress_queue.put(("completed", results))

        except Exception as e:
            self.progress_queue.put(("error", str(e)))

    def update_progress(self, message, progress=None):
        """Callback для обновления прогресса"""
        self.progress_queue.put(("progress", (message, progress)))

    def check_progress(self):
        """Проверка очереди обновлений от потока анализа"""
        try:
            while True:
                msg_type, data = self.progress_queue.get_nowait()

                if msg_type == "progress":
                    message, progress = data
                    self.status_var.set(message)
                    if progress is not None:
                        self.progress_var.set(progress)

                elif msg_type == "completed":
                    self.status_var.set("Анализ завершен")
                    self.progress_var.set(100)
                    self.display_results(data)

                elif msg_type == "error":
                    self.status_var.set(f"Ошибка: {data}")
                    messagebox.showerror("Ошибка анализа", data)

        except queue.Empty:
            pass

        # Продолжаем проверку
        if self.status_var.get() not in ["Анализ завершен", "Готов к работе"]:
            self.root.after(100, self.check_progress)

    def display_results(self, results):
        """Отображение результатов анализа"""
        self.current_results = results  # Сохраняем результаты для экспорта
        self.results_text.delete(1.0, tk.END)

        for scenario, scenario_results in results.items():
            self.results_text.insert(tk.END, f"Сценарий: {scenario}\n")
            self.results_text.insert(tk.END, "="*50 + "\n")

            for result in scenario_results:
                status = "Устойчиво" if result['stable'] else "Неустойчиво"
                self.results_text.insert(tk.END, f"Мощности: {result['powers']} МВт - {status}\n")

                if 'calc_time' in result:
                    self.results_text.insert(tk.END, f"Время расчета: {result['calc_time']:.2f} с\n")
                if 'comment' in result and result['comment']:
                    self.results_text.insert(tk.END, f"Комментарий: {result['comment']}\n")

                # Дополнительные параметры
                if 'parameters' in result:
                    for param_name, param_values in result['parameters'].items():
                        self.results_text.insert(tk.END, f"{param_name}: {param_values}\n")

                self.results_text.insert(tk.END, "\n")

            self.results_text.insert(tk.END, "\n")

    def stop_analysis(self):
        """Остановка анализа"""
        # TODO: реализовать остановку анализа
        self.status_var.set("Остановка анализа...")

    def export_results(self):
        """Экспорт результатов"""
        if not hasattr(self, 'current_results') or not self.current_results:
            messagebox.showwarning("Предупреждение", "Нет результатов для экспорта")
            return

        # Диалог выбора формата экспорта
        export_dialog = tk.Toplevel(self.root)
        export_dialog.title("Экспорт результатов")
        export_dialog.geometry("400x200")
        export_dialog.resizable(False, False)

        # Выбор формата
        ttk.Label(export_dialog, text="Выберите формат экспорта:").pack(pady=(20, 10))

        format_var = tk.StringVar(value="excel")
        ttk.Radiobutton(export_dialog, text="Excel (.xlsx)", variable=format_var, value="excel").pack(anchor=tk.W, padx=50)
        ttk.Radiobutton(export_dialog, text="CSV (.csv)", variable=format_var, value="csv").pack(anchor=tk.W, padx=50)
        ttk.Radiobutton(export_dialog, text="JSON (.json)", variable=format_var, value="json").pack(anchor=tk.W, padx=50)

        def do_export():
            try:
                # Генерируем имя файла с timestamp
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                if format_var.get() == "excel":
                    filename = f"results_{timestamp}.xlsx"
                    self.analyzer.export_to_excel(self.current_results, self.generators, filename)
                elif format_var.get() == "csv":
                    filename = f"results_{timestamp}.csv"
                    self.analyzer.export_to_csv(self.current_results, self.generators, filename)
                elif format_var.get() == "json":
                    filename = f"results_{timestamp}.json"
                    self.analyzer.export_to_json(self.current_results, self.generators, filename)

                messagebox.showinfo("Экспорт завершен", f"Результаты экспортированы в файл: {filename}")
                export_dialog.destroy()

            except Exception as e:
                messagebox.showerror("Ошибка экспорта", f"Не удалось экспортировать результаты: {e}")

        ttk.Button(export_dialog, text="Экспортировать", command=do_export).pack(pady=20)


def main():
    root = tk.Tk()

    # Настройка стилей
    style = ttk.Style()
    style.configure("Accent.TButton", font=("Arial", 10, "bold"))

    app = RastrAnalysisGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
