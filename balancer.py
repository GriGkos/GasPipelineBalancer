import sys
import os
import re
import time
import xlwings as xw
import numpy as np
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *

# Вспомогательный класс для виджета приоритетов
class DragDropListWidget(QListWidget):
    # Конструктор
    def __init__(self, items, parent=None):
        super().__init__(parent)
        self.addItems(items) # Добавляем список приоритетов в виджет
        self.setDragDropMode(QListWidget.DragDropMode.InternalMove)
    # Деструктор
    def __del__(self):
        del self

# Подкласс QMainWindow для настройки главного окна приложения
class MainWindow(QMainWindow):
    # Конструктор
    def __init__(self):
        super().__init__()

        # Объявляем  некоторые переменные
        self.setWindowTitle("Балансы")
        self.setWindowIcon(QIcon("images/window_icon.png"))
        self.setMinimumSize(320, 300)
        self.wb = None
        self.file_path = ''
        self.sheet_names = []
        self.zagruzka_sistemy = []
        self.option = 1
        self.count = 0

        # Вызываем метод добавления интерфейса
        self.add_ui()
        
    # Метод для добавления интерфейса
    def add_ui(self):
        # Строка выбора файла
        self.file_path_le = QLineEdit(self)
        self.file_path_le.setPlaceholderText("Выберите файл:")
        # Кнопка для открытия окна выбора файла
        self.browse_button = QPushButton("Browse...", self)
        # Коннектим кнпоку к слоту
        self.browse_button.clicked.connect(self.browse_files)

        # Горизонтальный layout для строки ввода и кнопки
        top_layout = QHBoxLayout()
        top_layout.addWidget(self.file_path_le)
        top_layout.addWidget(self.browse_button)

        # Надпись "Выберите режим:"
        self.option_label = QLabel("Выберите режим:", self)

        # Горизонтальный layout для кнопок выбора режима (год\сутки)
        self.option_layout = QHBoxLayout()
        # Группа кнопок, чтобы объединить кнопки выбора режима (год\сутки)
        self.option_group = QButtonGroup()
        # Кнопка выбора режима (год)
        self.year_check = QRadioButton("Год", self)
        self.year_check.setDisabled(True)
        self.option_group.addButton(self.year_check)
        self.option_layout.addWidget(self.year_check)
        # Кнопка выбора режима (сутки)
        self.day_check = QRadioButton("Сутки", self)
        self.day_check.setDisabled(True)
        self.option_group.addButton(self.day_check)
        self.option_layout.addWidget(self.day_check)    
        # Коннектим кнопки к слоту
        self.option_group.buttonClicked.connect(self.mode_selected)

        # Надпись "Выберите лист:"
        self.sheets_label = QLabel("Выберите лист:", self)

        # Комбобокс для списка листов
        self.sheets_cmb = QComboBox(self)
        self.sheets_cmb.setPlaceholderText("--- Выбрать ---")
        self.sheets_cmb.setDisabled(True)
        self.sheets_cmb.currentTextChanged.connect(self.update_sheet_name)

        # Устанавливаем выбранное в комбобоксе названием рабочего листа
        self.sheet_name = self.sheets_cmb.currentText()

        # Кнопка "Приоритет"
        self.btn_open_priority = QPushButton("Приоритет")
        self.btn_open_priority.setDisabled(True)
        self.btn_open_priority.clicked.connect(self.open_priority)
        
        # Кнопка "Старт"
        self.start_button = QPushButton("Старт", self)
        self.start_button.setDisabled(True)
        self.start_button.clicked.connect(self.before_balans)

        # Надписи, показывающиеся в конце вычислений
        self.calculation_label = QLabel("", self)
        self.time_label = QLabel("", self)

        # Создаем основной layout
        main_layout = QVBoxLayout()
        main_layout.addLayout(top_layout)
        main_layout.addWidget(self.sheets_label)
        main_layout.addWidget(self.sheets_cmb)
        main_layout.addWidget(self.option_label)
        main_layout.addLayout(self.option_layout)
        main_layout.addWidget(self.btn_open_priority)
        main_layout.addWidget(self.start_button)
        main_layout.addWidget(self.calculation_label)
        main_layout.addWidget(self.time_label)

        # Стретч, чтобы все было прижато к верхнему краю
        main_layout.addStretch()

        # Создаем контейнерный виджет
        container = QWidget()
        container.setLayout(main_layout)

        # Устанавливаем контейнер как центральный виджет
        self.setCentralWidget(container)

    # Вызывается при нажатии на кнопку "Browse..."
    def browse_files(self):
        # Открываем окно выбора файла и забираем путь к файлу
        self.file_path, _ = QFileDialog.getOpenFileName(
            None,                           # parent
            "Select Excel File",            # caption
            "",                             # directory
            "Excel Files (*.xlsx *.xlsm);;All Files (*)"  # filter
        )
        
        if self.file_path != "": # Проверяем, выбран ли файл

            if self.wb:
                self.wb.save()
                self.wb.close()
                print("Книга закрыта")

            self.file_path_le.setText(self.file_path) # Устанавливаем в строку около кнопки "Browse..." путь к файлу
            # Открываем выбраный файл
            self.wb = xw.Book(self.file_path) # type: ignore
            self.wb.app.visible = True
            # Паттерн для обрезки лишнего в названии листов 
            pattern = r'<Sheet \[.*?\]'

            # Делаем список листов
            self.sheet_names = list(re.sub(pattern, '', str(elem))[:-1] for elem in self.wb.sheets)
            # Добавляем в комбобокс выбора листа
            self.sheets_cmb.addItems(self.sheet_names)
            self.sheets_cmb.setEnabled(True)

    # Вызывается при выборе одного из режимов
    def mode_selected(self, button):
        if button.text() == "Год":
            try:
                self.option = 1
                self.sheet_prioritet = self.wb.sheets["Приоритет (год.)"] # type: ignore
            except Exception:
                QMessageBox.warning(self, "Ошибка ", "\nНет листа с названием \"Приоритет (год.)\"", QMessageBox.StandardButton.Ok)
                return
        else:
            try:
                self.option = 0
                self.sheet_prioritet = self.wb.sheets["Приоритет (сут.)"] # type: ignore
            except Exception:
                QMessageBox.warning(self, "Ошибка ", "\nНет листа с названием \"Приоритет (сут.)\"", QMessageBox.StandardButton.Ok)
                return

        # Последняя строка списка
        self.last_row = self.sheet_prioritet.range(self.sheet.cells.last_cell.row, 2).end('up').row
        # Диапазон списка подсистем + приоритеты (столбец справа)
        self.zagruzka_range = self.sheet_prioritet.range((1, 2), (self.last_row, 3))
        # Список названий подсистем
        self.podsist_names = [elem[0].replace('\xa0', ' ') for elem in self.zagruzka_range.value]

        # Получаем данные
        self.get_data()

    # Вызывается при выборе одного из листов
    def update_sheet_name(self, text):
        # Устанвливаем название выбранного рабочего листа 
        self.sheet_name = text
        # Открываем этот лист
        self.sheet = self.wb.sheets[self.sheet_name] # type: ignore

        if self.option_group.checkedButton():
            # Последняя строка списка
            self.last_row = self.sheet_prioritet.range(self.sheet.cells.last_cell.row, 2).end('up').row
            # Диапазон списка подсистем + приоритеты (столбец справа)
            self.zagruzka_range = self.sheet_prioritet.range((1, 2), (self.last_row, 3))
            # Список названий подсистем
            self.podsist_names = [elem[0].replace('\xa0', ' ') for elem in self.zagruzka_range.value]

            # Получаем данные
            self.get_data()

        # Разворачиваем скрытые ячейки
        self.sheet.api.Outline.ShowLevels(RowLevels=6)
        self.sheet.api.Outline.ShowLevels(ColumnLevels=3)

        self.year_check.setEnabled(True)
        self.day_check.setEnabled(True)
    
    # Вызывается при нажатии на кнопку "Приоритет"
    def open_priority(self):
        # Создаем диалоговое окно приоритетов
        self.dialog = QDialog(self)
        self.dialog.setWindowTitle('Список приоритетов')
        dialog_layout = QVBoxLayout()

        # Добавляем виджет списка в диалоговое окно
        self.drag_drop_list = DragDropListWidget(self.zagruzka_sistemy)
        dialog_layout.addWidget(self.drag_drop_list)

        # Кнопки Ok и Cancel
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.update_items_order)
        button_box.rejected.connect(self.dialog.reject)
        dialog_layout.addWidget(button_box)

        self.dialog.setLayout(dialog_layout)
        # Открываем диалоговое окно
        self.dialog.exec()

    # Вызывается при нажатии на кнопку Ok в окне приоритетов
    def update_items_order(self):
        # Обновляет список приоритетов
        self.items = [self.drag_drop_list.item(i).text() for i in range(self.drag_drop_list.count())] # type: ignore
        self.update_list()
        self.dialog.accept()
        
    # Обновляет порядок загрузки системы
    def update_list(self):
        self.zagruzka_sistemy.clear()
        self.zagruzka_sistemy = self.items.copy()
        self.update_prior_excel()

    # Обновляет приоритеты в таблице
    def update_prior_excel(self):
        new_dct= {key:value for value, key in enumerate(self.zagruzka_sistemy, start=1)}
        new_prioritety = [[name, new_dct[name]] if name in new_dct.keys() else [name, prior] for name, prior in self.prioritety]
        self.sheet_prioritet.range("B1").value = new_prioritety

    # Получение данных
    def get_data(self):
        self.btn_open_priority.setDisabled(True)
        self.start_button.setDisabled(True)

        # Определяем диапазон для поиска названий подсистем
        self.last_row = self.sheet.range(self.sheet.cells.last_cell.row, 1).end('up').row
        self.search_range = self.sheet.range((1, 1), (self.last_row, 8))

        ######
        ### Находим все строки коэффициентов поступления и распределения для балансировки
        ######

        # Создаем словарь, где {название под-мы: ее строка в таблице}
        self.podsist_row = {}
        for elem in self.podsist_names:
            found_cell = self.search_range.api.Find(elem)
            if found_cell is not None:
                self.podsist_row[elem] = found_cell.Row

        # Сортируем подсистемы в соответствии с приоритетом
        # Если нет приоритета - не берем 
        self.prioritety = [[elem1.replace('\xa0', ' '), elem2] for elem1, elem2 in self.zagruzka_range.value]
        self.zagruzka_sistemy = sorted([elem for elem in self.prioritety if elem[1] is not None \
                              and elem[0] in self.podsist_row.keys()], key=lambda x: x[1])
        self.zagruzka_sistemy = [key for key, value in self.zagruzka_sistemy]

        ######	
        ### Находим все дисбалансы по системам 
        ######

        found_disbalans = self.search_range.api.Find("Дисбаланс", LookIn=-4163)  # -4163 соответствует xlValues

        # Список для хранения содержимого найденных ячеек
        self.dibalans_row = []

        # Получить содержимое найденных ячеек
        if found_disbalans is not None:
            first_found_cell = found_disbalans
            while True:
                self.dibalans_row.append(found_disbalans.Row)
                found_disbalans = self.search_range.api.FindNext(found_disbalans)
                if found_disbalans.Address == first_found_cell.Address:	 # условие, чтобы не закцикливаться
                    break
            
        ######
        ### Находим все СН по КС
        ######

        self.sn_cs_cells = self.search_range.api.Find("СН КС", LookIn=-4163)  # -4163 соответствует xlValues

        # Список для хранения содержимого найденных ячеек
        self.sn_cs = []

        # Получить содержимое найденных ячеек
        if self.sn_cs_cells is not None:
            first_found_cell = self.sn_cs_cells
            while True:
                self.sn_cs.append(self.sn_cs_cells.Row)
                self.sn_cs_cells = self.search_range.api.FindNext(self.sn_cs_cells)
                if self.sn_cs_cells.Address == first_found_cell.Address:		# условие, чтобы не закцикливаться
                    break

        print("Коэффициенты найдены!")
        self.btn_open_priority.setEnabled(True)
        self.start_button.setEnabled(True)

    def before_balans(self):
        self.lap1_t = time.time()
        self.time_label.setText("")

        # Получение значений из диапазона "CP3"
        cp3_value = self.sheet.range("CP3").value

        self.first_coef = list(self.podsist_row.values())[0]
        self.years_last_col = self.sheet.range(self.first_coef, self.sheet.cells.last_cell.column).end('left').column

        if self.option:
            mask = np.array([i % 5 != 4 for i in range(self.years_last_col-8)])
        else:
            mask = np.zeros_like([False]*(self.years_last_col-8))

        # Обработка строк
        for element in self.sn_cs:
            row_values, row_formulas = self.read_row_values_and_formulas(self.sheet, element, self.years_last_col)
            updated_values = self.update_row_values(self.sheet, element, row_values, cp3_value, self.years_last_col)
            self.write_row_values(self.sheet, element, self.years_last_col, updated_values, row_formulas, mask)

        # Приводим коэффициенты подсистем в изначальное состояние (0)
        for elem in self.podsist_row.values():
            row_values, row_formulas = self.read_row_values_and_formulas(self.sheet, elem, self.years_last_col)
            row_values = np.zeros_like(row_values)
            self.write_row_values(self.sheet, elem, self.years_last_col, row_values, row_formulas, mask)

        print("Приведение значений в изначальное состояние закончено!")

        self.lap2_t = time.time()

        self.balans(mask)

    def balans(self, mask):
        
        # Функция балансировки
        def balans_god(sheet, disbalans_podsistema, postuplenie_v_podsistemy):
            
            target_range = sheet.range((disbalans_podsistema, 9), (disbalans_podsistema, self.years_last_col)) # Таргет - дисбаланс
            changing_range = sheet.range((self.podsist_row[postuplenie_v_podsistemy], 9), (self.podsist_row[postuplenie_v_podsistemy], self.years_last_col))
            
            postuplenie_formulas = changing_range.formula[0]
            
            # Если это коэффициент, то считаем его методом Ньютона
            if 'коэф.' in postuplenie_v_podsistemy.lower():
                self.goal_seek(target_range, changing_range, mask=mask)
                res = np.array(changing_range.value, dtype=object)
                res[mask] = np.clip(res[mask], 0, 1)
                changing_range.value = res

            # Если не коэффициент, а распределение газа, то просто сбрасываем столько, сколько надо (весь дисбаланс, если положительный,
            # но отрицательным он и не должен получаться)
            elif 'сброс' in postuplenie_v_podsistemy.lower():
                clipped_values = np.array(target_range.value, dtype=object)
                clipped_values[mask] = np.where(clipped_values[mask] < 0, 0, clipped_values[mask])
                changing_range.value = clipped_values
                
            elif 'поступление' in postuplenie_v_podsistemy.lower():
                reversed_values = np.array(target_range.value, dtype=object)
                reversed_values[mask] = np.where(reversed_values[mask] < 0, (-1)*reversed_values[mask], reversed_values[mask])
                changing_range.value = reversed_values

            changing_values = np.array(changing_range.value)
            
            # Записываем обновленные значения обратно в строки
            self.write_row_values(sheet, self.podsist_row[postuplenie_v_podsistemy], self.years_last_col, changing_values, postuplenie_formulas, mask)

        # Ставим в соответствие каждому коэффициенту дисбаланс, за который он отвечает
        # Словарь {номер строки коэф-та: номер строки дисбаланса подсистемы}
        postup_to_disbalans = {}
        for key, value in self.podsist_row.items():
            # Ищем минимальный элемент в dibalans_row, который больше значения словаря
            min_element = next((x for x in self.dibalans_row if x > value), None)
            postup_to_disbalans[value] = min_element

        # В соответствии с очередностью загрузки, начинаем балансировать
        for postup in self.zagruzka_sistemy:
            balans_god(self.sheet, postup_to_disbalans[self.podsist_row[postup]], postup)

        self.end_t = time.time()
        self.calculation_label.setText("Готово!")
        self.time_label.setText(f"Время приведения коэфф-ов в нач. сост. = {(self.lap2_t - self.lap1_t)/60:.{3}f} мин.\nВремя расчета коэффициентов = {(self.end_t - self.lap2_t)/60:.{3}f} мин.\nВсе время = {(self.end_t - self.lap1_t)/60:.{3}f} мин.")
        
    # Функция для преобразования строки в число, если возможно, иначе возвращает 0
    def convert_to_number_or_zero(self, elem):
        try:
            return float(elem)  # Пробуем преобразовать элемент в число
        except ValueError:
            return 0  # Если преобразование не удалось, возвращаем 0

    # Чтение значений и формул из строк
    def read_row_values_and_formulas(self, sheet, row, last_col):
        # Чтение формул из строки
        row_formulas = np.array(sheet.range((row, 9), (row, last_col)).formula)[0]
        # Векторизуем функцию для применения к массиву
        vectorized_conversion = np.vectorize(self.convert_to_number_or_zero)
        row_values = vectorized_conversion(row_formulas)
        return row_values, row_formulas

    # Запись значений по строкам в таблицу
    def write_row_values(self, sheet, row, last_col, values, formulas, mask):
        range_to_write = sheet.range((row, 9), (row, last_col))
        values = np.where(mask, values, formulas)
        range_to_write.value = values

    # Обновление значений строки
    def update_row_values(self, sheet, row, row_values, cp3_value, last_col):
        if cp3_value == 0:
            row_values = np.zeros_like(row_values)
        elif cp3_value == 1:
            next_row = np.array(sheet.range((row + 1, 9), (row + 1, last_col)).value) # type: ignore
            vectorized_conversion = np.vectorize(self.convert_to_number_or_zero)
            next_row = vectorized_conversion(next_row)
            row_values = np.where(next_row > 0, next_row * 0.0037, 0)
        return row_values
    
    # Функция поиска решений (метод Ньютона)
    # x_{n+1} = x_{n} - f(x_{n})/f'(x_{n})
    def goal_seek(self, target_range, changing_range, target_value=None, tolerance=None, max_iter=100, mask=None):
        iter_count = 0
        target_value = np.array([0]*(self.years_last_col-8))
        tolerance = np.array([1e-6]*(self.years_last_col-8))
        
        while iter_count < max_iter:
            current_value = np.array(target_range.value)
            # Если достигли нужной точности - выходим из функции
            if np.all(abs(current_value[mask] - target_value[mask]) < tolerance[mask]):
                return
            
            # Немного изменяем коэффициент, чтобы посчитать производную
            delta = np.full(len(current_value), 1e-6)
            initial_changing_value = np.array(changing_range.value, dtype=object)
            # Создаем копию массива и вносим изменения только в маскированные элементы
            initial_changing_value[mask] += delta[mask]
            changing_range.value = initial_changing_value
            new_value = np.array(target_range.value)
            initial_changing_value[mask] -= delta[mask]
            changing_range.value = initial_changing_value  # Возвращаем обратно

            # Если производная 0, значит это не коэф, а поступление или сброс,
            # просто сбрасываем сюда дисбаланс
            derivative = (new_value[mask] - current_value[mask]) / delta[mask]
            if not np.all(derivative):
                clipped_values = np.array(target_range.value, dtype=object)
                clipped_values[mask] = np.where(clipped_values[mask] < 0, 0, clipped_values[mask])
                changing_range.value = clipped_values
                return
            
            # Делаем шаг по формуле x_{n+1} = x_{n} - f(x_{n})/f'(x_{n})
            step = (target_value[mask] - current_value[mask]) / derivative
            initial_changing_value[mask] += step
            changing_range.value = initial_changing_value.tolist()
        
            iter_count += 1
        print(f"!!!Превышено кол-во итераций!!! {changing_range}")

    # Деструктор
    def __del__(self):
        # Если книга открыта, то закрыть
        if self.wb:
            self.wb.save()
            self.wb.close()
            print("Книга закрыта")

# Просто мейн. Все, что ниже, менять не надо
def main():
    app = QApplication(sys.argv)

    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()