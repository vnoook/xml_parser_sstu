import sys
import xml.etree.ElementTree as ET
import openpyxl
import PyQt5
import PyQt5.QtWidgets
import PyQt5.QtCore
import PyQt5.QtGui

# класс главного окна
class WindowMain(PyQt5.QtWidgets.QMainWindow):
    """Класс главного окна"""
    # описание главного окна
    def __init__(self):
        super().__init__()

        # переменные
        self.info_extention_open_file_xml = 'Файлы XML (*.xml)'
        self.info_path_open_file = None

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Парсер xml файлов для ССТУ')
        self.setGeometry(450, 100, 700, 180)
        self.setWindowFlags(PyQt5.QtCore.Qt.WindowStaysOnTopHint)

        # ОБЪЕКТЫ НА ФОРМЕ
        # label_select_file
        self.label_select_file = PyQt5.QtWidgets.QLabel(self)
        self.label_select_file.setObjectName('label_select_file')
        self.label_select_file.setText('Выберите файл XML')
        self.label_select_file.setGeometry(PyQt5.QtCore.QRect(10, 10, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_select_file.setFont(font)
        self.label_select_file.adjustSize()
        self.label_select_file.setToolTip(self.label_select_file.objectName())

        # toolButton_select_file_xml
        self.toolButton_select_file_xml = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_file_xml.setObjectName('toolButton_select_file_xml')
        self.toolButton_select_file_xml.setText('...')
        self.toolButton_select_file_xml.setGeometry(PyQt5.QtCore.QRect(10, 40, 50, 20))
        self.toolButton_select_file_xml.setFixedWidth(50)
        self.toolButton_select_file_xml.clicked.connect(self.select_file_xml)
        self.toolButton_select_file_xml.setToolTip(self.toolButton_select_file_xml.objectName())

        # label_path_file
        self.label_path_file = PyQt5.QtWidgets.QLabel(self)
        self.label_path_file.setObjectName('label_path_file')
        self.label_path_file.setEnabled(False)
        self.label_path_file.setText('')
        self.label_path_file.setGeometry(PyQt5.QtCore.QRect(10, 70, 400, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_path_file.setFont(font)
        self.label_path_file.adjustSize()
        self.label_path_file.setToolTip(self.label_path_file.objectName())

        # pushButton_parse_to_xls
        self.pushButton_parse_to_xls = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_parse_to_xls.setObjectName('pushButton_parse_to_xls')
        self.pushButton_parse_to_xls.setEnabled(False)
        self.pushButton_parse_to_xls.setText('Конвертировать файл в XLS')
        self.pushButton_parse_to_xls.setGeometry(PyQt5.QtCore.QRect(10, 100, 160, 25))
        # self.pushButton_parse_to_xls.setFixedWidth(130)
        self.pushButton_parse_to_xls.clicked.connect(self.parse_xml)
        self.pushButton_parse_to_xls.setToolTip(self.pushButton_parse_to_xls.objectName())

        # EXIT
        # button_exit
        self.button_exit = PyQt5.QtWidgets.QPushButton(self)
        self.button_exit.setObjectName('button_exit')
        self.button_exit.setText('Выход')
        self.button_exit.setGeometry(PyQt5.QtCore.QRect(10, 140, 180, 25))
        self.button_exit.setFixedWidth(50)
        self.button_exit.clicked.connect(self.click_on_btn_exit)
        self.button_exit.setToolTip(self.button_exit.objectName())

    # событие - нажатие на кнопку выбора файла
    def select_file_xml(self):
        # переменная для хранения информации из окна выбора файла
        data_of_open_file_name = None

        # запоминание старого значения пути выбора файлов
        old_path_of_selected_xml_file = self.label_path_file.text()

        # открытие окна выбора файла
        self.info_for_open_file = 'Выберите XML файл (.XML)'
        # непосредственное окно выбора файла и переменная для хранения пути файла
        data_of_open_file_name = PyQt5.QtWidgets.QFileDialog.getOpenFileName(self,
                                                                             self.info_for_open_file,
                                                                             self.info_path_open_file,
                                                                             self.info_extention_open_file_xml)

        # выбор только пути файла из data_of_open_file_name
        file_name = data_of_open_file_name[0]

        # # нажата кнопка выбора XLS файла
        # if self.sender().objectName() == self.toolButton_select_xls_file.objectName():
        #     if file_name == '':
        #         self.label_path_xls_file.setText(old_path_of_selected_xls_file)
        #         self.label_path_xls_file.adjustSize()
        #     else:
        #         old_path_of_selected_xls_file = self.label_path_xls_file.text()
        #         self.label_path_xls_file.setText(file_name)
        #         self.label_path_xls_file.adjustSize()
        #
        # # активация и деактивация объектов на форме зависящее от 'выбраны ли все файлы' и 'они разные'
        # if self.text_empty_path_file not in (self.label_path_html_file.text(), self.label_path_xls_file.text()):
        #     self.pushButton_send_mail.setEnabled(True)
        #     self.pushButton_send_test_mail.setEnabled(True)
        #     self.lineEdit_subject_letter.setEnabled(True)










    def parse_xml(self):
        file_xml = 'guid.xml'
        file_xls = 'guid.xlsx'
        list_xml = []

        wb = openpyxl.Workbook()
        wb_s = wb.active

        wb_s.append(["ID", "Name"])

        root_node = ET.parse(file_xml).getroot()

        for tag in root_node.findall('Department'):
            id_value = tag.get('ID')
            if not id_value:
                id_value = 'UNKNOWN DATA'
                print(id_value)

            name_value = tag.get('Name')
            if not name_value:
                name_value = 'UNKNOWN DATA'
                print(name_value)

            wb_s.append([id_value, name_value])

        wb.save(file_xls)
        wb.close()

    # событие - нажатие на кнопку Выход
    @staticmethod
    def click_on_btn_exit():
        sys.exit()


# создание основного окна
def main_app():
    app = PyQt5.QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    app_window_main = WindowMain()
    app_window_main.show()
    sys.exit(app.exec_())


# запуск основного окна
if __name__ == '__main__':
    main_app()
