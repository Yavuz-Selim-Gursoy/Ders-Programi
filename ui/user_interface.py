# ================================================
from PyQt5 import QtWidgets, QtGui
from pathlib import Path
from typing import TypeVar
from PyQt5.QtCore import pyqtSignal
from src.setget import set_percentage, set_column_name, get_percentage, get_column_names

import subprocess
import win32com.client as win32
import os
import shutil
import sys
import runpy

ROOT_DIR = Path(__file__).parent.parent.absolute()
DATA_DIR = f'{ROOT_DIR}\\data'
LESSONS_DIR = rf'{ROOT_DIR}\data\lessons'
STUDENTS_DIR = rf'{ROOT_DIR}\data\students'
inputButtonObject = TypeVar('inputButtonObject')
inputVariableToBeChecked = TypeVar('inputVariableToBeChecked')
# ================================================

# ================================================ UI Classes
class CreateLessonWindow(QtWidgets.QWidget):
    """
    Class for "Ders Oluştur" tab.

    Attributes:
        self.studentGradesPath: Path to the student grades table.
        self.table1Path: Path to the table 1.
        self.table2Path: Path to the table 2.
        self.programOutputPath: Path to the program output table.
        self.lessonOutputPath: Path to the lesson output table.

    Member Functions:
        Helper:
            __determine_color(checkInput: inputVariableToBeChecked, buttonObj: inputButtonObject) -> None: Function to determine color.

        Util:
            _initialize(self) -> None: Function to initialize the class.
            _validate_inputs(self) -> None: Function to validate input data.
            _upload_table(self, title: str) -> str | None: Function to upload table data.
            _clear_inputs(self) -> None: Function to clear input data.
            _show_error_message(self, title: str, message: str) -> None: Function to show error message.
            _show_info_message(self, title: str, message: str) -> None: Function to show info message.

        Main:
            create_lesson(self) -> None: Main function to create a new lesson.
            upload_student_grades(self) -> None: Main function to upload student grades.
            upload_table1(self) -> None: Main function to upload table 1.
            upload_table2(self) -> None: Main function to upload table 2.
            upload_program_output(self) -> None: Main function to upload program output table.
            upload_lesson_output(self) -> None: Main function to upload lesson output table.
    """

    lessonCreated = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._initialize()
        self.studentGradesPath = None
        self.table1Path = None
        self.table2Path = None
        self.programOutputPath = None
        self.lessonOutputPath = None

        # Call function "Temizle" after the window's initialization.
        # This way color coding would also be initialized without worrying about data loss.
        self._clear_inputs()

    # ================================================ Util
    def _initialize(self) -> None:
        """
        Utility function to initialize the window.

        Parameters: self: CreateLessonWindow instance
        Returns: None
        """

        # Add a textbox and a label to get the title of the lesson that is going to be created.
        self.lessonLabel = QtWidgets.QLabel("Dersin ismini girin (Ders kodu ile birlikte):", self)
        self.lessonInput = QtWidgets.QLineEdit(self)
        self.lessonInput.textChanged.connect(self._validate_inputs)

        # Add a button and a label to get the grades.xlsx file of the lesson that is going to be created.
        self.label1 = QtWidgets.QLabel("Öğrenci notlarının bulunduğu tabloyu yükleyin:", self)
        self.button1 = QtWidgets.QPushButton("Tablo Ekle", self)
        self.button1.clicked.connect(self.upload_student_grades)

        # Add a button and a label to get the table1.xlsx file of the lesson that is going to be created.
        self.label2 = QtWidgets.QLabel("Tablo1'i yükleyin:", self)
        self.button2 = QtWidgets.QPushButton("Tablo Ekle", self)
        self.button2.clicked.connect(self.upload_table1)

        # Add a button and a label to get the table2.xlsx file of the lesson that is going to be created.
        self.label3 = QtWidgets.QLabel("Tablo2'yi yükleyin:", self)
        self.button3 = QtWidgets.QPushButton("Tablo Ekle", self)
        self.button3.clicked.connect(self.upload_table2)

        # Add a button and a label to get the program output file.
        self.label4 = QtWidgets.QLabel("Program çıktı tablosunu yükleyin:", self)
        self.button4 = QtWidgets.QPushButton("Tablo Ekle", self)
        self.button4.clicked.connect(self.upload_program_output)

        # Add a button and a label to get the lesson output file.
        self.label5 = QtWidgets.QLabel("Ders çıktı tablosunu yükleyin:", self)
        self.button5 = QtWidgets.QPushButton("Tablo Ekle", self)
        self.button5.clicked.connect(self.upload_lesson_output)

        # Add the button "Dersi Oluştur".
        self.button6 = QtWidgets.QPushButton("Dersi Oluştur", self)
        self.button6.clicked.connect(self.create_lesson)

        # Add the button "Temizle".
        self.button7 = QtWidgets.QPushButton("Temizle", self)
        self.button7.clicked.connect(self._clear_inputs)

        # Add the button "Geri Dön".
        self.button8 = QtWidgets.QPushButton("Geri Dön", self)
        self.button8.clicked.connect(self.go_back_to_main_menu)

        # Create the layout.
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.lessonLabel)
        layout.addWidget(self.lessonInput)

        layout.addWidget(self.label1)
        layout.addWidget(self.button1)

        layout.addWidget(self.label2)
        layout.addWidget(self.button2)

        layout.addWidget(self.label3)
        layout.addWidget(self.button3)

        layout.addWidget(self.label4)
        layout.addWidget(self.button4)

        layout.addWidget(self.label5)
        layout.addWidget(self.button5)

        layout.addSpacing(25)
        layout.addWidget(self.button6)
        layout.addWidget(self.button7)
        layout.addWidget(self.button8)
        self.setLayout(layout)

        self.setWindowTitle("Ders Oluşturma Aracı")
        self.setFixedSize(300, 400)

    def _validate_inputs(self) -> None:
        """
        Utility function to validate the inputs using color codes.

        Parameters: self: CreateLessonWindow instance.
        Returns: None
        """
        def __determine_color(checkInput: inputVariableToBeChecked, buttonObj: inputButtonObject) -> None:
            """
            Helper function to determine the color codes.

            Parameters:
                checkInput (inputVariableToBeChecked): Input variable that is going to be checked.
                buttonObj (inputButtonObject): Button object that is going to change colors according to the input variable.
            Returns:
                None
            """

            # This function may seem inefficient or unnecessary,
            # but it actually cuts 25 lines of code from "_validate_inputs" function.
            green = "background-color: rgb(173, 255, 190);"
            red = "background-color: rgb(255, 200, 200);"

            if checkInput:
                buttonObj.setStyleSheet(green)
            else:
                buttonObj.setStyleSheet(red)

        __determine_color(self.lessonInput.text().strip(), self.lessonInput)
        __determine_color(self.studentGradesPath, self.button1)
        __determine_color(self.table1Path, self.button2)
        __determine_color(self.table2Path, self.button3)
        __determine_color(self.programOutputPath, self.button4)
        __determine_color(self.lessonOutputPath, self.button5)

    def _upload_table(self, title: str) -> str | None:
        """
        Utility function to upload the table to the file system.

        Parameters:
            self: CreateLessonWindow instance.
            title (str): Title of the table to be uploaded.
        Returns:
            str: Path to the uploaded file.
            None: If the file cannot be uploaded.
        """
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(self, title, "/", "Excel Tablosu (*.xlsx);;All Files (*)")
        return fileName if fileName else None

    def _clear_inputs(self) -> None:
        """
        Utility function to clear the inputs.

        Parameters: self: CreateLessonWindow instance.
        Returns: None
        """
        self.lessonInput.clear()
        self.studentGradesPath = None
        self.table1Path = None
        self.table2Path = None
        self.programOutputPath = None
        self.lessonOutputPath = None
        self._validate_inputs()

    def _show_error_message(self, title: str, message: str) -> None:
        """
        Utility function to show the error message.

        Parameters:
            self: CreateLessonWindow instance.
            title (str): Title of the error message.
            message (str): Error message.
        Returns: None
        """
        errorBox = QtWidgets.QMessageBox(self)
        errorBox.setIcon(QtWidgets.QMessageBox.Warning)
        errorBox.setWindowTitle(title)
        errorBox.setText(message)
        errorBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
        errorBox.exec_()

    def _show_info_message(self, title: str, message: str) -> None:
        """Utility function to show the info message.

        Parameters:
            self: CreateLessonWindow instance.
            title (str): Title of the info message.
            message (str): Info message.
        Returns: None
        """
        infoBox = QtWidgets.QMessageBox(self)
        infoBox.setIcon(QtWidgets.QMessageBox.Information)
        infoBox.setWindowTitle(title)
        infoBox.setText(message)
        infoBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
        infoBox.exec_()



    # ================================================ Main
    def create_lesson(self) -> None:
        """
        Main function that creates the lesson.

        Parameters: self: CreateLessonWindow instance.
        Returns: None
        """
        lessonTitle = self.lessonInput.text().strip()

        # Check if the inputs are already added.
        if not lessonTitle or not self.studentGradesPath or not self.table1Path or not self.table2Path or not self.programOutputPath or not self.lessonOutputPath:
            self._show_error_message("Eksik Bilgi", "Lütfen tüm bilgileri eksiksiz doldurun.")
            return

        # Create or replace lesson directory.
        lessonPath = Path(LESSONS_DIR) / lessonTitle

        try:
            if lessonPath.exists():
                # Clear the existing directory
                for file in lessonPath.iterdir():
                    if file.is_file(): file.unlink()
                    elif file.is_dir(): shutil.rmtree(file)
            else:
                os.makedirs(lessonPath)

            # Copy files to lesson directory, change their names so that main.py can use them.
            shutil.copy(self.studentGradesPath, lessonPath / "grades.xlsx")
            shutil.copy(self.table1Path, lessonPath / "table1.xlsx")
            shutil.copy(self.table2Path, lessonPath / "table2.xlsx")
            shutil.copy(self.programOutputPath, lessonPath / "program_output.xlsx")
            shutil.copy(self.lessonOutputPath, lessonPath / "lesson_output.xlsx")

            self._show_info_message("Başarılı", f"Ders başarıyla oluşturuldu: {lessonTitle}")

            # Ders oluşturulduğunda sinyal gönder
            self.lessonCreated.emit()

            self._clear_inputs()
            subprocess.call(f"python {ROOT_DIR}/src/main.py", shell=True)


        except Exception as e:
            self._show_error_message("Hata", f"Ders oluşturulurken bir hata oluştu: {str(e)}")

    def upload_student_grades(self) -> None:
        """Main function to upload grades.xlsx table."""
        self.studentGradesPath = self._upload_table("Öğrenci Notları Tablosunu Seçin")
        self._validate_inputs()

    def upload_table1(self) -> None:
        """Main function to upload table1.xlsx table."""
        self.table1Path = self._upload_table("Tablo1'i Seçin")
        self._validate_inputs()

    def upload_table2(self) -> None:
        """Main function to upload table2.xlsx table."""
        self.table2Path = self._upload_table("Tablo2'yi Seçin")
        self._validate_inputs()

    def upload_program_output(self) -> None:
        """Main function to upload program_output.xlsx table."""
        self.programOutputPath = self._upload_table("Program Çıktısını Seçin")
        self._validate_inputs()

    def upload_lesson_output(self) -> None:
        """Main function to upload lesson_output.xlsx table."""
        self.lessonOutputPath = self._upload_table("Ders Çıktısını Seçin")
        self._validate_inputs()

    def go_back_to_main_menu(self):
        # Geri Dön butonuna basıldığında ana menüye dön
        self.hide()  # Dersi Oluştur penceresini gizle
        if hasattr(self, 'parentWindow'):  # Eğer parentWindow özelliği varsa
            self.parentWindow.show()  # Ana menüyü tekrar göster

# =================================DÜZENLENMESİ GEREK============================================================================================
class EditLessonWindow(QtWidgets.QWidget):
    """
    Class for "Dersi Düzenle" tab.

    Attributes:
        self.lessonTitle: Title of the selected lesson.
        self.inputField1: First string input field.
        self.inputField2: Second string input field.
        self.inputField3: Third string input field.

    Member Functions:
        Helper:
            _initialize(self) -> None: Function to initialize the class.
            _clear_inputs(self) -> None: Function to clear input fields.
            _show_error_message(self, title: str, message: str) -> None: Function to show error message.
            _show_info_message(self, title: str, message: str) -> None: Function to show info message.

        Main:
            load_lesson(self) -> None: Main function to load a selected lesson.
            save_changes(self) -> None: Main function to save the changes made to the lesson.
            go_back_to_main_menu(self) -> None: Function to go back to the main menu.
    """

    def __init__(self):
        super().__init__()
        self._initialize()

        # Initialize the input fields
        self.inputField1 = None
        self.inputField2 = None
        self.inputField3 = None
        self.lessonTitle = None


    # ================================================ Util
    def _initialize(self) -> None:
        """
        Utility function to initialize the window.

        Parameters: self: EditLessonWindow instance
        Returns: None
        """
        # Add a combo box for selecting the lesson
        self.lessonLabel = QtWidgets.QLabel("Ders Seçin:", self)
        self.lessonComboBox = QtWidgets.QComboBox(self)
        self.lessonComboBox.addItems(os.listdir(LESSONS_DIR))  # Load lessons into combo box
        self.lessonComboBox.currentIndexChanged.connect(self.load_lesson)

        # Add three input fields for the lesson's attributes
        self.inputLabel1 = QtWidgets.QLabel("Yeni Ders Adı:", self)
        self.inputField1 = QtWidgets.QLineEdit(self)  # Initialize the input field

        self.inputLabel2 = QtWidgets.QLabel("Yeni Kolon İsimleri (Virgülle Ayrılmış):", self)
        self.inputField2 = QtWidgets.QLineEdit(self)  # Initialize the input field

        self.inputLabel3 = QtWidgets.QLabel("Yeni Yüzdelikler (Virgülle Ayrılmış):", self)
        self.inputField3 = QtWidgets.QLineEdit(self)  # Initialize the input field

        # Add Save, Clear, and Back buttons
        self.saveButton = QtWidgets.QPushButton("Kaydet", self)
        self.saveButton.clicked.connect(self.save_changes)

        self.clearButton = QtWidgets.QPushButton("Temizle", self)
        self.clearButton.clicked.connect(self.clear_inputs)

        self.backButton = QtWidgets.QPushButton("Geri Dön", self)
        self.backButton.clicked.connect(self.go_back_to_main_menu)

        # Layout setup
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.lessonLabel)
        layout.addWidget(self.lessonComboBox)

        layout.addWidget(self.inputLabel1)
        layout.addWidget(self.inputField1)

        layout.addWidget(self.inputLabel2)
        layout.addWidget(self.inputField2)

        layout.addWidget(self.inputLabel3)
        layout.addWidget(self.inputField3)

        layout.addWidget(self.saveButton)
        layout.addWidget(self.clearButton)
        layout.addWidget(self.backButton)

        self.setLayout(layout)

        self.setWindowTitle("Ders Düzenle")
        self.setFixedSize(300, 350)

    def load_lesson(self):
        selected_lesson = self.lessonComboBox.currentText()
        lesson_path = os.path.join(DATA_DIR, "lessons", selected_lesson)

        # Kontrol et, dosya yolu var mı?
        if not os.path.exists(lesson_path):
            print(f"Dosya yolu bulunamadı: {lesson_path}")
            self._show_error_message("Hata", "Seçilen ders dosyası bulunamadı.")
            return

        try:
            columns_list = get_column_names(selected_lesson)
            print(f"Kolon İsimleri: {columns_list}")

            # Sütun isimlerini sadece input alanı başlatılmışsa güncelle
            if self.inputField2:
                self.inputField2.setText(",".join(columns_list))  # Default column names

            percentages = [str(get_percentage(selected_lesson, col)) for col in columns_list]
            print(f"Yüzdelikler: {percentages}")

            # Yüzde bilgilerini sadece input alanı başlatılmışsa güncelle
            if self.inputField3:
                self.inputField3.setText(",".join(percentages))  # Default percentages

        except Exception as e:
            print(f"Bir hata oluştu: {e}")
            self._show_error_message("Hata", "Ders verilerini yüklerken bir hata oluştu.")

    def save_changes(self) -> None:
        try:
            selected_lesson = self.lessonComboBox.currentText()

            # 1. Dersin ismi
            new_lesson_name = self.inputField1.text().strip()
            if new_lesson_name:  # Eğer isim boş değilse
                # Yeni ismin, dersler içinde var olmadığını kontrol et
                if new_lesson_name in os.listdir(LESSONS_DIR):
                    self._show_error_message("Hata", f"{new_lesson_name} ismi zaten mevcut!")
                    return
                else:
                    # Ders ismini değiştir
                    os.rename(os.path.join(DATA_DIR, "lessons", selected_lesson),
                              os.path.join(DATA_DIR, "lessons", new_lesson_name))
                    print(f"Ders ismi {selected_lesson} -> {new_lesson_name} olarak değiştirildi.")

            # 2. Sütun isimleri
            new_column_names = self.inputField2.text().strip()
            if new_column_names:  # Eğer sütun isimleri boş değilse
                new_column_names_list = [name.strip() for name in new_column_names.split(",")]
                current_columns = get_column_names(selected_lesson)

                if len(new_column_names_list) == len(current_columns):
                    # Kolon isimlerini değiştir
                    for old_name, new_name in zip(current_columns, new_column_names_list):
                        set_column_name(selected_lesson, new_name, old_name)
                    print(f"Sütun isimleri değiştirildi: {new_column_names_list}")
                else:
                    self._show_error_message("Hata", "Yeni sütun isimleri mevcut kolon sayısıyla uyumsuz!")

            # 3. Yüzdelikler
            new_percentages = self.inputField3.text().strip()
            if new_percentages:  # Eğer yüzdeler boş değilse
                new_percentages_list = [int(p.strip()) for p in new_percentages.split(",")]
                current_columns = get_column_names(selected_lesson)

                if len(new_percentages_list) == len(current_columns):
                    # Yüzdeleri değiştir
                    for col, percentage in zip(current_columns, new_percentages_list):
                        set_percentage(selected_lesson, col, percentage)
                    print(f"Yüzdelikler değiştirildi: {new_percentages_list}")
                else:
                    self._show_error_message("Hata", "Yeni yüzdeler mevcut kolon sayısıyla uyumsuz!")

            self._show_success_message("Başarılı", "Değişiklikler kaydedildi.")
        except Exception as e:
            print(f"Hata oluştu: {e}")
            self._show_error_message("Hata", f"Bir hata oluştu: {e}")

    def clear_inputs(self):
        """Clear all input fields."""
        self.inputField1.clear()
        self.inputField2.clear()
        self.inputField3.clear()
        self.lessonComboBox.setCurrentIndex(0)

    def go_back_to_main_menu(self) -> None:
        """
        Go back to the main menu.

        Parameters: self: EditLessonWindow instance
        Returns: None
        """
        self.hide()  # Hide the edit lesson window
        if hasattr(self, 'parentWindow'):  # If parent window exists
            self.parentWindow.show()  # Show the main menu again

    def _show_error_message(self, title: str, message: str) -> None:
        """
        Utility function to show the error message.

        Parameters:
            self: EditLessonWindow instance.
            title (str): Title of the error message.
            message (str): Error message.
        Returns: None
        """
        errorBox = QtWidgets.QMessageBox(self)
        errorBox.setIcon(QtWidgets.QMessageBox.Warning)
        errorBox.setWindowTitle(title)
        errorBox.setText(message)
        errorBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
        errorBox.exec_()

    def _show_info_message(self, title: str, message: str) -> None:
        """Utility function to show the info message.

        Parameters:
            self: EditLessonWindow instance.
            title (str): Title of the info message.
            message (str): Info message.
        Returns: None
        """
        infoBox = QtWidgets.QMessageBox(self)
        infoBox.setIcon(QtWidgets.QMessageBox.Information)
        infoBox.setWindowTitle(title)
        infoBox.setText(message)
        infoBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
        infoBox.exec_()


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ana Menü")
        self.setFixedSize(300, 350)

        # Menü Çubuğu
        menuBar = self.menuBar()
        lessonMenu = menuBar.addMenu("Ders İşlemleri")

        # Menü Öğesi: "Ders Oluştur"
        createLessonAction = QtWidgets.QAction("Ders Oluştur", self)
        createLessonAction.triggered.connect(self.open_create_lesson_window)
        lessonMenu.addAction(createLessonAction)

        # Menü Çubuğu'na "Dersi Düzenle" menüsü eklemek
        editLessonAction = QtWidgets.QAction("Ders Düzenle", self)
        editLessonAction.triggered.connect(self.open_edit_lesson_window)
        lessonMenu.addAction(editLessonAction)

        # Merkez Widget ve Layout
        centralWidget = QtWidgets.QWidget()
        self.setCentralWidget(centralWidget)

        layout = QtWidgets.QVBoxLayout()

        self.label1 = QtWidgets.QLabel("Dersi seçin:", self)
        layout.addWidget(self.label1)

        # ComboBox Ekleniyor
        self.comboBox = QtWidgets.QComboBox()
        self.comboBox.addItems(os.listdir(LESSONS_DIR))  # Örnek dersler
        self.comboBox.currentIndexChanged.connect(self.on_combobox_selection)

        layout.addWidget(self.comboBox)

        # Butonlar
        self.buttons = [
            ("Program Çıktı Tablosu", lambda: self.open_table("program_output.xlsx")),
            ("Ders Çıktı Tablosu", lambda: self.open_table("lesson_output.xlsx")),
            ("Öğrenci Notları Tablosu", lambda: self.open_table("grades.xlsx")),
            ("Tablo1", lambda: self.open_table("table1.xlsx")),
            ("Tablo2", lambda: self.open_table("table2.xlsx")),
            ("Tablo3", lambda: self.open_table("table3.xlsx")),
            ("Tablo4", lambda: self.open_table("table4.xlsx")),
            ("Tablo5", lambda: self.open_table("table5.xlsx")),
        ]
        layout.addSpacing(20)

        self.label2 = QtWidgets.QLabel("Açmak istediğiniz tabloyu seçin:", self)
        layout.addWidget(self.label2)

        for button_text, function in self.buttons:
            button = QtWidgets.QPushButton(button_text)
            button.clicked.connect(function)
            layout.addWidget(button)

        centralWidget.setLayout(layout)

    def open_create_lesson_window(self):
        # Ders Oluşturma Penceresini Aç
        self.createLessonWindow = CreateLessonWindow()
        # CreateLessonWindow'dan gelen sinyali dinleyin
        self.createLessonWindow.lessonCreated.connect(self.update_combobox)
        self.createLessonWindow.show()

    def open_edit_lesson_window(self):
        # Dersi Düzenleme Penceresini Aç
        self.editLessonWindow = EditLessonWindow()
        self.editLessonWindow.show()

    def update_combobox(self):
        # ComboBox'ı güncelle
        self.comboBox.clear()
        self.comboBox.addItems(os.listdir(LESSONS_DIR))

    def on_combobox_selection(self, index):
        selectedItem = self.comboBox.currentText()

    def open_table(self, table_name):
        table_path = Path(f"{LESSONS_DIR}/{self.comboBox.currentText()}/{table_name}")
        try:
            # Excel uygulamasını başlat
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = True  # Excel uygulamasını görünür yap

            # Dosyayı aç
            workbook = excel.Workbooks.Open(table_path)
            print(f"Excel dosyası açıldı: {table_path}")
            return workbook
        except Exception as e:
            print(f"Excel dosyasını açarken bir hata oluştu: {e}")





if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())
