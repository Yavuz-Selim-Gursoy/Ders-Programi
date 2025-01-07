# ================================================
from PyQt5 import QtWidgets, QtGui
from pathlib import Path
from typing import TypeVar
import os
import shutil
import sys

ROOT_DIR = Path(__file__).parent.parent.absolute()
LESSONS_DIR = rf'{ROOT_DIR}\data\lessons'
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
        self.button8.clicked.connect(self.close)

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
            self._clear_inputs()

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


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = CreateLessonWindow()
    window.show()
    sys.exit(app.exec_())
