# ================================================
from PyQt5 import QtWidgets, QtGui
from pathlib import Path
import os
import shutil
import sys

ROOT_DIR = Path(__file__).parent.parent.absolute()
LESSONS_DIR = f'{ROOT_DIR}\\data\\lessons'
# ================================================

# ================================================ UI Classes  #TODO: add some comments to raise the readability.
class CreateLessonWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self._initialize()
        self.studentGradesPath = None
        self.table1Path = None
        self.table2Path = None
        self._clear_inputs()  # Pencere ilk açıldığında temizle fonksiyonunu çağır

    def _initialize(self) -> None:
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

        # Add the button "Dersi Oluştur".
        self.button4 = QtWidgets.QPushButton("Dersi Oluştur", self)
        self.button4.clicked.connect(self.create_lesson)

        self.button5 = QtWidgets.QPushButton("Temizle", self)
        self.button5.clicked.connect(self._clear_inputs)

        # Add the button "Geri Dön".
        self.button6 = QtWidgets.QPushButton("Geri Dön", self)
        self.button6.clicked.connect(self.close)

        # Create the layout.
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.lessonLabel)  # Label for string input
        layout.addWidget(self.lessonInput)  # String input
        layout.addWidget(self.label1)  # grades.xlsx input label
        layout.addWidget(self.button1)  # grades.xlsx input
        layout.addWidget(self.label2)  # table1.xlsx input label
        layout.addWidget(self.button2)  # table1.xlsx input
        layout.addWidget(self.label3)  # table2.xlsx input label
        layout.addWidget(self.button3)  # table2.xlsx input
        layout.addSpacing(25)
        layout.addWidget(self.button4)  # "Dersi Oluştur"
        layout.addWidget(self.button5)  # "Temizle"
        layout.addWidget(self.button6)  # "Geri Dön"
        self.setLayout(layout)

        self.setWindowTitle("Ders Oluşturma Aracı")
        self.setFixedSize(300, 350)

    def _validate_inputs(self) -> None:
        # Renklerin RGB değerleriyle tanımlanması
        green = "background-color: rgb(173, 255, 190);"
        red = "background-color: rgb(255, 200, 200);"

        # Ders ismi doğrulama
        if self.lessonInput.text().strip():
            self.lessonInput.setStyleSheet(green)
        else:
            self.lessonInput.setStyleSheet(red)

        # Öğrenci notları doğrulama
        if self.studentGradesPath:
            self.button1.setStyleSheet(green)
        else:
            self.button1.setStyleSheet(red)

        # Tablo1 doğrulama
        if self.table1Path:
            self.button2.setStyleSheet(green)
        else:
            self.button2.setStyleSheet(red)

        # Tablo2 doğrulama
        if self.table2Path:
            self.button3.setStyleSheet(green)
        else:
            self.button3.setStyleSheet(red)

    def _upload_table(self, title: str) -> None:
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(self, title, "/", "Excel Tablosu (*.xlsx);;All Files (*)")
        return fileName if fileName else None

    def _clear_inputs(self) -> None:
        self.lessonInput.clear()
        self.studentGradesPath = None
        self.table1Path = None
        self.table2Path = None
        self._validate_inputs()

    def _show_error_message(self, title: str, message: str) -> None:
        errorBox = QtWidgets.QMessageBox(self)
        errorBox.setIcon(QtWidgets.QMessageBox.Warning)
        errorBox.setWindowTitle(title)
        errorBox.setText(message)
        errorBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
        errorBox.exec_()

    def _show_info_message(self, title: str, message: str) -> None:
        infoBox = QtWidgets.QMessageBox(self)
        infoBox.setIcon(QtWidgets.QMessageBox.Information)
        infoBox.setWindowTitle(title)
        infoBox.setText(message)
        infoBox.setStandardButtons(QtWidgets.QMessageBox.Ok)
        infoBox.exec_()

    def create_lesson(self) -> None:  # TODO: you might want to rewrite those try/except blocks.
        lessonTitle = self.lessonInput.text().strip()

        # Assert and validate inputs
        if not lessonTitle or not self.studentGradesPath or not self.table1Path or not self.table2Path:
            self._show_error_message("Eksik Bilgi", "Lütfen tüm bilgileri eksiksiz doldurun.")
            return

        # Create or replace lesson directory
        lessonPath = Path(LESSONS_DIR) / lessonTitle

        try:
            if lessonPath.exists():
                # Clear the existing directory
                for file in lessonPath.iterdir():
                    if file.is_file():
                        file.unlink()
                    elif file.is_dir():
                        shutil.rmtree(file)
            else:
                os.makedirs(lessonPath)

            # Copy files to lesson directory
            shutil.copy(self.studentGradesPath, lessonPath / "grades.xlsx")
            shutil.copy(self.table1Path, lessonPath / "table1.xlsx")
            shutil.copy(self.table2Path, lessonPath / "table2.xlsx")

            self._show_info_message("Başarılı", f"Ders başarıyla oluşturuldu: {lessonTitle}")
            self._clear_inputs()
        except Exception as e:
            self._show_error_message("Hata", f"Bilinmeyen bir hata oluştu: {e}")

    def upload_student_grades(self) -> None:
        self.studentGradesPath = self._upload_table("Öğrenci notlarının bulunduğu tabloyu seçin")
        self._validate_inputs()

    def upload_table1(self) -> None:
        self.table1Path = self._upload_table("Tablo1 dosyasını seçin")
        self._validate_inputs()

    def upload_table2(self) -> None:
        self.table2Path = self._upload_table("Tablo2 dosyasını seçin")
        self._validate_inputs()




# Uygulamayı çalıştır
def main():
    app = QtWidgets.QApplication(sys.argv)
    window = CreateLessonWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

