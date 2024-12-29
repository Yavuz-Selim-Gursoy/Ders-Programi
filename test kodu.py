import unittest
import pandas as pd
import os
from pathlib import Path

# Varsayılan olarak Lesson ve Student sınıflarının bulunduğu modül
from src.main import (Lesson, Student, DATA_DIR,
                      ALL_LESSON_OBJECTS, LessonDataMutator,
                      LessonDataAccessor, StudentDataAccessor,
                      StudentDataMutator)


class TestLesson(unittest.TestCase):
    def setUp(self):
        """Her testten önce çalışacak ayarları yapar."""
        self.lesson_title = "BLM001"  # Test için kullanılacak ders başlığı
        self.lesson = Lesson(self.lesson_title)  # Lesson nesnesi oluşturur
        self.mutator = LessonDataMutator(self.lesson)
        self.accessor = LessonDataAccessor(self.lesson)

    def test_lesson_initialization(self):
        """Lesson nesnesinin doğru bir şekilde başlatıldığını test eder."""
        self.assertEqual(self.lesson.title, self.lesson_title)
        self.assertTrue(os.path.exists(self.lesson.inputFolderPath))
        self.assertIsInstance(self.lesson.tableOneDataFrame, pd.DataFrame)
        self.assertIsInstance(self.lesson.tableTwoDataFrame, pd.DataFrame)
        self.assertIsInstance(self.lesson.tableThreeDataFrame, pd.DataFrame)
        self.assertIsInstance(self.lesson.tableGradesDataFrame, pd.DataFrame)

    def test_create_folder_for_students(self):
        """Öğrenciler için klasörlerin oluşturulup oluşturulmadığını test eder."""
        self.lesson._create_folder_for_students()
        for student in self.lesson.lessonStudents:
            self.assertTrue(os.path.exists(f"{DATA_DIR}/students/{student}"))

    def test_create_df_from_lesson_table(self):
        """Ders tablolarından DataFrame'lerin doğru bir şekilde oluşturulup oluşturulmadığını test eder."""
        df1 = self.lesson._create_df_from_lesson_table(1)
        self.assertIsInstance(df1, pd.DataFrame)
        self.assertIn('İlişki Değeri', df1.columns)

    def test_check_tables(self):
        """Tabloların konsola yazdırılmasını test eder."""
        self.lesson._check_tables()  # Konsola yazdırma işlemi test edilir (görsel bir test).

    def test_setter_and_getter_functions(self):
        """Lesson sınıfındaki setter ve getter işlevlerini test eder."""
        # Table One: Update a specific value
        self.mutator.set_table_one_value("Ders çıktısı", 1, 0.99)
        value = self.accessor.get_table_one().iloc[1, 1]
        self.assertEqual(value, 0.99)

        # Table Two: Update a column
        new_column = pd.Series([1, 2, 3], name="Vize")
        self.mutator.set_table_two_column("Vize", new_column)
        column = self.accessor.get_table_two()["Vize"]
        pd.testing.assert_series_equal(column, new_column)

        # Table Three: Update an entire table
        new_table_three = pd.DataFrame({
            "Ders Çıktı": ["New1", "New2"],
            "Ağırlıklı Değerlendirme": [0.8, 0.9]
        })
        self.mutator.set_table_three(new_table_three)
        pd.testing.assert_frame_equal(self.accessor.get_table_three(), new_table_three)

class TestStudent(unittest.TestCase):
    def setUp(self):
        """Her testten önce çalışacak ayarları yapar."""
        self.student_id = 1  # Test için kullanılacak öğrenci ID'si
        self.student_folder_path = f'{DATA_DIR}/students/{self.student_id}'
        lesson_folder = f'{DATA_DIR}/students/{self.student_id}/BLM001'
        os.makedirs(lesson_folder, exist_ok=True)
        self.student = Student(self.student_id)  # Student nesnesi oluşturur
        self.student_mutator = StudentDataMutator(self.student)
        self.student_accessor = StudentDataAccessor(self.student)

    def test_student_initialization(self):
        """Student nesnesinin doğru bir şekilde başlatıldığını test eder."""
        self.assertEqual(self.student.id, self.student_id)
        self.assertTrue(os.path.exists(self.student.folderPath))
        self.assertIsInstance(self.student.lessons, list)

    def test_find_lessons(self):
        """Öğrencinin kayıtlı derslerini bulup bulamadığını test eder."""
        lessons = self.student._find_lessons()
        for lesson in lessons:
            self.assertIn(str(self.student.id), lesson.lessonStudents)

    def test_create_df_from_student_table(self):
        """Öğrenci tablolarından DataFrame'lerin doğru bir şekilde oluşturulup oluşturulmadığını test eder."""
        self.student._create_df_from_student_table()
        for lesson in self.student.lessons:
            self.assertTrue(os.path.exists(f"{DATA_DIR}/students/{self.student.id}/{lesson.title}/table4.xlsx"))
            self.assertTrue(os.path.exists(f"{DATA_DIR}/students/{self.student.id}/{lesson.title}/table5.xlsx"))

    def test_student_setter_and_getter_functions(self):
        """Student sınıfındaki setter ve getter işlevlerini test eder."""
        lesson_title = "BLM001"

        # Table4: Update an entire table
        new_table4 = pd.DataFrame({
            "Ders Çıktı": ["Outcome1", "Outcome2"],
            "Başarı": [80, 90]
        })
        self.student_mutator.set_table4(lesson_title, new_table4)
        pd.testing.assert_frame_equal(self.student_accessor.get_table4(lesson_title), new_table4)

        # Table5: Update a specific column
        new_column = pd.Series([50, 60], name="Başarı Oranı")
        self.student_mutator.set_table5_column(lesson_title, "Başarı Oranı", new_column)
        column = self.student_accessor.get_table5_column(lesson_title, "Başarı Oranı")
        pd.testing.assert_series_equal(column, new_column)

if __name__ == '__main__':
    unittest.main()
