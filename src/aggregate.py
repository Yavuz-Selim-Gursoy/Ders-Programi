# ===============================================
import os
import pandas as pd
from pathlib import Path

ROOT_DIR = Path(__file__).parent.parent.absolute()
DATA_DIR = rf'{ROOT_DIR}\data'
STUDENTS_DIR = os.path.join(DATA_DIR, "students")
LESSONS_DIR = os.path.join(DATA_DIR, "lessons")
# ===============================================


def aggregate():
    # Tüm ders klasörlerini al
    lesson_dirs = [d for d in os.listdir(LESSONS_DIR) if os.path.isdir(os.path.join(LESSONS_DIR, d))]

    for lesson in lesson_dirs:
        # Birleştirilmiş tablolar için boş liste
        aggregated_table4 = []
        aggregated_table5 = []

        # Öğrenci klasörlerini dolaş
        for student in os.listdir(STUDENTS_DIR):
            student_path = os.path.join(STUDENTS_DIR, student, lesson)
            if os.path.exists(student_path):
                table4_path = os.path.join(student_path, "table4.xlsx")
                table5_path = os.path.join(student_path, "table5.xlsx")

                # Eğer tablo4 mevcutsa, oku ve öğrenci numarası ile birlikte birleştir
                if os.path.exists(table4_path):
                    table4_data = pd.read_excel(table4_path)
                    table4_data.insert(0, "Öğrenci No", student)  # Öğrenci numarasını ekle
                    aggregated_table4.append(table4_data)
                    aggregated_table4.append(pd.DataFrame([[""] * len(table4_data.columns)], columns=table4_data.columns))  # Boş satır ekle

                # Eğer tablo5 mevcutsa, oku ve öğrenci numarası ile birlikte birleştir
                if os.path.exists(table5_path):
                    table5_data = pd.read_excel(table5_path)
                    table5_data.insert(0, "Öğrenci No", student)  # Öğrenci numarasını ekle
                    table5_column_count = len(table5_data.columns)
                    aggregated_table5.append(table5_data)
                    aggregated_table5.append(pd.DataFrame(columns=[""] * len(table5_data.columns)))  # Boş satır ekle

        # Birleştirilen tablo4'ü oluştur
        if aggregated_table4:
            final_table4 = pd.concat(aggregated_table4, ignore_index=True)
            final_table4_path = os.path.join(LESSONS_DIR, lesson, "table4.xlsx")
            final_table4.to_excel(final_table4_path, index=False)

        # Birleştirilen tablo5'i oluştur
        merged_df = pd.DataFrame(columns=[str(i + 1) for i in range(table5_column_count)])
        if aggregated_table5:
            for table5 in aggregated_table5:
                column_names_row = pd.DataFrame([table5.columns], columns=merged_df.columns)
                df_reset = pd.DataFrame(table5.values, columns=merged_df.columns)
                merged_df = pd.concat([merged_df, column_names_row, df_reset], ignore_index=True)

            final_table5_path = os.path.join(LESSONS_DIR, lesson, "table5.xlsx")
            merged_df.to_excel(final_table5_path, index=False)

if __name__ == "__main__":
    aggregate()
