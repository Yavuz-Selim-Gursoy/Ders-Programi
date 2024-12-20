# ==========================
import pandas as pd
import os
from pathlib import Path
import numpy as np

ROOT_DIR = Path(__file__).parent.parent.absolute()
DATA_DIR = f'{ROOT_DIR}\\data'
LESSONS_NAMES = os.listdir(f'{DATA_DIR}\\lessons')


# ==========================

# ========================== Main Classes
class Lesson:
    """
    Description: Lesson class.

    Attributes:
        title (str) : Lesson title.
        inputFolderPath (str) : Lesson folder name.
        tableOneDataFrame (pd.DataFrame) : Lesson's table one dataframe.
        tableTwoDataFrame (pd.DataFrame) : Lesson's table two dataframe.
        tableThreeDataFrame (pd.DataFrame) : Lesson's table three dataframe.
        tableGradesDataFrame (pd.DataFrame) : Lesson's grade table dataframe.
        lessonStudents (list) : Lesson's student list.

    Member Functions:
        Utils:
            TODO: setter and getter functions
            _create_dataframe_from_table(self: Lesson, tableNum:int) -> pd.Dataframe: Creates dataframe from tables (selects the table using 'tableNum'), updates tables after changes.
            _check_tables(self: Lesson) -> None: Prints dataframes on console.
    """

    def __init__(self, title: str):
        self.title = title
        self.inputFolderPath = f'{DATA_DIR}\\lessons\\{self.title}'
        self.tableOneDataFrame = self._create_df_from_lesson_table(1)
        self.tableTwoDataFrame = self._create_df_from_lesson_table(2)
        self.tableThreeDataFrame = self._create_df_from_lesson_table(3)
        self.tableGradesDataFrame = self._create_df_from_lesson_table(0)
        self.lessonStudents = [str(student) for student in self.tableGradesDataFrame.iloc[0:, 0].to_list()]
        self._create_folder_for_students()

    def _create_df_from_lesson_table(self, tableNum: int) -> pd.DataFrame:
        """
        Description: Function to create dataframe from tables.
        Parameters:
            tableNum (int) : Lesson's table number.
        Returns:
            pd.DataFrame : Lesson's table dataframe.
        """

        # Arrange the actions to be taken in each choice.

        if tableNum == 1:
            # Select a table using 'tableNum' and read its xlsx files.
            df = pd.read_excel(f"{self.inputFolderPath}\\table1.xlsx", sheet_name=0)

            for col in df.columns[1:-1]:
                for element in df[col][1:-1]:
                    assert 0 <= element <= 1, "Table1's values must be in between 0 and 1."

            # Calculate column 'İlişki Değeri'.
            df['İlişki Değeri'] = df.iloc[1:, 1:].sum(axis=1) / (df.shape[1] - 1)
            toplamIndex = df.columns.tolist().index('İlişki Değeri')
            df.iloc[0, toplamIndex] = "İlişki Değeri"

        if tableNum == 2:
            # Select a table using 'tableNum' and read its xlsx files.
            df = pd.read_excel(f"{self.inputFolderPath}\\table2.xlsx", sheet_name=0)

            # If there are fewer than 3 tasks in the season, assert ValueError.
            assert not len(df.columns) - 2 < 3, "Table2 must have at least 3 columns."

            # If the grading weights do not sum up to 100, assert ValueError.
            assert not df.iloc[0, 1:].sum() != 100, "Grading weights must sum to 100."

            # Calculate column 'TOPLAM'
            if 'TOPLAM' in df.columns:
                # If column 'TOPLAM' already exists, it'll cause recursion and the column values won't be correct.
                df['TOPLAM'] = df.iloc[2:, 1:-1].sum(axis=1)

            else:
                # If column doesn't exist, recursion won't happen.
                df['TOPLAM'] = df.iloc[2:, 1:].sum(axis=1)

            # We have to bring actual columns (in this case its 'TOPLAM') down in order to work with them.
            toplamIndex = df.columns.tolist().index('TOPLAM')
            df.iloc[1, toplamIndex] = "TOPLAM"

        if tableNum == 3:
            # Read table2.
            df2 = pd.read_excel(f"{self.inputFolderPath}\\table2.xlsx", sheet_name=0)

            # Cut the dataframe to use just its values.
            recalculatedDf2 = (df2.iloc[2:, 1:-1] * df2.iloc[0, 1:-1]) / 100

            # Column 'Ders çıktı' needs to be added manually, identify the column.
            dersCiktiColumn = df2.iloc[1:, 0].to_list()

            # The keys in the original table are incorrect, so they need to be added manually. Identify the key row.
            firstRow = df2.iloc[1, 1:-1].to_list()

            # Create the third dataframe (df3) using recalculatedDf2.
            # Column names are placeholders and should not be interpreted as actual keys.
            # The actual keys are stored in the first row. This step improves the table's readability in Excel.
            df3 = pd.DataFrame(recalculatedDf2.values,
                               columns=['TABLO 3'] + ['Ağırlıklı Değerlendirme'] * (len(df2.columns) - 3))

            # Insert the actual keys into first row.
            df3 = pd.concat([pd.DataFrame([firstRow], columns=df3.columns), df3], ignore_index=True)

            if 'TOPLAM' in df3.columns:
                # If column 'TOPLAM' already exists, it'll cause recursion and the column values won't be correct.
                df3['TOPLAM'] = df3.iloc[0:, 0:-1].sum(axis=1)

            else:
                # If column doesn't exist, recursion won't happen.
                df3['TOPLAM'] = df3.iloc[0:, 0:].sum(axis=1)

            # Bring down the column 'TOPLAM'.
            toplamIndex = df3.columns.tolist().index("TOPLAM")
            df3.iloc[0, toplamIndex] = None
            df3.iloc[0, toplamIndex] = "TOPLAM"

            # If column 'Ders Çıktı' already exists, pass.
            if 'Ders Çıktı' in df3.columns:
                pass

            # If column 'Ders Çıktı' does not exist, create one and insert it into the first column.
            else:
                df3.insert(0, 'Ders Çıktı', dersCiktiColumn)
            # Choose df3 as the output dataframe (df).
            df = df3

        if tableNum == 0:
            # Read the grades table.
            df = pd.read_excel(f"{self.inputFolderPath}\\grades.xlsx", sheet_name=0)

            if 'ORT' not in df.columns:
                # Read table2 and crop the values.
                df2 = pd.read_excel(f"{self.inputFolderPath}\\table2.xlsx", sheet_name=0)
                df2_cropped = df2.iloc[2:, 1:-1]

                # Identify the weights from table2.
                weights = df2.iloc[0, 1:-1].to_list()

                # Calculate the weighted scores using weights and grades.
                weighted_scores = df.iloc[:, 1:].mul(weights, axis=1) / 100

                # Create the column 'ORT' and insert the sum of the row's weighted scores.
                df["ORT"] = weighted_scores.sum(axis=1)

                # If table2 and grades table column counts doesn't match, assert an error.
                assert len(df2_cropped.columns) == (
                            len(df.columns) - 2), "Table2 and TableGrades column counts must be the same."

            else:
                pass

        # If wrong option is chosen, assert an error.
        else:
            assert 0 <= tableNum <= 3, "Wrong input table number. Options are: 0, 1, 2 or 3."

        # Remove 'unnamed' from tables.
        df.columns = [col if "Unnamed" not in col else "" for col in df.columns]
        resultDf = df

        # Rewrite the table and return the dataframe.
        if tableNum != 0:
            resultDf.to_excel(f"{self.inputFolderPath}\\table{tableNum}.xlsx", index=False)

        # If tableNum is 0, change its filename to grades.xlsx.
        elif tableNum == 0:
            resultDf.to_excel(f"{self.inputFolderPath}\\grades.xlsx", index=False)

        return resultDf

    def _check_tables(self) -> None:
        """Prints dataframes on console."""
        print(self.tableOneDataFrame.to_string())
        print('-' * 130)
        print(self.tableTwoDataFrame.to_string())
        print('-' * 130)
        print(self.tableThreeDataFrame.to_string())
        print('-' * 130)
        print(self.tableGradesDataFrame.to_string())

    def _create_folder_for_students(self) -> None:
        """Creates directory for every student registered to the lesson."""

        # For every student in self.lessonStudents,
        for studentID in self.lessonStudents:

            # If the student have their own folder, pass.
            if str(studentID) in os.listdir(f'{DATA_DIR}\\students\\'):
                pass

            # If not,
            elif str(studentID) not in os.listdir(f'{DATA_DIR}\\students\\'):

                # Create a folder named their id.
                os.mkdir(f"{DATA_DIR}\\students\\{studentID}")


class Student:
    """
    Description: Student class.
    Attributes:
        id (int): Student id.
        folderPath : Student folder path.
        allLessonObjects (list): Lesson objects list.
        lessons: Student lessons.
    Member Functions:
        TODO: setter and getter functions
    """

    def __init__(self, id: int):
        self.id = id
        self.folderPath = f'{DATA_DIR}\\students\\{self.id}'
        self.allLessonObjects = [Lesson(lessonTitle) for lessonTitle in LESSONS_NAMES]
        self.lessons = self._find_lessons()

    def _find_lessons(self) -> list:
        """Finds active lessons that is available for the student."""

        lessons = list()
        # For every lesson in all lessons,
        for lesson in self.allLessonObjects:

            # If student id is in the lesson's lessonStudents list,
            if str(self.id) in lesson.lessonStudents:
                # Append lesson to student's lessons list.
                lessons.append(lesson)
        return lessons

    def _create_df_from_student_table(self, tableNum: int) -> None:
        """
        Description: Creates dataframe from table.
        Parameters:
            tableNum (int): Table number.
        Returns:
            None
            """

        # For every lesson in student's lessons list,
        for lesson in self.lessons:

            # If lesson title is not in the student's own folder,
            if lesson.title not in os.listdir(f'{DATA_DIR}\\students\\{self.id}\\'):

                # Create a folder for the lesson
                os.mkdir(f"{DATA_DIR}\\students\\{self.id}\\{lesson.title}")

                # If the table choice is 4 and table4 does not exist in lesson's folder,
                if tableNum == 4 and 'table4.xlsx' not in os.listdir(f'{DATA_DIR}\\students\\{self.id}\\{lesson.title}'):
                    # TODO: create table 4.

                    # Get dataframes '3' and 'grades',
                    df3 = lesson.tableThreeDataFrame
                    dfGrades = lesson.tableGradesDataFrame

                    # Cut the dataframes to use just its values.
                    recalculatedDf3 = df3.iloc[1:, 1:-1]
                    recalculatedDfGrades = dfGrades.iloc[:, 1:-1]

                    # Get column 'TOPLAM' from dataframe 3 to calculate the 'MAX' column.
                    # Column 'Ders çıktı' needs to be added manually, get the column from dataframe 3.
                    dersCiktiColumn = df3.iloc[1:, 0].to_list()
                    columnToplamDf3 = df3.iloc[1:, -1]
                    columnToplamDf3Rounded = [round(num, 3) for num in columnToplamDf3]

                    # Create another list that is filled with hundreds using the length of columnMax
                    hundredsList = [100] * len(columnToplamDf3Rounded)
                    columnMaxDf4 = [int(x * y) for x, y in zip(columnToplamDf3Rounded, hundredsList)]

                    # Create the dataframe '4'.
                    df4 = pd.DataFrame(recalculatedDf3.values, columns=recalculatedDfGrades.columns.to_list())

                    # For every studentID in 'grades' dataframe,
                    for studentID in dfGrades.iloc[:, 0].to_list():

                        # If studentID matches with current student,
                        if studentID == self.id:

                            # Find the student's row index,
                            # Put the index into a buffer list, then get the index using [0] from the buffer list.
                            rowIndex = dfGrades.index[dfGrades["Öğrenci"] == self.id].to_list()[0]

                            # Using the rowIndex, find the current student's grades and insert them into a list.
                            # Multiply every single grade with their weight, change the datatype to integer.
                            studentGrades = recalculatedDfGrades.iloc[rowIndex, :].to_list()
                            tableFourDataFrame = df4.mul(studentGrades, axis=1)
                            tableFourDataFrame = tableFourDataFrame.astype(int)

                            # Create the column 'TOPLAM'.

                            tableFourDataFrame["TOPLAM"] = tableFourDataFrame.sum(axis=1)
                            tableFourDataFrame["MAX"] = columnMaxDf4
                            tableFourDataFrame["Başarı"] = round((tableFourDataFrame["TOPLAM"] / tableFourDataFrame["MAX"]) * 100, 1)
                            tableFourDataFrame.insert(0, 'Ders Çıktı', dersCiktiColumn)
                            resultDf = tableFourDataFrame

                # If the table choice is 4 and table4 does not exist in lesson's folder,
                if tableNum == 5 and 'table5.xlsx' not in os.listdir(
                        f'{DATA_DIR}\\students\\{self.id}\\{lesson.title}'):
                    # TODO: create table 5.
                    ...

                resultDf.to_excel(f'{DATA_DIR}\\students\\{self.id}\\{lesson.title}\\table{tableNum}.xlsx', index=False)

    def _check_tables(self) -> None:
        for lesson in self.lessons:
            # Get dataframes '3' and 'grades',
            df3 = lesson.tableThreeDataFrame
            dfGrades = lesson.tableGradesDataFrame

            # Cut the dataframes to use just its values.
            recalculatedDf3 = df3.iloc[1:, 1:-1]
            recalculatedDfGrades = dfGrades.iloc[:, 1:-1]

            # Get column 'TOPLAM' from dataframe 3 to calculate the 'MAX' column.
            # Column 'Ders çıktı' needs to be added manually, get the column from dataframe 3.
            dersCiktiColumn = df3.iloc[1:, 0].to_list()
            columnToplamDf3 = df3.iloc[1:, -1]
            columnToplamDf3Rounded = [round(num, 3) for num in columnToplamDf3]

            # Create another list that is filled with hundreds using the length of columnMax
            hundredsList = [100] * len(columnToplamDf3Rounded)
            columnMaxDf4 = [int(x * y) for x, y in zip(columnToplamDf3Rounded, hundredsList)]

            # Create the dataframe '4'.
            df4 = pd.DataFrame(recalculatedDf3.values, columns=recalculatedDfGrades.columns.to_list())

            # For every studentID in 'grades' dataframe,
            for studentID in dfGrades.iloc[:, 0].to_list():

                # If studentID matches with current student,
                if studentID == self.id:
                    # Find the student's row index,
                    # Put the index into a buffer list, then get the index using [0] from the buffer list.
                    rowIndex = dfGrades.index[dfGrades["Öğrenci"] == self.id].to_list()[0]

                    # Using the rowIndex, find the current student's grades and insert them into a list.
                    # Multiply every single grade with their weight, change the datatype to integer.
                    studentGrades = recalculatedDfGrades.iloc[rowIndex, :].to_list()
                    tableFourDataFrame = df4.mul(studentGrades, axis=1)
                    tableFourDataFrame = tableFourDataFrame.astype(int)

                    # Create the column 'TOPLAM'.

                    tableFourDataFrame["TOPLAM"] = tableFourDataFrame.sum(axis=1)
                    tableFourDataFrame["MAX"] = columnMaxDf4
                    tableFourDataFrame["Başarı"] = round((tableFourDataFrame["TOPLAM"] / tableFourDataFrame["MAX"]) * 100, 1)
                    tableFourDataFrame.insert(0, 'Ders Çıktı', dersCiktiColumn)
                    resultDf = tableFourDataFrame

                    print(tableFourDataFrame.to_string())


randomStudent = Student(220501003)

randomStudent._create_df_from_student_table(4)



