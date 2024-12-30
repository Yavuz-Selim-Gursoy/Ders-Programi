# ==========================
import pandas as pd
import os
from pathlib import Path
import numpy as np

ROOT_DIR = Path(__file__).parent.parent.absolute()
DATA_DIR = f'{ROOT_DIR}\\data'
LESSON_NAMES = os.listdir(f'{DATA_DIR}\\lessons')
ALL_LESSON_OBJECTS = list()
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
            _create_dataframe_from_table (self: Lesson, tableNum:int) -> pd.Dataframe: Creates dataframe from tables (selects the table using 'tableNum'), updates tables after changes.
            _check_tables (self: Lesson) -> None: Prints dataframes on console.
            _create_folder_for_students (self: Lesson) -> None: Creates folder for students.
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
            columnDersCikti = df2.iloc[1:, 0].to_list()

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
                df3.insert(0, 'Ders Çıktı', columnDersCikti)
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
        """Prints dataframes 1, 2, 3 and grades dataframe on console."""
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
        _find_lessons (self: Student) -> list: Finds registered lessons for every student.
        _create_df_from_student_table(self: Student) -> None: Creates dataframe from tables and writes them in an excel file.
    """

    def __init__(self, id: int):
        self.id = id
        self.folderPath = f'{DATA_DIR}\\students\\{self.id}'
        self.allLessonObjects = ALL_LESSON_OBJECTS
        self.lessons = self._find_lessons()
        self._create_df_from_student_table()


    def _find_lessons(self) -> list:
        """Finds registered lessons for the student."""

        lessons = list()
        # For every lesson in all lessons,
        for lesson in self.allLessonObjects:

            # If student id is in the lesson's lessonStudents list,
            if str(self.id) in lesson.lessonStudents:
                # Append lesson to student's lessons list.
                lessons.append(lesson)
        return lessons


    def _create_df_from_student_table(self) -> None:
        """
        Description: Creates dataframe from table.
        Parameters:
            self (Student): Student object.
        Returns:
            None
        """

        # For every lesson in student's lessons list,
        for lesson in self.lessons:

            # Create a folder for the lesson.
            if lesson.title not in os.listdir(f"{DATA_DIR}\\students\\{self.id}\\"):
                os.mkdir(f"{DATA_DIR}\\students\\{self.id}\\{lesson.title}")

            # Get dataframes '3' and 'grades',
            df3 = lesson.tableThreeDataFrame
            dfGrades = lesson.tableGradesDataFrame

            # Cut the dataframes to use just its values.
            recalculatedDf3 = df3.iloc[1:, 1:-1]
            recalculatedDfGrades = dfGrades.iloc[:, 1:-1]

            # Get column 'TOPLAM' from dataframe 3 to calculate the 'MAX' column.
            # Column 'Ders çıktı' needs to be added manually, get the column from dataframe 3.
            columnDersCikti = df3.iloc[1:, 0].to_list()
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
                    tableFourDataFrame.insert(0, 'Ders Çıktı', columnDersCikti)



            # Get dataframe 1 and sanitize it.
            df1 = lesson.tableOneDataFrame
            recalculatedDf1 = df1.iloc[1:, 1:-1]

            # Get column 'Başarı' from dataframe 4.
            columnBasari = tableFourDataFrame["Başarı"].to_list()

            # Get column 'Prg Çıktı' from dataframe 1.
            columnPrgCikti = df1.iloc[1:, 0].to_list()

            # Initialize dataframe 5 using dataframe 1's values.
            df5 = pd.DataFrame(recalculatedDf1.values, columns=columnBasari)

            # Multiply every element with column 'Başarı' from dataframe 4.
            tableFiveDataFrame = df5.mul(columnBasari, axis=1)

            # Create a phantom dataframe where every single grade is maximum.
            hundredsList = [100] * len(columnBasari)
            maxGradedDf5 = df5.mul(hundredsList, axis=1)
            maxGradedDf5["MAXBASARI"] = maxGradedDf5.sum(axis=1)

            # Calculate total weighted grades for student.
            columnTotalBasariDf5 = tableFiveDataFrame.sum(axis=1).to_list()

            # Round total weighted grades and insert them into another phantom dataframe.
            columnTotalBasariDf5Rounded = [round(i, 3) for i in columnTotalBasariDf5]
            df5["TOTALBASARI"] = columnTotalBasariDf5Rounded

            # Insert column 'Prg Çıktı' to dataframe 5.
            tableFiveDataFrame.insert(0, 'Prg Çıktı', columnPrgCikti)

            # Calculate column 'Başarı Oranı' using phantom dataframe df5's 'TOTALBASARI' and
            # phantom dataframe maxGradedDf5's 'MAXBASARI' columns, and insert it into the actual tableFiveDataFrame.
            tableFiveDataFrame["Başarı Oranı"] = round((df5["TOTALBASARI"] / maxGradedDf5["MAXBASARI"]) * 100, 1)

            # Write dataframes into their excel tables.
            if 'table4.xlsx' not in os.listdir(f'{DATA_DIR}\\students\\{self.id}\\{lesson.title}\\'):
                tableFourDataFrame.to_excel(f'{DATA_DIR}\\students\\{self.id}\\{lesson.title}\\table4.xlsx', index=False)

            if 'table5.xlsx' not in os.listdir(f'{DATA_DIR}\\students\\{self.id}\\{lesson.title}\\'):
                tableFiveDataFrame.to_excel(f'{DATA_DIR}\\students\\{self.id}\\{lesson.title}\\table5.xlsx', index=False)


# ============================ Setter Functions
class LessonDataMutator:
    """
    Description: Provides setter functions to update Lesson instances tables
    """
    def __init__(self, lesson_instance):
        """
        lesson_instance must be an object of the Lesson class.
        """
        self._lesson = lesson_instance

    def set_table_one(self, new_table: pd.DataFrame) -> None:
        """
        Updates the entire table1 dataframe
        """
        if not isinstance(new_table, pd.DataFrame):
            raise ValueError('new_table must be a pandas dataframe')
        self._lesson.tableOneDataFrame = new_table

    def set_table_one_column(self, column_name: str, new_column: pd.Series) -> None:
        """
        Updates a specific column in the table1 dataframe

        Parametres:
            column_name (str): The name of the Index
            new_column (pd.Series): The new values for the column

        Returns:
             None
        """
        table = self._lesson.tableOneDataFrame

        # Checks column Index
        if column_name not in table.columns:
            raise ValueError(f"Column '{column_name}' not found in Table1.")

        # Raise the error if came the different values
        if not isinstance(new_column, pd.Series):
            raise ValueError('new_column must be a pandas seires')

        # Update the row
        table[column_name] = new_column

    def set_table_one_row(self, row_index: int, new_row: pd.Series) -> None:
        """
        Updates a specific row in the table1 dataframe.

        Parameters:
            row_index (int): Index of the row to update.
            new_row (pd.Series): New values for the row.

        Returns:
            None
        """
        table = self._lesson.tableOneDataFrame

        # Checks row index
        if not (0 <= row_index < len(table)):
            raise IndexError(f"Row index '{row_index}' is out of range for Table1.")

        # Checks Type
        if not isinstance(new_row, pd.Series):
            raise ValueError("new_row must be a pandas Series.")

        # Column compatibility check
        if not all(new_row.index == table.columns):
            raise ValueError("new_row indices must match table columns.")

        # Update the row
        table.iloc[row_index] = new_row

    def set_table_one_value(self, column_name: str, row_index: int, new_value) -> None:
        """
        Updates a specific value in a column in table1 dataframe.

        Parameters:
            column_name (str): The column where the value is to be updated.
            row_index (int): The row index of the value to be updated.
            new_value: The new value to set.

        Returns:
            None
        """
        table = self._lesson.tableOneDataFrame

        # Check the presence of the column
        if column_name not in table.columns:
            raise ValueError(f"Column '{column_name}' not found in Table1")

        # Check if the row index is valid
        if not (0 <= row_index < len(table)):
            raise IndexError(f"Row index '{row_index}' is out of range for Table1")

        # Update value
        table.at[row_index, column_name] = new_value

    #--------------------------
    def set_table_two(self, new_table: pd.DataFrame) -> None:
        """
        Updates the entire table2 dataframe.
        """
        if not isinstance(new_table, pd.DataFrame):
            raise ValueError("new_table must be a pandas DataFrame.")
        self._lesson.tableTwoDataFrame = new_table

    def set_table_two_column(self, column_name: str, new_column: pd.Series) -> None:
        """
        Updates a specific column in the table2 dataframe

        Parametres:
            column_name (str): The name of the Index
            new_column (pd.Series): The new values for the column

        Returns:
             None
        """
        table = self._lesson.tableTwoDataFrame

        # Checks column Index
        if column_name not in table.columns:
            raise ValueError(f"Column '{column_name}' not found in Table2.")

        # Raise the error if came the different values
        if not isinstance(new_column, pd.Series):
            raise ValueError('new_column must be a pandas seires')

        # Update the row
        table[column_name] = new_column

    def set_table_two_row(self, row_index: int, new_row: pd.Series) -> None:
        """
        Updates a specific row in the table2 dataframe.

        Parameters:
            row_index (int): Index of the row to update.
            new_row (pd.Series): New values for the row.

        Returns:
            None
        """
        table = self._lesson.tableOneDataFrame

        # Checks row index
        if not (0 <= row_index < len(table)):
            raise IndexError(f"Row index '{row_index}' is out of range for Table2.")

        # Checks Type
        if not isinstance(new_row, pd.Series):
            raise ValueError("new_row must be a pandas Series.")

        # Column compatibility check
        if not all(new_row.index == table.columns):
            raise ValueError("new_row indices must match table columns.")

        # Update the row
        table.iloc[row_index] = new_row

    def set_table_two_value(self, column_name: str, row_index: int, new_value) -> None:
        """
        Updates a specific value in a column in table2 dataframe.

        Parameters:
            column_name (str): The column where the value is to be updated.
            row_index (int): The row index of the value to be updated.
            new_value: The new value to set.

        Returns:
            None
        """
        table = self._lesson.tableOneDataFrame

        # Check the presence of the column
        if column_name not in table.columns:
            raise ValueError(f"Column '{column_name}' not found in Table2")

        # Check if the row index is valid
        if not (0 <= row_index < len(table)):
            raise IndexError(f"Row index '{row_index}' is out of range for Table2")

        # Update value
        table.at[row_index, column_name] = new_value

    #--------------------------
    def set_table_three(self, new_table: pd.DataFrame) -> None:
        """
        Updates the entire table3 dataframe.
        """
        if not isinstance(new_table, pd.DataFrame):
            raise ValueError("new_table must be a pandas DataFrame.")
        self._lesson.tableTwoDataFrame = new_table

    def set_table_three_column(self, column_name: str, new_column: pd.Series) -> None:
        """
        Updates a specific column in the table3 dataframe

        Parametres:
            column_name (str): The name of the Index
            new_column (pd.Series): The new values for the column

        Returns:
             None
        """
        table = self._lesson.tableTwoDataFrame

        # Checks column Index
        if column_name not in table.columns:
            raise ValueError(f"Column '{column_name}' not found in Table3.")

        # Raise the error if came the different values
        if not isinstance(new_column, pd.Series):
            raise ValueError('new_column must be a pandas seires')

        # Update the row
        table[column_name] = new_column

    def set_table_three_row(self, row_index: int, new_row: pd.Series) -> None:
        """
        Updates a specific row in the table3 dataframe.

        Parameters:
            row_index (int): Index of the row to update.
            new_row (pd.Series): New values for the row.

        Returns:
            None
        """
        table = self._lesson.tableOneDataFrame

        # Checks row index
        if not (0 <= row_index < len(table)):
            raise IndexError(f"Row index '{row_index}' is out of range for Table3.")

        # Checks Type
        if not isinstance(new_row, pd.Series):
            raise ValueError("new_row must be a pandas Series.")

        # Column compatibility check
        if not all(new_row.index == table.columns):
            raise ValueError("new_row indices must match table columns.")

        # Update the row
        table.iloc[row_index] = new_row

    def set_table_three_value(self, column_name: str, row_index: int, new_value) -> None:
        """
        Updates a specific value in a column in table3 dataframe.

        Parameters:
            column_name (str): The column where the value is to be updated.
            row_index (int): The row index of the value to be updated.
            new_value: The new value to set.

        Returns:
            None
        """
        table = self._lesson.tableOneDataFrame

        # Check the presence of the column
        if column_name not in table.columns:
            raise ValueError(f"Column '{column_name}' not found in Table3")

        # Check if the row index is valid
        if not (0 <= row_index < len(table)):
            raise IndexError(f"Row index '{row_index}' is out of range for Table3")

        # Update value
        table.at[row_index, column_name] = new_value

    #--------------------------
    def set_table_grades(self, new_table: pd.DataFrame) -> None:
        """
        Updates the entire grades table dataframe.
        """
        if not isinstance(new_table, pd.DataFrame):
            raise ValueError("new_table must be a pandas DataFrame.")
        self._lesson.tableTwoDataFrame = new_table

    def set_table_grades_column(self, column_name: str, new_column: pd.Series) -> None:
        """
        Updates a specific column in the grades table dataframe

        Parametres:
            column_name (str): The name of the Index
            new_column (pd.Series): The new values for the column

        Returns:
             None
        """
        table = self._lesson.tableTwoDataFrame

        # Checks column Index
        if column_name not in table.columns:
            raise ValueError(f"Column '{column_name}' not found in TableGrades.")

        # Raise the error if came the different values
        if not isinstance(new_column, pd.Series):
            raise ValueError('new_column must be a pandas seires')

        # Update the row
        table[column_name] = new_column

    def set_table_grades_row(self, row_index: int, new_row: pd.Series) -> None:
        """
        Updates a specific row in the grades table dataframe.

        Parameters:
            row_index (int): Index of the row to update.
            new_row (pd.Series): New values for the row.

        Returns:
            None
        """
        table = self._lesson.tableOneDataFrame

        # Checks row index
        if not (0 <= row_index < len(table)):
            raise IndexError(f"Row index '{row_index}' is out of range for TableGrades.")

        # Checks Type
        if not isinstance(new_row, pd.Series):
            raise ValueError("new_row must be a pandas Series.")

        # Column compatibility check
        if not all(new_row.index == table.columns):
            raise ValueError("new_row indices must match table columns.")

        # Update the row
        table.iloc[row_index] = new_row

    def set_table_grades_value(self, column_name: str, row_index: int, new_value) -> None:
        """
        Updates a specific value in a column in grades table dataframe.

        Parameters:
            column_name (str): The column where the value is to be updated.
            row_index (int): The row index of the value to be updated.
            new_value: The new value to set.

        Returns:
            None
        """
        table = self._lesson.tableOneDataFrame

        # Check the presence of the column
        if column_name not in table.columns:
            raise ValueError(f"Column '{column_name}' not found in TableGrades ")

        # Check if the row index is valid
        if not (0 <= row_index < len(table)):
            raise IndexError(f"Row index '{row_index}' is out of range for TableGrades")

        # Update value
        table.at[row_index, column_name] = new_value


class StudentDataMutator:
    """
    Description: Provides setter functions to update Student instances' table4 and table5.
    """

    def __init__(self, student_instance):
        """
        student_instance must be an object of the Student class.
        """
        self._student = student_instance

    def set_table4(self, lesson_title: str, new_table: pd.DataFrame) -> None:
        """
        Updates the entire table4 dataframe for a given lesson.
        """
        file_path = f'{DATA_DIR}\\students\\{self._student.id}\\{lesson_title}\\table4.xlsx'
        if not isinstance(new_table, pd.DataFrame):
            raise ValueError("new_table must be a pandas DataFrame.")
        new_table.to_excel(file_path, index=False)

    def set_table4_column(self, lesson_title: str, column_name: str, new_column: pd.Series) -> None:
        """
        Updates a specific column in table4 for a given lesson.
        """
        table4 = StudentDataAccessor(self._student).get_table4(lesson_title)
        if column_name not in table4.columns:
            raise ValueError(f"Column '{column_name}' not found in Table4 for lesson '{lesson_title}'.")
        if not isinstance(new_column, pd.Series):
            raise ValueError("new_column must be a pandas Series.")
        table4[column_name] = new_column
        file_path = f'{DATA_DIR}\\students\\{self._student.id}\\{lesson_title}\\table4.xlsx'
        table4.to_excel(file_path, index=False)

    def set_table4_row(self, lesson_title: str, row_index: int, new_row: pd.Series) -> None:
        """
        Updates a specific row in table4 for a given lesson.
        """
        table4 = StudentDataAccessor(self._student).get_table4(lesson_title)
        if not 0 <= row_index < len(table4):
            raise IndexError(f"Row index '{row_index}' is out of range for Table4 in lesson '{lesson_title}'.")
        if not isinstance(new_row, pd.Series):
            raise ValueError("new_row must be a pandas Series.")
        table4.iloc[row_index] = new_row
        file_path = f'{DATA_DIR}\\students\\{self._student.id}\\{lesson_title}\\table4.xlsx'
        table4.to_excel(file_path, index=False)

    def set_table5(self, lesson_title: str, new_table: pd.DataFrame) -> None:
        """
        Updates the entire table5 dataframe for a given lesson.
        """
        file_path = f'{DATA_DIR}\\students\\{self._student.id}\\{lesson_title}\\table5.xlsx'
        if not isinstance(new_table, pd.DataFrame):
            raise ValueError("new_table must be a pandas DataFrame.")
        new_table.to_excel(file_path, index=False)

    def set_table5_column(self, lesson_title: str, column_name: str, new_column: pd.Series) -> None:
        """
        Updates a specific column in table5 for a given lesson.
        """
        table5 = StudentDataAccessor(self._student).get_table5(lesson_title)
        if column_name not in table5.columns:
            raise ValueError(f"Column '{column_name}' not found in Table5 for lesson '{lesson_title}'.")
        if not isinstance(new_column, pd.Series):
            raise ValueError("new_column must be a pandas Series.")
        table5[column_name] = new_column
        file_path = f'{DATA_DIR}\\students\\{self._student.id}\\{lesson_title}\\table5.xlsx'
        table5.to_excel(file_path, index=False)

    def set_table5_row(self, lesson_title: str, row_index: int, new_row: pd.Series) -> None:
        """
        Updates a specific row in table5 for a given lesson.
        """
        table5 = StudentDataAccessor(self._student).get_table5(lesson_title)
        if not 0 <= row_index < len(table5):
            raise IndexError(f"Row index '{row_index}' is out of range for Table5 in lesson '{lesson_title}'.")
        if not isinstance(new_row, pd.Series):
            raise ValueError("new_row must be a pandas Series.")
        table5.iloc[row_index] = new_row
        file_path = f'{DATA_DIR}\\students\\{self._student.id}\\{lesson_title}\\table5.xlsx'
        table5.to_excel(file_path, index=False)


# ========================== Getter Functions
class LessonDataAccessor:
    """
    Description: Provides getter functions from lesson instances tables
    """
    def __init__(self, lesson_instance):
        """
        lesson_instance must be object of Lesson class
        """
        self._lesson = lesson_instance

    def get_table_one(self) -> pd.DataFrame:
        """
        Return lesson_instances table1 dataframe
        """
        return self._lesson.tableOneDataFrame

    def get_table_one_column(self) -> pd.DataFrame:
        """
        Return lesson_instances table1 column
        """
        return self._lesson.tableOneDataFrame['İlişki Değeri']

    def get_table_two(self) -> pd.DataFrame:
        """
        Return lesson_instances table2 dataframe
        """
        return self._lesson.tableTwoDataFrame

    def get_table_two_column(self) -> pd.DataFrame:
        """
        Return lesson_instances table2 column
        """
        return self._lesson.tableTwoDataFrame['TOPLAM']

    def get_table_three(self) -> pd.DataFrame:
        """
        Return lesson_instances table3 dataframe
        """
        return self._lesson.tableThreeDataFrame

    def get_table_three_column(self) -> pd.DataFrame:
        """
        Return lesson_instances table3 column
        """
        return self._lesson.tableThreeDataFrame['TOPLAM']

    def get_tableGradesDataFrame(self) -> pd.DataFrame:
        """
        Return lesson_instances Grades table dataframe
        """
        return self._lesson.tableGradesDataFrame

    def get_tableGradesDataFrame_column(self) -> pd.DataFrame:
        """
        Return lesson_instances grades table column
        """
        return self._lesson.tableGradesDataFrame['ORT']


class StudentDataAccessor:
    """
    Description: Provides getter functions to access table4 and table5 dataframes
    or their specified columns/rows
    """
    def __init__(self, student_instance):
        self._student = student_instance

    def get_table4(self, lesson_title: str) -> pd.DataFrame:
        """
        Returns the entire table4 dataframe for a given lesson.
        """
        file_path = f'{DATA_DIR}\\students\\{self._student.id}\\{lesson_title}\\table4.xlsx'

        if os.path.exists(file_path):
            return pd.read_excel(file_path)
        else:
            raise FileNotFoundError(f"Table4 not found for lesson '{lesson_title}'.")

    def get_table4_column(self, lesson_title: str, column_name: str) -> pd.Series:
        """
        Returns a specific column from table4 for a given lesson.
        """
        table4 = self.get_table4(lesson_title)
        if column_name in table4.columns:
            return table4[column_name]
        else:
            raise ValueError(f"Column '{column_name}' not found in table4 for lesson '{lesson_title}'.")

    def get_table4_row(self, lesson_title: str, row_index: int) -> pd.Series:
        """
        Returns a specific row from table4 for a given lesson.
        """
        table4 = self.get_table4(lesson_title)
        if row_index in table4.index:
            return table4.iloc[row_index]
        else:
            raise IndexError(f"Row index '{row_index}' is out of range for Table4 in lesson '{lesson_title}'.")

    def get_table5(self, lesson_title: str) -> pd.DataFrame:
        """
        Returns the entire table5 dataframe for a given lesson.
        """
        file_path = f'{DATA_DIR}\\students\\{self._student.id}\\{lesson_title}\\table5.xlsx'
        if os.path.exists(file_path):
            return pd.read_excel(file_path)
        else:
            raise ValueError(f"Table4 not found for lesson '{lesson_title}'.")

    def get_table5_column(self, lesson_title: str, column_name: str) -> pd.Series:
        """
        Returns a specific column from table5 for a given lesson.
        """
        table5 = self.get_table5(lesson_title)
        if column_name in table5.columns:
            return table5[column_name]
        else:
            raise ValueError(f"Column '{column_name}' not found in Table5 for lesson '{lesson_title}'.")

    def get_table5_row(self, lesson_title: str, row_index: int) -> pd.Series:
        """
        Returns a specific row from table5 for a given lesson.
        """
        table5 = self.get_table5(lesson_title)
        if 0<= row_index < len(table5):
            return table5.iloc[row_index]
        else:
            raise ValueError(f"Row index '{row_index}' is out of range for Table5 in lesson '{lesson_title}'.")


# ========================== Main Functions
if __name__ == '__main__':
    # For every lesson title in LESSON_NAMES,
    for lessonTitle in LESSON_NAMES:

        # Create a 'Lesson' object to represent them and insert it into ALL_LESSON_OBJECTS.
        lessonObj = Lesson(lessonTitle)
        ALL_LESSON_OBJECTS.append(lessonObj)

    # After the 'Lesson' class is called, the 'students' folder will be filled with all students.
    # For every student folder in students folder,
    for studentID in os.listdir(f'{DATA_DIR}\\students'):

        # Create a 'Student' object to represent them.
        studentObj = Student(int(studentID))
