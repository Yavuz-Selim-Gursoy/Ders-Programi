# ==========================
import pandas as pd
import os
from pathlib import Path
import numpy as np
ROOT_DIR = Path(__file__).parent.parent.absolute()
DATA_DIR = f'{ROOT_DIR}\\data'
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
        self.lessonStudents = self.tableGradesDataFrame.iloc[0:, 0].to_list()


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

            # Calculate column 'Toplam'
            if 'Toplam' in df.columns:
                # If column 'Toplam' already exists, it'll cause recursion and the column values won't be correct.
                df['Toplam'] = df.iloc[2:, 1:-1].sum(axis=1)

            else:
                # If column doesn't exist, recursion won't happen.
                df['Toplam'] = df.iloc[2:, 1:].sum(axis=1)

            # We have to bring actual columns (in this case its 'Toplam') down in order to work with them.
            toplamIndex = df.columns.tolist().index('Toplam')
            df.iloc[1, toplamIndex] = "Toplam"

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
            df3 = pd.DataFrame(recalculatedDf2.values, columns=['TABLO 3'] + ['Ağırlıklı Değerlendirme'] * (len(df2.columns) - 3))

            # Insert the actual keys into first row.
            df3 = pd.concat([pd.DataFrame([firstRow], columns=df3.columns), df3], ignore_index=True)

            if 'Toplam' in df3.columns:
                # If column 'Toplam' already exists, it'll cause recursion and the column values won't be correct.
                df3['Toplam'] = df3.iloc[0:, 0:-1].sum(axis=1)

            else:
                # If column doesn't exist, recursion won't happen.
                df3['Toplam'] = df3.iloc[0:, 0:].sum(axis=1)

            # Bring down the column 'Toplam'.
            toplamIndex = df3.columns.tolist().index("Toplam")
            df3.iloc[0, toplamIndex] = None
            df3.iloc[0, toplamIndex] = "Toplam"

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

            if 'Ortalama' not in df.columns:
                # Read table2 and crop the values.
                df2 = pd.read_excel(f"{self.inputFolderPath}\\table2.xlsx", sheet_name=0)
                df2_cropped = df2.iloc[2:, 1:-1]

                # Identify the weights from table2.
                weights = df2.iloc[0, 1:-1].to_list()

                # Calculate the weighted scores using weights and grades.
                weighted_scores = df.iloc[:, 1:].mul(weights, axis=1) / 100

                # Create the column 'Ortalama' and insert the sum of the row's weighted scores.
                df["Ortalama"] = weighted_scores.sum(axis=1)

                # If table2 and grades table column counts doesn't match, assert an error.
                assert len(df2_cropped.columns) == (len(df.columns) - 2), "Table2 and TableGrades column counts must be the same."

            else:
                pass

        # If wrong option is chosen, assert an error.
        else:
            assert 0 <= tableNum <= 3, "Wrong input table number. Options are: 0, 1, 2 or 3."

        # Remove 'unnamed' from tables.
        df.columns = [col if "Unnamed" not in col else "" for col in df.columns]
        resultDf = df

        # Rewrite the table and return the dataframe.
        # If tableNum is 0, change its filename to grades.xlsx.
        if tableNum != 0:
            resultDf.to_excel(f"{self.inputFolderPath}\\table{tableNum}.xlsx", index=False)

        elif tableNum == 0:
            resultDf.to_excel(f"{self.inputFolderPath}\\grades.xlsx", index=False)
        
        return resultDf


    def _check_tables(self) -> None:
        print(self.tableOneDataFrame.to_string())
        print('-' * 130)
        print(self.tableTwoDataFrame.to_string())
        print('-' * 130)
        print(self.tableThreeDataFrame.to_string())
        print('-' * 130)
        print(self.tableGradesDataFrame.to_string())





class Student:
    """
    Description: Student class.
    Attributes:
        id (int): Student id.

    Member Functions:
    """
lesson1 = Lesson('BilgisayarMimarisi')
lesson1._check_tables()


