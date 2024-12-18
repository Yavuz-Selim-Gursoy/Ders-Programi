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
        id (int): Lesson id.
        title (str) : Lesson title.
        inputFolderName (str) : Lesson folder name.
        tableOneDataFrame (pd.DataFrame) : Lesson's table one dataframe.
        tableTwoDataFrame (pd.DataFrame) : Lesson's table two dataframe.
        TODO: tableThreeDataFrame (pd.DataFrame) : Lesson's table three dataframe. *MEMBER FUNC* builds this dataframe and writes it into 'table3.xlsx'.
    Member Functions:
        TODO: setter and getter functions
        _create_dataframe_from_table(self): Creates table three dataframe.

    """
    def __init__(self, id: int, title: str, inputFolderName: str):
        self.id = id
        self.title = title
        self.inputFolderPath = f'{DATA_DIR}\\lessons\\{inputFolderName}'
        self.tableOneDataFrame = self._create_dataframe_from_table(1)
        self.tableTwoDataFrame = self._create_dataframe_from_table(2)
        self.tableThreeDataFrame = self._create_dataframe_from_table(3)

    def _create_dataframe_from_table(self, tableNum) -> pd.DataFrame:
        """
        TODO: it would be awesome if we could create table3 using only this function too.
        Description: Function to create dataframe from tables
        Parameters:
            tableNum (int) : Lesson's table number.
        Returns:
            pd.DataFrame : Lesson's table dataframe."""

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

            # If the grading percentages do not sum up to 100, assert ValueError.
            assert not df.iloc[0, 1:].sum() != 100, "Grading percentages must sum to 100."

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
            df2 = pd.read_excel(f"{self.inputFolderPath}\\table2.xlsx", sheet_name=0)
            recalculatedDf2 = (df2.iloc[2:, 1:-1] * df2.iloc[0, 1:-1]) / 100
            dersCiktiColumn = df2.iloc[1:, 0].to_list()
            firstRow = df2.iloc[1, 1:-1].to_list()

            print('_____________________________')
            print(dersCiktiColumn)
            df = df2
            rowCount, colCount = len(df2), len(df2.columns)
            df3 = pd.DataFrame(recalculatedDf2.values, columns=['TABLO 3'] + ['Ağırlıklı Değerlendirme'] * (len(df2.columns) - 3))
            df3 = pd.concat([pd.DataFrame([firstRow], columns=df3.columns), df3], ignore_index=True)

            if 'Toplam' in df3.columns:
                # If column 'Toplam' already exists, it'll cause recursion and the column values won't be correct.
                df3['Toplam'] = df3.iloc[0:, 0:-1].sum(axis=1)

            else:
                # If column doesn't exist, recursion won't happen.
                df3['Toplam'] = df3.iloc[0:, 0:].sum(axis=1)

            toplamIndex = df3.columns.tolist().index("Toplam")
            df3.iloc[0, toplamIndex] = None  # Üstteki başlığı boş yap
            df3.iloc[0, toplamIndex] = "Toplam"  # Bir alt satıra taşı

            if 'Ders Çıktı' in df3.columns:
                pass

            else:
                df3.insert(0, 'Ders Çıktı', dersCiktiColumn)
            print(df3.to_string())






        # Remove 'unnamed' from tables.
        df.columns = [col if "Unnamed" not in col else "" for col in df.columns]
        resultDf = df

        # Rewrite the table and return the dataframe.
        resultDf.to_excel(f"{self.inputFolderPath}\\table{tableNum}.xlsx", index=False)
        return resultDf
        
    def _organize_directory(self, option) -> None:
        """option 'reorganize' creates reorganized files while option 'delete' deletes initial files."""
        if option == 'reorganize':
            if 'table1_reorganized.xlsx' not in os.listdir(self.inputFolderPath):
                self.tableOneDataFrame = pd.read_excel(f"{self.inputFolderPath}\\table1.xlsx", header=[0, 1])
                self.tableTwoDataFrame = pd.read_excel(f"{self.inputFolderPath}\\table2.xlsx", header=[0, 1])

            else:
                self.tableOneDataFrame = pd.read_excel(f"{self.inputFolderPath}\\table1_reorganized.xlsx",
                                                       header=[0, 1])
                self.tableTwoDataFrame = pd.read_excel(f"{self.inputFolderPath}\\table2_reorganized.xlsx",
                                                       header=[0, 1])
        if option == 'delete':
            if 'table1.xlsx' in os.listdir(self.inputFolderPath):
                os.remove(f"{self.inputFolderPath}\\table1.xlsx")
                os.remove(f"{self.inputFolderPath}\\table2.xlsx")



    def _organize_table_one_and_two(self) -> None:
        """Organizes table one and two dataframe, then rewrites table one and table two dataframe to their parent excel file."""
        # Appends column 'İlişki Değeri' to table1's dataframe.
        self.tableOneDataFrame['İlişki Değeri'] = self.tableOneDataFrame.iloc[1:, 1:].sum(axis=1) / (self.tableOneDataFrame.shape[1] - 1)

        # Appends column 'Toplam' to table2's dataframe.
        self.tableTwoDataFrame['Toplam'] = self.tableTwoDataFrame.iloc[2:, 1:].sum(axis=1)

        self.tableOneDataFrame.to_excel(f"{self.inputFolderPath}\\table1_reorganized.xlsx", index=True)
        self.tableTwoDataFrame.to_excel(f"{self.inputFolderPath}\\table2_reorganized.xlsx", index=True)


    def _create_table_three_dataframe(self) -> None:
        ...



    def check_tables(self) -> None:
        print(self.tableOneDataFrame.iloc[1:, 1:].sum(axis=1))
        print("Tablo1: \n", self.tableOneDataFrame.iloc[0:, 0:])
        print("Tablo2: \n", self.tableTwoDataFrame.iloc[0:, 0:])



lesson1 = Lesson(220, 'BilgisayarMimarisi', 'BilgisayarMimarisi')

class Student:
    """
    Description: Student class.
    Attributes:
    Member Functions:
    """



