# ===============================================
import os
import pandas as pd
from pathlib import Path

ROOT_DIR = Path(__file__).parent.parent.absolute()
DATA_DIR = f'{ROOT_DIR}\\data'
LESSON_NAMES = os.listdir(f'{DATA_DIR}\\lessons')
# ===============================================


def rearrange_dataframe(excelPath: str) -> pd.DataFrame:
    # Halihazırda bulunan yüzdelik (yüzdeliğin getter fonksiyonunu böyle yazabilirsin)
    dfUnedited = pd.read_excel(excelPath, sheet_name=0)

    # Yüzdelikleri al
    currentPercentages = dfUnedited.iloc[0, 1:-1].tolist()

    # Başa ve sona birer eleman ekle ki sütun sayısı ile yüzdelik sayısı aynı olsun
    currentPercentages.insert(0, pd.NA)
    currentPercentages.append(pd.NA)

    # Gerçek sütun isimleri
    dfActualColumnNames = dfUnedited.iloc[1, :].tolist()

    # Yüzdelik satırının df'e çevirilmiş hali
    currentPercentagesDf = pd.DataFrame([currentPercentages], columns=dfActualColumnNames)  # Tek satırlı bir DataFrame olarak oluştur

    # Yüzdelikleri içermeyen, sadece verileri içeren df
    dfEdited = pd.DataFrame(dfUnedited.iloc[2:, :].values, columns=dfActualColumnNames)

    # currentPercentagesDf'i dfEdited'in başına ekleyerek son df'i elde et
    dfFinal = pd.concat([currentPercentagesDf, dfEdited], ignore_index=True)

    return dfFinal

# =============================================== Setter/Getters
def get_column_names(lessonTitle: str) -> list:
    if lessonTitle in LESSON_NAMES:
        df = rearrange_dataframe(f"{DATA_DIR}\\lessons\\{lessonTitle}\\table2.xlsx")
        columnsList = df.columns.tolist()
        columnsList.pop(0), columnsList.pop(-1)
        return columnsList

def get_percentage(lessonTitle: str, columnName: str) -> int:
    if lessonTitle in LESSON_NAMES:
        df = rearrange_dataframe(f"{DATA_DIR}\\lessons\\{lessonTitle}\\table2.xlsx")
        currentPercentage = df.loc[0, columnName]
        return currentPercentage

def set_percentage(lessonTitle: str, columnName: str, percentage: int) -> None:
    if lessonTitle in LESSON_NAMES:
        df = rearrange_dataframe(f"{DATA_DIR}\\lessons\\{lessonTitle}\\table2.xlsx")
        df.loc[0, columnName] = percentage

def set_column_name(lessonTitle: str, new_name: str, old_name: str) -> pd.DataFrame:
    file_path = f"{DATA_DIR}\\lessons\\{lessonTitle}\\table2.xlsx"
    df = pd.read_excel(file_path)

    if old_name not in df.columns:
        print(f"'{old_name}' kolon adı bulunamadı!")
        return df  # Return the original DataFrame if the column is not found
    else:
        df.rename(columns={old_name: new_name}, inplace=True)
        return df

def write_to_excel(lessonTitle: str, df:pd.DataFrame) -> None:
    file_path = f"{DATA_DIR}\\lessons\\{lessonTitle}\\table2.xlsx"

    # indexleri sıfırlama bir nevi ignore yerine sayılabilir
    df.reset_index(drop=True, inplace=True)

    df.to_excel(file_path, index=False)
