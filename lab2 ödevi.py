import pandas as pd

class tablo1:
    """
       Tablo1 Sınıfı:

       Bu sınıf, bir Excel dosyasını okuyarak veriler üzerinde işlem yapar ve yeni bir
       "İliski Deger" sütunu ekler. Bu sütun, her satırdaki sayısal değerlerin ortalaması
       hesaplanarak elde edilir. İşlem tamamlandıktan sonra, sonuç belirtilen bir dosya
       yoluna kaydedilir. Ayrıca, kullanıcıların işlenmiş tabloya erişebilmesi ve
       "İliski Deger" sütununu ayrı olarak alabilmesi için yardımcı metotlar sunar.

       Özellikler:
           - file_path (str): Okunacak Excel dosyasının yolu.
           - output_file_path (str): İşlenmiş dosyanın kaydedileceği yol.
           - result_df (pd.DataFrame): İşlem görmüş veriyi tutar.

       Yöntemler:
           - create_tablo1(): Veriyi işler, ortalamaları hesaplar, sütunu ekler ve kaydeder.
           - get_iliski_deger(): İşlenmiş tablodaki "İliski Deger" sütununu döner.
       """
    def __init__(self, file_path: str, output_file_path: str):
        self.file_path = file_path
        self.output_file_path = output_file_path
        self.result_df = None

    def create_tablo1(self):
        df = pd.read_excel(self.file_path, sheet_name=0)

        # Ilk satır ve sütunu geçme
        proccessed_df = df.iloc[1:, 1:]

        # Iliski değeri hesapla
        proccessed_df = proccessed_df.apply(
            lambda row: row.map(lambda x: x if isinstance(x, (int, float)) else 0).sum() / len(row), axis=1)

        # yeni satırı ekleme
        self.result_df = df.loc[1:, 'İliski Deger'] = proccessed_df.values

        # gereksiz unnamed sütunlarını kaldırır
        df.columns = [col if "Unnamed" not in col else "" for col in df.columns]
        self.result_df = df

        # Yeni dosya çıktısı
        self.result_df.to_excel(self.output_file_path, index=False)

    def get_ılıski_Deger(self):
        """Tablonun ilişki değer sütununu döner"""
        if self.result_df is not None and 'İliski Deger' in self.result_df.columns:
            return self.result_df['İliski Deger']
        return None




class tablo3:
    """
        Tablo3 Sınıfı:

        Bu sınıf, verilen bir Excel dosyasını okuyarak ağırlık değerlerini işler, her satır için
        ağırlıklı ortalama hesaplar ve "Toplam" sütunu ekler. Kullanıcıların tablo içerisindeki
        belirli verilere erişebilmesi, ağırlıkları alabilmesi veya toplam sütununa ulaşabilmesi için
        yardımcı metotlar sağlar.

        Özellikler:
            - file_path (str): Okunacak Excel dosyasının yolu.
            - weights (pd.Series): İlk satırdan alınan ağırlık değerleri.
            - result_df (pd.DataFrame): İşlenmiş veri tablosunu tutar.

        Yöntemler:
            - from_tablo2_to_tablo3(output_file_path: str):
              Excel dosyasını işler, ağırlıklı değerleri hesaplar ve sonucu yeni bir dosyaya kaydeder.
            - get_weights(): Tablo içindeki ağırlık değerlerini döner.
            - get_toplam_coulmn(): "Toplam" sütununu döner.
            - get_specified_place(row: int, column: int):
              Tablo içindeki belirtilen bir satır ve sütundaki değeri döner. (Başlıklar dikkate alınmaz.)
    """
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.weights = None
        self.result_df = None

    def from_tablo2_to_tablo3(self, output_file_path: str) -> None:
        """
        """

        df_tablo2 = pd.read_excel(self.file_path, sheet_name=0)

        # ağırlıkları almak
        self.weights = pd.to_numeric(df_tablo2.iloc[0,1:], errors='coerce')

        # temizlemek
        df_clean = df_tablo2.iloc[2:].reset_index(drop=True)
        df_clean.columns = ['Ders Çıktı'] + list(self.weights.index)
        df_clean.iloc[:, 1:] = df_clean.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')

        # ağırlıkları hesaplamak
        weighted_values = df_clean.iloc[:, 1:].multiply(self.weights, axis=1).div(100)
        df_clean['Toplam'] = weighted_values.sum(axis=1)

        self.result_df = pd.concat([df_clean.iloc[:,:1], weighted_values, df_clean[['Toplam']]], axis=1)

        # gpt kodu
        self.result_df.columns = [col if "Unnamed" not in col else "" for col in self.result_df.columns]

        self.result_df.to_excel(output_file_path, index=False)

    def get_weights(self):
        """Tablodaki ağırlıkları Döner"""
        return self.weights

    def get_toplam_coulmn(self):
        """Toplam kolonunu döner"""
        if self.result_df is not None and 'Toplam' in self.result_df.columns:
            return self.result_df['Toplam']

    def get_specified_place(self, row: int, column: int):
        """tablodaki belli bir spesifik bir yeri döner
        Ama YAVUZ tablonun en başındaki başlıkları görmezden gelerek
        bir spesifik yer vermen gerekiyor"""
        if self.result_df is not None:
            try:
                return self.result_df.iloc[row, column]
            except Exception as err:
                return err









