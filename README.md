# Ders Programı
 Programlama Lab. Ödev 2

Tek kural, fonksiyonlara bir "parameters" ve "returns" yorum kısmı eklenmeli ve yorumlar kod içerisinde bolca bulunmalı. Örnek fonksiyon:

    def create_ball(canvas: Canvas, ballSize: tuple, ballColor: str) -> None:
        Description: Helper function that creates a ball using entered parameters.
        [AFFECTS GLOBAL SCOPE VARIABLES] -> BALL_OBJECTS

        Parameters:
            canvas (Canvas): The canvas on which the ball is created.
            ballSize (tuple): Size of the ball as a tuple of coordinates (x0, y0, x1, y1).
            ballColor (str): Color of the ball.

        Returns:
            None

Yapılacaklar:
1) Tablo1 ve Tablo2 çekilerek dataframe'e koyulacak. 
2) Tablolar için setter ve getter fonkisyonları yazılacak.
3) Tablo1 deki "İlişki Değeri" sütunu için satırdaki elemanların toplamının sütun sayısına bölünmesi gerek. Bunun için fonksiyon yazılmalı.
4) Tablo2'de minimum 3 sütun olabilir, birinci satır yüzdelik ağırlıkları belirtirken sütun sayısı dönem içerisindeki değerlendirme görevlerini(ödev, sınav gibi) tutmalı. Tablo2'deki her bir satır bir ders çıktısına eşit, bu çıktıların hangi değerlendirme görevleri ile ne kadar ilişkili olduğunu tutuyor. Toplam sütunu satırdaki elemanlar toplanarak bulunuyor.
5) Tablo3'teki değerler Tablo2'deki ödevin yüzdeliği ile ders çıktısının ölçüldüğü karoların çarpılmasıyla elde ediliyor. Tablo3'ü buna göre doldurmalıyız. Toplam sütunu satırdaki elemanlar toplanarak bulunuyor.
6) TabloNot öğrencilerin dönem içerisindeki görevlerden aldığı notları içeriyor. Ortalama sütunu satırın aritmetik ortalaması ile bulunuyor.
7) Tablo4 her öğrenci için ayrı yapılmalı, Tablo3 ve TabloNot tablolarındaki değerler çarpılarak karolara yerleştirilecek. MAX sütunu alınabilecek maksimum notu, Başarı sütunu ise alınmış not ile max notun oranıyla bulunuyor.
8) Tablo4 teki başarı sütunu Tablo5'te ana satıra dönüşüyor (yani Tablo5'te her öğrenciye özel oluyor). Tablo5 doldurulurken Tablo1'deki satırlar, Tablo5'in ana satırı ve 100 çarpılıyor. Buradaki başarı oranı ise Tablo5 satır toplamı / (Tablo1 satırı * 100) şeklinde hesaplanıyor.

1,2,3 -> Emrah
4,5,6,7
