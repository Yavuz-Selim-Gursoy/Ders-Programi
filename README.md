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
