from math import radians, sin, cos
from terms import *


class MoonCorrection:
    def __init__(self, t):
        # T dalam TD

        self._t         = t
        
        #bujur rata-rata bulan
        self.L_aksen    =radians((218.3164591 + 481267.88134236*t - 0.0013268*t**2 + t**3/538841 - t**4/65194000) % 360)
        
        #elongasi rata-rata bulan
        self.d          = radians((297.8502042 + 445267.1115168*t - 0.00163*t**2 + t**3/545868 - t**4/113065000) % 360)
        
        # Anomali rata-rata matari
        self.m          = radians((357.5291092 + 35999.0502909*t - 0.0001536*t**2 + t**3/24490000) % 360)
        
        # Anomali rata-rata bulan
        self.m_aksen    = radians((134.9634114 + 477198.8676313*t + 0.008997*t**2 + t**3/69699 - t**4/14712000) % 360)

        # argumen bujur bulan
        self.f          = radians((93.2720993 + 483202.0175273*t - 0.0034029*t**2 - t**3/3526000 + t**4/863310000) % 360)

        # eksentrisitas orbit
        self.e          = 1 - 0.002516*t - 0.0000074*t**2

        self.A1         = radians((119.75 + 131.849 * t) % 360)
        self.A2         = radians((53.09 + 479264.29 * t) % 360)
        self.A3         = radians((313.45 + 481266.484 * t) % 360)


    @property
    def koreksi_bujur_bulan(self):
        kb = list_koreksi_bujur_bulan
        total = 0

        for i in range(len(kb)):
            total += kb[i][4] * (self.e**(abs(kb[i][1]))) * sin(kb[i][0] * self.d+kb[i][1] * self.m + kb[i][2] * self.m_aksen+kb[i][3] * self.f)
        
        total_koreksi_bujur_bulan = total

        return (total_koreksi_bujur_bulan + 3958 * sin(self.A1)+1962 * sin(self.L_aksen-self.f) + 318 * sin(self.A2)) / 1000000


    @property
    def lintang_bulan(self):
        #koreksi lintang bulan = lintang bulan
        kl = list_koreksi_lintang_bulan

        total_koreksi_lintang_bulan = 0
        for i in range(len(kl)):
            total_koreksi_lintang_bulan += kl[i][4] * (self.e**(abs(kl[i][1]))) * sin(kl[i][0] * self.d+kl[i][1] * self.m + kl[i][2] * self.m_aksen+kl[i][3] * self.f)
        
        return radians((total_koreksi_lintang_bulan - 2235 * sin(self.L_aksen) + 382*sin(self.A3) + 175*sin(self.A1 - self.f) + 175*sin(self.A1 + self.f) + 127 * sin(self.L_aksen - self.m_aksen) - 115 * sin(self.L_aksen + self.m_aksen))/1000000)


    @property
    def jarak_bumi_bulan(self):
        kj = list_koreksi_jarak_bumi_bulan

        total_koreksi_jarak_bumi_bulan = 0
        for i in range(len(kj)):
            total_koreksi_jarak_bumi_bulan += kj[i][4] * (self.e**(abs(kj[i][1]))) * cos(kj[i][0] * self.d + kj[i][1] * self.m + kj[i][2] * self.m_aksen + kj[i][3] * self.f)

        koreksi_jarak_bumi_bulan = total_koreksi_jarak_bumi_bulan / 1000
        return 385000.56 + koreksi_jarak_bumi_bulan
