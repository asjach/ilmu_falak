from math import radians, sin, cos, tan, asin, acos, atan2, degrees
from nutation import *



class Moon:
    def __init__(self, t_td, lst_nampak, latitude) -> None:
        self.t_td = t_td
        T = t_td
        self.d                      = radians((297.8502042 + 445267.1115168*T - 0.00163*T*T + T*T*T/545868 - T*T*T*T/113065000) % 360)
        self.m                      = radians((357.5291092 + 35999.0502909*T - 0.0001536*T*T + T*T*T/24490000) % 360)
        self.m_aksen                = radians((134.9634114 + 477198.8676313*T + 0.008997*T*T + T*T*T/69699 - T*T*T*T/14712000) % 360)
        self.f                      = radians((93.2720993 + 483202.0175273*T - 0.0034029*T*T - T*T*T/3526000 + T*T*T*T/863310000) % 360)
        self.eksentrisitas_orbit    = 1 - 0.002516*T - 0.0000074*T*T
        self.nutasi                 = Nutation(self.t_td)
        self.koreksi_delta_psi      = self.nutasi.delta_psi_total/3600
        self.lst_nampak             = lst_nampak
        self.latitude               = latitude                          

        #Argument
        self.a1 = radians((119.75 + 131.849 * T) % 360)
        self.a2 = radians((53.09 + 479264.29 * T) % 360)
        self.a3 = radians((313.45 + 481266.484 * T) % 360)


    @property
    def bujur_rata_rata_bulan(self):
        t = self.t_td
        hasil = radians((218.3164591 + 481267.88134236*t - 0.0013268*t*t + t*t*t/538841 - t*t*t*t/65194000) % 360)
        return hasil


    @property
    def bujur_bulan(self):
        
        return (degrees(self.bujur_rata_rata_bulan) + self.koreksi_bujur_bulan)%360


    @property
    def bujur_bulan_nampak(self):
        return radians((self.bujur_bulan + self.koreksi_delta_psi) % 360)


    @property
    def koreksi_bujur_bulan(self):
        a1 = self.a1
        a2 = self.a2
        f = self.f
        hasil = (self.total_koreksi_bujur_bulan + 3958 * sin(a1)+1962 * sin(self.bujur_rata_rata_bulan-f) + 318 * sin(a2)) / 1000000
        return hasil

    @property
    def total_koreksi_bujur_bulan(self):
        kb = list_koreksi_bujur_bulan
        d = self.d
        m = self.m
        m_aksen = self.m_aksen
        f = self.f
        eksentrisitas = self.eksentrisitas_orbit

        total = 0
        for i in range(len(kb)):
            total += kb[i][4] * (eksentrisitas**(abs(kb[i][1]))) * sin(kb[i][0] * d+kb[i][1] * m + kb[i][2] * m_aksen+kb[i][3] * f)
        return total

    @property
    def lintang_bulan(self):
        m_aksen = self.m_aksen
        f = self.f
        L = self.bujur_rata_rata_bulan
        hasil = (self.total_koreksi_lintang_bulan - 2235 * sin(L) + 382*sin(self.a3) + 175*sin(self.a1 - f) + 175*sin(self.a1 + f) + 127*sin(L - m_aksen) - 115*sin(L + m_aksen))/1000000
        return radians(hasil)


    @property
    def total_koreksi_lintang_bulan(self):
        kl = list_koreksi_lintang_bulan
        d = self.d
        m = self.m
        m_aksen = self.m_aksen
        f = self.f
        eksentrisitas = self.eksentrisitas_orbit
        hasil = 0
        for i in range(len(kl)):
            hasil += kl[i][4] * (eksentrisitas**(abs(kl[i][1]))) * sin(kl[i][0] * d+kl[i][1] * m + kl[i][2] * m_aksen+kl[i][3] * f)
        return hasil


    @property
    def total_koreksi_jarak_bumi_bulan(self):
        kj = list_koreksi_jarak_bumi_bulan
        eksentrisitas = self.eksentrisitas_orbit
        d = self.d
        m = self.m
        m_aksen = self.m_aksen
        f = self.f

        total = 0
        for i in range(len(kj)):
            total += kj[i][4] * (eksentrisitas**(abs(kj[i][1]))) * cos(kj[i][0] * d + kj[i][1] * m + kj[i][2] * m_aksen + kj[i][3] * f)
        return total

    @property
    def koreksi_jarak_bumi_bulan(self):
        hasil = self.total_koreksi_jarak_bumi_bulan / 1000
        return hasil


    @property
    def jarak_bumi_bulan(self):
        return 385000.56+self.koreksi_jarak_bumi_bulan


    @property
    def sudut_parallaks(self):
        return degrees(asin(6378.14/self.jarak_bumi_bulan))


    @property
    def sudut_jari_jari_bulan(self):
        return 358473400/(self.jarak_bumi_bulan*3600)


    @property
    def alpha(self):
        x_num = cos(self.bujur_bulan_nampak)
        y_num = sin(self.bujur_bulan_nampak) * cos(self.nutasi.epsilon) - tan(self.lintang_bulan) * sin(self.nutasi.epsilon)
        hasil = (degrees(atan2(y_num, x_num))) % 360
        #hasil = ((atan2(y_num, x_num))) % 360
        return hasil


    @property
    def delta(self):
        hasil =degrees(asin(sin(self.lintang_bulan)*cos(self.nutasi.epsilon)+cos(self.lintang_bulan)*sin(self.nutasi.epsilon)*sin(self.bujur_bulan_nampak)))
        return hasil

    @property
    def hour_angle(self):
        return self.lst_nampak * 15 - self.alpha


    @property
    def azimut_selatan(self):
        y = sin(radians(self.hour_angle))
        x = cos(radians(self.hour_angle))*sin(radians(self.latitude))-tan(radians(self.delta))*cos(radians(self.latitude))
        return degrees(atan2(y, x))


    @property
    def azimut_bulan(self):
        return (self.azimut_selatan + 180) % 360


    @property
    def altitude(self):
        lat = radians(self.latitude)
        delta = radians(self.delta)
        ha = radians(self.hour_angle)
        return degrees(asin(sin(lat)*sin(delta)+cos(lat)*cos(delta)*cos(ha)))


    @property
    def sudut_fai(self):
        delta_b = radians(self.delta)
        delta_m = 00000000000000000000000000000000000000
        return acos(sin(delta_b)*sin(delta_m) + cos(delta_b)*cos(delta_m)*cos(delta_b-delta_m))


    @property
    def sudut_fase(self):
        pass


    @property
    def illuminasi(self):
        pass



list_koreksi_bujur_bulan = [
    [0, 0, 1, 0, 6288774],
    [2, 0, -1, 0, 1274027],
    [2, 0, 0, 0, 658314],
    [0, 0, 2, 0, 213618],
    [0, 1, 0, 0, -185116],
    [0, 0, 0, 2, -114332],
    [2, 0, -2, 0, 58793],
    [2, -1, -1, 0, 57066],
    [2, 0, 1, 0, 53322],
    [2, -1, 0, 0, 45758],
    [0, 1, -1, 0, -40923],
    [1, 0, 0, 0, -34720],
    [0, 1, 1, 0, -30383],
    [2, 0, 0, -2, 15327],
    [0, 0, 1, 2, -12528],
    [0, 0, 1, -2, 10980],
    [4, 0, -1, 0, 10675],
    [0, 0, 3, 0, 10034],
    [4, 0, -2, 0, 8548],
    [2, 1, -1, 0, -7888],
    [2, 1, 0, 0, -6766],
    [1, 0, -1, 0, -5163],
    [1, 1, 0, 0, 4987],
    [2, -1, 1, 0, 4036],
    [2, 0, 2, 0, 3994],
    [4, 0, 0, 0, 3861],
    [2, 0, -3, 0, 3665],
    [0, 1, -2, 0, -2689],
    [2, 0, -1, 2, -2602],
    [2, -1, -2, 0, 2390],
    [1, 0, 1, 0, -2348],
    [2, -2, 0, 0, 2236],
    [0, 1, 2, 0, -2120],
    [0, 2, 0, 0, -2069],
    [2, -2, -1, 0, 2048],
    [2, 0, 1, -2, -1773],
    [2, 0, 0, 2, -1595],
    [4, -1, -1, 0, 1215],
    [0, 0, 2, 2, -1110],
    [3, 0, -1, 0, -892],
    [2, 1, 1, 0, -810],
    [4, -1, -2, 0, 759],
    [0, 2, -1, 0, -713],
    [2, 2, -1, 0, -700],
    [2, 1, -2, 0, 691],
    [2, -1, 0, -2, 596],
    [4, 0, 1, 0, 549],
    [0, 0, 4, 0, 537],
    [4, -1, 0, 0, 520],
    [1, 0, -2, 0, -487],
    [2, 1, 0, -2, -399],
    [0, 0, 2, -2, -381],
    [1, 1, 1, 0, 351],
    [3, 0, -2, 0, -340],
    [4, 0, -3, 0, 330],
    [2, -1, 2, 0, 327],
    [0, 2, 1, 0, -323],
    [1, 1, -1, 0, 299],
    [2, 0, 3, 0, 294]
]


list_koreksi_lintang_bulan = [
    [0, 0, 0, 1, 5128122],
    [0, 0, 1, 1, 280602],
    [0, 0, 1, -1, 277693],
    [2, 0, 0, -1, 173237],
    [2, 0, -1, 1, 55413],
    [2, 0, -1, -1, 46271],
    [2, 0, 0, 1, 32573],
    [0, 0, 2, 1, 17198],
    [2, 0, 1, -1, 9266],
    [0, 0, 2, -1, 8822],
    [2, -1, 0, -1, 8216],
    [2, 0, -2, -1, 4324],
    [2, 0, 1, 1, 4200],
    [2, 1, 0, -1, -3359],
    [2, -1, -1, 1, 2463],
    [2, -1, 0, 1, 2211],
    [2, -1, -1, -1, 2065],
    [0, 1, -1, -1, -1870],
    [4, 0, -1, -1, 1828],
    [0, 1, 0, 1, -1794],
    [0, 0, 0, 3, -1749],
    [0, 1, -1, 1, -1565],
    [1, 0, 0, 1, -1491],
    [0, 1, 1, 1, -1475],
    [0, 1, 1, -1, -1410],
    [0, 1, 0, -1, -1344],
    [1, 0, 0, -1, -1335],
    [0, 0, 3, 1, 1107],
    [4, 0, 0, -1, 1021],
    [4, 0, -1, 1, 833],
    [0, 0, 1, -3, 777],
    [4, 0, -2, 1, 671],
    [2, 0, 0, -3, 607],
    [2, 0, 2, -1, 596],
    [2, -1, 1, -1, 491],
    [2, 0, -2, 1, -451],
    [0, 0, 3, -1, 439],
    [2, 0, 2, 1, 422],
    [2, 0, -3, -1, 421],
    [2, 1, -1, 1, -366],
    [2, 1, 0, 1, -351],
    [4, 0, 0, 1, 331],
    [2, -1, 1, 1, 315],
    [2, -2, 0, -1, 302],
    [0, 0, 1, 3, -283],
    [2, 1, 1, -1, -229],
    [1, 1, 0, -1, 223],
    [1, 1, 0, 1, 223],
    [0, 1, -2, -1, -220],
    [2, 1, -1, -1, -220],
    [1, 0, 1, 1, -185],
    [2, -1, -2, -1, 181],
    [0, 1, 2, 1, -177],
    [4, 0, -2, -1, 176],
    [4, -1, -1, -1, 166],
    [1, 0, 1, -1, -164],
    [4, 0, 1, -1, 132],
    [1, 0, -1, -1, -119],
    [4, -1, 0, -1, 115],
    [2, -2, 0, 1, 107]
]

list_koreksi_jarak_bumi_bulan = [
    [0, 0, 1, 0, -20905355],
    [2, 0, -1, 0, -3699111],
    [2, 0, 0, 0, -2955968],
    [0, 0, 2, 0, -569925],
    [2, 0, -2, 0, 246158],
    [2, -1, 0, 0, -204586],
    [2, 0, 1, 0, -170733],
    [2, -1, -1, 0, -152138],
    [0, 1, -1, 0, -129620],
    [1, 0, 0, 0, 108743],
    [0, 1, 1, 0, 104755],
    [0, 0, 1, -2, 79661],
    [0, 1, 0, 0, 48888],
    [4, 0, -1, 0, -34782],
    [2, 1, 0, 0, 30824],
    [2, 1, -1, 0, 24208],
    [0, 0, 3, 0, -23210],
    [4, 0, -2, 0, -21636],
    [1, 1, 0, 0, -16675],
    [2, 0, -3, 0, 14403],
    [2, -1, 1, 0, -12831],
    [4, 0, 0, 0, -11650],
    [2, 0, 2, 0, -10445],
    [2, 0, 0, -2, 10321],
    [2, -1, -2, 0, 10056],
    [2, -2, 0, 0, -9884],
    [2, 0, -1, -2, 8752],
    [1, 0, -1, 0, -8379],
    [0, 1, -2, 0, -7003],
    [1, 0, 1, 0, 6322],
    [0, 1, 2, 0, 5751],
    [2, -2, -1, 0, -4950],
    [0, 0, 2, -2, -4421],
    [2, 0, 1, -2, 4130],
    [4, -1, -1, 0, -3958],
    [3, 0, -1, 0, 3258],
    [0, 0, 0, 2, -3149],
    [2, 1, 1, 0, 2616],
    [2, 2, -1, 0, 2354],
    [0, 2, -1, 0, -2117],
    [4, -1, -2, 0, -1897],
    [1, 0, -2, 0, -1739],
    [4, -1, 0, 0, -1571],
    [4, 0, 1, 0, -1423],
    [0, 2, 1, 0, 1165],
    [0, 0, 4, 0, -1117]
]