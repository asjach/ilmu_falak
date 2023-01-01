from math import acos, asin, atan2, degrees, radians, sin, cos, tan
from nutation import Nutation
from sun import SunProperty
from sun import SunEquatorial
from terms import *


class MoonCorrection:
    def __init__(self, t):
        self._t = t
        self.L_aksen = radians((218.3164591 + 481267.88134236*t - 0.0013268*t**2 + t**3/538841 - t**4/65194000) % 360)  # bujur rata-rata bulan
        self.d = radians((297.8502042 + 445267.1115168*t - 0.00163*t**2 + t**3/545868 - t**4/113065000) % 360)  # elongasi rata-rata bulan
        self.m = radians((357.5291092 + 35999.0502909*t - 0.0001536*t**2 + t**3/24490000) % 360)  # Anomali rata-rata matari
        self.m_aksen = radians((134.9634114 + 477198.8676313*t + 0.008997*t**2 + t**3/69699 - t**4/14712000) % 360)  # Anomali rata-rata bulan
        self.f = radians((93.2720993 + 483202.0175273*t - 0.0034029*t**2 - t**3/3526000 + t**4/863310000) % 360)  # argumen bujur bulan
        self.e = 1 - 0.002516*t - 0.0000074*t**2  # eksentrisitas orbit

        self.A1 = radians((119.75 + 131.849*t) % 360)
        self.A2 = radians((53.09 + 479264.29*t) % 360)
        self.A3 = radians((313.45 + 481266.484*t) % 360)

    @property
    def koreksi_bujur_bulan(self):
        kb = list_koreksi_bujur_bulan
        total = 0
        for i in range(len(kb)):
            total += kb[i][4] * (self.e**(abs(kb[i][1]))) * sin(kb[i][0] * self.d+kb[i][1] * self.m + kb[i][2] * self.m_aksen+kb[i][3] * self.f)
        total_koreksi_bujur_bulan = total
        return (total_koreksi_bujur_bulan + 3958*sin(self.A1) + 1962*sin(self.L_aksen-self.f) + 318*sin(self.A2))/1000000


    @property
    def lintang_bulan(self):
        #koreksi lintang bulan = lintang bulan
        kl = list_koreksi_lintang_bulan

        total_koreksi_lintang_bulan = 0
        for i in range(len(kl)):
            total_koreksi_lintang_bulan += kl[i][4] * (self.e**(abs(kl[i][1])))*sin(kl[i][0]*self.d+kl[i][1]*self.m + kl[i][2]*self.m_aksen+kl[i][3]*self.f)
        
        hasil = (total_koreksi_lintang_bulan - 2235*sin(self.L_aksen) + 382*sin(self.A3) + 175*sin(self.A1 - self.f) + 175*sin(self.A1 + self.f) + 127*sin(self.L_aksen - self.m_aksen) - 115*sin(self.L_aksen + self.m_aksen)) /1000000

        return hasil


    @property
    def moon_distance(self):
        kj = list_koreksi_jarak_bumi_bulan

        total_koreksi_jarak_bumi_bulan = 0
        for i in range(len(kj)):
            total_koreksi_jarak_bumi_bulan += kj[i][4]*(self.e**(abs(kj[i][1])))*cos(kj[i][0]*self.d + kj[i][1]*self.m + kj[i][2]*self.m_aksen + kj[i][3]*self.f)

        koreksi_jarak_bumi_bulan = total_koreksi_jarak_bumi_bulan/1000
        return 385000.56 + koreksi_jarak_bumi_bulan


class MoonEcliptic():
    def __init__(self, T):
        moon_correction = MoonCorrection(T)
        self.l_aksen = moon_correction.L_aksen
        self.koreksi_bujur = moon_correction.koreksi_bujur_bulan
        self.latitude = moon_correction.lintang_bulan
        nutation = Nutation(T)
        self.delta_psi = nutation.delta_psi

    
    @property
    def apparent_longitude(self):
        longitude = (degrees(self.l_aksen) + self.koreksi_bujur)%360
        return radians((longitude + self.delta_psi)%360)

    @property
    def apparent_latitude(self):
        return (self.latitude)


class MoonEquatorial():
    def __init__(self, T, lst) -> None:
        moon_ecliptic = MoonEcliptic(T)
        self.apparent_longitude = moon_ecliptic.apparent_longitude
        self.apparent_latitude = moon_ecliptic.latitude
        self.lst = lst

        nutation = Nutation(T)
        self.epsilon = nutation.epsilon
        self.delta_psi = nutation.delta_psi
        

    @property
    def right_ascension(self):
        x_num = cos(self.apparent_longitude)
        y_num = sin(self.apparent_longitude) * cos(self.epsilon) - tan(radians(self.apparent_latitude)) * sin(self.epsilon)
        hasil = (atan2(y_num, x_num)) % 360
        return hasil


    @property
    def declination(self):
        latitude = radians(self.apparent_latitude)
        hasil = asin(sin(latitude)*cos(self.epsilon)+cos(latitude)*sin(self.epsilon)*sin(self.apparent_longitude))
        return hasil


    @property
    def hour_angle(self):
        return radians(self.lst * 15 - degrees(self.right_ascension))


class MoonHorizon():
    def __init__(self, T, lst, latitude):
        moon_equatorial = MoonEquatorial(T, lst)
        self.hour_angle = moon_equatorial.hour_angle
        self.declination = moon_equatorial.declination
        self.latitude = latitude


    @property
    def azimut_selatan(self):
        y = sin(self.hour_angle)
        x = cos((self.hour_angle))*sin(radians(self.latitude))-tan((self.declination))*cos(radians(self.latitude))
        return degrees(atan2(y, x))


    @property
    def azimut_bulan(self):
        return (self.azimut_selatan + 180) % 360


    @property
    def altitude(self):
        lat = radians(self.latitude)
        delta = radians(self.declination)
        ha = radians(self.hour_angle)
        return degrees(asin(sin(lat)*sin(delta)+cos(lat)*cos(delta)*cos(ha)))


class MoonProperty():
    def __init__(self, T, lst) -> None:
        self.lst = lst

        moon_correction = MoonCorrection(T)
        self.moon_distance = moon_correction.moon_distance

        moon_equatorial = MoonEquatorial(T, lst)
        self.right_ascension = moon_equatorial.right_ascension
        self.declination = moon_equatorial.declination

        sun_property = SunProperty(T)
        self.sun_distance = sun_property.sun_distance_km

        self.sun_equtorial = SunEquatorial(T)


    @property
    def moon_distance_km(self):
        return self.moon_distance


    @property
    def moon_distance_au(self):
        return self.moon_distance

    @property
    def sudut_parallaks(self):
        return degrees(asin(6378.14/self.moon_distance))


    @property
    def sudut_jari_jari_bulan(self):
        return 358473400/(self.moon_distance*3600)


    @property
    def sudut_fai(self):
        delta_b = (self.declination)
        alpha_b = (self.right_ascension)
        alpha_m = (self.sun_equtorial.right_acsension)
        delta_m = (self.sun_equtorial.declination)
        return acos(sin(delta_b)*sin(delta_m) + cos(delta_b)*cos(delta_m)*cos(alpha_b-alpha_m))


    @property
    def sudut_fase(self):
        x = self.moon_distance_km - self.sun_distance*cos(self.sudut_fai)
        y = self.sun_distance*sin(self.sudut_fai)
        return atan2(y, x)


    @property
    def illuminasi(self):
        return 100*(1 + cos(self.sudut_fase))/2