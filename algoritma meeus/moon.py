from math import radians, sin, cos, tan, asin, acos, atan2, degrees
from nutation import *
from moon_correction import MoonCorrection
from sun_correction import SunCorrection
#from sun import *


class Moon:
    def __init__(self, t_td, lst_nampak, latitude, a_matahari, d_matahari):
        self.t_td = t_td
        T = t_td
        self.nutasi                 = Nutation(t_td)
        koreksi_bulan               = MoonCorrection(t_td)
        self.bujur_rata_rata_bulan  = koreksi_bulan.L_aksen
        self.koreksi_bujur_bulan    = koreksi_bulan.koreksi_bujur_bulan
        self.lintang_bulan          = koreksi_bulan.lintang_bulan
        self.jarak_bumi_bulan       = koreksi_bulan.jarak_bumi_bulan
        self.delta_psi              = self.nutasi.delta_psi
        self.lst_nampak             = lst_nampak
        self.latitude               = latitude
        self.jarak_bumi_matahari    = SunCorrection(t_td).jarak_bumi_matahari * 149598000
        self.a_matahari             = a_matahari
        self.d_matahari             = d_matahari                    


    @property
    def bujur_bulan(self):
        
        return (degrees(self.bujur_rata_rata_bulan) + self.koreksi_bujur_bulan)%360


    @property
    def bujur_bulan_nampak(self):
        return radians((self.bujur_bulan + self.delta_psi) % 360)


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
        delta_m = (self.d_matahari)
        alpha_b = radians(self.alpha)
        alpha_m = radians(self.a_matahari)
        return acos(sin(delta_b)*sin(delta_m) + cos(delta_b)*cos(delta_m)*cos(alpha_b-alpha_m))


    @property
    def sudut_fase(self):
        x = self.jarak_bumi_bulan - self.jarak_bumi_matahari*cos(self.sudut_fai)
        y = self.jarak_bumi_matahari*sin(self.sudut_fai)
        return atan2(y,x)


    @property
    def illuminasi(self):
        return 100*(1 + cos(self.sudut_fase))/2
