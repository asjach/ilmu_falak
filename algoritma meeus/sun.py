from math import radians, sin, cos, tan, asin, atan, atan2, degrees
from nutation import *
from moon_correction import MoonCorrection
from sun_correction import SunCorrection

class Sun:
    def __init__(self, t_td, lst_nampak, latitude):
        self.t                          = t_td
        self.tau                        = t_td/10
        
        self.nutasi                     = Nutation(t_td)
        self.delta_psi                  = self.nutasi.delta_psi
        self.epsilon                    = self.nutasi.epsilon
        
        koreksi_bulan                   = MoonCorrection(t_td)
        self.T                          = koreksi_bulan.koreksi_bujur_bulan

        koreksi_matahari                = SunCorrection(t_td)
        self.theta_terkoreksi           = koreksi_matahari.theta_terkoreksi
        self.jarak_bumi_matahari        = koreksi_matahari.jarak_bumi_matahari
        self.lintang_matahari_tampak    = koreksi_matahari.lintang_matahari_tampak

        self.lst_nampak                 = lst_nampak
        self.latitude                   = radians(latitude)

    
    @property
    def koreksi_aberasi(self):
        return -20.4898/(3600*self.jarak_bumi_matahari)


    @property
    def bujur_matahari_nampak(self):
        return radians((self.theta_terkoreksi + self.delta_psi + self.koreksi_aberasi) % 360)

    
    @property
    def sudut_jari_jari_matahari(self):
        return (959.63/3600)/self.jarak_bumi_matahari


    @property
    def alpha_matahari(self):
        y = sin((self.bujur_matahari_nampak))*cos(self.epsilon) -tan(self.lintang_matahari_tampak)*sin(self.epsilon)
        x = cos((self.bujur_matahari_nampak))
        return (degrees(atan2(y,x)) % 360)


    @property
    def delta_matahari(self):
        return asin(sin(self.lintang_matahari_tampak)*cos(self.epsilon) + cos(self.lintang_matahari_tampak)*sin(self.epsilon)*sin(self.bujur_matahari_nampak))

    @property
    def hour_angle_matahari(self):
        return radians(self.lst_nampak*15-self.alpha_matahari)


    @property
    def azimut_selatan(self):
        y = sin(self.hour_angle_matahari)
        x = cos(self.hour_angle_matahari)*sin(self.latitude) - tan(self.delta_matahari)*cos(self.latitude)
        return degrees((atan2(y,x)))

    @property
    def azimut_matahari(self):
        return (self.azimut_selatan + 180) % 360


    @property
    def altitude(self):
        return degrees(asin(sin(self.latitude)*sin(self.delta_matahari)+cos(self.latitude)*cos(self.delta_matahari)*cos(self.hour_angle_matahari)))


    @property
    def sudut_parallaks_matahari(self):
        return degrees(atan(6378.14/(self.jarak_bumi_matahari*149598000)))



