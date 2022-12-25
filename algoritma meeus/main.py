from math import *
from utilities import *
from calcutation import *
from place import Place
#from functions.nutation import Nutation
from moon import *
from nutation import *

class DateTime:
    def __init__ (self, year, month, day, hour=0, minute=0, second=0, timezone=0, longitude=0):
        self.year       = year
        self.month      = month
        self.day        = day
        self.hour       = hour
        self.minute     = minute
        self.second     = second
        self.timezone   = timezone
        self.JD_ut      = jd_ut(self.year, self.month, self.day, self.hour, self.minute, self.second, self.timezone)
        self.T_ut       = (self.JD_ut-2451545)/36525
        self.delta_t    = delta_T(self.year, self.month, self.day)/86400
        self.JDE        = self.JD_ut + self.delta_t
        self.T_td       = (self.JDE - 2451545)/36525
        self.tau        = self.T_td/10
        self.GST        = ((280.46061837 + 360.98564736629*(self.JD_ut - 2451545) + 0.000387933*self.T_ut**2 - self.T_ut**3/38710000) % 360)/15
        nutasi          = Nutation(self.T_td)
        self.GST_nampak = self.GST + nutasi.delta_psi * cos(nutasi.epsilon)/15
        self.LST_nampak = (self.GST_nampak + longitude/15) % 24
        self.koreksi_dpsi = nutasi.delta_psi_total






coordinate = Place(6,57,0,"LS", 107, 37,0,"BT")
waktu = DateTime(2022, 12, 14, 0, 0, 0,7, coordinate.longitude_dec)
data_bulan = Moon(waktu.T_td)

print(f"""
Data Bulan
Bujur rata-rata bulan   : {degrees(data_bulan.bujur_rata_rata_bulan)}°  
                        : {data_bulan.bujur_rata_rata_bulan} rads                         
Koreksi bujur bulan     : {data_bulan.koreksi_bujur_bulan} rads
Bujur Bulan             : {data_bulan.bujur_bulan}° 
Koreksi Delta Psi       : {data_bulan.koreksi_delta_psi}°
Bujur Bulan Nammpak     : {data_bulan.bujur_bulan_nampak}°

Lintang Bulan           : {data_bulan.lintang_bulan}°
Jarak Bulan-Bumi        : {data_bulan.jarak_bumi_bulan} km
Sudut Parallaks         : {data_bulan.sudut_parallaks}°
Sudut jari-jari bulan   : {data_bulan.sudut_jari_jari_bulan}°
Alpha                   : {data_bulan.alpha}°
Delta                   : {data_bulan.delta}°
Hour Angle              : {data_bulan.hour_angle}°

""")


