from math import *
from utilities import *
from place import Place
from moon import *
from sun import *
from nutation import *
from delta_t import delta_T
from moon_correction import *
from time_convertion import KonversiWaktu


class Tanggal:
    def __init__(self, year, month, day):
        self.year       = year
        self.month      = month
        self.day        = day
    

    @property
    def year_decimal(self):
        return self.year + (self.month-1)/12 + self.day/365

    @property
    def ymd(self):
        return {'y': self.year, 'm': self.month, 'd': self.day}


class Waktu:
    def __init__(self, hour, minute, second, timezone):
        self.hour       = hour
        self.minute     = minute
        self.second     = second
        self.timezone   = timezone


    @property
    def hms(self):
        #return (self.hour, self.minute, self.second)
        return {'h': self.hour, 'm':self.minute, 's':self.second}

#INPUT
lokasi1 = Place("Bandung", 6,57,0,"LS", 107, 37,0,"BT")
tanggal = Tanggal(2022,12,14)
waktu = Waktu(14, 0, 0, 7)

#PROSES
T = KonversiWaktu(tanggal.year, tanggal.month, tanggal.day, waktu.hour, waktu.minute, waktu.second, waktu.timezone, lokasi1.longitude_dec)


# OUTPUT
data_matahari = Sun(T.T_td, T.LST_nampak, lokasi1.latitude_dec)
data_bulan = Moon(T.T_td, T.LST_nampak, lokasi1.latitude_dec, data_matahari.alpha_matahari, data_matahari.delta_matahari)


print(f"""
Data Bulan
Lokasi                  : {lokasi1.nama_lokasi}, {lokasi1.koordinat}        
Bujur rata-rata bulan   : {degrees(data_bulan.bujur_rata_rata_bulan)}°                       
Koreksi bujur bulan     : {data_bulan.koreksi_bujur_bulan} rads
Bujur Bulan             : {data_bulan.bujur_bulan}° 
Koreksi Delta Psi       : {data_bulan.delta_psi}°
Bujur Bulan Nampak      : {data_bulan.bujur_bulan_nampak}°
Lintang Bulan           : {data_bulan.lintang_bulan} rad
Jarak Bulan-Bumi        : {data_bulan.jarak_bumi_bulan} km
Sudut Parallaks         : {data_bulan.sudut_parallaks}°
Sudut jari-jari bulan   : {data_bulan.sudut_jari_jari_bulan}°
Alpha                   : {data_bulan.alpha}°
Delta                   : {data_bulan.delta}°
Hour Angle              : {data_bulan.hour_angle}°
Azimuth dari Selatan    : {data_bulan.azimut_selatan}°
Azimuth                 : {data_bulan.azimut_bulan}°
Altitude                : {data_bulan.altitude}°
Sudut Fai               : {data_bulan.sudut_fai}°
Sudut Fase (i)          : {data_bulan.sudut_fase}°
Illuminasi Bulan        : {data_bulan.illuminasi}°

DATA MATAHARI
Theta_terkoreksi        : {data_matahari.theta_terkoreksi}°
Delta Psi               : {data_matahari.delta_psi}°
Koreksi aberasi         : {data_matahari.koreksi_aberasi}
Bujur matahari nampak   : {data_matahari.bujur_matahari_nampak}
Lintang matahari nampak : {data_matahari.lintang_matahari_tampak}
Jarak bumi-matahari =   : {data_matahari.jarak_bumi_matahari} ({data_matahari.jarak_bumi_matahari*149598000}km)
Sudut jari-jari matahari: {data_matahari.sudut_jari_jari_matahari}
Alpha                   : {data_matahari.alpha_matahari}
Delta                   : {data_matahari.delta_matahari}
Hour Angle              : {data_matahari.hour_angle_matahari}
Azimuth dari selatan    : {data_matahari.azimut_selatan}
Azimuth                 : {data_matahari.azimut_matahari}
Altitude                : {data_matahari.altitude}
Sudut paralaks matahari : {data_matahari.sudut_parallaks_matahari}
""")



