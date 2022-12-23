import math
from typing import Self
from main import *
from delta_t_meeus import *
from obliquity_nutasi_meeus import *

class Koordinat:
    def __init__(self, lat_deg, lat_min, lat_sec, lat, lon_deg, lon_min, lon_sec, lon):
        self._lat_deg = lat_deg
        self._lat_min = lat_min
        self._lat_sec = lat_sec
        self._lat = lat
        self._lon_deg = lon_deg
        self._lon_min = lon_min
        self._lon_sec = lon_sec
        self._lon = lon

    
    def lintang(self):
        if self._lat == "N" or self._lat == "LU":
            return self._lat_deg + self._lat_min/60 + self._lat_sec/3600
        return (self._lat_deg + self._lat_min/60 + self._lat_sec/3600) * (-1)

    
    def bujur(self):
        if self._lon == "E" or self._lon == "BT":
            return self._lon_deg + self._lon_min/60 + self._lon_sec/3600
        return (self._lon_deg + self._lon_min/60 + self._lon_sec/3600) * (-1)

    
    def cetak_koordinat(self):
        _lintang = to_dms(self.lintang())
        _bujur = to_dms(self.bujur())
        print(f"{_lintang} {self._lat} {_bujur} {self._lon}")

   
    def lat_rad(self):
        return math.radians(self.lintang())
        
    
    def lot_rad(self):
        return math.radians(self.bujur())
    

class Tanggal:
    def __init__(self, hari, bulan, tahun):
        self.hari = hari
        self.bulan = bulan
        self.tahun = tahun

        self.tahun_desimal = self.tahun + (self.bulan-1)/12 + self.hari/365


class Waktu:
    def __init__(self, jam, menit, detik, timezone):
        self.jam = jam
        self.menit = menit
        self.detik = detik
        self.timezone = timezone
    

    def time_dec(self):
        return (self.jam + self.menit/60 + self.detik/3600 + self.timezone)/24

class DetailPerhitungan:
    def __init__(self, tahun, bulan, hari, jam, menit, detik, timezone):
        self.tahun = tahun
        self.bulan = bulan
        self.hari = hari
        self.jam = jam
        self.menit = menit
        self.detik = detik
        self.timezone = timezone
        _m = self.bulan + 12 if self.bulan < 3 else self.bulan
        _y = self.tahun - 1 if self.bulan < 3 else self.tahun
        _a = int(_y/100)
        if self.tahun < 1583:
            if self.bulan < 11:
                if self.hari < 4:
                    b = 0
                    if self.hari > 14:
                        b = 2 + int(_a/4)-_a
                    else:
                        print("tanggal salah")
            else:
                b = 2 + int(_a/4)-_a
        else:
            b   = 2 + int(_a/4) - _a
        jd_ut            = 1720994.5 + int(365.25 * _y) + int(30.60001 * (_m + 1)) + b + self.hari + (self.jam + self.menit/60 + self.detik/3600)/24 - self.timezone/24
        _t_ut            = (jd_ut-2451545)/36525
        _delta           = delta_t(tgl.tahun_desimal)
        _jde             = jd_ut + _delta/86400
        self.t_td            = (_jde - 2451545)/36525
        _tau             = self.t_td/10
        _gst             = ((280.46061837 + 360.98564736629*(jd_ut - 2451545) + 0.000387933*_t_ut**2 - _t_ut**3/38710000) % 360)/15
        _delta_e         = delta_epsilon_total(self.t_td)
        _u               = u(self.t_td)
        _epsilon_z       = epsilon_zero(_u)
        _delta_e_total   = delta_epsilon_total(self.t_td)/10000
        _epsilon         = math.radians(epsilon(_epsilon_z, _delta_e_total))
        _delta_psi_total = delta_psi_total(self.t_td)
        _delta_psi       = _delta_psi_total/3600
        _gst_nampak      = _gst + _delta_psi * math.cos(_epsilon)/15
        _lst_nampak      = (_gst_nampak + bandung.bujur()/15) % 24


class DataMatahari:
    pass


class DataBulan:
    def __init__ (self, t_td):
        self.t_td       = t_td
        l_aksen         = (218.3164591 + 481267.88134236*self.t_td - 0.0013268*self.t_td*self.t_td + self.t_td*self.t_td*self.t_td/538841 - self.t_td*self.t_td*self.t_td*self.t_td/65194000) % 360
        print(l_aksen)
        koreksi_bujur_bulan = 0




            

#test


bandung = Koordinat(6,57,0,"LS", 107, 37,0,"BT")
#tgl = Tanggal(14, 7, 2016)   
tgl = Tanggal(14, 10, 2022)   
pukul = Waktu(9,0,0,7)
calc = DetailPerhitungan(tgl.tahun, tgl.bulan, tgl.hari, pukul.jam, pukul.menit, pukul.detik, pukul.timezone)
bulan =DataBulan(calc.t_td)

#print(to_dms(xxxxx))
