import math
from typing import Self
from main import *
from delta_t_meeus import *
from nutasi_meeus import *

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


class DataMatahari:
    def __init__(self, tahun, bulan, hari, jam, menit, detik, timezone):
        self.tahun = tahun
        self.bulan = bulan
        self.hari = hari
        self.jam = jam
        self.menit = menit
        self.detik = detik
        self.timezone = timezone

        m = self.bulan + 12 if self.bulan < 3 else self.bulan
        y = self.tahun - 1 if self.bulan < 3 else self.tahun
        a = int(y/100)
        if self.tahun < 1583:
            if self.bulan < 11:
                if self.hari < 4:
                    b = 0
                    if self.hari > 14:
                        b = 2 + int(a/4)-a
                    else:
                        print("tanggal salah")
            else:
                b = 2 + int(a/4)-a
        else:
            b   = 2 + int(a/4) - a
        jd_ut   = 1720994.5 + int(365.25 * y) + int(30.60001 * (m + 1)) + b + self.hari + (self.jam + self.menit/60 + self.detik/3600)/24 - self.timezone/24
        t_ut    = (jd_ut-2451545)/36525
        delta   = delta_t(tgl.tahun_desimal)
        jde     = jd_ut + delta/86400
        t_td    = (jde - 2451545)/36525
        tau     = t_td/10
        gst     = ((280.46061837 + 360.98564736629*(jd_ut - 2451545) + 0.000387933*t_ut**2 - t_ut**3/38710000) % 360)/15
        delta_e= delta_epsilon(t_td)


class DataBulan:
    pass
#test


bandung = Koordinat(6,57,0,"LS", 107, 37,0,"BT")
#tgl = Tanggal(14, 7, 2016)   
tgl = Tanggal(14, 10, 2022)   
pukul = Waktu(9,0,0,7)
data_matahari = DataMatahari(tgl.tahun, tgl.bulan, tgl.hari, pukul.jam, pukul.menit, pukul.detik, pukul.timezone)

# bandung.cetak_koordinat()
