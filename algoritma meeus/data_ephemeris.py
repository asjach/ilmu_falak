# from main import Tanggal, Waktu
from utilities import *
from math import degrees, radians
from moon import Moon
from place import Place
from sun import Sun
from time_convertion import KonversiWaktu


# tempat = Place("bandung", 6, 57, 0, "LS", 107, 37, 0, "BT")
# tanggal = Tanggal(2022,12,14)
# waktu = Waktu(14,0,0,7)


class DataEphemeris:
    def __init__(self, nama,lat_deg, lat_min, lat_sec, lat_dir, lon_deg, lon_min, lon_sec, lon_dir, year, month, day, hour, minute, second, timezone):
        self._tempat = Place(nama, lat_deg, lat_min, lat_sec,lat_dir, lon_deg, lon_min, lon_sec, lon_dir )
        self._t = KonversiWaktu(year, month, day, hour, minute, second, timezone, self._tempat.longitude_dec)
        self._data_matahari = Sun(self._t.T_td, self._t.LST_nampak,self._tempat.latitude_dec)
        self._data_bulan = Moon(self._t.T_td, self._t.LST_nampak,self._tempat.latitude_dec, self._data_matahari.alpha_matahari, self._data_matahari.hour_angle_matahari)
    
 
    def bujur_bulan(self, format='dms'):
        if format == 'deg':
            output = degrees(self._data_bulan.bujur_bulan_nampak)
        elif format == 'rad':
            output = radians(self._data_bulan.bujur_bulan)
        elif format == 'dms':
            output = (to_dms(degrees(self._data_bulan.bujur_bulan_nampak)))
        return output

data = DataEphemeris('Bandung', 6, 57, 0, "LS", 107, 37, 0, "BT", 2022,12, 27, 18, 10,17, 7)
print(data.bujur_bulan('rad'))