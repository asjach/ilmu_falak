from math import radians
from functions import to_dms


class Place:
    def __init__(self, nama_lokasi, lat_deg, lat_min, lat_sec, lat_direction, lon_deg, lon_min, lon_sec, lon_direction):
        self.nama_lokasi = nama_lokasi
        self._lat_deg = lat_deg
        self._lat_min = lat_min
        self._lat_sec = lat_sec
        self._lat_direction = lat_direction
        self._lon_deg = lon_deg
        self._lon_min = lon_min
        self._lon_sec = lon_sec
        self._lon_direction = lon_direction
    

    @property
    def koordinat(self):
        _koordinat = f'{self._lat_deg}°{self._lat_min}\'{self._lat_sec}\" {self._lat_direction} {self._lon_deg}°{self._lon_min}\'{self._lon_sec}\" {self._lon_direction}'
        return _koordinat

    @property
    def latitude_dec(self):
        if self._lat_direction == "N" or self._lat_direction == "LU":
            return self._lat_deg + self._lat_min/60 + self._lat_sec/3600
        return (self._lat_deg + self._lat_min/60 + self._lat_sec/3600) * (-1)
    

    @property
    def latitude_dms(self):
        return to_dms(self.latitude_dec)
    
    
    @property
    def latitude_rad(self):
        return radians(self.latitude_dec)


    @property
    def longitude_dec(self):
        if self._lon_direction == "E" or self._lon_direction == "BT":
            return self._lon_deg + self._lon_min/60 + self._lon_sec/3600
        return (self._lon_deg + self._lon_min/60 + self._lon_sec/3600) * (-1)


    @property
    def longitude_dms(self):
        return to_dms(self.longitude_dec)


    @property
    def longitude_rad(self):
        return radians(self.longitude_dec)