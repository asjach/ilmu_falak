import math
from fungsi_bantu import to_dms

class Coordinate:
    def __init__(self, lat_deg, lat_min, lat_sec, lat_direction, lon_deg, lon_min, lon_sec, lon_direction):
        self._lat_deg = lat_deg
        self._lat_min = lat_min
        self._lat_sec = lat_sec
        self._lat = lat_direction
        self._lon_deg = lon_deg
        self._lon_min = lon_min
        self._lon_sec = lon_sec
        self._lon = lon_direction
    

    @property
    def latitude_dec(self):
        if self._lat == "N" or self._lat == "LU":
            return self._lat_deg + self._lat_min/60 + self._lat_sec/3600
        return (self._lat_deg + self._lat_min/60 + self._lat_sec/3600) * (-1)
    

    @property
    def latitude_dms(self):
        return to_dms(self.latitude_dec)
    
    
    @property
    def latitude_rad(self):
        return math.radians(self.latitude_dec)


    @property
    def longitude_dec(self):
        if self._lon == "E" or self._lon == "BT":
            return self._lon_deg + self._lon_min/60 + self._lon_sec/3600
        return (self._lon_deg + self._lon_min/60 + self._lon_sec/3600) * (-1)


    @property
    def longitude_dms(self):
        return to_dms(self.latitude_dec)


    @property
    def longitude_rad(self):
        return math.radians(self.longitude_dec)


class DateTime:
    def __init__ (self, year, month, day, hour, minute, second, timezone):
        self._year = year
        self._month = month
        self._day = day
        self._hour = hour
        self._minute = minute
        self._second = second
        self._timezone = timezone
        
    @property
    def JD_ut(self):
        _m = self._month + 12 if self._month < 3 else self._month
        _y = self._year - 1 if self._month < 3 else self._year
        _a = int(_y/100)
        if self._year < 1583:
            if self._month < 11:
                if self._day < 4:
                    b = 0
                    if self._day > 14:
                        b = 2 + int(_a/4)-_a
                    else:
                        print("tanggal salah")
            else:
                b = 2 + int(_a/4)-_a
        else:
            b   = 2 + int(_a/4) - _a 
        return 1720994.5 + int(365.25 * _y) + int(30.60001 * (_m + 1)) + b + self._day + (self._hour + self._minute/60 + self._second/3600)/24 - self._timezone/24


    
    

class Calc:
    pass


koordinat = Coordinate(6,57,0,"LS", 107, 37,0,"BT")
waktu = DateTime(2022, 12, 25, 19, 1, 0,7)
print('---- Lintang ----')
print(koordinat.latitude_dec)
print(koordinat.latitude_rad)
print(koordinat.latitude_dms)
print('---- Bujur ----')
print(koordinat.longitude_dec)
print(koordinat.longitude_rad)
print(koordinat.longitude_dms)
print('---- JD dalam UT ----')
print(waktu.JD_ut)
print(waktu.jd_ut)



