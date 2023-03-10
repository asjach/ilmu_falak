from math import cos
from VSOP87_ELP2000.functions import dms_to_dec, jd, delta_t, print_hms
from nutation import Nutation
from moon import *
from sun import *


#  simulasi input
# tanggal dan waktu
tahun = 2023
bulan = 1
tgl = 3
jam = 12
menit = 0
detik = 0
timezone = 0

# lokasi
lintang_d = 6
lintang_m = 57
lintang_s = 0
lintang_arah = "LS"
bujur_d = 107
bujur_m = 37
bujur_s = 0
bujur_arah = "BT"


class Ephemeris():
    def __init__(self, year, month, day, hour = 0, minute = 0, second = 0, lat_d = 0, lat_m = 0, lat_s = 0, lat_dir = '', lon_d = 0, lon_m = 0, lon_s = 0, lon_dir = '', timezone = 0) -> None:
        self.year = year
        self.month = month
        self.day = day
        self.hour = hour
        self.minute = minute
        self.second = second
        self.lat_d = lat_d
        self.lat_m = lat_m
        self.lat_s = lat_s
        self.lat_dir = lat_dir
        self.lon_d = lon_d
        self.lon_m = lon_m
        self.lon_s = lon_s
        self.lon_dir = lon_dir
        self.timezone = timezone
        self.nutation = Nutation(self.T)
        self.moon_ecliptic = MoonEcliptic(self.T)
        self.moon_equatorial = MoonEquatorial(self.T, self.lst)
        self.moon_property = MoonProperty(self.T, self.lst)
        self.moon_horizontal = MoonHorizon(self.T, self.lst, dms_to_dec(lat_d, lat_m, lat_s, lat_dir))
        self.sun_ecliptic = SunEcliptic(self.T)
        self.sun_equatorial = SunEquatorial(self.T)
        self.sun_horizon = SunHorizon(self.T, dms_to_dec(lat_d, lat_m, lat_s, lat_dir))
        self.sun_property = SunProperty(self.T)

    @property
    def _jd(self):
        return jd(self.year, self.month, self.day, self.hour, self.minute, self.second, self.timezone)

    @property
    def jde(self):
        decimal_year = self.year + (self.month-1)/12 + self.day/365
        deltaT = delta_t(decimal_year)
        return self._jd + deltaT
        
    @property
    def t(self):
        # T is Julian Century (JC)
        # T in UT (Universal Time)
        return (self._jd - 2451545)/36525


    @property
    def T(self):
        # T is Julian Century Ephemeris (JCE)
        # T in TD (Dynamic Time) / 
        return (self.jde - 2451545)/36525
    
    @property
    def lst(self):
        lon_dec = dms_to_dec(self.lat_d, self.lat_m, self.lat_s, self.lat_dir)
        gst = ((280.46061837 + 360.98564736629*(self._jd - 2451545) + 0.000387933*self.t**2 - self.t**3/38710000) % 360)/15
        gstA = gst + self.nutation.delta_psi * cos(self.nutation.epsilon)/15
        lst = (gstA + (lon_dec)/15) % 24
        return lst


    @property
    def moon_longitude(self):
        return degrees(self.moon_ecliptic.apparent_longitude)


    @property
    def moon_latitude(self):
        return self.moon_ecliptic.apparent_latitude


    @property
    def moon_distance(self):
        return self.moon_property.moon_distance_km
    

    @property
    def moon_right_ascension(self):
        return degrees(self.moon_equatorial.right_ascension)


    @property
    def moon_declination(self):
        return degrees(self.moon_equatorial.declination)


    @property
    def moon_hour_angle(self):
        return degrees(self.moon_equatorial.hour_angle)


    @property
    def moon_azimuth (self):
        return self.moon_horizontal.azimut_bulan


    @property
    def moon_altitude(self):
        return self.moon_horizontal.altitude


    @property
    def moon_parallax(self):
        return self.moon_property.sudut_parallaks


    @property
    def sun_declination(self):
        return self.sun_equatorial.declination


    @property
    def eot(self):
        return self.sun_property.equation_of_time

moon = Ephemeris(tahun, bulan, tgl, jam, menit, detik, lintang_d, lintang_m, lintang_s, lintang_arah, bujur_d, bujur_m, bujur_s, bujur_arah, timezone)
print(f'''
MOON
longitude: {moon.moon_longitude}
latitude: {moon.moon_latitude}
jarak ke bumi: {moon.moon_distance}
ra bulan: {moon.moon_right_ascension}
deklinasi bulan: {moon.moon_declination}
hour angle bulan: {moon.moon_hour_angle}
azimut bulan: {moon.moon_azimuth}
altitude bulan: {moon.moon_altitude}
Horizontal Parallax: {moon.moon_parallax}


sudut fai : {moon.moon_property.sudut_fai}
sudut fase: {moon.moon_property.sudut_fase}



eot: {print_hms(moon.eot/60)}
''')