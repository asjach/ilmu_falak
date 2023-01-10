from math import radians, sin
from VSOP87_ELP2000.functions import latitude_dec, longitude_dec, to_hms, julian_day



jakarta = (6, 12, 0, "LS", 106, 49, 0, "BT", 7)
surabaya = (7, 14, 0, "LS", 112, 44, 0, "BT", 7)
banda_aceh = (5, 33, 0, "LU", 95, 19, 0, "BT", 7)


def kwd(lon_d, lon_m, lon_s, lon_dir, timezone):
    bujur_daerah = (timezone*15)
    bujur_tempat = longitude_dec(lon_d, lon_m, lon_s, lon_dir)
    hasil = (bujur_tempat - bujur_daerah)/15
    return hasil


def eot(tahun, bulan, hari, tz):
    jd = julian_day(tahun, bulan, hari, timezone=tz)
    u = (jd - 2451545)/36525
    lo = 280.46607 + 36000.7698*u
    result = -(1789 + 237*u)*sin(radians(lo)) - (7146 - 62*u)

x = kwd(*jakarta[4:9])
print(to_hms(x))