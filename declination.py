import math
from main import *
from konstanta import *

#  ========== Pendekatan Cooper ==========
# http://soal-olim-astro.blogspot.com/2013/06/mencari-deklinasi-matahari.html

# data awal
def n(tgl, bln, thn):
    if cek_kabisat(thn):
        return hari_kabisat[bln-1] + tgl
    return hari_biasa[bln-1] + tgl

N =n(26,7,2011)

# == Cooper I ==
def declination_cooper1(n):
    res = 23.45 * math.sin(math.radians((360/365*(284+n))))
    return res

# == Cooper II ==
def declination_cooper2(N):
    res = 23.45 * math.sin(math.radians(360/365*(N-81)))
    return res



# ===== Algoritma PSA ===== #
# http://www.psa.es/sdg/sunpos.htm

# menghitung Julian Day
def sun_declination_psa(year, month, day, hour, minute, second, timezone):
    hour = hour + (minute + second/60)/60
    li_aux1 = int((month-14)/12)
    li_aux2 = int((1461 * (year + 4800 + li_aux1))/4 + (367 * (month - 2 - 12 * li_aux1))/12 - (3 * ((year + 4900 + li_aux1)/100))/4 + day - 32075)
    julian_date = li_aux2 + 0.5 + hour/24 - timezone/24
    julian_days = julian_date - 2451545

    print(julian_date)
# hitung koordinat ekliptika
    omega = 2.1429 - 0.0010394594 * julian_days
    mean_longitude = 4.8950630+ 0.017202791698 * julian_days # dalam radians
    mean_anomaly = 6.2400600 + 0.0172019699 * julian_days
    ecliptic_longitude = mean_longitude + 0.03341607 * math.sin(mean_anomaly) + 0.00034894 * math.sin ( 2 * mean_anomaly ) - 0.0001134 - 0.0000203 * math.sin(omega)
    ecliptic_obliquity = 0.4090928 - 6.2140e-9 * julian_days + 0.0000396 * math.cos(omega)
    #ecliptic_oblique = 0.4090928 - 6.2140e-9 * julian_days + 0.0000396 * math.cos(omega)

# hitung deklinasi
    sin_ecliptic_longitude = math.sin(ecliptic_longitude)
    declination = math.asin(math.sin(ecliptic_obliquity) * sin_ecliptic_longitude)
    return declination

x = sun_declination_psa(2016, 5, 14, 9, 0, 0, 7)
print(to_dms(math.degrees(x)))
