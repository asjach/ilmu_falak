from functools import reduce
from math import acos, asin, atan, atan2, cos, degrees, sin, radians, sqrt, tan
from functions import dms_to_dec, print_dms
from terms import terms_delta_psi, terms_delta_epsilon
from vsop87 import*


def julian_datum(year, month, day):
    '''
    Julian Day (JD)
    '''
    aa = 10000 * year + 100 * month + day
    if month < 3:
        month = month + 12
        year = year -1
    if aa <= 15821004.99999:
        B = 0
    else:
        A = int(year/100)
        B = 2 - A + int(A/4)
    return int(365.25*(year + 4716)) + int(30.6001 * (month + 1)) + day + B - 1524.5



def julian_ephemeris_day(year, month, day):
    '''
    Julian Ephemeris Day (JDE)
    '''

    jd = julian_datum(year, month, day)
    deltaT = delta_t(year, month, day)
    return jd + deltaT


# def julian_ephemeris_day(year, month, day, hour, minute, second, timezone):
#     '''
#     Julian Ephemeris Day (JDE)
#     '''
#     jd = julian_datum(year, month, day, hour, minute, second, timezone)
#     deltaT = delta_t(year, month, day, hour, minute, second, timezone, "D")
#     return jd + deltaT


def delta_t(year, month, day) -> float:
    """ 
    Delta T
    """
    initial_year = year, 1, 0
    year_decimal = year + (julian_datum(year, month, day) - julian_datum(year, 1, 0))/365.25

    delta = 0
    if year_decimal <= -500:
        c = year_decimal/100
        delta = -20 + 32 * (c - 18.2)**2
    elif year_decimal > -500 and year_decimal <= 500:
        c = year_decimal/100
        delta = 10583.6 - 1014.41*c + 33.78311*c**2 - 5.952053*c**3 - 0.1798452*c**4 + 0.022174192*c**5 + 0.0090316521*c**6
    elif year_decimal > 500 and year_decimal <= 1600:
        c = year_decimal/100
        delta =  1574.2 - 556.01*(c-10) + 71.23472*(c-10)**2 + 0.319781*(c-10)**3 - 0.8503463*(c-10)**4 - 0.005050998*(c-10)**5 + 0.0083572073*(c-10)**6
    elif year_decimal > 1600 and year_decimal <= 1700:
        c = year_decimal - 1600
        delta = 120 - 0.9808*c - 0.01532*c**2 + c**3/7129
    elif year_decimal > 1700 and year_decimal <= 1800:
        c = year_decimal-1700
        delta = 8.83 + 0.1603*c - 0.0059285*c**2 + 0.00013336*c**3 - c**4/1174000
    elif year_decimal > 1800 and year_decimal <= 1860:
        c = year_decimal-1800
        delta = 13.72 - 0.332447*c + 0.0068612*c**2 + 0.0041116*c**3 - 0.00037436*c**4 + 0.0000121272*c**5 - 0.0000001699*c**6 + 0.000000000875*c**7
    elif year_decimal > 1860 and year_decimal <= 1900:
        c = year_decimal - 1860
        delta = 7.62 + 0.5737*c - 0.251754*c**2 + 0.01680668*c**3 - 0.0004473624*c**4 + c**5/233174
    elif year_decimal > 1900 and year_decimal <= 1920:
        c = year_decimal-1900
        delta = -2.79 + 1.494119*c - 0.0598939*c**2 + 0.0061966*c**3 - 0.000197*c**4
    elif year_decimal > 1920 and year_decimal <= 1941:
        c = year_decimal-1920
        delta = 21.2 + 0.84493*c - 0.0761*c**2 + 0.0020936*c**3
    elif year_decimal > 1941 and year_decimal <= 1961:
        c = year_decimal-1950
        delta = 29.07 + 0.407*c - c**2/233 + c**3/2547
    elif year_decimal > 1961 and year_decimal <= 1986:
        c = year_decimal-1975
        delta = 45.45 + 1.067*c - c**2/260 - c**3/718
    elif year_decimal > 1986 and year_decimal <= 2005:
        c = year_decimal-2000
        delta = 63.86 + 0.3345*c - 0.060374*c*c + 0.0017275*c**3 + 0.000651814*c**4 + 0.00002373599*c**5
    elif year_decimal > 2005 and year_decimal <= 2050:
        c = year_decimal-2000
        delta = 62.92 + 0.32217*c + 0.005589*c**2
    elif year_decimal > 2050 and year_decimal <= 2150:
        c = (year_decimal - 1820)/100
        delta = -20 + 32*c**2 - 0.5628*(2150 - year_decimal)
    elif year_decimal > 2150:
        c = (year_decimal - 1820)/100
        delta = -20 + 32*c**2
    else:
        return f'wrong input'

    return delta/86400


def julian_century(year, month, day):
    '''
    Julian Century (J2000 1 Januari pukul 12.00 UT)
    '''
    jd = julian_datum(year, month, day)
    return (jd - 2451545)/36525

# def julian_century(year, month, day, hour, minute, second, timezone):
#     '''
#     Julian Century (J2000 1 Januari pukul 12.00 UT)
#     '''
#     jd = julian_datum(year, month, day, hour, minute, second, timezone)
#     return (jd - 2451545)/36525


def julian_ephemeris_century(year, month, day):
    '''
    Julian Ephemeris Century (J2000 1 Januari 12.00 UT)
    '''
    jde = julian_ephemeris_day(year, month, day)
    return (jde - 2451545)/36525


def jd_to_ymd(jd):
    z = int(jd + 0.5)
    if z < 2299161:
        a = z  #+ 1524
    else:
        alpha = int((z - 1867216.25)/36524.25)
        a = z + 1 + alpha - int(alpha/4)
    b = a + 1524
    c = int((b -122.1)/365.25)
    d = int(365.25*c)
    e = int((b - d)/30.6001)

    tgl = int(b - d - int(30.6001 * e))
    jam = jd + 0.5 - z

    if e < 14:
        bulan = e - 1
    if e == 14 or e == 15:
        bulan = e - 13
    
    if bulan > 2:
        tahun = int(c - 4716)
    if bulan == 1 or bulan == 2: 
        tahun = int(c - 4715)
    h = int(jam*24)
    m = int((jam - h/24)*1440)
    s = round((jam - h/24 - m/1440)*86400,2)
    return (tahun, bulan, tgl, h, m, s)


def greenwich_mean_sidereal_time(year, month, day):
    '''
    Greenwich Mean Sidereal Time (GST)
    '''
    jd = julian_datum(year, month, day)
    t = julian_ephemeris_century(year, month, day)
    gmst = 280.46061837 + 360.98564736629*(jd - 2451545) + 0.000387933*t**2 - t**3/38710000
    gmst = gmst % 360 / 15
    return gmst


def greenwich_hour_angle(year, month, day):
    '''
    Greenwich Hour Angle (GHA)
    '''
    epsilon = obliquity_of_ecliptic(year, month, day)
    d_psi = nutation_in_longitude(year, month, day)/3600
    gmst = greenwich_mean_sidereal_time(year, month, day)
    gha = gmst + d_psi * cos(radians(epsilon))/15
    return gha % 24


# def apparent_local_sidereal_time(year, month, day, hour, minute, second, timezone, lon_d, lon_m, lon_s, lon_dir):
#     '''
#     LST_nampak
#     '''

#     bujur = dms_to_dec(lon_d, lon_m, lon_s, lon_dir)
#     gha = greenwich_hour_angle(year, month, day, hour, minute, second, timezone)
#     hasil = (gha + bujur/15) % 24
#     return  hasil


def apparent_local_sidereal_time(bujur, year, month, day):
    '''
    LST_nampak
    '''
    bujur = dms_to_dec(bujur)
    gha = greenwich_hour_angle(year, month, day)
    hasil = (gha + bujur/15) % 24
    return  hasil


#############   ARAH QIBLAT  #############
def arah_qibla(lintang, bujur):
    '''
    Penentuan Arah Qiblat dengan rumus segitiga bola
    '''
    lintang_kabah = 21 + 25/60 + 21.03/3600
    bujur_kabah = 39 + 49/60 + 34.31/3600
    lintang = dms_to_dec(lintang)
    bujur = dms_to_dec(bujur)

    A = radians(90-lintang)
    B = radians(90 - lintang_kabah)
    c = radians(bujur - bujur_kabah)
    sB = sin(B)*sin(c)
    cB = cos(B)*sin(A) - cos(A)*sin(B)*cos(c)
    Bb = atan2(sB, cB)

    arah_qibla = 360 - degrees(Bb)
    while arah_qibla > 360:
        arah_qibla = arah_qibla - 360
        if arah_qibla < 360: break

    while arah_qibla < 0:
        arah_qibla = arah_qibla + 360
        if arah_qibla > 0: break
    return arah_qibla


def jarak_qibla(lintang, bujur):
    '''
    jarak satu titik ke ka'bah dalam satuan kilometer (km)
    '''
    lintang_kabah = 21 + 25/60 + 21.03/3600
    bujur_kabah = 39 + 49/60 + 34.31/3600
    lintang = dms_to_dec(lintang)
    bujur = dms_to_dec(bujur)

    phi1 = radians(lintang)
    phi2 = radians(lintang_kabah)
    dl = radians(bujur - bujur_kabah)
    return acos(sin(phi1)*sin(phi2) + cos(phi1)*cos(phi2)*cos(dl))*6371.137


def arah_qibla_with_ellipsoid_correction(lintang, bujur):
    lintang_kabah = 21 + 25/60 + 21.03/3600
    bujur_kabah = 39 + 49/60 + 34.31/3600
    lintang = dms_to_dec(lintang)
    bujur = dms_to_dec(bujur)
    
    E = 0.0066943800229
    lintang_kabah_terkoreksi = degrees(atan((1 - E)*tan(radians(lintang_kabah))))
    lintang_tempat_terkoreksi = degrees(atan((1 - E)*tan(radians(lintang))))

    A = radians(90 - lintang_tempat_terkoreksi)
    B = radians(90 - lintang_kabah_terkoreksi)
    c = radians(bujur - bujur_kabah)

    sB = sin(B)*sin(c)
    cB = cos(B)*sin(A) - cos(A)*sin(B)*cos(c)
    Bb = atan2(sB, cB)

    arah_qibla = 360 - degrees(Bb)
    while arah_qibla > 360:
        arah_qibla = arah_qibla - 360
        if arah_qibla < 360: break

    while arah_qibla < 0:
        arah_qibla = arah_qibla + 360
        if arah_qibla > 0: break
    return arah_qibla


def jarak_qibla_with_ellipsoid_correction(lintang, bujur):
    lintang_kabah = 21 + 25/60 + 21.03/3600
    bujur_kabah = 39 + 49/60 + 34.31/3600
    lintang = dms_to_dec(lintang)
    bujur = dms_to_dec(bujur)

    A = 6378.14
    fo = 1/298.257
    F = (lintang + lintang_kabah)/2
    G = (lintang - lintang_kabah)/2
    L = (bujur - bujur_kabah)/2
    S = (sin(radians(G))*cos(radians(L)))**2 + (cos(radians(F))*sin(radians(L)))**2
    c = (cos(radians(G))*cos(radians(L)))**2 + (sin(radians(F))*sin(radians(L)))**2
    w  = atan(sqrt(S/c))
    R = sqrt(S*c)/w
    D = 2*w*A
    H1 = (3*R - 1)/(2*c)
    H2 = (3*R + 1)/(2*S)
    return D *(1 + fo*H1*(sin(radians(F))*cos(radians(G)))**2 - fo*H2*(cos(radians(F))*sin(radians(G)))**2)


def arah_qibla_vincenty(lintang, bujur):
    lintang_kabah = 21 + 25/60 + 21.03/3600
    bujur_kabah = 39 + 49/60 + 34.31/3600
    lintang = dms_to_dec(lintang)
    bujur = dms_to_dec(bujur)

    F = 1/298.257223563
    ae = 6378137
    be = 6356752.314

    U1 = atan((1 - F)*tan(radians(lintang_kabah)))
    U2 = atan((1 - F)*tan(radians(lintang)))
    L = radians(bujur - bujur_kabah)    

    Lambda = L
    Lambda0 = 500
    iter_limit = 0

    while True:
        iter_limit += 1
        Lambda0 = Lambda
        sSigma = sqrt((cos(U2)*sin(Lambda))**2 + (cos(U1)*sin(U2) - sin(U1)*cos(U2)*cos(Lambda))**2)
        cSigma = sin(U1)*sin(U2) + cos(U1)*cos(U2)*cos(Lambda)
        sigma = atan2(sSigma, cSigma)

        sAlpha = cos(U1)*cos(U2)*sin(Lambda)/sSigma
        c2Alpha = 1 - sAlpha**2
        c2SigmaM = cSigma - 2*sin(U1)*sin(U2)/c2Alpha
        c = F/16*c2Alpha*(4+F*(4 - 3*c2Alpha))
        Lambda = L + (1-c)*F*sAlpha*(sigma + c*sSigma*(c2SigmaM + c*cSigma*(-1 + 2*c2SigmaM**2)))
        if abs(Lambda0 - Lambda) < 0.000000000001 or iter_limit > 100:
            break

    up2 = c2Alpha*(ae**2 - be**2)/be**2
    A = 1 + up2/16384*(4096 + up2*(-768 + up2*(320 - 175*up2)))
    B = up2/1024*(256 + up2*(-128 + up2*(74 - 47*up2)))

    dSigma = B * sSigma * (c2SigmaM + 0.25*B*(cSigma*(-1 + 2*c2SigmaM**2) - 1/6*B*c2SigmaM*(-3 + 4*sSigma**2)*(-3 + 4*c2SigmaM**2)))
    #  S = be*A*(sigma - dSigma)
    #  alpha1 = atan2(cos(U2)*sin(Lambda), cos(U1)*sin(U2) - sin(U1)*cos(U2)*cos(Lambda))
    alpha2 = atan2(cos(U1)*sin(Lambda), -sin(U1)*cos(U2) + cos(U1)*sin(U2)*cos(Lambda))
    
    arah_kibla_vincenty = -180 + degrees(alpha2)
    arah_qibla_vincenty = (arah_kibla_vincenty) % 360
    return arah_qibla_vincenty


def jarak_qibla_vincenty(lintang, bujur):
    lintang_kabah = 21 + 25/60 + 21.03/3600
    bujur_kabah = 39 + 49/60 + 34.31/3600
    lintang = dms_to_dec(lintang)
    bujur = dms_to_dec(bujur)

    F = 1/298.257223563
    ae = 6378137
    be = 6356752.314

    U1 = atan((1 - F)*tan(radians(lintang_kabah)))
    U2 = atan((1 - F)*tan(radians(lintang)))
    L = radians(bujur - bujur_kabah)    

    Lambda = L
    Lambda0 = 500
    iter_limit = 0

    while True:
        iter_limit += 1
        Lambda0 = Lambda
        sSigma = sqrt((cos(U2)*sin(Lambda))**2 + (cos(U1)*sin(U2) - sin(U1)*cos(U2)*cos(Lambda))**2)
        cSigma = sin(U1)*sin(U2) + cos(U1)*cos(U2)*cos(Lambda)
        sigma = atan2(sSigma, cSigma)

        sAlpha = cos(U1)*cos(U2)*sin(Lambda)/sSigma
        c2Alpha = 1 - sAlpha**2
        c2SigmaM = cSigma - 2*sin(U1)*sin(U2)/c2Alpha
        c = F/16*c2Alpha*(4+F*(4 - 3*c2Alpha))
        Lambda = L + (1-c)*F*sAlpha*(sigma + c*sSigma*(c2SigmaM + c*cSigma*(-1 + 2*c2SigmaM**2)))
        if abs(Lambda0 - Lambda) < 0.000000000001 or iter_limit > 100:
            break

    up2 = c2Alpha*(ae**2 - be**2)/be**2
    A = 1 + up2/16384*(4096 + up2*(-768 + up2*(320 - 175*up2)))
    B = up2/1024*(256 + up2*(-128 + up2*(74 - 47*up2)))

    dSigma = B * sSigma * (c2SigmaM + 0.25*B*(cSigma*(-1 + 2*c2SigmaM**2) - 1/6*B*c2SigmaM*(-3 + 4*sSigma**2)*(-3 + 4*c2SigmaM**2)))
    S = be*A*(sigma - dSigma)
    return S/1000
    

def shadow_ratio_vsop87(lintang, bujur, year, month, day):
    h_matahari = apparent_sun_altitude_vsop87(lintang, bujur, year, month, day)
    if h_matahari < 0:
        return 999
    else:
        return 1/tan(radians(h_matahari))


def shadow_direction_vsop87(lintang, bujur, year, month, day): 
    return (sun_azimuth_vsop87(lintang, bujur, year, month, day) + 180) % 360


def rashdul_qibla(lintang, bujur, timezone, year, month, day, opsi = 1):
    la = radians(lintang)
    if opsi == 1:
        qb = arah_qibla_vincenty(lintang, bujur)
    if opsi == 2:
        qb = arah_qibla_with_ellipsoid_correction(lintang, bujur)
    if opsi == 3:
        qb = arah_qibla(lintang, bujur)

    mDay = day
    while True:
        rq_old = rq
        D = radians(apparent_sun_declination_vsop87(year, month, mDay))
        E = equation_of_time(year, month, mDay)
        if qb > 180:
            U = atan(1/(tan(radians(360 - qb))*sin(la)))
            T = degrees(acos(tan(D)*cos(U)/tan(la)) + U)/15
            rq = 12 - E - T + kwd(timezone, bujur)
        else:
            U = atan(1/(tan(radians(qb))*sin(la)))
            T = degrees(acos(tan(D)*cos(U)/tan(la)) + U)/15
            rq = 12 - E - T + kwd(timezone, bujur)
        mDay = int(day) + (rq - timezone)/24
        if abs(rq - rq_old) < (1/3600):break

    direction = 0
    delta = (shadow_direction_vsop87(lintang, bujur, year, month, mDay) - qb)*3600

    if abs(delta) > 5:
        direction = 1
        delta = (shadow_direction_vsop87(lintang, bujur, year, month, mDay) + 180 - qb)*3600
    
    while abs(delta) > 1:
        inv = 20/86400
        _delta = (shadow_direction_vsop87(lintang, bujur, year, month, mDay + inv) + 180*direction - qb)*3600
        mDay = mDay + inv*delta/(delta - _delta)
        delta = (shadow_direction_vsop87(lintang, bujur, year, month, mDay) + 180*direction - qb)*3600
        if abs(delta) < 1: break
    
    rq = (mDay*24 + timezone) % 24
    if apparent_sun_altitude_vsop87(lintang, bujur, year, month, mDay) < 0:
        rq = 999
    return rq


############# SHALAT ###############

def equation_of_time(year, month, day):
    '''
    EQUATION OF TIME
    dalam satuan jam
    '''
    Lambda = sun_mean_longitude(year, month, day)
    alpha = apparent_sun_right_ascension_vsop87(year, month, day)
    dpsi = nutation_in_longitude(year, month, day)/3600
    epsilon = obliquity_of_ecliptic(year, month, day)
    eot = ((Lambda - 0.0057183 - alpha) + dpsi *cos(radians(epsilon)))/15
    if eot >23:
        eot = eot-24
    return eot


def shubuh_time(lintang, bujur, timezone, year, month, day, alt = -20, ikhtiyat = 2):
    return fajr_time(lintang, bujur, timezone, year, month, day, alt, ikhtiyat)


def fajr_time(lintang, bujur, timezone, year, month, day, alt = -20, ikhtiyat = 2):
    fajr_time = 999
    mDay = day
    if alt < 0:
        E = equation_of_time(year, month, mDay)
        h = sun_hour_angle_vsop87(lintang, alt, year, month, mDay)
        if h != 999:
            fajr_time = 12 - E - h + kwd(timezone, bujur) + ikhtiyat/60
    else:
        h_sun = sun_altitude_at_sunset_or_sunrise(year, month, mDay)
        terbit = sunrise(lintang, bujur, timezone, year, month, mDay, h_sun)
        if terbit != 999:
            fajr_time = terbit + alt/60
    
    if fajr_time == 999:exit()
    mDay = int(day) + (fajr_time-timezone)/24
    if alt < 0:
        E = equation_of_time(year, month, mDay)
        h = sun_hour_angle_vsop87(lintang, alt, year, month, mDay)
        if h !=999:
            fajr_time = 12 - E - h + kwd(timezone, bujur) + ikhtiyat/60
    else:
        h_sun = sun_altitude_at_sunset_or_sunrise(year, month, mDay)
        terbit = sunset(lintang, bujur, timezone, year, month, int(day) + (terbit - timezone)/24, h_sun)
        if terbit !=999:
            fajr_time = terbit + alt/60
    return fajr_time


def duhr_time(lintang, bujur, timezone, year, month, day, ikhtiyat=2):
    E = equation_of_time(year, month, day)
    duhr_time = 12 - E + kwd(timezone, bujur) + ikhtiyat/60

    E = equation_of_time(year, month, int(day) + (duhr_time - timezone)/24)
    return 12 - E + kwd(timezone, bujur) + ikhtiyat/60


def altitude_of_sun_at_ashr(lintang, bujur, timezone, year, month, day, rasio = 1, is_shadow_duhr = True, is_refraction = False):
    latitude = dms_to_dec(lintang)

    sun_dec = radians(apparent_sun_declination_vsop87(year, month, day))
    A = asin(cos(sun_dec)*cos(radians(latitude)) + sin(sun_dec)*sin(radians(latitude)))
    if is_refraction:
        A = A - radians(refraction_true_altitude(degrees(A), 1010, 27)/60)
    if is_shadow_duhr:
        altitude_of_sun_at_ashr = degrees(atan(1/(1/tan(A) + rasio)))
    else:
        altitude_of_sun_at_ashr = degrees(atan(1/rasio))
    if is_refraction:
        altitude_of_sun_at_ashr = altitude_of_sun_at_ashr - refraction_apparent_altitude(altitude_of_sun_at_ashr, 1010, 27)/60
    return altitude_of_sun_at_ashr


def ashr_time(lintang, bujur, timezone, year, month, day, alt = 1, ikhtiyat = 2):
    ashr_time = 999
    if alt == 1:
        alt = altitude_of_sun_at_ashr(lintang, bujur, timezone, year, month, day)

    E = equation_of_time(year, month, day)
    h = sun_hour_angle_vsop87(lintang, alt, year, month, day)
    if h != 999:
        ashr_time = 12 - E + h + kwd(timezone, bujur) + ikhtiyat/60
    if ashr_time == 999:
        exit
    E = equation_of_time(year, month, int(day) + (ashr_time - timezone)/24)
    h = sun_hour_angle_vsop87(lintang, alt, year, month, int(day) + (ashr_time - timezone)/24)
    if h !=999:
        ashr_time = 12 - E + h + kwd(timezone, bujur) + ikhtiyat/60
    return ashr_time


def maghrib_time(lintang, bujur, timezone, year, month, day, alt = -1, ikhtiyat = 2):
    terbenam = sunset(lintang, bujur, timezone, year, month, day, alt)
    maghrib_time = terbenam
    if maghrib_time !=999:
        maghrib_time = maghrib_time + ikhtiyat/60
    return maghrib_time

def isya_time(lintang, bujur, timezone, year, month, day, alt = -18, ikhtiyat = 2):
    isya_time = 999
    mDay = day
    if alt < 0:
        E = equation_of_time(year, month, mDay)
        h = sun_hour_angle_vsop87(lintang, alt, year, month, mDay)
        if h != 999:
            isya_time = 12 - E + h + kwd(timezone, bujur)
    else:
        h_sun = sun_altitude_vsop87(year, month, mDay)
        terbenam = sunrise(lintang, bujur, timezone, year, month, mDay, h_sun)
        if terbenam != 999:
            isya_time = terbenam + alt/60

    if isya_time == 999:
        exit
    mDay = int(day) + (isya_time - timezone)/24
    if alt < 0:
        E = equation_of_time(year, month, mDay)
        h = sun_hour_angle_vsop87(lintang, alt, year, month, mDay)
        if h != 999:
            isya_time = 12 - E + h + kwd(timezone, bujur) + ikhtiyat/60
    else:
        h_sun = sun_altitude_at_sunset_or_sunrise(year, month, mDay)
        terbenam = sunset(lintang, bujur, timezone, year, month, int(day) + (terbenam - timezone)/24, h_sun)
        if terbenam != 999:
            isya_time = terbenam + alt/60
    return isya_time


def sun_mean_longitude(year, month, day):
    t = julian_ephemeris_century(year, month, day)
    tau = t/10
    return (280.4664567 + 360007.6982779*tau + 0.03032028*tau**2 + tau**3/49931 - tau**4/15299 - tau**5/1988000) % 360


def sun_mean_anomaly(year, month, day):
    t = julian_ephemeris_century(year, month, day)
    return 357.5291092 + 35999.0502909*t - 0.0001536*t**2 + t**3/24490000


def sun_longitude_low_accuracy(year, month, day):
    t = julian_ephemeris_century(year, month, day)
    m = radians(357.5291 + 35999.0503*t - 0.0001559*t**2 - 0.00000048*t**3)
    c = (1.9146 - 0.004817*t - 0.000014*t**2) * sin(m) + (0.019993 - 0.000101*t) * sin(2*m) + 0.00029*sin(3*m)
    lo = (280.46645 + 36000.76983*t + 0.0003032*t**2) % 360
    return (lo + c) % 360


def sun_latitude_low_accuracy(year, month, day):
    return 0


def distance_to_the_sun_low_accuracy(year, month, day):
    e = eccentricity_of_earth_orbit(year, month, day)
    m = radians(sun_mean_anomaly(year, month, day))
    c = radians(sun_eq_center(year, month, day))
    v = m + c
    return 1.000001018*(1 - e**2)/(1 + e*cos(v))


def apparent_sun_longitude_low_accuracy(year, month, day):
    result = sun_longitude_low_accuracy(year, month, day) + sun_aberration(year, month, day)/3600 + nutation_in_longitude(year, month, day)/3600
    return result

def _theta(year, month, day):
    t = julian_ephemeris_century(year, month, day)
    tau = t/10
    L0 = terms_L0
    L1 = terms_L1
    L2 = terms_L2
    L3 = terms_L3
    L4 = terms_L4
    L5 = terms_L5

    totalL0 = totalL1 = totalL2 = totalL3 = totalL4 = totalL5 = 0

    for i in range(len(L0)):
        totalL0 += L0[i][0] * cos(L0[i][1] + L0[i][2] * tau)
    for i in range(len(L1)):
        totalL1 += L1[i][0] * cos(L1[i][1] + L1[i][2] * tau)
    for i in range(len(L2)):
        totalL2 += L2[i][0] * cos(L2[i][1] + L2[i][2] * tau)
    for i in range(len(L3)):
        totalL3 += L3[i][0] * cos(L3[i][1] + L3[i][2] * tau)
    for i in range(len(L4)):
        totalL4 += L4[i][0] * cos(L4[i][1] + L4[i][2] * tau)
    for i in range(len(L5)):
        totalL5 += L5[i][0] * cos(L5[i][1] + L5[i][2] * tau)

    return (totalL0 + totalL1*tau + totalL2*tau**2 + totalL3*tau**3 + totalL4*tau**4 + totalL5*tau**5)
    ...

def _beta(year, month, day):
    t = julian_ephemeris_century(year, month, day)
    tau = t/10
    B0 = terms_B0
    B1 = terms_B1
    B2 = terms_B2
    B3 = terms_B3
    B4 = terms_B4

    totalB0 = totalB1 = totalB2 = totalB3 = totalB4 = 0
    for i in range(len(B0)):
        totalB0 += B0[i][0] * cos(B0[i][1] + B0[i][2] * tau)
    for i in range(len(B1)):
        totalB1 += B1[i][0] * cos(B1[i][1] + B1[i][2] * tau)
    for i in range(len(B2)):
        totalB2 += B2[i][0] * cos(B2[i][1] + B2[i][2] * tau)
    for i in range(len(B3)):
        totalB3 += B3[i][0] * cos(B3[i][1] + B3[i][2] * tau)
    for i in range(len(B4)):
        totalB4 += B4[i][0] * cos(B4[i][1] + B4[i][2] * tau)
    return (totalB0 + totalB1*tau + totalB2*tau**2 + totalB3*tau**3 + totalB4*tau**4)


def sun_longitude_vsop87(year, month, day):
    '''
    
    '''
    t = julian_ephemeris_century(year, month, day)
    tau = t/10

    theta = _theta(year, month, day)
    beta = _beta(year, month, day)
    beta = -(degrees(beta))
    theta = 180 + degrees(theta)
    theta =  (theta) % 360
    lambda1 = theta - 1.397*t - 0.00031*t**2
    delta_beta = 0.03916*(cos(radians(lambda1)) - sin(radians(lambda1)))/3600
    beta = beta + delta_beta
    delta_theta = (-0.09033 + 0.03916*(cos(radians(lambda1)) + sin(radians(lambda1)))*tan(beta))/3600
    return (theta + delta_theta) % 360


def sun_latitude_vsop87(year, month, day):
    '''
    
    '''
    t = julian_ephemeris_century(year, month, day)
    tau = t/10

    theta = _theta(year, month, day)
    beta = _beta(year, month, day)
    beta = -(degrees(beta))
    theta = 180 + degrees(theta)
    theta = (theta) % 360
    lambda1 = theta - 1.397*t - 0.00031*t**2
    delta_beta = 0.03916*(cos(radians(lambda1)) - sin(radians(lambda1)))/3600
    return beta + delta_beta


def distance_to_the_sun_vsop87(year, month, day):
    '''belum selesai, masih mencari terms R0-R5'''
    tau = julian_ephemeris_century(year, month, day)/10
    R0 = terms_R0
    R1 = terms_R1
    R2 = terms_R2
    R3 = terms_R3
    R4 = terms_R4
    R5 = terms_R5

    totalR0 = totalR1 = totalR2 = totalR3 = totalR4 = totalR5 = 0

    for i in range(len(R0)):
        totalR0 += R0[i][0] * cos(R0[i][1] + R0[i][2] * tau)
    for i in range(len(R1)):
        totalR1 += R1[i][0] * cos(R1[i][1] + R1[i][2] * tau)
    for i in range(len(R2)):
        totalR2 += R2[i][0] * cos(R2[i][1] + R2[i][2] * tau)
    for i in range(len(R3)):
        totalR3 += R3[i][0] * cos(R3[i][1] + R3[i][2] * tau)
    for i in range(len(R4)):
        totalR4 += R4[i][0] * cos(R4[i][1] + R4[i][2] * tau)
    for i in range(len(R5)):
        totalR5 += R5[i][0] * cos(R5[i][1] + R5[i][2] * tau)

    return (totalR0 + totalR1*tau + totalR2*tau**2 + totalR3*tau**3 + totalR4*tau**4 + totalR5*tau**5)

    
def apparent_sun_longitude_vsop87(year, month, day):
    return sun_longitude_vsop87(year, month, day) + sun_aberration(year, month, day)/3600 + nutation_in_longitude(year, month, day)/3600


def sun_declination_low_accuracy(year, month, day):
    Lambda = radians(sun_longitude_low_accuracy(year, month, day))
    beta = radians(sun_latitude_low_accuracy(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))
    sin_delta = sin(beta)*cos(epsilon) + cos(beta)*sin(epsilon)*sin(Lambda)
    return degrees(asin(sin_delta))


def sun_declination_vsop87(year, month, day):
    Lambda = radians(sun_longitude_vsop87(year, month, day))
    beta = radians(sun_latitude_vsop87(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))
    sin_delta = sin(beta)*cos(epsilon) + cos(beta)*sin(epsilon)*sin(Lambda)
    return degrees(asin(sin_delta))


def apparent_sun_declination_low_accuracy(year, month, day):
    Lambda = radians(apparent_sun_longitude_low_accuracy(year, month, day))
    beta = radians(sun_latitude_low_accuracy(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))
    sin_delta = sin(beta)*cos(epsilon) + cos(beta)*sin(epsilon)*sin(Lambda)
    return degrees(asin(sin_delta))


def apparent_sun_declination_vsop87(year, month, day):
    Lambda = radians(apparent_sun_longitude_vsop87(year, month, day))
    beta = radians(sun_latitude_vsop87(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))
    sin_delta = sin(beta)*cos(epsilon) + cos(beta)*sin(epsilon)*sin(Lambda)
    return degrees(asin(sin_delta))


def sun_right_ascension_low_accuracy(year, month, day):
    b = radians(sun_latitude_low_accuracy(year, month, day))
    l = radians(sun_longitude_low_accuracy(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))
    sin_alpha = -sin(b)*sin(epsilon) + cos(b)*cos(epsilon)*sin(l)
    cos_alpha = cos(b)*cos(l)
    return (degrees(atan2(sin_alpha, cos_alpha))) % 360


def sun_right_ascension_vsop87(year, month, day):
    b = radians(sun_latitude_vsop87(year, month, day))
    l = radians(sun_longitude_vsop87(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))

    sin_alpha = -sin(b)*sin(epsilon) + cos(b)*cos(epsilon)*sin(l)
    cos_alpha = cos(b)*cos(l)
    return (degrees(atan2(sin_alpha, cos_alpha))) % 360


def apparent_sun_right_ascension_low_accuracy(year, month, day):
    b = radians(sun_latitude_low_accuracy(year, month, day))
    l = radians(apparent_sun_longitude_low_accuracy(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))
    sin_alpha = -sin(b)*sin(epsilon) + cos(b) * cos(epsilon)*sin(l)
    cos_alpha = cos(b)*cos(l)
    return (degrees(atan2(sin_alpha, cos_alpha))) % 360


def apparent_sun_right_ascension_vsop87(year, month, day):
    b = radians(sun_latitude_vsop87(year, month, day))
    l = radians(apparent_sun_longitude_vsop87(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))
    sin_alpha = -sin(b)*sin(epsilon) + cos(b)*cos(epsilon)*sin(l)
    cos_alpha = cos(b)*cos(l)
    return (degrees(atan2(sin_alpha, cos_alpha))) % 360


def sun_altitude_vsop87(lintang, bujur, year, month, day):
    latitude = radians(dms_to_dec(lintang))
    h = radians(hour_angle_of_the_sun(bujur, year, month, day)*15)
    dec = radians(apparent_sun_declination_vsop87(year, month, day))
    sun_altitude = cos(h)*cos(dec) * cos(latitude) + sin(dec) * sin(latitude)
    hasil = degrees(asin(sun_altitude))
    return hasil


def apparent_sun_altitude_vsop87(lintang, bujur, year, month, day):
    h = sun_altitude_vsop87(lintang, bujur, year, month, day)
    r = refraction_true_altitude(h, 1010, 27)/60
    return h + r


def sun_azimuth_vsop87(lintang, bujur, year, month, day): 
    latitude = dms_to_dec(lintang)
    h = radians(hour_angle_of_the_sun(bujur, year, month, day)*15)
    dec = radians(apparent_sun_declination_vsop87(year, month, day))
    cosA = cos(h)*sin(radians(latitude)) - tan(dec)*cos(radians(latitude))
    hasil = (degrees(atan2(sin(h), cosA)) + 180) % 360
    return hasil


def sun_altitude_at_sunset_or_sunrise(year, month, day, p = 1010, t = 27, h = 0): 
    sd = (sun_semi_diameter_vsop87(year, month, day+0.5))
    hasil = -(sd + refraction_apparent_altitude(0, p, t) + koreksi_dip(h))
    return hasil


def sunset(lintang, bujur, timezone, year, month, day, alt = -1): 
    E = equation_of_time(year, month, day)
    h = sun_hour_angle_vsop87(lintang, alt, year, month, day)
    terbenam = 999
    if h != 999:
        terbenam = 12 - E + h + kwd(timezone, bujur)
    if terbenam != 999:
        E = equation_of_time(year, month, int(day) + (terbenam - timezone)/24)
        h = sun_hour_angle_vsop87(lintang, alt, year, month, int(day) + (terbenam - timezone)/24)
        terbenam = 12 - E + h + kwd(timezone, bujur)
    return terbenam


def sunrise(lintang, bujur, timezone, year, month, day, alt = -1):
    E = equation_of_time(year, month, day)
    h = sun_hour_angle_vsop87(lintang, alt, year, month, day)
    sunrise = 999
    if h != 999:
        sunrise = 12 - E - h + kwd(timezone, bujur)
    if sunrise !=999:
        E = equation_of_time(year, month, int(day) + (sunrise - timezone)/24)
        h = sun_hour_angle_vsop87(lintang, alt, year, month, int(day) + (sunrise - timezone)/24)
        sunrise = 12 - E - h + kwd(timezone, bujur)
    return sunrise

def sun_semi_diameter_low_accuracy(year, month, day): 
    R = distance_to_the_sun_low_accuracy(year, month, day)
    return (959.63/R)/60


def sun_semi_diameter_vsop87(year, month, day):
    R = distance_to_the_sun_vsop87(year, month, day)
    return (959.63/R)/60


def sun_eq_center(year, month, day):
    T = julian_ephemeris_century(year, month, day)
    M = radians(sun_mean_anomaly(year, month, day))
    suneqcenter = (1.9146 - 0.004817*T - 0.000014*T**2)*sin(M) + (0.019993 - 0.000101*T)*sin(2*M) + 0.00029*sin(3*M)
    return suneqcenter

def eccentricity_of_earth_orbit(year, month, day):
    T = julian_ephemeris_century(year, month, day)
    return 0.016708617 - 0.000042037*T - 0.0000001236*T**2


def hour_angle_of_the_sun(bujur, year, month, day):
    apparent_lst = apparent_local_sidereal_time(bujur, year, month, day)
    apparent_ra_vsop87 = apparent_sun_right_ascension_vsop87(year, month, day)
    return (apparent_lst - apparent_ra_vsop87/15) % 24


def sun_hour_angle_vsop87(lintang, alt, year, month, day):
    lat = dms_to_dec(lintang)
    delta = radians(apparent_sun_declination_vsop87(year, month, day))
    cosh = (sin(radians(alt)) - sin(delta)*sin(radians(lat)))/(cos(delta)*cos(radians(lat)))
    return degrees(acos(cosh))*24/360

def moon_mean_longitude(year, month, day):
    T = julian_ephemeris_century(year, month, day)
    moonmeanlongitude = 218.3164591 + 481267.88134236*T - 0.0013268*T**2 + T**3/538841 - T**4/65194000
    return moonmeanlongitude

def ijtimak(year, month, day):
    return new_moon(year, month, day)


def konjungsi(year, month, day):
    return new_moon(year, month, day)


def new_moon(year, month, day):
    Y = year
    M = month
    D = day
    jde = julian_datum(Y, M, D)
    while True:
        jdeo = nilai_jde(Y, M, D, 0)
        CP = koreksi_new_moon(Y, M, D)
        c = koreksi_planetary_arguments(Y, M, D, 0)
        new_moon = jdeo + CP + c
        toymd = jd_to_ymd(new_moon)
        Y = toymd[0]
        M = toymd[1]
        D = toymd[2] + (toymd[3]/24 + toymd[4]/1440 + toymd[5]/86400) + 29.2
        if new_moon < jde: break
    
    while (new_moon - jde) < 30:
        jdeo = nilai_jde(Y, M, D, 0)
        CP = koreksi_new_moon(Y, M, D)
        c = koreksi_planetary_arguments(Y, M, D, 0)
        new_moon = jdeo + CP + c
        Y = toymd[0]
        M = toymd[1]
        D = toymd[2] + (toymd[3]/24 + toymd[4]/1440 + toymd[5]/86400) - 29.2
        if (new_moon - jde) > 30:break
    new_moon = new_moon - delta_t(year, month, day)
    return new_moon


def first_quarter(year, month, day):
    Y = year
    M = month
    D = day
    jde = julian_datum(Y, M, D)
    while True:
        jdeo = nilai_jde(Y, M, D, 0.25)
        CP = koreksi_new_moon(Y, M, D)
        w = koreksi_w(Y, M, D, 0.25)
        c = koreksi_planetary_arguments(Y, M, D, 0.25)
        first_quarter = jdeo + w + CP + c
        toymd = jd_to_ymd(first_quarter)
        Y = toymd[0]
        M = toymd[1]
        D = toymd[2] + (toymd[3]/24 + toymd[4]/1440 + toymd[5]/86400) + 29.2
        if first_quarter < jde: break
    
    while (first_quarter - jde) < 30:
        jdeo = nilai_jde(Y, M, D, 0.25)
        CP = koreksi_first_quarter(Y, M, D)
        w = koreksi_w(Y, M, D, 0.25)
        c = koreksi_planetary_arguments(Y, M, D, 0.25)
        first_quarter = jdeo + w + CP + c
        toymd = jd_to_ymd(first_quarter)
        Y = toymd[0]
        M = toymd[1]
        D = toymd[2] + (toymd[3]/24 + toymd[4]/1440 + toymd[5]/86400) - 29.2
        if (first_quarter - jde) > 30:break
    first_quarter = first_quarter - delta_t(year, month, day)
    return first_quarter    


def full_moon(year, month, day):
    Y = year
    M = month
    D = day
    jde = julian_datum(Y, M, D)
    while True:
        jdeo = nilai_jde(Y, M, D, 0.5)
        CP = koreksi_new_moon(Y, M, D)
        c = koreksi_planetary_arguments(Y, M, D, 0.5)
        full_moon = jdeo + CP + c
        toymd = jd_to_ymd(full_moon)
        Y = toymd[0]
        M = toymd[1]
        D = toymd[2] + (toymd[3]/24 + toymd[4]/1440 + toymd[5]/86400) + 29.2
        if full_moon < jde: break
    
    while (full_moon - jde) < 30:
        jdeo = nilai_jde(Y, M, D, 0.5)
        CP = koreksi_full_moon(Y, M, D)
        c = koreksi_planetary_arguments(Y, M, D, 0.5)
        full_moon = jdeo + CP + c
        Y = toymd[0]
        M = toymd[1]
        D = toymd[2] + (toymd[3]/24 + toymd[4]/1440 + toymd[5]/86400) - 29.2
        if (full_moon - jde) > 30:break
    full_moon = full_moon - delta_t(year, month, day)
    return full_moon

def last_quarter(year, month, day):
    Y = year
    M = month
    D = day
    jde = julian_datum(Y, M, D)
    while True:
        jdeo = nilai_jde(Y, M, D, 0.75)
        CP = koreksi_new_moon(Y, M, D)
        w = koreksi_w(Y, M, D, 0.75)
        c = koreksi_planetary_arguments(Y, M, D, 0.75)
        last_quarter = jdeo + w + CP + c
        toymd = jd_to_ymd(last_quarter)
        Y = toymd[0]
        M = toymd[1]
        D = toymd[2] + (toymd[3]/24 + toymd[4]/1440 + toymd[5]/86400) + 29.2
        if last_quarter < jde: break
    
    while (last_quarter - jde) < 30:
        jdeo = nilai_jde(Y, M, D, 0.75)
        CP = koreksi_last_quarter(Y, M, D)
        w = koreksi_w(Y, M, D, 0.75)
        c = koreksi_planetary_arguments(Y, M, D, 0.75)
        last_quarter = jdeo + w + CP + c
        toymd = jd_to_ymd(last_quarter)
        Y = toymd[0]
        M = toymd[1]
        D = toymd[2] + (toymd[3]/24 + toymd[4]/1440 + toymd[5]/86400) - 29.2
        if (last_quarter - jde) > 30:break
    last_quarter = last_quarter - delta_t(year, month, day)
    return last_quarter


def moon_longitude(year, month, day):
    moon_longitude = (moon_mean_longitude(year, month, day) + sigma_l(year, month, day)/100000) % 360
    return moon_longitude


def apparent_moon_longitude(year, month, day):
    apparent_moon_longitude = (moon_longitude(year, month, day) + nutation_in_longitude(year, month, day)) % 360
    return apparent_moon_longitude


def moon_latitude(year, month, day):
    moon_latitude = sigma_b(year, month, day)/1000000
    return moon_latitude


def distance_to_the_moon(year, month, day):
    distance_to_the_moon = 385000.56 + sigma_r(year, month, day)/1000
    return distance_to_the_moon


def moon_mean_elongation(year, month, day):
    t = julian_ephemeris_century(year, month, day)
    return 297.8502042 + 445267.1115168*t - 0.00163 *t**2 + t**3/545868 - t**4/113065000


def moon_declination(year, month, day):
    B = radians(moon_latitude(year, month, day))
    L = radians(moon_longitude(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))
    sin_dec = sin(B)*cos(epsilon) + cos(B)* sin(epsilon)*sin(L)
    moon_declination = degrees(asin(sin_dec))
    return moon_declination


def apparent_moon_declination(year, month, day):
    B = radians(moon_latitude(year, month, day))
    L = radians(apparent_moon_longitude(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))
    sin_dec = sin(B)*cos(epsilon) + cos(B)* sin(epsilon)*sin(L)
    apparent_moon_declination = degrees(asin(sin_dec))
    return apparent_moon_declination


def moon_right_ascension(year, month, day):
    dec = radians(moon_declination(year, month, day))
    B = radians(moon_latitude(year, month, day))
    L = radians(moon_longitude(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))

    sin_alpha = -sin(B)*sin(epsilon) + cos(B)*cos(epsilon)*sin(L)
    sin_alpha = sin_alpha/cos(dec)
    cos_alpha = cos(B)*cos(L)/cos(dec)
    moon_right_ascension = degrees(atan2(sin_alpha, cos_alpha))
    return moon_right_ascension


def apparent_moon_right_ascension(year, month, day):
    dec = radians(apparent_moon_declination(year, month, day))
    B = radians(moon_latitude(year, month, day))
    L = radians(apparent_moon_longitude(year, month, day))
    epsilon = radians(obliquity_of_ecliptic(year, month, day))

    sin_alpha = -sin(B)*sin(epsilon) + cos(B)*cos(epsilon)*sin(L)
    sin_alpha = sin_alpha/cos(dec)
    cos_alpha = cos(B)*cos(L)/cos(dec)
    apparent_moon_right_ascension = degrees(atan2(sin_alpha, cos_alpha))
    return apparent_moon_right_ascension


def moon_elongation(year, month, day):
    d0 = radians(apparent_sun_declination_vsop87(year, month, day))
    d1 = radians(apparent_moon_declination(year, month, day))
    a0 = radians(apparent_sun_right_ascension_vsop87(year, month, day))
    A1 = radians(apparent_moon_right_ascension(year, month, day))
    cos_elo = sin(d0)*sin(d1) + cos(d0)*cos(d1)*cos(a0 - A1)
    moon_elongation = degrees(acos(cos_elo))
    return moon_elongation


def moon_elongation_at_sunset(lintang, bujur, timezone, year, month, day, alt = -1):
    terbenam = sunset(lintang, bujur, timezone, year, month, day, alt)
    moon_elongation_at_sunset = moon_elongation(year, month, int(day) + (terbenam - timezone)/24)


def moon_age(year, month, day):
    jd = julian_datum(year, month, day)
    jdi = new_moon(year, month, day)
    if jdi > jd:
        toymd = jd_to_ymd(jdi)
        Y = toymd[0]
        M = toymd[1]
        D = toymd[2] + (toymd[3]/24 + toymd[4]/1440 + toymd[5]/86400) - 30
        jdi = new_moon(Y, M, D)
    moon_age = jd - jdi
    return moon_age


def moon_age_at_sunset(lintang, bujur, timezone, year, month, day, alt = -1):
    terbenam = sunset(lintang, bujur, timezone, year, month, day, alt)
    jd_terbenam = julian_datum(year, month, int(day) + (terbenam - timezone)/24)
    jd_ijtimak = new_moon(year, month, day)
    if jd_ijtimak > jd_terbenam:
        to_ymd = jd_to_ymd(jd_ijtimak)
        Y = to_ymd[0]
        M = to_ymd[1]
        D = to_ymd[2] + (to_ymd[3]/24 + to_ymd[4]/1440 + to_ymd[5]/86400) - 30
        jd_ijtimak = new_moon(Y, M, D)
    moon_age_at_sunset = jd_terbenam - jd_ijtimak
    return moon_age_at_sunset


def moon_illumination(year, month, day):
    elo = radians(moon_elongation(year, month, day))
    R = distance_to_the_sun_vsop87(year, month, day)*149597870.691
    D = distance_to_the_moon(year, month, day)
    i = atan2(R*sin(elo), D - R*cos(elo))
    k = (1 + cos(i))/2
    moon_illumination = k
    return moon_illumination


def moon_illumination_at_sunset(lintang, bujur, timezone, year, month, day, alt = -1):
    terbenam = sunset(lintang, bujur, timezone, year, month, day, alt)
    moon_illumination_at_sunset = moon_illumination(year, month, int(day) + (terbenam - timezone)/24)
    return moon_illumination_at_sunset


def crescent_width(year, month, day):
    mSD = moon_semi_diameter(year, month, day)
    illum = moon_illumination(year, month, day)
    crescent_width = illum*2*mSD
    return crescent_width


def crescent_width_at_sunset(lintang, bujur, timezone, year, month, day):
    alt = sun_altitude_at_sunset_or_sunrise(year, month, day, 1010, 27, 0)/60
    m_terbenam = sunset(lintang, bujur, timezone, year, month, day, alt)
    mSD = moon_semi_diameter(year, month, int(day) + (m_terbenam - timezone)/24)
    illum = moon_illumination(year, month, int(day) + (m_terbenam - timezone)/24)
    crescent_width_at_sunset = illum*2*mSD
    return crescent_width_at_sunset


def azimuth_difference_at_sunset(lintang, bujur, timezone, year, month, day):
    alt = sun_altitude_at_sunset_or_sunrise(year, month, day, 1010, 27, 0)/60
    m_terbenam = sunset(lintang, bujur, timezone, year, month, day, alt)
    m_az = sun_azimuth_vsop87(lintang, bujur, year, month, int(day) + (m_terbenam - timezone)/24)
    b_az = moon_azimuth(lintang, bujur, year, month, int(day) + (m_terbenam - timezone)/24)
    azimuth_difference_at_sunset = b_az - m_az
    return azimuth_difference_at_sunset


def azimuth_difference(lintang, bujur, timezone, year, month, day):
    m_az = sun_azimuth_vsop87(lintang, bujur, year, month, day)
    b_az = moon_azimuth(lintang, bujur, year, month, day)
    azimuth_difference = b_az - m_az
    return azimuth_difference


def moon_set(lintang, bujur, timezone, year, month, day):
    la = radians(dms_to_dec(lintang))
    dt = delta_t(year, month, day)
    hset = 0.7275*horizontal_moon_parallax(year, month, day) - refraction_apparent_altitude(0, 1010, 27)/60

    jde = julian_ephemeris_day(year, month, day)
    gmst = greenwich_mean_sidereal_time(year, month, day)*15

    dec = apparent_moon_declination(year, month, day)
    ra = apparent_moon_right_ascension(year, month, day)

    h0 = acos((sin(radians(hset)) - sin(la)*sin(radians(dec)))/(cos(la)*cos(radians(dec))))

    A1 = dec - apparent_moon_declination(year, month, day-1)
    B1 = apparent_moon_declination(year, month, day+1) - dec
    c1 = B1 - A1
    A2 = ra - apparent_moon_right_ascension(year, month, day - 1)
    B2 = apparent_moon_right_ascension(year, month, day + 1) - ra
    c2 = B2 - A2

    m0 = ((ra - dms_to_dec(bujur) - gmst)/360) % 1
    M = m0 + degrees(h0)/360

    dm = 1
    while dm > 0.0000001:
        theta = gmst + 360.985647 * M
        n = M + dt
        D = radians(dec + 0.5*n*(A1 + B1 + n*c1))
        h = radians(theta + dms_to_dec(bujur) - (ra + 0.5*n*(A2 + B2 + n*c2)))
        hh = degrees(asin(sin(la) * sin(D) + cos(la)*cos(D)*cos(h)))
        dm = (hh - hset)/(360*cos(la)*cos(D)*sin(h))
        M = M + dm
        if dm < 0.0000001: break

    moonset = (jde + M + timezone/24 - dt) - julian_datum(year, month, int(day))
    moonset = moonset*24
    return moonset


def moon_rise(lintang, bujur, timezone, year, month, day):
    la = radians(dms_to_dec(lintang))
    dt = delta_t(year, month, day)
    hrise = 0.7275*horizontal_moon_parallax(year, month, day) - refraction_apparent_altitude(0, 1010, 27)/60

    jde = julian_ephemeris_day(year, month, day)
    gmst = greenwich_mean_sidereal_time(year, month, day)*15

    dec = apparent_moon_declination(year, month, day)
    ra = apparent_moon_right_ascension(year, month, day)

    h0 = acos((sin(radians(hrise)) - sin(la)*sin(radians(dec)))/(cos(la)*cos(radians(dec))))

    A1 = dec - apparent_moon_declination(year, month, day - 1)
    B1 = apparent_moon_declination(year, month, day + 1) - dec
    c1 = B1 - A1
    A2 = ra - apparent_moon_right_ascension(year, month, day - 1)
    B2 = apparent_moon_right_ascension(year, month, day + 1) - ra
    c2 = B2 - A2

    m0 = ((ra - dms_to_dec(bujur) - gmst)/360) % 1
    M = m0 - degrees(h0)/360

    dm = 1
    while dm > 0.0000001:
        theta = gmst + 360.985647 * M
        n = M + dt
        D = radians(dec + 0.5*n*(A1 + B1 + n*c1))
        h = radians(theta + dms_to_dec(bujur) - (ra + 0.5*n*(A2 + B2 + n*c2)))
        hh = degrees(asin(sin(la) * sin(D) + cos(la)*cos(D)*cos(h)))
        dm = (hh - hrise)/(360*cos(la)*cos(D)*sin(h))
        M = M + dm
        if dm < 0.0000001: break

    moonrise = (jde + M + timezone/24 - dt) - julian_datum(year, month, int(day))
    moonrise = moonrise*24
    return moonrise


def moon_transit(lintang, bujur, timezone, year, month, day): 
    jde = julian_ephemeris_day(year, month, day)
    gmst = greenwich_mean_sidereal_time(year, month, day)*15

    dt = delta_t(year, month, day)
    RA = apparent_moon_right_ascension(year, month, day)
    A = RA - apparent_moon_right_ascension(year, month, day - 1)
    B = apparent_moon_right_ascension(year, month, day + 1) - RA

    M = (RA - dms_to_dec(bujur) - gmst) % 1

    dm = 1
    while dm > 0.0000001:
        theta = gmst + 360.985647*M
        n = M + dt
        h = theta + dms_to_dec(bujur) - (RA + 0.5*n*(A + B + n*c))
        dm = -h/360
        M = M + dm
        if dm < 0.0000001: break

    moontransit = jde + M + timezone/24 - dt
    return moontransit


def moon_altitude(lintang, bujur, year, month, day): 
    la = radians(dms_to_dec(lintang))
    h = radians(hour_angle_of_the_moon(bujur, year, month, day)*15)
    dec = radians(apparent_moon_declination(year, month, day))
    moon_altitude = cos(h)*cos(dec)*cos(la) + sin(dec)*sin(la)
    moon_altitude = degrees(asin(moon_altitude))
    return moon_altitude


def apparent_moon_altitude(lintang, bujur, year, month, day): 
    h = moon_altitude(lintang, bujur, year, month, day)
    apparent_moon_altitude = h + refraction_true_altitude(h, 1010, 27)/60
    return apparent_moon_altitude


def moon_azimuth(lintang, bujur, year, month, day): 
    la = dms_to_dec(lintang)
    h = radians(hour_angle_of_the_moon(bujur, year, month, day)*15)
    dec = radians(apparent_moon_declination(year, month, day))
    cos_a = cos(h)*sin(radians(la)) - tan(dec)*cos(radians(la))
    moon_azimuth = (degrees(atan2(sin(h), cos_a)) + 180) % 360
    return moon_azimuth


def hilal_duration(lintang, bujur, timezone, year, month, day): 
    alt = sun_altitude_at_sunset_or_sunrise(year, month, day, 1010, 27, 0)/60
    m_terbenam = sunset(lintang, bujur, timezone, year, month, day, alt)
    b_terbenam = moon_set(lintang, bujur, timezone, year, month, day)
    hilal_duration = (b_terbenam - m_terbenam)*60
    if abs(hilal_duration) > 100:
        hilal_duration = 999
    return hilal_duration


def moon_semi_diameter(year, month, day): 
    R = distance_to_the_moon(year, month, day)
    moonsemidiameter = (358473400/R)/60
    return moonsemidiameter


def moon_mean_anomaly(year, month, day):
    t = julian_ephemeris_century(year, month, day)
    return 134.9634114 + 477198.8676313*t + 0.008997*t**2 + t**3/69699 - t**4/14712000


def jde_of_max_north_moon_declination(year, month, day):
    k = value_k(year, month, day)
    T = value_t(year, month, day)
    D = radians(152.2029 + 333.0705546 * k - 0.0004025 * T**2 + 0.00000011 * T**3)
    M = radians(14.8591 + 26.9281592 * k - 0.0000544 * T**2 - 0.0000001 * T**3)
    M1 = radians(4.6881 + 356.9562795 * k + 0.0103126 * T**2 + 0.00001251 * T**3)
    F = radians(325.8867 + 1.4467806 * k - 0.0020708 * T**2 - 0.00000215 * T**3)
    E = nilai_e(year, month, day)
   
    P = 0.8975 * cos(F)
    P = P - 0.4726 * sin(M1)
    P = P - 0.103 * sin(2 * F)
    P = P - 0.0976 * sin(2 * D - M1)
    P = P - 0.0462 * cos(M1 - F)
    P = P - 0.0461 * cos(M1 + F)
    P = P - 0.0438 * sin(2 * D)
    P = P + 0.0162 * E * sin(M)
    P = P - 0.0157 * cos(3 * F)
    P = P + 0.0145 * sin(M1 + 2 * F)
    P = P + 0.0136 * cos(2 * D - F)
    P = P - 0.0095 * cos(2 * D - M1 - F)
    P = P - 0.0091 * cos(2 * D - M1 + F)
    P = P - 0.0089 * cos(2 * D + F)
    P = P + 0.0075 * sin(2 * M1)
    P = P - 0.0068 * sin(M1 - 2 * F)
    P = P + 0.0061 * cos(2 * M1 - F)
    P = P - 0.0047 * sin(M1 + 3 * F)
    P = P - 0.0043 * E * sin(2 * D - M - M1)
    P = P - 0.004 * cos(M1 - 2 * F)
    P = P - 0.0037 * sin(2 * D - 2 * M1)
    P = P + 0.0031 * sin(F)
    P = P + 0.003 * sin(2 * D + M1)
    P = P - 0.0029 * cos(M1 + 2 * F)
    P = P - 0.0029 * E * sin(2 * D - M)
    P = P - 0.0027 * sin(M1 + F)
    P = P + 0.0024 * E * sin(M - M1)
    P = P - 0.0021 * sin(M1 - 3 * F)
    P = P + 0.0019 * sin(2 * M1 + F)
    P = P + 0.0018 * cos(2 * D - 2 * M1 - F)
    P = P + 0.0018 * sin(3 * F)
    P = P + 0.0017 * cos(M1 + 3 * F)
    P = P + 0.0017 * cos(2 * M1)
    P = P - 0.0014 * cos(2 * D - M1)
    P = P + 0.0013 * cos(2 * D + M1 + F)
    P = P + 0.0013 * cos(M1)
    P = P + 0.0012 * sin(3 * M1 + F)
    P = P + 0.0011 * sin(2 * D - M1 + F)
    P = P - 0.0011 * cos(2 * D - 2 * M1)
    P = P + 0.001 * cos(D + F)
    P = P + 0.001 * E * sin(M + M1)
    P = P - 0.0009 * sin(2 * D - 2 * F)
    P = P + 0.0007 * cos(2 * M1 + F)
    P = P - 0.0007 * cos(3 * M1 + F)
    JDE = 2451562.5897 + 27.321582241 * k + 0.000100695 * T ^ 2 - 0.000000141 * T ^ 3 + P
    toymd = jd_to_ymd(JDE)
    Tahun = toymd[0]
    Bulan = toymd[1]
    Tanggal = toymd[2] + (toymd[3] + toymd[4]/60 + toymd[5]/3600)/24
    hasil = JDE - delta_t(Tahun, Bulan, Tanggal) / 86400
    return hasil


def jde_of_max_south_moon_declination(year, month, day):
    k = value_k(year, month, day)
    T = value_t(year, month, day)
    D = radians(345.6676 + 333.0705546 * k - 0.0004025 * T ^ 2 + 0.00000011 * T ^ 3)
    M = radians(1.3951 + 26.9281592 * k - 0.0000544 * T ^ 2 - 0.0000001 * T ^ 3)
    M1 = radians(186.21 + 356.9562795 * k + 0.0103126 * T ^ 2 + 0.00001251 * T ^ 3)
    F = radians(145.1633 + 1.4467806 * k - 0.0020708 * T ^ 2 - 0.00000215 * T ^ 3)
    E = nilai_e(year, month, day)
    
    P = -0.8975 * cos(F) - 0.4726 * sin(M1)
    P = P - 0.103 * sin(2 * F)
    P = P - 0.0976 * sin(2 * D - M1)
    P = P + 0.0541 * cos(M1 - F)
    P = P + 0.0516 * cos(M1 + F)
    P = P - 0.0438 * sin(2 * D)
    P = P + 0.0112 * sin(M) * E
    P = P + 0.0157 * cos(3 * F)
    P = P + 0.0023 * sin(M1 + 2 * F)
    P = P - 0.0136 * cos(2 * D - F)
    P = P + 0.011 * cos(2 * D - M1 - F)
    P = P + 0.0091 * cos(2 * D - M1 + F)
    P = P + 0.0089 * cos(2 * D + F)
    P = P + 0.0075 * sin(2 * M1)
    P = P - 0.003 * sin(M1 - 2 * F)
    P = P - 0.0061 * cos(2 * M1 - F)
    P = P - 0.0047 * sin(M1 + 3 * F)
    P = P - 0.0043 * sin(2 * D - M - M1) * E
    P = P + 0.004 * cos(M1 - 2 * F)
    P = P - 0.0037 * sin(2 * D - 2 * M1)
    P = P - 0.0031 * sin(F)
    P = P + 0.003 * sin(2 * D + M1)
    P = P + 0.0029 * cos(M1 + 2 * F)
    P = P - 0.0029 * sin(2 * D - M) * E
    P = P - 0.0027 * sin(M1 + F)
    P = P + 0.0024 * sin(M - M1) * E
    P = P - 0.0021 * sin(M1 - 3 * F)
    P = P - 0.0019 * sin(2 * M1 + F)
    P = P - 0.0006 * cos(2 * D - 2 * M1 - F)
    P = P - 0.0018 * sin(3 * F)
    P = P - 0.0017 * cos(M1 + 3 * F)
    P = P + 0.0017 * cos(2 * M1)
    P = P + 0.0014 * cos(2 * D - M1)
    P = P - 0.0013 * cos(2 * D + M1 + F)
    P = P - 0.0013 * cos(M1)
    P = P + 0.0012 * sin(3 * M1 + F)
    P = P + 0.0011 * sin(2 * D - M1 + F)
    P = P + 0.0011 * cos(2 * D - 2 * M1)
    P = P + 0.001 * cos(D + F)
    P = P + 0.001 * sin(M + M1) * E
    P = P - 0.0009 * sin(2 * D - 2 * F)
    P = P - 0.0007 * cos(2 * M1 + F)
    P = P - 0.0007 * cos(3 * M1 + F)

    JDE = 2451548.9289 + 27.321582241 * k + 0.000100695 * T**2 - 0.000000141 * T**3 + P
    toymd = jd_to_ymd(JDE)
    Tahun = toymd[0]
    Bulan = toymd[1]
    Tanggal = toymd[2] + (toymd[3] + toymd[4]/60 + toymd[5]/3600)/24
    hasil = JDE - delta_t(Tahun, Bulan, Tanggal) / 86400
    return hasil


def max_north_moon_declination(year, month, day):
    k = value_k(year, month, day)
    T = value_t(year, month, day)
    D = radians(152.2029 + 333.0705546 * k - 0.0004025 * T**2 + 0.00000011 * T**3)
    M = radians(14.8591 + 26.9281592 * k - 0.0000544 * T**2 - 0.0000001 * T**3)
    M1 = radians(4.6881 + 356.9562795 * k + 0.0103126 * T**2 + 0.00001251 * T**3)
    F = radians(325.8867 + 1.4467806 * k - 0.0020708 * T**2 - 0.00000215 * T**3)
    E = nilai_e(year, month, day)
    
    P = 5.1093 * sin(F)
    P = P + 0.2658 * cos(2 * F)
    P = P + 0.1448 * sin(2 * D - F)
    P = P - 0.0322 * sin(3 * F)
    P = P + 0.0133 * cos(2 * D - 2 * F)
    P = P + 0.0125 * cos(2 * D)
    P = P - 0.0124 * sin(M1 - F)
    P = P - 0.0101 * sin(M1 + 2 * F)
    P = P + 0.0097 * cos(F)
    P = P - 0.0087 * E * sin(2 * D + M - F)
    P = P + 0.0074 * sin(M1 + 3 * F)
    P = P + 0.0067 * sin(D + F)
    P = P + 0.0063 * sin(M1 - 2 * F)
    P = P + 0.006 * E * sin(2 * D - M - F)
    P = P - 0.0057 * sin(2 * D - M1 - F)
    P = P - 0.0056 * cos(M1 + F)
    P = P + 0.0052 * cos(M1 + 2 * F)
    P = P + 0.0041 * cos(2 * M1 + F)
    P = P - 0.004 * cos(M1 - 3 * F)
    P = P + 0.0038 * cos(2 * M1 - F)
    P = P - 0.0034 * cos(M1 - 2 * F)
    P = P - 0.0029 * sin(2 * M1)
    P = P + 0.0029 * sin(3 * M1 + F)
    P = P - 0.0028 * E * cos(2 * D + M - F)
    P = P - 0.0028 * cos(M1 - F)
    P = P - 0.0023 * cos(3 * F)
    P = P - 0.0021 * sin(2 * D + F)
    P = P + 0.0019 * cos(M1 + 3 * F)
    P = P + 0.0018 * cos(D + F)
    P = P + 0.0017 * sin(2 * M1 - F)
    P = P + 0.0015 * cos(3 * M1 + F)
    P = P + 0.0014 * cos(2 * D + 2 * M1 + F)
    P = P - 0.0012 * sin(2 * D - 2 * M1 - F)
    P = P - 0.0012 * cos(2 * M1)
    P = P - 0.001 * cos(M1)
    P = P - 0.001 * sin(2 * F)
    P = P + 0.0006 * sin(M1 + F)
    MaxNorthMoonDeclination = 23.6961 - 0.013004 * T + P
    return MaxNorthMoonDeclination


def max_south_moon_declination(year, month, day):
    k = value_k(year, month, day)
    T = value_t(year, month, day)
    D = radians(345.6676 + 333.0705546 * k - 0.0004025 * T**2 + 0.00000011 * T**3)
    M = radians(1.3951 + 26.9281592 * k - 0.0000544 * T**2 - 0.0000001 * T**3)
    M1 = radians(186.21 + 356.9562795 * k + 0.0103126 * T**2 + 0.00001251 * T**3)
    F = radians(145.1633 + 1.4467806 * k - 0.0020708 * T**2 - 0.00000215 * T**3)
    E = nilai_e(year, month, day)
    
    P = -5.1093 * sin(F)
    P = P + 0.2658 * cos(2 * F)
    P = P - 0.1448 * sin(2 * D - F)
    P = P + 0.0322 * sin(3 * F)
    P = P + 0.0133 * cos(2 * D - 2 * F)
    P = P + 0.0125 * cos(2 * D)
    P = P - 0.0015 * sin(M1 - F)
    P = P + 0.0101 * sin(M1 + 2 * F)
    P = P - 0.0097 * cos(F)
    P = P + 0.0087 * E * sin(2 * D + M - F)
    P = P + 0.0074 * sin(M1 + 3 * F)
    P = P + 0.0067 * sin(D + F)
    P = P - 0.0063 * sin(M1 - 2 * F)
    P = P - 0.006 * E * sin(2 * D - M - F)
    P = P + 0.0057 * sin(2 * D - M1 - F)
    P = P - 0.0056 * cos(M1 + F)
    P = P - 0.0052 * cos(M1 + 2 * F)
    P = P - 0.0041 * cos(2 * M1 + F)
    P = P - 0.004 * cos(M1 - 3 * F)
    P = P - 0.0038 * cos(2 * M1 - F)
    P = P + 0.0034 * cos(M1 - 2 * F)
    P = P - 0.0029 * sin(2 * M1)
    P = P + 0.0029 * sin(3 * M1 + F)
    P = P + 0.0028 * E * cos(2 * D + M - F)
    P = P - 0.0028 * cos(M1 - F)
    P = P + 0.0023 * cos(3 * F)
    P = P + 0.0021 * sin(2 * D + F)
    P = P + 0.0019 * cos(M1 + 3 * F)
    P = P + 0.0018 * cos(D + F)
    P = P - 0.0017 * sin(2 * M1 - F)
    P = P + 0.0015 * cos(3 * M1 + F)
    P = P + 0.0014 * cos(2 * D + 2 * M1 + F)
    P = P + 0.0012 * sin(2 * D - 2 * M1 - F)
    P = P - 0.0012 * cos(2 * M1)
    P = P + 0.001 * cos(M1)
    P = P - 0.001 * sin(2 * F)
    P = P + 0.0037 * sin(M1 + F)
    MaxSouthMoonDeclination = 23.6961 - 0.013004 * T + P 
    return MaxSouthMoonDeclination


def moon_argument_of_latitude(year, month, day):
    t = julian_ephemeris_century(year, month, day)
    return 93.2720993 + 483202.0175273*t - 0.0034029*t**2 - t**3/3526000 + t**4/863310000


def sigma_b(year, month, day): 
    L1 = radians(moon_mean_longitude(year, month, day))
    D = radians(moon_mean_elongation(year, month, day))
    M = radians(sun_mean_anomaly(year, month, day))
    M1 = radians(moon_mean_anomaly(year, month, day))
    F = radians(moon_argument_of_latitude(year, month, day))
    A1 = radians(argument_a1(year, month, day))
    A2 = radians(argument_a2(year, month, day))
    A3 = radians(argument_a3(year, month, day))
    E = nilai_e(year, month, day)
   
    sB = 0#
    sB = sB + 5128122 * sin(F)
    sB = sB + 280602 * sin(M1 + F)
    sB = sB + 277693 * sin(M1 - F)
    sB = sB + 173237 * sin(2 * D - F)
    sB = sB + 55413 * sin(2 * D - M1 + F)
    sB = sB + 46271 * sin(2 * D - M1 - F)
    sB = sB + 32573 * sin(2 * D + F)
    sB = sB + 17198 * sin(2 * M1 + F)
    sB = sB + 9266 * sin(2 * D + M1 - F)
    sB = sB + 8822 * sin(2 * M1 - F)
    sB = sB + 8216 * E * sin(2 * D - M - F)
    sB = sB + 4324 * sin(2 * D - 2 * M1 - F)
    sB = sB + 4200 * sin(2 * D + M1 + F)
    sB = sB - 3359 * E * sin(2 * D + M - F)
    sB = sB + 2463 * E * sin(2 * D - M - M1 + F)
    sB = sB + 2211 * E * sin(2 * D - M + F)
    sB = sB + 2065 * E * sin(2 * D - M - M1 - F)
    sB = sB - 1870 * E * sin(M - M1 - F)
    sB = sB + 1828 * sin(4 * D - M1 - F)
    sB = sB - 1794 * E * sin(M + F)
    sB = sB - 1749 * sin(3 * F)
    sB = sB - 1565 * E * sin(M - M1 + F)
    sB = sB - 1491 * sin(D + F)
    sB = sB - 1475 * E * sin(M + M1 + F)
    sB = sB - 1410 * E * sin(M + M1 - F)
    sB = sB - 1344 * E * sin(M - F)
    sB = sB - 1335 * sin(D - F)
    sB = sB + 1107 * sin(3 * M1 + F)
    sB = sB + 1021 * sin(4 * D - F)
    sB = sB + 833 * sin(4 * D - M1 + F)
    sB = sB + 777 * sin(M1 - 3 * F)
    sB = sB + 671 * sin(4 * D - 2 * M1 + F)
    sB = sB + 607 * sin(2 * D - 3 * F)
    sB = sB + 596 * sin(2 * D + 2 * M1 - F)
    sB = sB + 491 * E * sin(2 * D - M + M1 - F)
    sB = sB - 451 * sin(2 * D - 2 * M1 + F)
    sB = sB + 439 * sin(3 * M1 - F)
    sB = sB + 422 * sin(2 * D + 2 * M1 + F)
    sB = sB + 421 * sin(2 * D - 3 * M1 - F)
    sB = sB - 366 * E * sin(2 * D + M - M1 + F)
    sB = sB - 351 * E * sin(2 * D + M + F)
    sB = sB + 331 * sin(4 * D + F)
    sB = sB + 315 * E * sin(2 * D - M + M1 + F)
    sB = sB + 302 * E ^ 2 * sin(2 * D - 2 * M - F)
    sB = sB - 283 * sin(M1 + 3 * F)
    sB = sB - 229 * E * sin(2 * D + M + M1 - F)
    sB = sB + 223 * E * sin(D + M - F)
    sB = sB + 223 * E * sin(D + M + F)
    sB = sB - 220 * E * sin(M - 2 * M1 - F)
    sB = sB - 220 * E * sin(2 * D + M - M1 - F)
    sB = sB - 185 * sin(D + M1 + F)
    sB = sB + 181 * E * sin(2 * D - M - 2 * M1 - F)
    sB = sB - 177 * E * sin(M + 2 * M1 + F)
    sB = sB + 176 * sin(4 * D - 2 * M1 - F)
    sB = sB + 166 * E * sin(4 * D - M - M1 - F)
    sB = sB - 164 * sin(D + M1 - F)
    sB = sB + 132 * sin(4 * D + M1 - F)
    sB = sB - 119 * sin(D - M1 - F)
    sB = sB + 115 * E * sin(4 * D - M - F)
    sB = sB + 107 * E ^ 2 * sin(2 * D - 2 * M + F)
    sB = sB - 2235 * sin(L1) + 382 * sin(A3) + 175 * sin(A1 - F) + 175 * sin(A1 + F) + 127 * sin(L1 - M1) - 115 * sin(L1 + M1)
    SigmaB = sB
    return SigmaB


def sigma_l(year, month, day): 
    L1 = radians(moon_mean_longitude(year, month, day))
    D = radians(moon_mean_elongation(year, month, day))
    M = radians(sun_mean_anomaly(year, month, day))
    M1 = radians(moon_mean_anomaly(year, month, day))
    F = radians(moon_argument_of_latitude(year, month, day))
    A1 = radians(argument_a1(year, month, day))
    A2 = radians(argument_a2(year, month, day))
    A3 = radians(argument_a3(year, month, day))
    E = nilai_e(year, month, day)

    
    sL = 0#
    sL = sL + 6288774 * sin(M1)
    sL = sL + 1274027 * sin(2 * D - M1)
    sL = sL + 658314 * sin(2 * D)
    sL = sL + 213618 * sin(2 * M1)
    sL = sL - 185116 * E * sin(M)
    sL = sL - 114332 * sin(2 * F)
    sL = sL + 58793 * sin(2 * D - 2 * M1)
    sL = sL + 57066 * E * sin(2 * D - M - M1)
    sL = sL + 53322 * sin(2 * D + M1)
    sL = sL + 45758 * E * sin(2 * D - M)
    sL = sL - 40923 * E * sin(M - M1)
    sL = sL - 34720 * sin(D)
    sL = sL - 30383 * E * sin(M + M1)
    sL = sL + 15327 * sin(2 * D - 2 * F)
    sL = sL - 12528 * sin(M1 + 2 * F)
    sL = sL + 10980 * sin(M1 - 2 * F)
    sL = sL + 10675 * sin(4 * D - M1)
    sL = sL + 10034 * sin(3 * M1)
    sL = sL + 8548 * sin(4 * D - 2 * M1)
    sL = sL - 7888 * E * sin(2 * D + M - M1)
    sL = sL - 6766 * E * sin(2 * D + M)
    sL = sL - 5163 * sin(D - M1)
    sL = sL + 4987 * E * sin(D + M)
    sL = sL + 4036 * E * sin(2 * D - M + M1)
    sL = sL + 3994 * E * sin(2 * D + 2 * M1)
    sL = sL + 3861 * sin(4 * D)
    sL = sL + 3665 * sin(2 * D - 3 * M1)
    sL = sL - 2689 * E * sin(M - 2 * M1)
    sL = sL - 2602 * sin(2 * D - M1 + 2 * F)
    sL = sL + 2390 * E * sin(2 * D - M - 2 * M1)
    sL = sL - 2348 * sin(D + M1)
    sL = sL + 2236 * E ^ 2 * sin(2 * D - 2 * M)
    sL = sL - 2120 * E * sin(M + 2 * M1)
    sL = sL - 2069 * E ^ 2 * sin(2 * M)
    sL = sL + 2048 * E ^ 2 * sin(2 * D - 2 * M - M1)
    sL = sL - 1773 * sin(2 * D + M1 - 2 * F)
    sL = sL - 1595 * sin(2 * D + 2 * F)
    sL = sL + 1215 * E * sin(4 * D - M - M1)
    sL = sL - 1110 * sin(2 * M1 + 2 * F)
    sL = sL - 892 * sin(3 * D - M1)
    sL = sL - 810 * E * sin(2 * D + M + M1)
    sL = sL + 759 * E * sin(4 * D - M - 2 * M1)
    sL = sL - 713 * E ^ 2 * sin(2 * M - M1)
    sL = sL - 700 * E ^ 2 * sin(2 * D + 2 * M - M1)
    sL = sL + 691 * E * sin(2 * D + M - 2 * M1)
    sL = sL + 596 * E * sin(2 * D - M - 2 * F)
    sL = sL + 549 * sin(4 * D + M1)
    sL = sL + 537 * sin(4 * M1)
    sL = sL + 520 * E * sin(4 * D - M)
    sL = sL - 487 * sin(D - 2 * M1)
    sL = sL - 399 * E * sin(2 * D + M - 2 * F)
    sL = sL - 381 * sin(2 * M1 - 2 * F)
    sL = sL + 351 * E * sin(D + M + M1)
    sL = sL - 340 * sin(3 * D - 2 * M1)
    sL = sL + 330 * sin(4 * D - 3 * M1)
    sL = sL + 327 * E * sin(2 * D - M + 2 * M1)
    sL = sL - 323 * E ^ 2 * sin(2 * M + M1)
    sL = sL + 299 * E * sin(D + M - M1)
    sL = sL + 294 * sin(2 * D + 3 * M1)
    SigmaL = sL + 3958 * sin(A1) + 1962 * sin(L1 - F) + 318 * sin(A2)
    return SigmaL


def sigma_r(year, month, day): 
    L1 = radians(moon_mean_longitude(year, month, day))
    D = radians(moon_mean_elongation(year, month, day))
    M = radians(sun_mean_anomaly(year, month, day))
    M1 = radians(moon_mean_anomaly(year, month, day))
    F = radians(moon_argument_of_latitude(year, month, day))
    A1 = radians(argument_a1(year, month, day))
    A2 = radians(argument_a2(year, month, day))
    A3 = radians(argument_a3(year, month, day))
    E = nilai_e(year, month, day)
    
    mD = 0#
    mD = mD - 20905355 * cos(M1)
    mD = mD - 3699111 * cos(2 * D - M1)
    mD = mD - 2955968 * cos(2 * D)
    mD = mD - 569925 * cos(2 * M1)
    mD = mD + 48888 * E * cos(M)
    mD = mD - 3149 * cos(2 * F)
    mD = mD + 246158 * cos(2 * D - 2 * M1)
    mD = mD - 152138 * E * cos(2 * D - M - M1)
    mD = mD - 170733 * cos(2 * D + M1)
    mD = mD - 204586 * E * cos(2 * D - M)
    mD = mD - 129620 * E * cos(M - M1)
    mD = mD + 108743 * cos(D)
    mD = mD + 104755 * E * cos(M + M1)
    mD = mD + 10321 * cos(2 * D - 2 * F)
    mD = mD + 79661 * cos(M1 - 2 * F)
    mD = mD - 34782 * cos(4 * D - M1)
    mD = mD - 23210 * cos(3 * M1)
    mD = mD - 21636 * cos(4 * D - 2 * M1)
    mD = mD + 24208 * E * cos(2 * D + M - M1)
    mD = mD + 30824 * E * cos(2 * D + M)
    mD = mD - 8379 * cos(D - M1)
    mD = mD - 16675 * E * cos(D + M)
    mD = mD - 12831 * E * cos(2 * D - M + M1)
    mD = mD - 10445 * cos(2 * D + 2 * M1)
    mD = mD - 11650 * cos(4 * D)
    mD = mD + 14403 * cos(2 * D - 3 * M1)
    mD = mD - 7003 * E * cos(M - 2 * M1)
    mD = mD + 10056 * E * cos(2 * D - M - 2 * M1)
    mD = mD + 6322 * cos(D + M1)
    mD = mD - 9884 * E ^ 2 * cos(2 * D - 2 * M)
    mD = mD + 5751 * E * cos(M + 2 * M1)
    mD = mD - 4950 * E ^ 2 * cos(2 * D - 2 * M - M1)
    mD = mD + 4130 * cos(2 * D + M1 - 2 * F)
    mD = mD - 3958 * E * cos(4 * D - M - M1)
    mD = mD + 3258 * cos(3 * D - M1)
    mD = mD + 2616 * E * cos(2 * D + M + M1)
    mD = mD - 1897 * E * cos(4 * D - M - 2 * M1)
    mD = mD - 2117 * E ^ 2 * cos(2 * M - M1)
    mD = mD + 2354 * E ^ 2 * cos(2 * D + 2 * M - M1)
    mD = mD - 1423 * cos(4 * D + M1)
    mD = mD - 1117 * cos(4 * M1)
    mD = mD - 1571 * E * cos(4 * D - M)
    mD = mD - 1739 * cos(D - 2 * M1)
    mD = mD - 4421 * cos(2 * M1 - 2 * F)
    mD = mD + 1165 * E ^ 2 * cos(2 * M + M1)
    mD = mD + 8752 * cos(2 * D - M1 - 2 * F)
    SigmaR = mD
    return SigmaR


def hour_angle_of_the_moon(bujur, year, month, day): 
    hour_angle_of_the_moon = (apparent_local_sidereal_time(bujur, year, month, day) - apparent_moon_right_ascension(year, month, day)) % 24
    return hour_angle_of_the_moon


def moon_longitude_elp2000(year, month, day): 
    jd = julian_ephemeris_day(year, month, day)
    T = julian_ephemeris_century(year, month, day)
    


def apparent_moon_longitude_elp2000(): 
    pass


def moon_latitude_elp2000(): 
    pass


def apparent_moon_latitude_elp2000(): 
    pass


def distance_to_the_moon_elp2000(): 
    pass


def ijtimak_elp2000_vsop87(): 
    pass


def konjungsi_elp2000_vsop87(): 
    pass


def new_moon_elp2000_vsop87(): 
    pass


def first_quarter_elp2000_vsop87(): 
    pass


def full_moon_elp2000_vsop87(): 
    pass


def last_quarter_elp2000_vsop87(): 
    pass


def new_moon_to_new_moon_elp2000_vsop87(): 
    pass


def d_latitude(lat_d, lat_m, lat_s, lat_dir): 
    pass


def sun_aberration(year, month, day):
    tau = julian_ephemeris_century(year, month, day)/10
    dL = 3548.193
    dL += 118.568*sin(radians(87.5287 + 359993.7286*tau))
    dL += 2.476*sin(radians(85.0561 + 719987.4571*tau))
    dL += 1.376*sin(radians(27.8502 + 4452671.1152*tau))
    dL += 0.119*sin(radians(73.1375 + 450368.8564*tau))
    dL += 0.114*sin(radians(337.2264 + 329644.6718*tau))
    dL += 0.086*sin(radians(222.54 + 659289.3436*tau))
    dL += 0.078*sin(radians(162.8136 + 9224659.7915*tau))
    dL += 0.054*sin(radians(82.5823 + 1079981.1857*tau))
    dL += 0.052*sin(radians(171.5189 + 225184.4282*tau))
    dL += 0.034*sin(radians(30.3214 + 4092677.3866*tau))
    dL += 0.033*sin(radians(119.8105 + 337181.4711*tau))
    dL += 0.023*sin(radians(247.5418 + 299295.6151*tau))
    dL += 0.023*sin(radians(325.1526 + 315559.556*tau))
    dL += 0.021*sin(radians(155.1241 + 675553.2846*tau))
    dL += 7.311*sin(radians(333.4515 + 359993.7286*tau))
    dL += 0.305*sin(radians(330.9814 + 719987.4571*tau))
    dL += 0.01*sin(radians(328.517 + 1079981.1857*tau))
    dL += 0.309*sin(radians(241.4518 + 359993.7286*tau))
    dL += 0.021*sin(radians(205.0482 + 719987.4571*tau))
    dL += 0.004*sin(radians(297.861 + 4452671.1152*tau))
    dL += 0.01*sin(radians(154.7066 + 359993.7286*tau))

    R = distance_to_the_sun_vsop87(year, month, day)
    return -0.005775518*R*dL


def nutation_in_longitude(year, month, day):
    '''
    Delta Psi

    '''
    t = julian_ephemeris_century(year, month, day)
    d = radians(moon_mean_elongation(year, month, day))
    m = radians(sun_mean_anomaly(year, month, day))
    m1 = radians(moon_mean_anomaly(year, month, day))
    f = radians(moon_argument_of_latitude(year, month, day))
    omega = radians(nilai_omega(year, month, day))
    
    # TERM delta PSI
    dpsi = terms_delta_psi
    total = 0
    for i in range(len(dpsi)):
        total += (dpsi[i][5] + dpsi[i][6]*t)*sin(dpsi[i][0]*d + dpsi[i][1]*m + dpsi[i][2]*m1 + dpsi[i][3]*f + dpsi[i][4]*omega)
    return total/(10000*3600)


def nutation_in_obliquity(year, month, day):
    '''
    Delta Epsilon
    '''
    t = julian_ephemeris_century(year, month, day)
    d = radians(moon_mean_elongation(year, month, day))
    m = radians(sun_mean_anomaly(year, month, day))
    m1 = radians(moon_mean_anomaly(year, month, day))
    f = radians(moon_argument_of_latitude(year, month, day))
    omega = radians(nilai_omega(year, month, day))

    depsilon = terms_delta_epsilon
    total = 0
    for i in range(len(depsilon)):
        total += (depsilon[i][5] + depsilon[i][6]*t)*cos(depsilon[i][0]*d + depsilon[i][1]*m + depsilon[i][2]*m1 + depsilon[i][3]*f + depsilon[i][4]*omega)
    return total/(10000*3600)


def horizontal_moon_parallax(): ...


def refraction_apparent_altitude(h0, p, T):
    R = 1/tan(radians(h0 + 7.31/(h0 + 4.4))) + 0.0013515
    dR1 = -0.06*sin(radians(14.7*R/60 + 13))
    dR2 = (p/1010)*(283/(273 + T))
    return (R + dR1/60)*dR2


def refraction_true_altitude(h, p, T):
    R = 1/tan(radians(h + 10.3/(h + 5.11))) + 0.0019279
    dR = (p/1010)*(283/(273 + T))
    return R * dR

def kwd (timezone, bujur): 
    longitude = dms_to_dec(bujur)
    return (timezone*15 - longitude)/15


def koreksi_dip(h):
    D = acos(6371008/(6371008 + h))
    return degrees(D)*60

def koreksi_new_moon(): 
    pass


def koreksi_first_quarter(): 
    pass


def koreksi_full_moon(): 
    pass


def koreksi_last_quarter(): 
    pass


def koreksi_w(): 
    pass


def koreksi_planetary_arguments(): 
    pass


def mean_obliquity_of_ecliptic(year, month, day):
    '''
    Epsilon Zero
    '''
    t = julian_ephemeris_century(year, month, day)
    return 23 + 26/60 + (21.448 - 46.815*t - 0.00059*t**2 + 0.001813*t**3)/3600


def obliquity_of_ecliptic(year, month, day):
    '''
    Epsilon
    '''
    return mean_obliquity_of_ecliptic(year, month, day) + nutation_in_obliquity(year, month, day)


def nilai_k(): ...


def value_k(): ...


def nilai_t(): ...


def value_t(): ...


def nilai_e(): ...


def value_e(): ...


def nilai_m(): ...


def nilai_m1(): ...


def nilai_f(): ...


def nilai_o(): ...


def nilai_omega(year, month, day):
    t = julian_ephemeris_century(year, month, day)
    return 125.04452 - 1934.136261*t + 0.0020708*t**2 + t**3 / 450000


def nilai_jde(): ...


def argument_a1(): ...


def argument_a2(): ...


def argument_a3(): ...

