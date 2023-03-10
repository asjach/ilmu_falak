# Belum menggunakan Method/Function Overloading

def print_dec(d: int, m: int = 0, s:float = 0):
    '''
    Fungsi untuk merubah bentuk DMS menjadi desimal  
    '''
    if d >= 0:
        _dec = d + abs(m)/60 + abs(s)/3600
        if _dec == int(_dec):
            _dec = int(d + abs(m)/60 + abs(s)/3600)
    else:
        _dec = d - abs(m)/60 - abs(s)/3600
        if _dec == int(_dec):
            _dec = int(d - abs(m)/60 - abs(s)/3600)
    return _dec


def print_dms(decimal_value):
    '''
    fungsi untuk mengubah Desimal menjadi DMS
    '''
    _d = int(decimal_value)
    _m = int((decimal_value - _d)*60)
    _s = ((decimal_value - _d)*60 - _m)*60
    return f"{_d}° {abs(_m)}' {round(abs(_s),2)}\""


def print_hms(decimal_value):
    '''
    fungsi untuk mengubah Desimal menjadi H:M:S
    '''
    _d = int(decimal_value)
    _m = int((decimal_value - _d)*60)
    _s = ((decimal_value - _d)*60 - _m)*60
    return f"{_d}:{abs(_m)}:{round(abs(_s),2)}"


def cek_kabisat(thn):
    if thn % 4 == 0:
        if thn % 100 == 0:
            if thn % 400 == 0:
                return True
            return False
        return True
    return False


def dms_to_dec(input):
    plus = ["LU", "N", "BT", "E", 'lu', 'n', 'bt', 'e']
    minus = ['LS', 'S', 'BB', 'W', 'ls', 's', 'bb', 'w']
    d = input[0]
    m = input[1]
    s = input[2]
    dir = input[3]
    hasil = d + m/60 + s/3600
    if dir in plus:
        return hasil
    elif dir in minus:
        return -hasil
    return plus


def jd(*input)-> float:
    """
    Julian Day (JD) dalam Universal Time (UT)

    Input:
    - [3] year, month, day (Y, M, D 0) UT
    - [4] year, month, day, hour ( Y, M, D, h) UT
    - [5] year, month, day, hour, timezone ( Y, M, D, h, timezone) Local Time
    - [6] year, month, day, hour, minute, second (Y, M, D, h, m, s) UT
    - [7] year, month, day, hour, minute, second, timezone (Y, M, D, h, m, s, timezone) Local Time
    """
    hour = minute = second = timezone = 0
    if len(input) == 3:
        year = input[0]; month = input[1]; day = input[2]
    elif len(input) == 4:
        year = input[0]; month = input[1]; day = input[2]
        hour = input[3]        
    elif len(input) == 5:
        year = input[0]; month = input[1]; day = input[2]
        hour = input[3]
        timezone = input[4]
    elif len(input) == 6:
        year = input[0]; month = input[1]; day = input[2]
        hour = input[3]; minute = input[4]; second = input[5]
    elif len(input) == 7:
        year = input[0]; month = input[1]; day = input[2]
        hour = input[3]; minute = input[4]; second = input[5]
        timezone = input[6]
    else:
        return f'wrong input'


    m = month + 12 if month < 3 else month
    y = year - 1 if month < 3 else year
    a = int(y/100)
    if year < 1583:
        if year == 1982:
            if month == 10:
                if day > 4 and day < 15:
                    print("Tidak ada tanggal 5-14 Oktober 1582 dalam sejarah")
                elif day > 14:
                    b = 2 + int(a/4)-a
            elif month > 10:
                b = 2 + int(a/4)-a
            else:
                b = 0 
        b = 0               
    else:
        b = 2 + int(a/4) - a
    JD = 1720994.5 + int(365.25 * y) + int(30.60001 * (m + 1)) + b + day + (hour + minute/60 + second/3600)/24 - timezone/24
    return JD


def delta_t(*input) -> float:
    """ Delta T

    Parameter: 
    - Year in decimal (1 input)
    - Year in decimal, D | S (1 input, output dalam Day atau dalan Second)
    - Year, Month, D | S (2 input, output D atau S)
    - Year, Month, Day (3 input, output otomatis D)
    - Year, Month, Day, D | S (3 input, output D | S)

    Pilihan output antara D (day) atau S (second), jika tidak diinput/diinput selain nilai tersebut maka akan mengembalikan 'D'
    """

    if len(input) == 3:
        if type(input[-1]) == str:
            _year = input[0]
            _month = input[1]
            year_decimal = _year + (_month-1)/12
        else:
            _year = input[0]
            _month = input[1]
            _day = input[2]
            year_decimal = _year + (_month-1)/12 + _day/365
    elif len(input) == 4:
        _year = input[0]
        _month = input[1]
        _day = input[2]
        year_decimal = _year + (_month-1)/12 + _day/365
    else:
        year_decimal = input[0]

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

    if len(input)> 1 and type(input[-1]) == str:
        if input[-1] == 'S' or input[-1] == 's':      
            return delta
    return delta/86400    

def jde(*input):
    '''
    Julian Ephemeris Day

    Parameter:
    - [1] jd
    - [2] Y, M, D (UT)
    - [4] Y, M, D, h (UT)
    - [5] Y, M, D, h, timezone (LT)
    - [6] Y, M, D, H, M, S (UT)
    - [7] Y, M, D, h, m, s, timezone (LT)
    '''
    if len(input) == 1:
        julian_day = input[0]
        deltaT = 0 # harus buat dulu fungsi konversi jd ke YMDhms

    elif len(input) == 3:
        year = input[0]; month = input[1]; day = input[2]
        julian_day = jd(year, month, day)  # UT
        deltaT = delta_t(year, month, day)

    elif len(input) == 4:
        year = input[0]; month = input[1]; day = input[2]
        hour = input[3]
        julian_day = jd(year, month, day, hour)  # UT
        deltaT = delta_t(year, month, day) 

    elif len(input) == 5:
        year = input[0]; month = input[1]; day = input[2]
        hour = input[3]; timezone = input[4]
        julian_day = jd(year, month, day, hour, timezone)  # LT
        deltaT = delta_t(year, month, day)

    elif len(input) == 6:
        year = input[0]; month = input[1]; day = input[2]
        hour = input[3]; minute = input[4]; second = input[5]
        julian_day = jd(year, month, day, hour, minute, second)  # UT
        deltaT = delta_t(year, month, day)

    elif len(input) == 7:
        year = input[0]; month = input[1]; day = input[2]
        hour = input[3]; minute = input[4]; second = input[5]; timezone = input[6]
        julian_day = jd(year, month, day, hour, minute, second, timezone)  # LT
        deltaT = delta_t(year, month, day)

    else:
        return f'wrong input'

    return julian_day + deltaT


def t_ut(*input):
    '''
    Julian Century dalam Universal Time belum dikoreksi oleh delta T
    digunakan untuk menentukan Greenwich Sidereal Time (GST) dan Local Sidereal Time (LST).
    GST dan LST digunakan untuk konversi koordinat Equatorial menjadi koordinat Horizon

    Parameter:
    - [1] Julian Day (1 input)
    - [3] Year, month, day (3 input)
    - [7] Year, month, day, hour, minute, second, timezone (7 input)
    '''
    if len(input) == 1:
        _jd = input[0]
    elif len(input) >= 3 and len(input) <= 7:
        _jd = jd(*input)
    elif len(input) > 7:
        return f'wrong input'
    return (_jd - 2451545)/36525

def t_td(*input):
    '''
    Julian Ephemeris Century dalam UT sudah dikoreksi dengan Delta T
    digunakan dalam perhitungan koreksi Nutasi, dan koreksi bulan
    '''
    _jde = jde(*input)
    return (_jde - 2451545)/36525

def tau(*input):
    '''
    Julian Ephemeris Millenia dalam UT sudah dikoreksi dengan Delta T
    digunakan dalam perhitungan koreksi Matahari
    '''
    _jde = jde(*input)
    return (_jde - 2451545)/365250



