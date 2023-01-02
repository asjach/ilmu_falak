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


def dms_to_dec(d=None, m=None, s=None, dir=None):
    plus = ["LU", "N", "BT", "E", 'lu', 'n', 'bt', 'e']
    minus = ['LS', 'S', 'BB', 'W', 'ls', 's', 'bb', 'w']

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
    - year, month, day (Y, M, D 0) UT
    - year, month, day, hour ( Y, M, D, h) UT
    - year, month, day, hour, timezone ( Y, M, D, h, timezone) Local Time
    - year, month, day, hour, minute, second (Y, M, D, h, m, s) UT
    - year, month, day, hour, minute, second, timezone (Y, M, D, h, m, s, timezone) Local Time
    """
    hour = minute = second = timezone = 0
    if len(input) == 3:
        year = input[0]
        month = input[1]
        day = input[2]
    elif len(input) == 4:
        year = input[0]
        month = input[1]
        day = input[2]
        hour = input[3]        
    elif len(input) == 5:
        year = input[0]
        month = input[1]
        day = input[2]
        hour = input[3]
        timezone = input[4]
    elif len(input) == 6:
        year = input[0]
        month = input[1]
        day = input[2]
        hour = input[3]
        minute = input[4]
        second = input[5]
    elif len(input) == 7:
        year = input[0]
        month = input[1]
        day = input[2]
        hour = input[3]
        minute = input[4]
        second = input[5]
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
        delta = -20 + 32 * (year_decimal/100 - 18.2) * (year_decimal/100 - 18.2)
    if year_decimal > 500 and year_decimal <= 1600:
        delta =  1574.2 - 556.01*(year_decimal/100-10) + 71.23472*(year_decimal/100-10)*(year_decimal/100-10) 
        delta += 0.319781*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10) 
        delta -= 0.8503463*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10) 
        delta -= 0.005050998*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10) 
        delta += 0.0083572073*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)

    if year_decimal > 1600 and year_decimal <= 1700:
        delta = 120 - 0.9808*(year_decimal-1600) - 0.01532*(year_decimal-1600)*(year_decimal-1600) 
        delta += (year_decimal-1600)*(year_decimal-1600)*(year_decimal-1600)/7129

    if year_decimal > 1700 and year_decimal <= 1800:
        delta = 8.83 + 0.1603*(year_decimal-1700) - 0.0059285*(year_decimal-1700)*(year_decimal-1700) 
        delta += 0.00013336*(year_decimal-1700)*(year_decimal-1700)*(year_decimal-1700) 
        delta -= (year_decimal-1700)*(year_decimal-1700)*(year_decimal-1700)*(year_decimal-1700)/1174000

    if year_decimal > 1800 and year_decimal <= 1860:
        delta = 13.72 - 0.332447*(year_decimal-1800) + 0.0068612*(year_decimal-1800)*(year_decimal-1800) 
        delta += 0.0041116*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800) 
        delta -= 0.00037436*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800) 
        delta += 0.0000121272*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800) 
        delta -= 0.0000001699*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800) 
        delta += 0.000000000875*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)

    if year_decimal > 1860 and year_decimal <= 1900:
        delta = 7.62 + 0.5737*(year_decimal-1860) - 0.251754*(year_decimal-1860)*(year_decimal-1860) 
        delta += 0.01680668*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860) 
        delta -= 0.0004473624*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860) 
        delta += (year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860)/233174

    if year_decimal > 1900 and year_decimal <= 1920:
        delta = -2.79 + 1.494119*(year_decimal-1900) - 0.0598939*(year_decimal-1900)*(year_decimal-1900) 
        delta += 0.0061966*(year_decimal-1900)*(year_decimal-1900)*(year_decimal-1900) 
        delta -= 0.000197*(year_decimal-1900)*(year_decimal-1900)*(year_decimal-1900)*(year_decimal-1900)

    if year_decimal > 1920 and year_decimal <= 1941:
        delta = 21.2 + 0.84493*(year_decimal-1920) 
        delta -= 0.0761*(year_decimal-1920)*(year_decimal-1920) 
        delta += 0.0020936*(year_decimal-1920)*(year_decimal-1920)*(year_decimal-1920)

    if year_decimal > 1941 and year_decimal <= 1961:
        delta = 29.07 + 0.407*(year_decimal-1950) 
        delta -= (year_decimal-1950)*(year_decimal-1950)/233 
        delta += (year_decimal-1950)*(year_decimal-1950)*(year_decimal-1950)/2547

    if year_decimal > 1961 and year_decimal <= 1986:
        delta = 45.45 + 1.067*(year_decimal-1975) 
        delta -= (year_decimal-1975)*(year_decimal-1975)/260 
        delta -= (year_decimal-1975)*(year_decimal-1975)*(year_decimal-1975)/718

    if year_decimal > 1986 and year_decimal <= 2005:
        delta = 63.86 + 0.3345*(year_decimal-2000) 
        delta -= 0.060374*(year_decimal-2000)*(year_decimal-2000) 
        delta += 0.0017275*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000) 
        delta += 0.000651814*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000) 
        delta += 0.00002373599*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000)

    if year_decimal > 2005 and year_decimal <= 2050:
        delta = 62.92 + 0.32217*(year_decimal-2000) 
        delta += 0.005589*(year_decimal-2000)*(year_decimal-2000)

    if year_decimal > 2050 and year_decimal <= 2150:
        delta = -20 + 32*((year_decimal - 1820)/100)*((year_decimal - 1820)/100) - 0.5628*(2150 - year_decimal)

    if year_decimal > 2150:
        delta = -20 + 32*((year_decimal - 1820)/100)*((year_decimal - 1820)/100)


    if len(input)> 1 and type(input[-1]) == str:
        if input[-1] == 'S' or input[-1] == 's':      
            return delta
    return delta/86400    


def jde(*input):
    '''
    Julian Ephemeris Day

    Parameter:
    - jd
    - Y, M, D
    - Y, M, D, Timezone
    - Y, M, D, H, M, S
    - Y, M, D, h, m, s, timezone

    '''
    if len(input) == 1:
        ...
    elif len(input) == 3:
        ...
    elif len(input) == 4:
        ...
    elif len(input) == 5:
        ...
    elif 

    jd = jd(*input)
    deltaT = delta_t(input[0], input[1], input[2])

    return jd + deltaT




def t_ut(*input):
    '''
    Julian Century dalam Universal Time

    Parameter:
    - Julian Day (1 input)
    - Year, month, day (3 input)
    - Year, month, day, hour, minute, second, timezone (7 input)
    '''
    if len(input) == 1:
        jd = input[0]
        return (jd - 2451545)/36525
    elif len(input) >= 3 and len(input) <= 7:
        jd = jd(*input)
        return (jd - 2451545)/36525
    elif len(input) > 7:
        return f'wrong input'


def t_td(*input):
    ...


def tau(*input):
    ...