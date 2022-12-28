def to_dec(d: int, m: int = 0, s:float = 0):
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


def to_dms(dec):
    '''
    fungsi untuk mengubah Desimal menjadi DMS
    '''
    _d = int(dec)
    _m = int((dec - _d)*60)
    _s = ((dec - _d)*60 - _m)*60
    return f"{_d}° {abs(_m)}' {round(abs(_s),2)}\""


def cek_kabisat(thn):
    if thn % 4 == 0:
        if thn % 100 == 0:
            if thn % 400 == 0:
                return True
            return False
        return True
    return False



# fungsi year_decimal
def year_dec(year, month, day):
    '''
    fungsi untuk mengubah tahun, bulan, tanggal menjadi bentuk desimal
    dalam modul ini diperlukan untuk menemukan Delta T
    '''
    return year + month/12 + day/365


def JD_ut(year:int, month:int, day:int, hour:int, minute:int, second:int, timezone:float)-> float:
    """
    fungsi untuk mengubah waktu menjadi Julian Day
    """
    m = month + 12 if month < 3 else month
    y = year - 1 if month < 3 else year
    a = int(y/100)
    if year < 1583:
        if month < 11:
            if day < 4:
                b = 0
                if day > 14:
                    b = 2 + int(a/4)-a
                else:
                    print("tanggal salah")
        else:
            b = 2 + int(a/4) - a
    else:
        b   = 2 + int(a/4) - a
    
    return 1720994.5 + int(365.25 * y) + int(30.60001 * (m + 1)) + b + day + (hour + minute/60 + second/3600)/24 - timezone/24


    
