def to_dec(d, m=0, s=0):
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