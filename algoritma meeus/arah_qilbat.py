from math import sin, cos, tan, acos, atan, radians, degrees

from ephemeris import Ephemeris
from functions import to_dms


class Lokasi():
    def __init__(self, latD, latM, latS, latDir, lonD, lonM, lonS, lonDir):
        self.latD = latD
        self.latM = latM
        self.latS = latS
        self.lonD = lonD
        self.lonM = lonM
        self.lonS = lonS
        self.latDir = latDir
        self.lonDir = lonDir
        self.arah_lintang = latDir
        self.arah_bujur = lonDir

        
    @property
    def lintang(self):
        if self.arah_lintang == "LU" or self.arah_lintang == "N":
            result = self.latD + self.latM/60 + self.latS/3600
            return result
        return -(self.latD + self.latM/60 + self.latS/3600) #* (-1)
    
    
    @property
    def bujur(self):
        if self.arah_bujur == "BT" or self.arah_bujur == "E":
            return self.lonD + self.lonM/60 + self.lonS/3600
        return -(self.lonD + self.lonM/60 + self.lonS/3600) #* (-1)


def arah_qiblat(lat, lon):
    a = (90 - lat)
    b = (90 - (21 + 25/60 + 21/3600))
    C = (lon - (39 + 49/60 + 34/3600))

    B = atan(1/(1/tan(radians(b))*sin(radians(a))/sin(radians(C))-cos(radians(a))*1/tan(radians(C))))
    return degrees(B)


def bayang_bayang_matahari(deklinasi, lintang, bujur, eot):

    kwd = (120 - bujur)/15
    a = (90 - deklinasi)
    b = (90 - lintang)
    B = (arah_qiblat(lintang, bujur))
    p = degrees(atan(1/(cos(radians(b)) * tan(radians(B)))))
    test = to_dms(eot)
    test2 = to_dms(kwd)
    c_p = degrees(acos(1/tan(radians(a)) * tan(radians(b)) * cos(radians(p))))
    test3 = to_dms(c_p)
    tets4 = to_dms(p)
    t1 = ((c_p + p)/15) + 12 - (eot/60) + kwd
    return (t1)

bandung = Lokasi(3, 19, 42, "LS", 114, 36, 51.97, "BT")

deklinasi = degrees(Ephemeris(2012, 7, 30, 8).sun_declination)
eot = Ephemeris(2012, 7, 30, 8).eot
x = bayang_bayang_matahari(deklinasi, bandung.lintang, bandung.bujur, eot)
print(to_dms(x))

    