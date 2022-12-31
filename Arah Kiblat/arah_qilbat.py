from math import sin, cos, tan, atan, radians, degrees

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
        self.lintang = self.latD + self.latM/60 + self.latS/3600
        self.bujur = self.lonD + self.lonM/60 + self.lonS/3600
        self.arah_lintang = latDir
        self.arah_bujur = lonDir
    
    @property
    def lintang(self):
        if self.arah_lintang == "LU" or self.arah_lintang == "N":
            return self.lintang
        return self.lintang * (-1)
    
    
    @property
    def bujur(self):
        if self.arah_bujur == "BT" or self.arah_bujur == "E":
            return self.bujur
        return self.bujur * (-1)

    # @lintang.setter
    # def set_lintang(self):
    #     return self.latD + self.latM/60 + self.latS/3600

    # @bujur.setter
    # def set_bujur(self):
    #     return self.lonD + self.lonM/60 + self.lonS/3600


@property
def arah_qiblat(lat, lon):
    a = radians(90 - lat)
    b = radians(90 - (21 + 25/60))
    C = radians(lon - (39 + 50/60))

    B = atan(1/(1/tan(b)*sin(a)/sin(C)-cos(a)*1/tan(C)))
    return degrees(B)


bandung = Lokasi(6, 57, 0, "LS", 107, 37, 0, "BT")
x = arah_qiblat(bandung.latD, bandung.lonD)
print(x)

    