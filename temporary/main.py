class Latitude:
    def __init__(self, d, m, s):
        self.d = d
        self.m = m
        self.s = s
    
    @property    
    def dms(self):
        return [self.d, self.m, self.s]


# class Longitude:
#     def __init__(self, d, m, s):
#         self
#         pass

class Tempat():
    def __init__(self, nama, lintang):
        self.nama = nama
        self.d = lintang[0]
        self.m = lintang[1]
        self.s = lintang[2]
        
    def cetak(self):
        print(self.d, '\n', type(self.m), self.s)
    
lintang = Latitude(-6, 30, 34)    
Bandung = Tempat("Bandung", lintang.dms)
Bandung.cetak()
print(lintang.dms)