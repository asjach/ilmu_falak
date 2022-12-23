# class DeltaT:
#     def __init__ (self, tahun, bulan, hari):
#         self.tahun = tahun
#         self.bulan = bulan
#         self.hari = hari
#         tahun_desimal= self.tahun + (self.bulan-1)/12 + self.hari/365
#         print(tahun_desimal)

def delta_t(tahun_desimal):
    delta = 0
    if tahun_desimal <= -500:
        delta = -20 + 32 * (tahun_desimal/100 - 18.2) * (tahun_desimal/100 - 18.2)
    if tahun_desimal > 500 and tahun_desimal <= 1600:
        delta =  1574.2 - 556.01*(tahun_desimal/100-10) + 71.23472*(tahun_desimal/100-10)*(tahun_desimal/100-10) 
        delta += 0.319781*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10) 
        delta -= 0.8503463*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10) 
        delta -= 0.005050998*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10) 
        delta += 0.0083572073*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10)*(tahun_desimal/100-10)

    if tahun_desimal > 1600 and tahun_desimal <= 1700:
        delta = 120 - 0.9808*(tahun_desimal-1600) - 0.01532*(tahun_desimal-1600)*(tahun_desimal-1600) 
        delta += (tahun_desimal-1600)*(tahun_desimal-1600)*(tahun_desimal-1600)/7129

    if tahun_desimal > 1700 and tahun_desimal <= 1800:
        delta = 8.83 + 0.1603*(tahun_desimal-1700) - 0.0059285*(tahun_desimal-1700)*(tahun_desimal-1700) 
        delta += 0.00013336*(tahun_desimal-1700)*(tahun_desimal-1700)*(tahun_desimal-1700) 
        delta -= (tahun_desimal-1700)*(tahun_desimal-1700)*(tahun_desimal-1700)*(tahun_desimal-1700)/1174000

    if tahun_desimal > 1800 and tahun_desimal <= 1860:
        delta = 13.72 - 0.332447*(tahun_desimal-1800) + 0.0068612*(tahun_desimal-1800)*(tahun_desimal-1800) 
        delta += 0.0041116*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800) 
        delta -= 0.00037436*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800) 
        delta += 0.0000121272*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800) 
        delta -= 0.0000001699*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800) 
        delta += 0.000000000875*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)*(tahun_desimal-1800)

    if tahun_desimal > 1860 and tahun_desimal <= 1900:
        delta = 7.62 + 0.5737*(tahun_desimal-1860) - 0.251754*(tahun_desimal-1860)*(tahun_desimal-1860) 
        delta += 0.01680668*(tahun_desimal-1860)*(tahun_desimal-1860)*(tahun_desimal-1860) 
        delta -= 0.0004473624*(tahun_desimal-1860)*(tahun_desimal-1860)*(tahun_desimal-1860)*(tahun_desimal-1860) 
        delta += (tahun_desimal-1860)*(tahun_desimal-1860)*(tahun_desimal-1860)*(tahun_desimal-1860)*(tahun_desimal-1860)/233174

    if tahun_desimal > 1900 and tahun_desimal <= 1920:
        delta = -2.79 + 1.494119*(tahun_desimal-1900) - 0.0598939*(tahun_desimal-1900)*(tahun_desimal-1900) 
        delta += 0.0061966*(tahun_desimal-1900)*(tahun_desimal-1900)*(tahun_desimal-1900) 
        delta -= 0.000197*(tahun_desimal-1900)*(tahun_desimal-1900)*(tahun_desimal-1900)*(tahun_desimal-1900)

    if tahun_desimal > 1920 and tahun_desimal <= 1941:
        delta = 21.2 + 0.84493*(tahun_desimal-1920) 
        delta -= 0.0761*(tahun_desimal-1920)*(tahun_desimal-1920) 
        delta += 0.0020936*(tahun_desimal-1920)*(tahun_desimal-1920)*(tahun_desimal-1920)

    if tahun_desimal > 1941 and tahun_desimal <= 1961:
        delta = 29.07 + 0.407*(tahun_desimal-1950) 
        delta -= (tahun_desimal-1950)*(tahun_desimal-1950)/233 
        delta += (tahun_desimal-1950)*(tahun_desimal-1950)*(tahun_desimal-1950)/2547

    if tahun_desimal > 1961 and tahun_desimal <= 1986:
        delta = 45.45 + 1.067*(tahun_desimal-1975) 
        delta -= (tahun_desimal-1975)*(tahun_desimal-1975)/260 
        delta -= (tahun_desimal-1975)*(tahun_desimal-1975)*(tahun_desimal-1975)/718

    if tahun_desimal > 1986 and tahun_desimal <= 2005:
        delta = 63.86 + 0.3345*(tahun_desimal-2000) 
        delta -= 0.060374*(tahun_desimal-2000)*(tahun_desimal-2000) 
        delta += 0.0017275*(tahun_desimal-2000)*(tahun_desimal-2000)*(tahun_desimal-2000) 
        delta += 0.000651814*(tahun_desimal-2000)*(tahun_desimal-2000)*(tahun_desimal-2000)*(tahun_desimal-2000) 
        delta += 0.00002373599*(tahun_desimal-2000)*(tahun_desimal-2000)*(tahun_desimal-2000)*(tahun_desimal-2000)*(tahun_desimal-2000)

    if tahun_desimal > 2005 and tahun_desimal <= 2050:
        delta = 62.92 + 0.32217*(tahun_desimal-2000) 
        delta += 0.005589*(tahun_desimal-2000)*(tahun_desimal-2000)

    if tahun_desimal > 2050 and tahun_desimal <= 2150:
        delta = -20 + 32*((tahun_desimal - 1820)/100)*((tahun_desimal - 1820)/100) - 0.5628*(2150 - tahun_desimal)

    if tahun_desimal > 2150:
        delta = -20 + 32*((tahun_desimal - 1820)/100)*((tahun_desimal - 1820)/100)

    return delta


