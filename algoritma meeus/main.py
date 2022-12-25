from math import *
from functions.utilities import *
from functions.calcutation import *
from place import Place
from functions.nutation import Nutation
from functions.moon import *

class DateTime:
    def __init__ (self, year, month, day, hour=0, minute=0, second=0, timezone=0, longitude=0):
        self.year       = year
        self.month      = month
        self.day        = day
        self.hour       = hour
        self.minute     = minute
        self.second     = second
        self.timezone   = timezone
        self.JD_ut      = jd_ut(self.year, self.month, self.day, self.hour, self.minute, self.second, self.timezone)
        self.T_ut       = (self.JD_ut-2451545)/36525
        self.delta_t    = delta_T(self.year, self.month, self.day)/86400
        self.JDE        = self.JD_ut + self.delta_t
        self.T_td       = (self.JDE - 2451545)/36525
        self.tau        = self.T_td/10
        self.GST        = ((280.46061837 + 360.98564736629*(self.JD_ut - 2451545) + 0.000387933*self.T_ut**2 - self.T_ut**3/38710000) % 360)/15
        nutasi          = Nutation(self.T_td)
        self.GST_nampak = self.GST + nutasi.delta_psi * cos(nutasi.epsilon)/15
        self.LST_nampak = (self.GST_nampak + longitude/15) % 24
        self.koreksi_dpsi = nutasi.delta_psi_total
        

class DataBulan():
    def __init__(self, t_td):
        self.t_td                   = t_td
        self.L_deg                  = (218.3164591 + 481267.88134236*t_td - 0.0013268*t_td*t_td + t_td*t_td*t_td/538841 - t_td*t_td*t_td*t_td/65194000) % 360
        self.L_rad                  = radians((218.3164591 + 481267.88134236*t_td - 0.0013268*t_td*t_td + t_td*t_td*t_td/538841 - t_td*t_td*t_td*t_td/65194000) % 360)
        moon                        = Moon(self.t_td)
        self.koreksi_bujur_bulan    = moon.koreksi_bujur_bulan
        self.bujur_bulan            = (moon.bujur_rata_rata_bulan + moon.koreksi_bujur_bulan)%360
        pass



coordinate = Place(6,57,0,"LS", 107, 37,0,"BT")
waktu = DateTime(2022, 12, 14, 0, 0, 0,7, coordinate.longitude_dec)
data_bulan = DataBulan(waktu.T_td)
#bulan   = Moon(waktu.T_td)

# print('---- Koordinat ----')
# # print(koordinat.latitude_dec)
# # print(koordinat.latitude_rad)
# print(coordinate.latitude_dms)
# print('\n---- Bujur ----')
# # print(koordinat.longitude_dec)
# # print(koordinat.longitude_rad)
# print(coordinate.longitude_dms)
# print('\n\n---- Hasil Perhitungan ----')
# print("JD dalam UT\t:", waktu.JD_ut)
# print("Delta T\t\t:",waktu.delta_t)
# print("T dalam UT\t:",waktu.T_ut)
# print("T dalam TD\t:",waktu.T_td)
# print("JDE:",waktu.JDE)
# print("LST nampak:",waktu.LST_nampak)


print('---- DATA BULAN ----')
print("Bujur rata-rata bulan (L')\t:", "Derajat: ", data_bulan.L_deg, "\tRadian:", data_bulan.L_rad)
print("Koreksi Bujur Bulan\t\t:", data_bulan.koreksi_bujur_bulan)
print("Bujur Bulan\t\t:", data_bulan.bujur_bulan)
# print("T dalam TD Psi\t\t\t:", waktu.T_td)
# print("L'\t\t\t\t:", degrees(bulan.bujur_rata_rata_bulan))

# print("Koreksi Delta Psi\t\t:", waktu.koreksi_dpsi)
# print("Koreksi Lintang Bulan\t\t:", bulan.koreksi_lintang_bulan)
# print("Koreksi Jarak Bumi-Bulan\t:", bulan.koreksi_jarak_bumi_bulan)


