from math import cos
from delta_t import delta_T
from nutation import Nutation

class KonversiWaktu:
    def __init__(self,year, month, day, hour, minute, second, timezone, longitude ):
        self.year           = year
        self.month          = month
        self.day            = day
        self.hour           = hour
        self.minute         = minute
        self.second         = second
        self.timezone       = timezone
        self.T_ut           = (self.JD_ut-2451545)/36525
        delta_t             = delta_T(self.year_decimal)/86400
        JDE                 = self.JD_ut + delta_t
        self.T_td           = (JDE - 2451545)/36525
        GST                 = ((280.46061837 + 360.98564736629*(self.JD_ut - 2451545) + 0.000387933*self.T_ut**2 - self.T_ut**3/38710000) % 360)/1
        nutasi              = Nutation(self.T_td)
        GST_nampak          = GST + nutasi.delta_psi * cos(nutasi.epsilon)/15
        self.LST_nampak     = (GST_nampak + longitude/15) % 24

    @property
    def JD_ut(self):
        m = self.month + 12 if self.month < 3 else self.month
        y = self.year - 1 if self.month < 3 else self.year
        a = int(y/100)
        if self.year < 1583:
            if self.month < 11:
                if self.day < 4:
                    b = 0
                    if self.day > 14:
                        b = 2 + int(a/4)-a
                    else:
                        print("tanggal salah")
            else:
                b = 2 + int(a/4) - a
        else:
            b   = 2 + int(a/4) - a
        
        return 1720994.5 + int(365.25 * y) + int(30.60001 * (m + 1)) + b + self.day + (self.hour + self.minute/60 + self.second/3600)/24 - self.timezone/24


    @property
    def year_decimal(self):
            return self.year + (self.month-1)/12 + self.day/365

