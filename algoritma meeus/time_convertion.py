from math import cos
from nutation import Nutation
from functions import julian_day, delta_T

class KonversiWaktu:
    def __init__(self,year, month, day, hour, minute, second, timezone):
        self.year           = year
        self.month          = month
        self.day            = day
        self.hour           = hour
        self.minute         = minute
        self.second         = second
        self.timezone       = timezone
        self.JD             = julian_day(self.year, self.month, self.day, self.hour, self.minute, self.second, self.timezone)
        self.T_ut           = (self.JD_ut-2451545)/36525
        tahun_desimal       = self.year + (self.month-1)/12 + self.day/365
        delta_t             = delta_T(tahun_desimal)/86400
        JDE                 = self.JD_ut + delta_t
        self.T_td           = (JDE - 2451545)/36525


class T_td:
    '''
    Julian Ephemeris Century
    '''
    def __init__(self,year, month, day, hour, minute, second, timezone):
        self.year           = year
        self.month          = month
        self.day            = day
        self.hour           = hour
        self.minute         = minute
        self.second         = second
        self.timezone       = timezone
        JD                  = julian_day(self.year, self.month, self.day, self.hour, self.minute, self.second, self.timezone)
        tahun_desimal       = self.year + (self.month-1)/12 + self.day/365
        delta_t             = delta_T(tahun_desimal)/86400
        JDE                 = JD + delta_t
        self.T_td           = (JDE - 2451545)/36525


class LST:
    '''
    local_sidere
    '''
    def __init__(self,year, month, day, hour, minute, second, timezone, longitude ):
        self.year           = year
        self.month          = month
        self.day            = day
        self.hour           = hour
        self.minute         = minute
        self.second         = second
        self.timezone       = timezone
        JD                  = julian_day(self.year, self.month, self.day, self.hour, self.minute, self.second, self.timezone)
        self.T_ut           = (JD-2451545)/36525
        tahun_desimal       = self.year + (self.month-1)/12 + self.day/365
        delta_t             = delta_T(tahun_desimal)/86400
        JDE                 = JD + delta_t
        T_td                = (JDE - 2451545)/36525
        T_td                = (JDE - 2451545)/36525
        GST                 = ((280.46061837 + 360.98564736629*(self.JD_ut - 2451545) + 0.000387933*self.T_ut**2 - self.T_ut**3/38710000) % 360)/1
        nutasi              = Nutation(T_td)
        GST_nampak          = GST + nutasi.delta_psi * cos(nutasi.epsilon)/15
        self.LST_nampak     = (GST_nampak + longitude/15) % 24