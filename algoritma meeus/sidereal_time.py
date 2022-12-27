from math import cos
from nutation import Nutation
from time_convertion import KonversiWaktu

class SiderealTime(KonversiWaktu):
    def __init__(self, year, month, day, hour, minute, second, timezone, longitude):
        super().__init__(year, month, day, hour, minute, second, timezone)
        self.longitude = longitude

        nutasi              = Nutation(self.T_td)
        self.GST_nampak     = self.GST + nutasi.delta_psi * cos(nutasi.epsilon)/15
        self.LST_nampak     = (self.GST_nampak + longitude/15) % 24
        self.koreksi_dpsi   = nutasi.delta_psi_total