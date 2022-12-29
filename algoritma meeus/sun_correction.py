from math import degrees, radians, sin, cos
from moon_correction import MoonCorrection
from terms import *

class SunCorrection:
    def __init__(self, t):
        #  Julian Millenial Ephemeris
        self.tau = t/10
        _theta = (degrees(self.L()) + 180) % 360
        _delta_theta = -0.09033/3600  
        self.theta_terkoreksi = (_theta + _delta_theta) % 360
        self._lambda_aksen = radians((_theta - 1.397 * t - 0.00031*t**2) % 360)


    def L(self):
        tau = self.tau
        L0 = list_L0
        L1 = list_L1
        L2 = list_L2
        L3 = list_L3
        L4 = list_L4
        L5 = list_L5

        totalL0 = 0
        totalL1 = 0
        totalL2 = 0
        totalL3 = 0
        totalL4 = 0
        totalL5 = 0
        for i in range(len(L0)):
            totalL0 += L0[i][0]*cos(L0[i][1] + L0[i][2] * tau)

        for i in range(len(L1)):
            totalL1 += L1[i][0]*cos(L1[i][1] + L1[i][2] * tau)

        for i in range(len(L2)):
            totalL2 += L2[i][0]*cos(L2[i][1] + L2[i][2] * tau)

        for i in range(len(L3)):
            totalL3 += L3[i][0]*cos(L3[i][1] + L3[i][2] * tau)

        for i in range(len(L4)):
            totalL4 += L4[i][0]*cos(L4[i][1] + L4[i][2] * tau)

        for i in range(len(L5)):
            totalL5 += L5[i][0]*cos(L5[i][1] + L5[i][2] * tau)

        return (totalL0 + totalL1*tau + totalL2*tau**2 + totalL3*tau**3 + totalL4*tau**4 + totalL5*tau**5)/100000000


    @property
    def jarak_bumi_matahari(self):
        tau = self.tau
        R0 = list_R0
        R1 = list_R1
        R2 = list_R2
        R3 = list_R3
        R4 = list_R4

        totalR0 = totalR1 = totalR2 = totalR3 = totalR4 = 0

        for i in range(len(R0)):
            totalR0 += R0[i][0] * cos(R0[i][1] + R0[i][2] * tau)
        for i in range(len(R1)):
            totalR1 += R1[i][0] * cos(R1[i][1] + R1[i][2] * tau)
        for i in range(len(R2)):
            totalR2 += R2[i][0] * cos(R2[i][1] + R2[i][2] * tau)
        for i in range(len(R3)):
            totalR3 += R3[i][0] * cos(R3[i][1] + R3[i][2] * tau)
        for i in range(len(R4)):
            totalR4 += R4[i][0] * cos(R4[i][1] + R4[i][2] * tau)

        return (totalR0 + totalR1 * tau + totalR2*tau**2 + totalR3*tau**3 + totalR4*tau**4)/100000000

    # di class matahari disebut beta/lintang matahari tampak
    @property
    def lintang_matahari_tampak(self):
        tau = self.tau
        B0 = list_B0
        B1 = list_B1
        totalB0 = 0
        totalB1 = 0

        for i in range(len(B0)):
            totalB0 += B0[i][0] * cos(B0[i][1] + B0[i][2] * tau)

        for i in range(len(B1)):
            totalB1 += B1[i][0] * cos(B1[i][1] + B1[i][2] * tau)
        totalB1 *= tau
        B = degrees((totalB0 + totalB1)/100000000)*3600
        beta = -B
        delta_beta = 0.03916 * (cos(self._lambda_aksen)-sin(self._lambda_aksen))
        beta_terkoreksi = beta + delta_beta
        return radians(beta_terkoreksi/3600)

    

