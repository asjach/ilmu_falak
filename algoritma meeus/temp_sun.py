from math import asin, atan, atan2, degrees, radians, sin, cos, tan
#from temp_moon import MoonCorrection
from nutation import Nutation
from terms import *

class SunCorrection:
    def __init__(self, T):
        #  Julian Millenial Ephemeris
        self.tau = T/10
        _theta = (degrees(self.Lambda()) + 180) % 360
        _delta_theta = -0.09033/3600  
        self.theta_terkoreksi = (_theta + _delta_theta) % 360
        self._lambda_aksen = ((_theta - 1.397 * T - 0.00031*T**2) % 360)


    def Lambda(self):
        tau = self.tau
        L0 = list_L0
        L1 = list_L1
        L2 = list_L2
        L3 = list_L3
        L4 = list_L4
        L5 = list_L5

        totalL0 = totalL1 = totalL2 = totalL3 = totalL4 = totalL5 = 0
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
    def sun_distance(self):
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
    def apparent_latitude(self):
        tau = self.tau
        B0 = list_B0
        B1 = list_B1
        totalB0 = totalB1 = 0

        for i in range(len(B0)):
            totalB0 += B0[i][0] * cos(B0[i][1] + B0[i][2] * tau)

        for i in range(len(B1)):
            totalB1 += B1[i][0] * cos(B1[i][1] + B1[i][2] * tau)
        totalB1 *= tau
        B = degrees((totalB0 + totalB1)/100000000)*3600
        beta = -B
        delta_beta = 0.03916 * (cos(radians(self._lambda_aksen))-sin(radians(self._lambda_aksen)))
        beta_terkoreksi = beta + delta_beta
        return radians(beta_terkoreksi/3600)


class SunEcliptic():
    def __init__(self, T) -> None:
        self.sun_correction = SunCorrection(T)
        self.nutation =Nutation(T)
    
    
    @property
    def koreksi_aberasi(self):
        return -20.4898/(3600*self.sun_correction.sun_distance)


    @property
    def apparent_longitude(self):
        return radians((self.sun_correction.theta_terkoreksi + self.nutation.delta_psi + self.koreksi_aberasi) % 360)


    @property
    def apparent_latitude(self):
        hasil = self.sun_correction.apparent_latitude
        return (hasil)


class SunEquatorial():
    def __init__(self, T) -> None:
        self.sun_ecliptic = SunEcliptic(T)
        self.nutation = Nutation(T)


    @property
    def right_acsension(self):
        y = sin(self.sun_ecliptic.apparent_longitude)*cos(self.nutation.epsilon) - tan(self.sun_ecliptic.apparent_latitude)*sin(self.nutation.epsilon)
        x = cos((self.sun_ecliptic.apparent_longitude))
        hasil =(degrees(atan2(y,x)) % 360)
        return radians(hasil)


    @property
    def declination(self):
        return asin(sin(self.sun_ecliptic.apparent_latitude)*cos(self.nutation.epsilon) + cos(self.sun_ecliptic.apparent_latitude)*sin(self.nutation.epsilon)*sin(self.sun_ecliptic.apparent_longitude))


    @property
    def hour_angle(self):
        return self.hour_angle


    @hour_angle.setter
    def hour_angle(self, lst):
        return radians(self.lst*15-self.right_acsension)


class SunHorizon():
    def __init__(self, T, latitude) -> None:
        self.sun_equatorial = SunEquatorial(T)
        self.latitude = radians(latitude)
    
    
    @property
    def azimut_south(self):
        return self.azimut_south

    @azimut_south.setter
    def azimut_south(self):
        y = sin(self.sun_equatorial.hour_angle)
        x = cos(self.sun_equatorial.hour_angle)*sin(self.latitude) - tan(self.sun_equatorial.declination)*cos(self.latitude)
        return degrees((atan2(y,x)))
    
    
    @property
    def azimut_matahari(self):
        return (self.azimut_south + 180) % 360


    @property
    def altitude(self):
        return degrees(asin(sin(self.latitude)*sin(self.sun_equatorial.declination)+cos(self.latitude)*cos(self.sun_equatorial.declination)*cos(self.sun_equatorial.hour_angle)))


class SunProperty():
    def __init__(self, T) -> None:
        self.tau = T/10
        self.sun_correction = SunCorrection(T)
        self.sun_equatorial = SunEquatorial(T)
        self.nutation = Nutation(T)
    

    @property
    def sun_distance_au(self):
        return self.sun_correction.sun_distance

    @property
    def sun_distance_km(self):
        return self.sun_correction.sun_distance*149598000


    @property
    def parallax(self):
        return degrees(atan(6378.14/(self.sun_distance_km)))


    @property
    def equation_of_time(self):
        L0 = (280.4664567 + 360007.6982779 * self.tau + 0.03032028 * self.tau **2 + self.tau**3/49931 - self.tau**4/15300 - self.tau**5/2000000) % 360
        eot = L0 - 0.0057183 - self.sun_equatorial.right_acsension + self.nutation.delta_psi * cos(radians(self.nutation.epsilon))
        return eot


    @property
    def sudut_jari_jari_matahari(self):
        return (959.63/3600)/self.sun_distance_km






