from math import radians, cos, sin
from terms import terms_delta_epsilon, terms_delta_psi

#from terms import delta_psi, delta_epsilon

class Nutation:
    def __init__(self, t_td):
        self.t = t_td
        t = t_td
        self.d = radians((297.85036 + 445267.11148*t - 0.0019142*t**2 + t**3/189474) % 360)
        self.m = radians((357.52772 + 35999.05034*t - 0.0001603*t*t - t*t*t/300000) % 360)
        self.m_aksen = radians((134.96298 + 477198.867398*t + 0.0086972*t*t + t*t*t/56250) % 360)
        self.f = radians((93.27191 + 483202.017538*t - 0.0036825*t*t + t*t*t/327270) % 360)
        self.omega = radians((125.04452 - 1934.136261*t + 0.0020708*t*t + t*t*t/450000) % 360)

    @property
    def epsilon(self):
        hasil = self.epsilon_zero + (self.delta_epsilon_total)/3600
        hasil = radians(hasil)
        return hasil
    
    @property
    def epsilon_zero(self):
        t = self.t
        u = t/100
        hasil =  23 + 26/60 + 21.448/3600 + (-4680.93*u - 1.55*u**2 + 1999.25*u**3 - 51.38*u**4 - 249.67*u**5 - 39.05*u**6+ 7.12*u**7 + 27.87*u**8 + 5.79*u**9 + 2.45*u**10)/3600
        return hasil
    
    @property
    def delta_epsilon_total(self):
        t = self.t
        d = self.d
        m = self.m
        m_aksen = self.m_aksen
        f = self.f
        omega = self.omega

        total = 0
        for i in range(len(terms_delta_epsilon)):
            total += (terms_delta_epsilon[i][5] + terms_delta_epsilon[i][6]*t) * cos(terms_delta_epsilon[i][0] * d + terms_delta_epsilon[i][1] * m + terms_delta_epsilon[i][2] * m_aksen + terms_delta_epsilon[i][3] * f + terms_delta_epsilon[i][4] * omega)
        return total/10000

    @property
    def delta_psi(self):
        hasil = self.delta_psi_total/3600
        return hasil

    @property
    def delta_psi_total(self):
        t = self.t
        d = self.d
        m = self.m
        m_aksen = self.m_aksen
        f = self.f
        omega = self.omega
        total = 0
        for i in range(len(terms_delta_psi)):
            total += (terms_delta_psi[i][5] + terms_delta_psi[i][6]*t) * sin(terms_delta_psi[i][0] * d + terms_delta_psi[i][1] * m + terms_delta_psi[i][2] * m_aksen + terms_delta_psi[i][3] * f + terms_delta_psi[i][4] * omega)
        return total/10000

        pass
