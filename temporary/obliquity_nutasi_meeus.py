import math
from temporary.datalist import delta_epsilon as de
from temporary.datalist import delta_psi as dp


def u(t_td):
    return t_td/100


def epsilon_zero(u):
    epsilon_0 = 23 + 26/60 + 21.448/3600 
    epsilon_0 += (-4680.93*u - 1.55*u**2 + 1999.25*u**3 - 51.38*u**4 - 249.67*u**5 - 39.05*u**6+ 7.12*u**7 + 27.87*u**8 + 5.79*u**9 + 2.45*u**10)/3600
    return epsilon_0


def delta_epsilon_total(t_td):
    t = t_td
    d = math.radians((297.85036 + 445267.11148*t - 0.0019142*t**2 + t**3/189474) % 360)
    m = math.radians((357.52772 + 35999.05034*t - 0.0001603*t*t - t*t*t/300000) % 360)
    m_aksen = math.radians((134.96298 + 477198.867398*t + 0.0086972*t*t + t*t*t/56250) % 360)
    f = math.radians((93.27191 + 483202.017538*t - 0.0036825*t*t + t*t*t/327270) % 360)
    omega = math.radians((125.04452 - 1934.136261*t + 0.0020708*t*t + t*t*t/450000) % 360)
    total = 0
    for i in range(len(de)):
            total += (de[i][5] + de[i][6]*t) * math.cos(de[i][0] * d + de[i][1] * m + de[i][2] * m_aksen + de[i][3] * f + de[i][4] * omega)
    return total
    

def epsilon(epsilon_zero, delta_epsilon_total):
    return epsilon_zero + delta_epsilon_total/3600


def delta_psi_total(t_td):
    t = t_td
    d = math.radians((297.85036 + 445267.11148*t - 0.0019142*t**2 + t**3/189474) % 360)
    m = math.radians((357.52772 + 35999.05034*t - 0.0001603*t*t - t*t*t/300000) % 360)
    m_aksen = math.radians((134.96298 + 477198.867398*t + 0.0086972*t*t + t*t*t/56250) % 360)
    f = math.radians((93.27191 + 483202.017538*t - 0.0036825*t*t + t*t*t/327270) % 360)
    omega = math.radians((125.04452 - 1934.136261*t + 0.0020708*t*t + t*t*t/450000) % 360)
    total = 0
    for i in range(len(dp)):
            total += (dp[i][5] + dp[i][6]*t) * math.sin(dp[i][0] * d + dp[i][1] * m + dp[i][2] * m_aksen + dp[i][3] * f + dp[i][4] * omega)
    return total/10000

def delta_psi(delta_psi_total):
    delta_psi_total/3600