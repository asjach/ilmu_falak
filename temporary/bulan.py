import math
from temporary.datalist import koreksi_bujur_bulan as kb
from temporary.datalist import koreksi_lintang_bulan as kl
from temporary.datalist import koreksi_jarak_bumi_bulan as kj
from temporary.meeus import *

# def L(calc.T):
#     return math.radians((218.3164591 + 481267.88134236*t_td - 0.0013268*t_td*t_td + t_td*t_td*t_td/538841 - t_td*t_td*t_td*t_td/65194000) % 360)


# ==========================================





def koreksi_bulan(t_td):
    T       = t_td
    L       = math.radians((218.3164591 + 481267.88134236*t_td - 0.0013268*t_td*t_td + t_td*t_td*t_td/538841 - t_td*t_td*t_td*t_td/65194000) % 360)
    d       = math.radians((297.8502042 + 445267.1115168*T - 0.00163*T*T + T*T*T/545868 - T*T*T*T/113065000) % 360)
    m       = math.radians((357.5291092 + 35999.0502909*T - 0.0001536*T*T + T*T*T/24490000) % 360)
    m_aksen = math.radians((134.9634114 + 477198.8676313*T + 0.008997*T*T + T*T*T/69699 - T*T*T*T/14712000) % 360)
    f       = math.radians((93.2720993 + 483202.0175273*T - 0.0034029*T*T - T*T*T/3526000 + T*T*T*T/863310000) % 360)
    eksentrisitas_orbit = 1 - 0.002516*T - 0.0000074*T*T
    total_koreksi_bujur_bulan = 0
    for i in range(len(kb)):
        total_koreksi_bujur_bulan += kb[i][4] * (eksentrisitas_orbit**(abs(kb[i][1]))) * math.sin(kb[i][0] * d+kb[i][1] * m + kb[i][2] * m_aksen+kb[i][3] * f)
    
    a1 = math.radians((119.75 + 131.849 * T) % 360)
    a2 = math.radians((53.09 + 479264.29 * T) % 360)
    a3 = math.radians((313.45 + 481266.484 * T) % 360)

    koreksi_bujur_bulan = (total_koreksi_bujur_bulan + 3958 * math.sin(a1)+1962 * math.sin(L-f) + 318 * math.sin(a2)) / 1000000
    total_koreksi_lintang_bulan = 0

    for i in range(len(kl)):
        total_koreksi_lintang_bulan += kl[i][4] * (eksentrisitas_orbit**(abs(kl[i][1]))) * math.sin(kl[i][0] * d+kl[i][1] * m + kl[i][2] * m_aksen+kl[i][3] * f)


    koreksi_lintang_bulan = (total_koreksi_lintang_bulan - 2235 * math.sin(L) + 382*math.sin(a3) + 175*math.sin(a1 - f) + 175*math.sin(a1 + f) + 127*math.sin(L - m_aksen) - 115*math.sin(L + m_aksen))/1000000


    total_koreksi_jarak_bumi_bulan = 0
    for i in range(len(kj)):
        total_koreksi_jarak_bumi_bulan += kj[i][4] * (eksentrisitas_orbit**(abs(kj[i][1]))) * math.cos(kj[i][0] * d + kj[i][1] * m + kj[i][2] * m_aksen + kj[i][3] * f)


    koreksi_jarak_bumi_bulan = total_koreksi_jarak_bumi_bulan/1000


    test = koreksi_jarak_bumi_bulan


    print(test)





   