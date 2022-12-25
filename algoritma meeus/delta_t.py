def delta_T(year_decimal):
    #year_decimal = year + (month-1)/12 + day/365
    delta = 0
    if year_decimal <= -500:
        delta = -20 + 32 * (year_decimal/100 - 18.2) * (year_decimal/100 - 18.2)
    if year_decimal > 500 and year_decimal <= 1600:
        delta =  1574.2 - 556.01*(year_decimal/100-10) + 71.23472*(year_decimal/100-10)*(year_decimal/100-10) 
        delta += 0.319781*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10) 
        delta -= 0.8503463*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10) 
        delta -= 0.005050998*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10) 
        delta += 0.0083572073*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)*(year_decimal/100-10)

    if year_decimal > 1600 and year_decimal <= 1700:
        delta = 120 - 0.9808*(year_decimal-1600) - 0.01532*(year_decimal-1600)*(year_decimal-1600) 
        delta += (year_decimal-1600)*(year_decimal-1600)*(year_decimal-1600)/7129

    if year_decimal > 1700 and year_decimal <= 1800:
        delta = 8.83 + 0.1603*(year_decimal-1700) - 0.0059285*(year_decimal-1700)*(year_decimal-1700) 
        delta += 0.00013336*(year_decimal-1700)*(year_decimal-1700)*(year_decimal-1700) 
        delta -= (year_decimal-1700)*(year_decimal-1700)*(year_decimal-1700)*(year_decimal-1700)/1174000

    if year_decimal > 1800 and year_decimal <= 1860:
        delta = 13.72 - 0.332447*(year_decimal-1800) + 0.0068612*(year_decimal-1800)*(year_decimal-1800) 
        delta += 0.0041116*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800) 
        delta -= 0.00037436*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800) 
        delta += 0.0000121272*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800) 
        delta -= 0.0000001699*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800) 
        delta += 0.000000000875*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)*(year_decimal-1800)

    if year_decimal > 1860 and year_decimal <= 1900:
        delta = 7.62 + 0.5737*(year_decimal-1860) - 0.251754*(year_decimal-1860)*(year_decimal-1860) 
        delta += 0.01680668*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860) 
        delta -= 0.0004473624*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860) 
        delta += (year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860)*(year_decimal-1860)/233174

    if year_decimal > 1900 and year_decimal <= 1920:
        delta = -2.79 + 1.494119*(year_decimal-1900) - 0.0598939*(year_decimal-1900)*(year_decimal-1900) 
        delta += 0.0061966*(year_decimal-1900)*(year_decimal-1900)*(year_decimal-1900) 
        delta -= 0.000197*(year_decimal-1900)*(year_decimal-1900)*(year_decimal-1900)*(year_decimal-1900)

    if year_decimal > 1920 and year_decimal <= 1941:
        delta = 21.2 + 0.84493*(year_decimal-1920) 
        delta -= 0.0761*(year_decimal-1920)*(year_decimal-1920) 
        delta += 0.0020936*(year_decimal-1920)*(year_decimal-1920)*(year_decimal-1920)

    if year_decimal > 1941 and year_decimal <= 1961:
        delta = 29.07 + 0.407*(year_decimal-1950) 
        delta -= (year_decimal-1950)*(year_decimal-1950)/233 
        delta += (year_decimal-1950)*(year_decimal-1950)*(year_decimal-1950)/2547

    if year_decimal > 1961 and year_decimal <= 1986:
        delta = 45.45 + 1.067*(year_decimal-1975) 
        delta -= (year_decimal-1975)*(year_decimal-1975)/260 
        delta -= (year_decimal-1975)*(year_decimal-1975)*(year_decimal-1975)/718

    if year_decimal > 1986 and year_decimal <= 2005:
        delta = 63.86 + 0.3345*(year_decimal-2000) 
        delta -= 0.060374*(year_decimal-2000)*(year_decimal-2000) 
        delta += 0.0017275*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000) 
        delta += 0.000651814*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000) 
        delta += 0.00002373599*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000)*(year_decimal-2000)

    if year_decimal > 2005 and year_decimal <= 2050:
        delta = 62.92 + 0.32217*(year_decimal-2000) 
        delta += 0.005589*(year_decimal-2000)*(year_decimal-2000)

    if year_decimal > 2050 and year_decimal <= 2150:
        delta = -20 + 32*((year_decimal - 1820)/100)*((year_decimal - 1820)/100) - 0.5628*(2150 - year_decimal)

    if year_decimal > 2150:
        delta = -20 + 32*((year_decimal - 1820)/100)*((year_decimal - 1820)/100)
    return delta