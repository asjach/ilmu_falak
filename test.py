from math import sin, cos, tan, radians, degrees, asin

# Fungsi untuk menghitung nilai tanggal Julian
def julian_date(year, month, day, hour, minute, second):
    a = (14 - month) // 12
    y = year + 4800 - a
    m = month + 12 * a - 3
    jd = day + (153 * m + 2) // 5 + 365 * y + y // 4 - y // 100 + y // 400 - 32045
    jd += (hour - 12) / 24 + minute / 1440 + second / 86400
    return jd

# Fungsi untuk menghitung nilai deklinasi Matahari
def sun_declination(year, month, day, hour, minute, second):
    # Hitung nilai tanggal Julian
    jd = julian_date(year, month, day, hour, minute, second)
    
    # Hitung nilai lama periode rotasi Bumi terhadap sumbunya (eksentrisitas orbit Bumi)
    T = jd - 2451545.0
    e = 0.016708634 - 0.000042037 * T - 0.0000001267 * T**2
    
    # Hitung nilai mean anomaly (anomali rata-rata) Matahari pada tanggal yang akan dihitung
    M = radians(357.52910 + 35999.05030 * T - 0.0001559 * T**2 - 0.00000048 * T**3)
    
    # Hitung nilai longitude of the ascending node (longitude dari ascending node) Matahari pada tanggal yang akan dihitung
    Omega = radians(125.04 - 1934.136 * T)
    
    # Hitung nilai deklinasi Matahari pada tanggal yang akan dihitung
    sin_delta = sin(Omega) * sin(M)
    delta = degrees(sin_delta)
    
    return delta

def declination_angle(year, month, day, hour, minute, second, timezone):
  # Tentukan nilai tanggal Julian (Julian Date) dari tanggal Gregorian yang akan dihitung
  a = (14 - month) // 12
  y = year + 4800 - a
  m = month + 12 * a - 3
  jd = day + (153 * m + 2) // 5 + 365 * y + y // 4 - y // 100 + y // 400 - 32045
  jd += (hour - timezone) / 24.0 + minute / 1440.0 + second / 86400.0
  
  # Tentukan nilai lama periode rotasi Bumi terhadap sumbunya (eksentrisitas orbit Bumi)
  T = jd - 2451545.0
  e = 0.016708634 - 0.000042037 * T - 0.0000001267 * T**2
  
  # Tentukan nilai mean anomaly (anomali rata-rata) Matahari pada tanggal yang akan dihitung
  M = radians(357.52911 + 35999.05029 * T - 0.0001537 * T**2)
  
  # Tentukan nilai true anomaly (anomali sebenarnya) Matahari pada tanggal yang akan dihitung
  C = radians(1.914602 - 0.004817 * T - 0.000014 * T**2)
  v = M + C
  
  # Tentukan nilai longitude of the ascending node (longitude dari ascending node) pada tanggal yang akan dihitung
  omega = radians(125.04 - 1934.136 * T)
  
  # Tentukan nilai deklinasi Matahari pada tanggal yang akan dihitung
  delta = asin(sin(omega) * sin(v))
  
  return delta

# Contoh penggunaan fungsi sun_declination
x = declination_angle(2022, 12, 22, 00, 00, 00,7)
print(f"Nilai deklinasi Matahari pada tanggal 22 Desember 2020 pukul 12:00:00 adalah {x:.2f} derajat.")