import re

test_string = '''
Dalam Buku Ephemeris Hisab Rukyat tahun 2009 pada Tanggal 20 Agustus 2009 dapat diturunkan data-data sbb:
1.      FIB (Fraction Illumination Bulan) terkecil pada tanggal 20 Agustus 2009 adalah = 0.00045 jam 10:00 GMT 
2.      ELM (Eliptic Longitude Matahari)  pada pukul 10:00 GMT = 147o 31’ 30”
3.      ALB (Apparent Longitude Bulan)    pada pukul 10:00 GMT = 147o 29’ 52”
4.      Menghitung Sabaq Matahari (SM) per Jam (Harga Mutlak)

'''

pattern = re.compile(r'\d')
matches = pattern.finditer(test_string)
for match in matches:
    print(match)