# def JD_UT(year, month, day, hour, minute, second, timezone):
#     m = month + 12 if month < 3 else month
#     y = year - 1 if month < 3 else year
#     a = int(y/100)
#     if year < 1583:
#         if month < 11:
#             if day < 4:
#                 b = 0
#                 if day > 14:
#                     b = 2 + int(a/4)-a
#                 else:
#                     print("tanggal salah")
#         else:
#             b = 2 + int(a/4) - a
#     else:
#         b   = 2 + int(a/4) - a
    
#     jd_ut = 1720994.5 + int(365.25 * y) + int(30.60001 * (m + 1)) + b + day + (hour + minute/60 + second/3600)/24 - timezone/24

#     return jd_ut


