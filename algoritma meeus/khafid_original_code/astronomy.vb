Option Explicit

' KONVERSI SISTEM WAKTU (15)
' ---------------------
'* Public Function JulianDatum(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function JulianEphemerisDay(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function JulianCentury(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function JulianEphemerisCentury(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function CalDat(ByVal JD As Double) As String
'* Public Function JD2Year(ByVal JD As Double) As integer
'* Public Function JD2Month(ByVal JD As Double) As integer
'* Public Function JD2Date(ByVal JD As Double) As integer
'* Public Function JD2Time(ByVal JD As Double) As Double
'* Public Function NamaHari(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer) As String
'* Public Function NamaBulan(ByVal Month As Integer) As String
'* Public Function DeltaT(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Double) As Double
'* Public Function GreenwichMeanSiderealTime(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function GreenwichHourAngle(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentLocalSiderealTime(ByVal Bujur As Double, ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double

' ARAH KIBLAT (7)
' --------------
'
'* Public Function ArahQibla(ByVal Lintang As Double, ByVal Bujur As Double) As Double
'* Public Function JarakQibla(ByVal Lintang As Double, ByVal Bujur As Double) As Double
'* Public Function ArahQiblaWithEllipsoidCorrection(ByVal Lintang As Double, ByVal Bujur As Double) As Double
'* Public Function JarakQiblaWithEllipsoidCorrection(ByVal Lintang As Double, ByVal Bujur As Double) As Double
'* Public Function ArahQiblaVincenty(ByVal Lintang As Double, ByVal Bujur As Double) As Double
'* Public Function JarakQiblaVincenty(ByVal Lintang As Double, ByVal Bujur As Double) As Double
'* Public Function ShadowRatioVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ShadowDirectionVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function RashdulQibla(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Double) As Double

' JADWAL SHALAT (8)
' -------------
'* Public Function EquationOfTime(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ShubuhTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
'                  Optional ByVal Alt As Double = 20#, Optional ByVal Ihtiyat As Double = 2#) As Double
'* Public Function FajrTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
'                  Optional ByVal Alt As Double = -20#, Optional ByVal Ihtiyat As Double = 2#) As Double
'* Public Function DuhrTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
'                  Optional ByVal Ihtiyat As Double = 2#) As Double
'* Public Function AltitudeOfSunAtAshr(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
'                  Optional ByVal Rasio As Double = 1#, Optional ByVal IsShadowDuhr As Boolean = True) As Double
'* Public Function AshrTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
'                  Optional ByVal Alt As Double = 1#, Optional ByVal Ihtiyat As Double = 2#) As Double''
'* Public Function MaghribTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
'                  Optional ByVal Alt As Double = -1#, Optional ByVal Ihtiyat As Double = 2#) As Double
'* Public Function IsyaTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
'                  Optional ByVal Alt As Double = -18#, Optional ByVal Ihtiyat As Double = 2#) As Double

' KOORDINAT MATAHARI (30)
' ------------------
'* Public Function SunMeanLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunMeanAnomaly(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunLongitudeLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunLatitudeLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function DistanceToTheSunLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentSunLongitudeLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunLongitudeVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunLatitudeVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function DistanceToTheSunVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentSunLongitudeVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunDeclinationLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunDeclinationVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentSunDeclinationLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentSunDeclinationVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunRightAscensionLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunRightAscensionVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentSunRightAscensionLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentSunRightAscensionVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunAltitudeVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
                   ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentSunAltitudeVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
                   ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunAzimuthVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
                   ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunAltitudeAtSunsetOrSunrise(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
'                  Optional ByVal p As Double = 1010, Optional ByVal T As Double = 27, Optional ByVal h As Double = 0) As Double
'* Public Function Sunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Alt As Double = -1#) As Double
'* Public Function Sunrise(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Alt As Double = -1#) As Double
'* Public Function SunSemiDiameterLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunSemiDiameterVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunsEqCenter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function EccentricityOfEarthOrbit(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SunHourAngleVSOP87(ByVal Lintang As Double, ByVal Alt As Double, ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function HourAngleOfTheSun(ByVal Bujur As Double, ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double

' POSISI BULAN METODE JEAN MEEUS (36)
' -----------------------------------
'* Public Function MoonMeanLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function Ijtimak(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function Konjungsi(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function NewMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function FirstQuarter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function FullMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function LastQuarter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentMoonLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonLatitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function DistanceToTheMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonRightAscention(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentMoonRightAscention(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonAltitude(ByVal Lintang As Double, ByVal Bujur As Double, _
                   ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentMoonAltitude(ByVal Lintang As Double, ByVal Bujur As Double, _
                   ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonAzimuth(ByVal Lintang As Double, ByVal Bujur As Double, _
                   ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonElongation(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonElongationAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                   ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                   Optional ByVal Alt As Double = -1#) As Double
'* Public Function MoonsAge(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) as Double
'* Public Function MoonsAgeJd(ByVal Jd As double) as Double
'* Public Function MoonsAgeAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
'                  Optional ByVal Alt As Double = -1#) As Double
'* Public Function MoonsAgeAtSunsetJd(Byval Jd as double, ByVal Lintang As Double, ByVal Bujur As Double, _
'                  Optional ByVal Alt As Double = -1#) As Double
'* Public Function MoonIllumination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonIlluminationAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                   ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                   Optional ByVal Alt As Double = -1#) As Double
'* Public Function CrescentWidth(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function CrescentWidthAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function AzimuthDifferenceAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function AzimuthDifference(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function Moonset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function Moonrise(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonTransit(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function HilalDuration(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
'                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonSemidiameter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonMeanAnomaly(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonMeanElongation(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function JDEofMaxNorthMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function JDEofMaxSouthMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MaxNorthMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MaxSouthMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonArgumentOfLatitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SigmaB(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SigmaL(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function SigmaR(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function HourAngleOfTheMoon(ByVal Bujur As Double, ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double

'* POSISI BULAN METODE ELP 2000/82
'---------------------------------
'* Public Function MoonLongitudeELP2000(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ApparentMoonLongitudeELP2000(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function MoonLatitudeELP2000(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function DistanceToTheMoonELP2000(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double


'* FASE-FASE BULAN METODE ELP2000/82 dan VSOP87
'----------------------------------------------
'* Public Function NewMoonELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function FirstQuarterELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function FullMoonELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function LastQuarterELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function NewMoonToNewMoonELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double

' KOREKSI-KOREKSI (14)
' ---------------
'* Public Function dLatitude(ByVal Lintang As Double)
'* Public Function SunAberration(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function NutationInLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function NutationInObliquity(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function HorizontalMoonParallax(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function RefractionApparentAltitude(ByVal h0 As Double, ByVal p As Double, ByVal T As Double)
'* Public Function kwd(ByVal TZ As Double, ByVal Bujur As Double) As Double
'* Public Function KoreksiDIP(Optional ByVal h As Double = 0#) As Double
'* Public Function KoreksiNewMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function KoreksiFirstQuarter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function KoreksiFullMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function KoreksiLastQuarter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function Koreksi_W(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'* Public Function Koreksi_PlanetaryArguments(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double

' KOMPONEN_KOMPONEN PEMBANTU (13)
' --------------------------
'* Public Function MeanObliquityOfEcliptic(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function ObliquityOfEcliptic(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function nilai_K(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'* Public Function value_K(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function nilai_T(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'* Public Function value_T(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function nilai_E(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'* Public Function nilai_M(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'* Public Function nilai_M1(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'* Public Function nilai_F(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'* Public Function nilai_O(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'* Public Function nilai_Omega(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function nilai_JDE(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'* Public Function Argument_A1(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function Argument_A2(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'* Public Function Argument_A3(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double

Public Function JulianDatum(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function JulianDatum dimaksudkan untuk menghitung Julian Datum
' dengan epoch J2000 pada 1 Januari 2010 jam 12.00 UT
'
' INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
' OUTPUT:
'     Julian Datum dalam satuan hari
'
' CATATAN:
'     Nilai Julian Datum akan bergantung dari Epoch yang dipilih.
'     The following method is valid for positive as well as for negative years, but not for negative JD.
'
' Contoh pemakaian fungsi:
'     Menghitung Julian datum pada tanggal 12 Februari 2010, jam 13:30:15 UT sbb:
'
'     maka :
'        Y = 2010
'        M = 2
'        D = 12.56267361111111
'        x = JulianDatum(Y,M,D)
'        x =
'            2455240.06267361'
'
'  Referensi:
'     Astronomical Algorithm, hal. 61
'----------------------------
' Cibinong, 04 Februari 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim aa, A, B, c As Double
   
   aa = 10000# * Year + 100# * Month + Day
   If (Month < 3) Then
       Month = Month + 12
       Year = Year - 1
   End If
   If (aa <= 15821004.99999) Then
      B = 0#
   Else
      A = Int(Year / 100#)
      B = 2 - A + Int(A / 4#)
   End If
   JulianDatum = Int(365.25 * (Year + 4716)) + Int(30.6001 * (Month + 1)) + Day + B - 1524.5
End Function

Public Function JulianEphemerisDay(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function JulianEphemerisDay dimaksudkan untuk menghitung Julian Ephemris Day
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 13:30:15 UT
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  OUTPUT:
'     Julian Ephemeris dalam satuan hari dihitung dari Epoch
'
' CATATAN:
'    Yang membedakan fungsi ini dengan fungsi JulianDatum adalah koreksi Delta T, sehingga perhitungan
'    dilakukan dalam waktu dinamis (Dynamical Time), sehingga hasil hitungannya menjadi Julian Ephemeris Day.
'    The following method is valid for positive as well as for negative years, but not for negative JD.
'
' Contoh pemakaian fungsi:
'     Menghitung Julian Ephemeris Day pada tanggal 12 Februari 2010, jam 13:30:15 UT sbb:
'
'     maka :
'        Y = 2010
'        M = 2
'        D = 12.56267361111111
'        x = JulianEphemerisDay(Y,M,D)
'        x =
'            2455240.06344621
'
'  Referensi:
'     Astronomical Algorithm, hal. 61
'----------------------------
' Cibinong, 04 Februari 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   JulianEphemerisDay = JulianDatum(Year, Month, Day) + DeltaT(Year, Month, Day) / 86400#
End Function


Public Function JulianCentury(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function JulianCentury dimaksudkan untuk menghitung Julian Century
' dengan epoch J2000 pada 1 Janauri jam 12.00 UT
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 13:30:15 UT
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  OUTPUT:
'     Julian Century dalam satuan abad (36525 hari)
'
'  CATATAN:
'     A Julian century is a time interval of 36525 days.
'     The following method is valid for positive as well as for negative years, but not for negative JD.
'
' Contoh pemakaian fungsi:
'     Menghitung Julian Ephemeris Day pada tanggal 12 Februari 2010, jam 13:30:15 UTsbb:
'
'     maka :
'        Y = 2010
'        M = 2
'        D = 12.56267361111111
'        x = JulianCentury(Y,M,D)
'        x =
'            0.101165302494486
'
'  Referensi:
'     Astronomical Algorithm, hal. 83
'----------------------------
' Cibinong, 04 Februari 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim JD As Double
   
   JD = JulianDatum(Year, Month, Day)
   JulianCentury = (JD - 2451545#) / 36525#
End Function

Public Function JulianEphemerisCentury(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function JulianEphemerisCentury dimaksudkan untuk menghitung Julian Ephemeris Century
' dengan epoch J2000 pada 1 Janauri jam 12.00 UT
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 13:30:15 UT
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
' OUTPUT:
'     Julian Ephemeris Century dalam satuan abad (36525 hari)
'
' CATATAN:
'     A Julian century is a time interval of 36525 days.
'     The following method is valid for positive as well as for negative years, but not for negative JD.
'
' Contoh pemakaian fungsi:
'     Menghitung Julian Ephemeris Century pada tanggal 12 Februari 2010, jam 13:30:15 UTsbb:
'
'     maka :
'        Y = 2010
'        M = 2
'        D = 12.56267361111111
'        x = JulianEphemerisCentury(Y,M,D)
'        x =
'            0.101165323647059
'
'  Referensi:
'     Astronomical Algorithm, hal. 83
'----------------------------
' Cibinong, 04 Februari 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim JDE As Double
   
   JDE = JulianEphemerisDay(Year, Month, Day)
   JulianEphemerisCentury = (JDE - 2451545#) / 36525#
End Function

Public Function CalDat(ByVal JD As Double) As String
'  Function CalDat dipakai untuk mengkonversi Julian Datum ke penanggalan yang dipakai sehari-hari
'  Fungsi ini merupakan kebalikan dari fungsi JulianDatum, yakni Calculation of the Calendar Date from the JD
'
'  INPUT:
'      - JD  : Julian Datum (JD) bisa juga Juluan Ephemeris Day (JDE)
'
'  OUTPUT:
'    Konversi CalDat dalam text string, memuat hari, tanggal, bulan, tahun, jam, menit dan detik
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Example 7.c — Calculate the calendar date corresponding to JD 2436 116.31.
'
'     maka :
'        JD = 2436 116.31
'        x = CalDat(JD)
'        x =
'            Jum'at, 04 Oktober 1957 19:26:24
'
'  CATATAN:
'  The following method is valid for positive as well as for
'  negative years, but not for negative Julian Day numbers.
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 63
' -------------------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Tgl, Bln, Thn, Jam As Double
  Dim alpha, A, B, c, D, E, F, Z As Double
  
    Z = Int(JD + 0.5)
    If Z < 2299161# Then
      A = Z '+ 1524#
    Else
      alpha = Int((Z - 1867216.25) / 36524.25)
      A = Z + 1 + alpha - Int(alpha / 4)
    End If
    
    B = A + 1524
    c = Int((B - 122.1) / 365.25)
    D = Int(365.25 * c)
    E = Int((B - D) / 30.6001)
    
    Tgl = Int(B - D - Int(30.6001 * E))
    Jam = JD + 0.5 - Z
    
    If E < 14 Then Bln = E - 1
    If E = 14 Or E = 15 Then Bln = E - 13
    
    If Bln > 2 Then Thn = c - 4716
    If Bln = 1 Or Bln = 2 Then Thn = c - 4715
    
    CalDat = NamaHari(Thn, Bln, Tgl) & Format(Tgl, ", 00 ") & _
             NamaBulan(Bln) & Format(Thn, " 0000 ") & Format(Jam, "hh:mm:ss")
End Function

Public Function JD2Year(ByVal JD As Double) As Integer
'  Function JD2Year dipakai untuk mengkonversi Julian Datum ke Tahun, lihat CalDat.
'  Fungsi ini sama dengan fungsi CalDat, namun hanya diambil tahunnya saja
'
'  INPUT:
'      - JD  : Julian Datum (JD) bisa juga Juluan Ephemeris Day (JDE)
'
'  OUTPUT:
'     Tahun dalam bilangan bulat (integer)
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Example 7.c — Calculate the calendar date corresponding to JD 2436 116.31.
'
'     maka :
'        JD = 2436116.31
'        x = JD2Year(JD)
'        x =
'            1957
'
'  CATATAN:
'    The following method is valid for positive as well as for
'    negative years, but not for negative Julian Day numbers.
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 63
' -------------------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Tgl, Bln, Thn, Jam As Double
  Dim alpha, A, B, c, D, E, F, Z As Double
  
    Z = Int(JD + 0.5)
    If Z < 2299161# Then
      A = Z '+ 1524#
    Else
      alpha = Int((Z - 1867216.25) / 36524.25)
      A = Z + 1 + alpha - Int(alpha / 4)
    End If
    
    B = A + 1524
    c = Int((B - 122.1) / 365.25)
    D = Int(365.25 * c)
    E = Int((B - D) / 30.6001)
    
    Tgl = Int(B - D - Int(30.6001 * E))
    Jam = JD + 0.5 - Z
    
    If E < 14 Then Bln = E - 1
    If E = 14 Or E = 15 Then Bln = E - 13
    
    If Bln > 2 Then Thn = c - 4716
    If Bln = 1 Or Bln = 2 Then Thn = c - 4715
    
    JD2Year = Thn
End Function

Public Function JD2Month(ByVal JD As Double) As Integer
'  Function JD2Month dipakai untuk mengkonversi Julian Datum ke bulan, lihat CalDat.
'  Fungsi ini sama dengan fungsi CalDat, namun hanya diambil bulannya saja
'
'  INPUT:
'      - JD  : Julian Datum (JD) bisa juga Juluan Ephemeris Day (JDE)
'
'  OUTPUT:
'     Bulan dalam bilangan bulat (integer)
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Example 7.c — Calculate the calendar date corresponding to JD 2436 116.31.
'
'     maka :
'        JD = 2436116.31
'        x = JD2Month(JD)
'        x =
'            10
'
'  CATATAN:
'  The following method is valid for positive as well as for
'  negative years, but not for negative Julian Day numbers.
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 63
' -------------------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Tgl, Bln, Thn, Jam As Double
  Dim alpha, A, B, c, D, E, F, Z As Double
  
    Z = Int(JD + 0.5)
    If Z < 2299161# Then
      A = Z
    Else
      alpha = Int((Z - 1867216.25) / 36524.25)
      A = Z + 1 + alpha - Int(alpha / 4)
    End If
    
    B = A + 1524
    c = Int((B - 122.1) / 365.25)
    D = Int(365.25 * c)
    E = Int((B - D) / 30.6001)
    
    Tgl = Int(B - D - Int(30.6001 * E))
    Jam = JD + 0.5 - Z
    
    If E < 14 Then Bln = E - 1
    If E = 14 Or E = 15 Then Bln = E - 13
    
    JD2Month = Bln
End Function

Public Function JD2Date(ByVal JD As Double) As Integer
'  Function JD2Date dipakai untuk mengkonversi Julian Datum ke Tanggal, lihat CalDat.
'  Fungsi ini sama dengan fungsi CalDat, namun hanya diambil Tanggalnya saja
'
'  INPUT:
'      - JD  : Julian Datum (JD) bisa juga Juluan Ephemeris Day (JDE)
'
'  OUTPUT:
'     Tanggal dalam bilangan bulat (integer)
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Example 7.c — Calculate the calendar date corresponding to JD 2436 116.31.
'
'     maka :
'        JD = 2436116.31
'        x = JD2Date(JD)
'        x =
'            4
'
'  CATATAN:
'    The following method is valid for positive as well as for
'    negative years, but not for negative Julian Day numbers.
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 63
' -------------------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Tgl, Bln, Thn, Jam As Double
  Dim alpha, A, B, c, D, E, F, Z As Double
  
    Z = Int(JD + 0.5)
    If Z < 2299161# Then
      A = Z '+ 1524#
    Else
      alpha = Int((Z - 1867216.25) / 36524.25)
      A = Z + 1 + alpha - Int(alpha / 4)
    End If
    
    B = A + 1524
    c = Int((B - 122.1) / 365.25)
    D = Int(365.25 * c)
    E = Int((B - D) / 30.6001)
    
    Tgl = Int(B - D - Int(30.6001 * E))
    JD2Date = Tgl
End Function

Public Function JD2Time(ByVal JD As Double) As Double
'  Function JD2Time dipakai untuk mengkonversi Julian Datum ke Time (Waktu), lihat CalDat.
'  Fungsi ini sama dengan fungsi CalDat, namun hanya diambil Waktunya saja
'
'  INPUT:
'      - JD  : Julian Datum (JD) bisa juga Juluan Ephemeris Day (JDE)
'
'  OUTPUT:
'     Waktu dalam bilangan desimal dalam satuan hari.
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Example 7.c — Calculate the calendar date corresponding to JD 2436 116.31.
'
'     maka :
'        JD = 2436116.31
'        x = JD2Time(JD)
'        x =
'            0.81
'
'  CATATAN:
'    The following method is valid for positive as well as for
'    negative years, but not for negative Julian Day numbers.
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 63
' -------------------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Z, Jam As Double
  
    Z = Int(JD + 0.5)
    Jam = JD + 0.5 - Z
    JD2Time = Jam

End Function

Public Function NamaHari(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer) As String
'  Function NamaHari dipakai untuk menentukan hari pada Tanggal, Bulan dan Tahun tertentu
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : tanggal
'
'  OUTPUT:
'     Nama Hari dalam Bahasa Indonesia dengan format string
'
'
' Contoh pemakaian fungsi:
'     Tanggal 15 Februari 1977 hari apa?
'      maka :
'        Y = 1977
'        M = 2
'        D = 15
'        x = NamaHari(Y, M, D)
'        x =
'            Selasa
'
' -------------------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Hari As Integer
  
  Hari = Modulus(Int(JulianDatum(Year, Month, Day) + 3), 7)
  Select Case Hari
     Case 1: NamaHari = "Ahad"
     Case 2: NamaHari = "Senin"
     Case 3: NamaHari = "Selasa"
     Case 4: NamaHari = "Rabu"
     Case 5: NamaHari = "Kamis"
     Case 6: NamaHari = "Jum'at"
     Case 0: NamaHari = "Sabtu"
  End Select
End Function

Public Function NamaBulan(ByVal Month As Integer) As String
'  Function NamaBulan dipakai untuk mengganti bulan dalam angka menjadi teks string
'
'  INPUT:
'      - Month : bulan Masehi (1-12)
'
'  OUTPUT:
'     Nama Bulan dalam Bahasa Indonesia dengan format string
'
'
' Contoh pemakaian fungsi:
'      Bulan ke 5 namanya apa?
'      maka :
'        M = 5
'        x = NamaBulan(M)
'        x =
'            Mei
'
' -------------------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Select Case Month
    Case 1: NamaBulan = "Januari"
    Case 2: NamaBulan = "Februari"
    Case 3: NamaBulan = "Maret"
    Case 4: NamaBulan = "April"
    Case 5: NamaBulan = "Mei"
    Case 6: NamaBulan = "Juni"
    Case 7: NamaBulan = "Juli"
    Case 8: NamaBulan = "Agustus"
    Case 9: NamaBulan = "September"
    Case 10: NamaBulan = "Oktober"
    Case 11: NamaBulan = "November"
    Case 12: NamaBulan = "Desember"
  End Select
End Function

Public Function DeltaT(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Double) As Double
' Fungsi DeltaT ini dimaksudkan untuk mrnghitung koreksi Delta T.
'
'           Delta T = TD - UT
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Delta T dalam detik
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Delta T Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan :
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi DeltaT.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = DeltaT(Y, M, D)
'    Hasil:
'         x = 66.7519393437262
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 71
' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' 18 Juni 2011
'
  Dim myYear, c As Double
  
  myYear = Year + (JulianDatum(Year, Month, Day) - JulianDatum(Year, 1, 0)) / 365.25

' untuk tahun seblum 500 B.C
     If (myYear <= -500) Then
        c = myYear / 100#
        DeltaT = -20 + 32 * (c - 18.2) ^ 2
     End If
     
' untuk tahun 500 B.C sampai 500 C.
     If (myYear > -500 And myYear <= 500) Then
        c = myYear / 100#
        DeltaT = 10583.6 - 1014.41 * c + 33.78311 * c ^ 2
        DeltaT = DeltaT - 5.952053 * c ^ 3 - 0.1798452 * c ^ 4
        DeltaT = DeltaT + 0.022174192 * c ^ 5 + 0.0090316521 * c ^ 6
     End If
' untuk tahun 500 C sampai 1600 C.
     If (myYear > 500 And myYear <= 1600) Then
        c = myYear / 100 - 10
        DeltaT = 1574.2 - 556.01 * c + 71.23472 * c ^ 2 + 0.319781 * c ^ 3
        DeltaT = DeltaT - 0.8503463 * c ^ 4 - 0.005050998 * c ^ 5 + 0.0083572073 * c ^ 6
     End If
' untuk tahun 1600 C sampai 1700 C.
     If (myYear > 1600 And myYear <= 1700) Then
        c = myYear - 1600
        DeltaT = 120 - 0.9808 * c - 0.01532 * c ^ 2 + c ^ 3 / 7129
     End If
      
' untuk tahun 1700 C sampai 1800 C.
     If (myYear > 1700 And myYear <= 1800) Then
        c = myYear - 1700
        DeltaT = 8.83 + 0.1603 * c - 0.0059285 * c ^ 2 + 0.00013336 * c ^ 3 - c ^ 4 / 1174000
     End If
' untuk tahun 1800 C sampai 1860 C.
     If (myYear > 1800 And myYear <= 1860) Then
        c = myYear - 1800
        DeltaT = 13.72 - 0.332447 * c + 0.0068612 * c ^ 2 + 0.0041116 * c ^ 3 - 0.00037436 * c ^ 4
        DeltaT = DeltaT + 0.0000121272 * c ^ 5 - 0.0000001699 * c ^ 6 + 0.000000000875 * c ^ 7
     End If
' untuk tahun 1860 C sampai 1900 C.
     If (myYear > 1860 And myYear <= 1900) Then
        c = myYear - 1860
        DeltaT = 7.62 + 0.5737 * c - 0.251754 * c ^ 2 + 0.01680668 * c ^ 3 - 0.0004473624 * c ^ 4 + c ^ 5 / 233174
     End If
' untuk tahun 1900 C sampai 1920 C.
     If (myYear > 1900 And myYear <= 1920) Then
        c = myYear - 1900
        DeltaT = -2.79 + 1.494119 * c - 0.0598939 * c ^ 2 + 0.0061966 * c ^ 3 - 0.000197 * c ^ 4
     End If
' untuk tahun 1920 C sampai 1941 C.
     If (myYear > 1920 And myYear <= 1941) Then
        c = myYear - 1920
        DeltaT = 21.2 + 0.84493 * c - 0.0761 * c ^ 2 + 0.0020936 * c ^ 3
     End If
' untuk tahun 1941 C sampai 1961 C.
     If (myYear > 1941 And myYear <= 1961) Then
        c = myYear - 1950
        DeltaT = 29.07 + 0.407 * c - c ^ 2 / 233 + c ^ 3 / 2547
     End If
     
' untuk tahun 1961 C sampai 1986 C.
     If (myYear > 1961 And myYear <= 1986) Then
        c = myYear - 1975
        DeltaT = 45.45 + 1.067 * c - c ^ 2 / 260 - c ^ 3 / 718
     End If
' untuk tahun 1986 C sampai 2005 C.
     If (myYear > 1986 And myYear <= 2005) Then
        c = myYear - 2000
        DeltaT = 63.86 + 0.3345 * c - 0.060374 * c ^ 2 + 0.0017275 * c ^ 3 + 0.000651814 * c ^ 4 + 0.00002373599 * c ^ 5
     End If
' untuk tahun 2005 C sampai 2050 C.
     If (myYear > 2005 And myYear <= 2050) Then
        c = myYear - 2000
        DeltaT = 62.92 + 0.32217 * c + 0.005589 * c ^ 2
     End If
     
' untuk tahun 2050 C sampai 2150 C.
     If (myYear > 2050 And myYear <= 2150) Then
        c = (myYear - 1820) / 100
        DeltaT = -20 + 32 * c ^ 2 - 0.5628 * (2150 - myYear)
     End If
     
' untuk tahun 2150 C dan berikutnya.
     If (myYear > 2150) Then
        c = (myYear - 1820) / 100
        DeltaT = -20 + 32 * c ^ 2
     End If
End Function

Public Function GreenwichMeanSiderealTime(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function GreenwichMeanSiderealTime dimaksudkan untuk menghitung Greenwich Mean Sidereal Time.
' Waktu Sideris Rata-rata di Greenwich(q0), pada saat jam 0 Waktu Universal pada tanggal yang
' dimaksudakan dapat diperoleh sebagai berikut (rumus ini diadopsi IAU /International
' Astronomical Union pada tahun 1982) :
'
' GMST = 280.46061837 + 360.98564736629 * (JD - 2451545) + 0.000387933 * T ^ 2 - T ^ 3 / 38710000
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Greenwich Sidereal Time dalam jam
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Greenwich Sidereal Time Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi GreenwichMeanSiderealTime.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = GreenwichMeanSiderealTime(Y, M, D)
'    Hasil:
'         x = 12.9760819006711
'           = 12:58:34
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 11. hal. 83-84
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim JD, T, GMST As Double
   
   JD = JulianDatum(Year, Month, Day)
   T = JulianEphemerisCentury(Year, Month, Day)
  
   GMST = 280.46061837 + 360.98564736629 * (JD - 2451545) + 0.000387933 * T ^ 2 - T ^ 3 / 38710000
   GMST = Modulus(GMST, 360#) / 15#
   GreenwichMeanSiderealTime = GMST
End Function

Public Function GreenwichHourAngle(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function GreenwichHourAngle dimaksudkan untuk menghitung Greenwich Hour Angle
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Greenwich Hour Angle dalam jam
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Greenwich Sidereal Time Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi GreenwichHourAngle.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = GreenwichHourAngle(Y, M, D)
'    Hasil:
'         x = 12.9946749514878
'           = 12:59:41
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 11. hal. 84
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Eps, dPsi, GHA, test As Double
  
  Eps = ObliquityOfEcliptic(Year, Month, Day)
  dPsi = NutationInLongitude(Year, Month, Day) / 3600#
  test = GreenwichMeanSiderealTime(Year, Month, Day)
  GHA = GreenwichMeanSiderealTime(Year, Month, Day) + dPsi * Cos(Radians(Eps)) / 15#
  GreenwichHourAngle = Modulus(GHA, 24)
End Function

Public Function ApparentLocalSiderealTime(ByVal Bujur As Double, ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function ApparentLocalSiderealTime dimaksudkan untuk menghitung Apparent Local Sidereal Time
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Local Sidereal Time dalam jam
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Local Sidereal Time Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' di Jakarta dengan koordinat (6° 8' S 106° 45'E) Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Menyiapkan Bujur Tempat (Observer)
'    3. Jalankan fungsi ApparentLocalSiderealTime.
'           B = 106. + 45./60.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentLocalSiderealTime(B, Y, M, D)
'    Hasil:
'         x = 20.0927485673378
'           = 20:05:34
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 11. hal. 84
'----------------------------
'*Demak, 18 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  ApparentLocalSiderealTime = Modulus(GreenwichHourAngle(Year, Month, Day) + Bujur / 15#, 24#)
End Function

Public Function ArahQibla(ByVal Lintang As Double, ByVal Bujur As Double) As Double
' Fungsi Arah Kiblat dimaksudkan untuk menghitung arah kiblat
' berdasarkan rumus segitiga bola (Spherical Trigonometry).

' INPUT:
'    Lintang : koordinat lintang suatu tempat dalam koordinat geografik.
'    Bujur   : Koordinat Bujur suatu tempat dalam koordinat geografik
'
' OUTPUT:
'    Arah Kiblat sebagai sudut arah/azimuth dari arah utara searah
'    jarum jam dalam satuan derajat
'
' DATA EMBEDDED dalam fungsi ini:
'    Koordinat Ka'bah diperoleh dari GoogleEarth :
'      Lintang Ka'bah : 21° 25' 21.03"
'      Bujur Ka'bah   : 39° 49' 34.31"
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Arah Kiblat dari Jakarta dengan koordinat (6° 8' S 106° 45'E)
' dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur,
'    2. Jalankan fungsi ArahQibla.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           x = ArahQibla(Lintang, Bujur)
'    Hasil:
'         x = 295.152500797225
'           = 295° 09' 09.003"
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985

' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' 15 Juli 2010

    Dim A, B, c, E As Double
    Dim sB, cB, Bb As Double
    Dim LintangKabah, BujurKabah As Double
    
' Koordinat Kabah
    LintangKabah = 21 + 25 / 60 + 21.03 / 3600
    BujurKabah = 39 + 49 / 60 + 34.31 / 3600
   
' Perhitungan arah dengan rumus segitiga Bola

    A = Radians(90# - Lintang)
    B = Radians(90# - LintangKabah)
    c = Radians(Bujur - BujurKabah)
    sB = Sin(B) * Sin(c)
    cB = Cos(B) * Sin(A) - Cos(A) * Sin(B) * Cos(c)
    Bb = Atn2(cB, sB)
    
' Arah Kiblat diukur dari arah utara searah jarum jam
' Sudut bernilai positif antara 0 sampai 360 derajat
    
    ArahQibla = 360# - Degrees(Bb)
    Do Until (ArahQibla < 360#)
      ArahQibla = ArahQibla - 360#
    Loop
    Do Until (ArahQibla > 0#)
      ArahQibla = ArahQibla + 360#
    Loop
End Function

Public Function JarakQibla(ByVal Lintang As Double, ByVal Bujur As Double) As Double
' Fungsi JarakQibla dimaksudkan untuk menghitung Jarak suatu Tempat ke ka'bah
' berdasarkan rumus segitiga bola (Spherical Trigonometry).
'
' INPUT:
'    Lintang : koordinat lintang suatu tempat dalam koordinat geografik.
'    Bujur   : Koordinat Bujur suatu tempat dalam koordinat geografik
'
' OUTPUT:
'    Jarak ke Kabah dinyatakan dalam km
'
' DATA EMBEDDED dalam fungsi ini:
'    Koordinat Ka'bah diperoleh dari GoogleEarth :
'      Lintang Ka'bah : 21° 25' 21.03"
'      Bujur Ka'bah   : 39° 49' 34.31"
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Jarak ke Kabah dari Jakarta dengan koordinat (6° 8' S 106° 45'E)
' dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur,
'    2. Jalankan fungsi JarakQibla.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           x = JarakQibla(Lintang, Bujur)
'    Hasil:
'         x = 7907.16424875021 km
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985

' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
'* 15 Juli 2010

    Dim phi1, phi2, dL As Double
    Dim LintangKabah, BujurKabah As Double
    
' Koordinat Kabah
    LintangKabah = 21 + 25 / 60 + 21.03 / 3600
    BujurKabah = 39 + 49 / 60 + 34.31 / 3600
   
' Perhitungan arah dengan rumus segitiga Bola

    phi1 = Radians(Lintang)
    phi2 = Radians(LintangKabah)
    dL = Radians(Bujur - BujurKabah)
    JarakQibla = Acos(Sin(phi1) * Sin(phi2) + Cos(phi1) * Cos(phi2) * Cos(dL)) * 6371.137
End Function

Public Function ArahQiblaWithEllipsoidCorrection(ByVal Lintang As Double, ByVal Bujur As Double) As Double
' Fungsi ArahQiblaWithEllipsoidCorrection dimaksudkan untuk menghitung Jarak suatu Tempat ke ka'bah
' berdasarkan rumus segitiga bola (Spherical Trigonometry) dengan koreksi dari koordinat geografik ke
' ke koordinat geosentrik.
'
' INPUT:
'    Lintang : koordinat lintang suatu tempat dalam koordinat geografik.
'    Bujur   : Koordinat Bujur suatu tempat dalam koordinat geografik
'
' OUTPUT:
'    Arah ke Qibla dinyatakan dalam derajat
'
' DATA EMBEDDED dalam fungsi ini:
'    Koordinat Ka'bah diperoleh dari GoogleEarth :
'      Lintang Ka'bah : 21° 25' 21.03"
'      Bujur Ka'bah   : 39° 49' 34.31"
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Jarak ke Kabah dari Jakarta dengan koordinat (6° 8' S 106° 45'E)
' dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur,
'    2. Jalankan fungsi ArahQiblaWithEllipsoidCorrection.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           x = ArahQiblaWithEllipsoidCorrection(Lintang, Bujur)
'    Hasil:
'         x = 295.006428253695
'           = 295° 00' 23.142"
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985

' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
'* 15 Juli 2010

    Dim A, B, c, E As Double
    Dim LintangTempatTerkoreksi As Double
    Dim LintangKabahTerkoreksi As Double
    Dim sB, cB, Bb, ArahQibla As Double
    Dim LintangKabah, BujurKabah As Double
    
' Koordinat Kabah
    LintangKabah = 21 + 25 / 60 + 21.03 / 3600
    BujurKabah = 39 + 49 / 60 + 34.31 / 3600
   
' Koreksi koordinat dari geografik (sistem bumi sebagai ellipsoid)
' ke koordinat geosentrik (sistem bumi sebagai bola). Koreksi
' dilakukan dengan menggunakan parameter WGS-84. Parameter ini biasa
' dipakai dalam pengukuran GPS (Global Positioning System)
' excentrisitas (e) =   6.6943800229e-3
    E = 0.0066943800229
    LintangKabahTerkoreksi = Degrees(Atn((1 - E) * Tan(Radians(LintangKabah))))
    LintangTempatTerkoreksi = Degrees(Atn((1 - E) * Tan(Radians(Lintang))))

   
' Perhitungan arah dengan rumus segitiga Bola

    A = Radians(90# - LintangTempatTerkoreksi)
    B = Radians(90# - LintangKabahTerkoreksi)
    c = Radians(Bujur - BujurKabah)
    
    
    sB = Sin(B) * Sin(c)
    cB = Cos(B) * Sin(A) - Cos(A) * Sin(B) * Cos(c)
    Bb = Atn2(cB, sB)
    
' Arah Kiblat diukur dari arah utara searah jarum jam
' Sudut bernilai positif antara 0 sampai 360 derajat
    
    ArahQibla = 360# - Degrees(Bb)
    Do Until (ArahQibla < 360#)
      ArahQibla = ArahQibla - 360#
    Loop
    Do Until (ArahQibla > 0#)
      ArahQibla = ArahQibla + 360#
    Loop
    ArahQiblaWithEllipsoidCorrection = ArahQibla
End Function

Public Function JarakQiblaWithEllipsoidCorrection(ByVal Lintang As Double, ByVal Bujur As Double) As Double
' Fungsi JarakQiblaWithEllipsoidCorrection dimaksudkan untuk menghitung Jarak suatu Tempat ke ka'bah
' berdasarkan rumus segitiga bola (Spherical Trigonometry) dengan koreksi ellipsoid untuk konversi
' dari koordinat geografik ke koordinat geosentrik.
'
' INPUT:
'    Lintang : koordinat lintang suatu tempat dalam koordinat geografik.
'    Bujur   : Koordinat Bujur suatu tempat dalam koordinat geografik
'
' OUTPUT:
'    Jarak ke Kabah dinyatakan dalam km
'
' DATA EMBEDDED dalam fungsi ini:
'    Koordinat Ka'bah diperoleh dari GoogleEarth :
'      Lintang Ka'bah : 21° 25' 21.03"
'      Bujur Ka'bah   : 39° 49' 34.31"
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Jarak ke Kabah dari Jakarta dengan koordinat (6° 8' S 106° 45'E)
' dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur,
'    2. Jalankan fungsi JarakQiblaWithEllipsoidCorrection.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           x = JarakQiblaWithEllipsoidCorrection(Lintang, Bujur)
'    Hasil:
'         x = 7909.09859810563 km
'
' Referensi :
'    Astronomical Algorithm, Jean Meeus, hal. 81
'
' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
'* 15 Juli 2010
    Dim fo, F, G, L, S, c, w, R, D, A, H1, H2 As Double
    Dim LintangKabah, BujurKabah As Double
    
' Koordinat Kabah
    LintangKabah = 21 + 25 / 60 + 21.03 / 3600
    BujurKabah = 39 + 49 / 60 + 34.31 / 3600
    
' Berdasarkan WGS-84, semi major axis bumi (a): 6371.137 km
' Penggepengan (Flattening) : fo = 1 / 298.257223563
    
     
    A = 6378.14
    fo = 1 / 298.257
    F = (Lintang + LintangKabah) / 2#
    G = (Lintang - LintangKabah) / 2#
    L = (Bujur - BujurKabah) / 2#
    S = (Sin(Radians(G)) * Cos(Radians(L))) ^ 2 + (Cos(Radians(F)) * Sin(Radians(L))) ^ 2
    c = (Cos(Radians(G)) * Cos(Radians(L))) ^ 2 + (Sin(Radians(F)) * Sin(Radians(L))) ^ 2
    w = Atn(Sqr(S / c))
    R = Sqr(S * c) / w
    D = 2 * w * A
    H1 = (3 * R - 1) / (2 * c)
    H2 = (3 * R + 1) / (2 * S)
    JarakQiblaWithEllipsoidCorrection = D * (1 + fo * H1 * (Sin(Radians(F)) * Cos(Radians(G))) ^ 2 - _
                                        fo * H2 * (Cos(Radians(F)) * Sin(Radians(G))) ^ 2)
End Function

Public Function ArahQiblaVincenty(ByVal Lintang As Double, ByVal Bujur As Double) As Double
' Fungsi ArahQiblaVincenty dimaksudkan untuk menghitung Arah suatu Tempat ke ka'bah
' berdasarkan rumus Vincenty (Geodetic Approach).

' INPUT:
'    Lintang : koordinat lintang suatu tempat dalam koordinat geografik.
'    Bujur   : Koordinat Bujur suatu tempat dalam koordinat geografik
'
' OUTPUT:
'    Arah Kiblat dinyatakan dalam derajat
'
' DATA EMBEDDED dalam fungsi ini:
'    Koordinat Ka'bah diperoleh dari GoogleEarth :
'      Lintang Ka'bah : 21° 25' 21.03"
'      Bujur Ka'bah   : 39° 49' 34.31"
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Arah ke Kiblat dari Jakarta dengan koordinat (6° 8' S 106° 45'E)
' dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur,
'    2. Jalankan fungsi ArahQiblaVincenty.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           x = ArahQiblaVincenty(Lintang, Bujur)
'    Hasil:
'         x = 295.025569375748
'           = 295° 01' 32.050"
'
' Referensi :
'   Vincenty, T. (April 1975). "Direct and Inverse Solutions of Geodesics
'   on the Ellipsoid with application of nested equations". Survey Review XXIII
'   (misprinted as XXII) (176): 88–93. http://www.ngs.noaa.gov/PUBS_LIB/inverse.pdf.
'   Retrieved 2009-07-11.
' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
'* 15 Juli 2010

    Dim S, A, B, c, E, F As Double
    Dim LintangTempatTerkoreksi As Double
    Dim LintangKabahTerkoreksi, c2SigmaM, dSigma As Double
    Dim sigma, sAlpha, cAlpha, c2Alpha, c2sigma, Bb As Double
    Dim ae, be, U1, U2, L, Lambda0, Lambda, sSigma, cSigma As Double
    Dim iterlimit As Integer
    Dim LintangKabah, BujurKabah, up2, alpha1, alpha2 As Double
    
' Koordinat Kabah
    LintangKabah = 21 + 25 / 60 + 21.03 / 3600
    BujurKabah = 39 + 49 / 60 + 34.31 / 3600
   
    F = 1# / 298.257223563
    ae = 6378137#
    be = 6356752.314
    
'  LintangKabah = -(37 + 57 / 60# + 3.7203 / 3600#)
'  BujurKabah = 144 + 25 / 60# + 29.5244 / 3600#
'  Lintang = -(37 + 39 / 60# + 10.1561 / 3600#)
'  Bujur = 143 + 55 / 60# + 35.3839 / 3600#

' Koreksi koordinat dari geografik (sistem bumi sebagai ellipsoid)
' ke koordinat geosentrik (sistem bumi sebagai bola). Koreksi
' dilakukan dengan menggunakan parameter WGS-84. Parameter ini biasa
' dipakai dalam pengukuran GPS (Global Positioning System)
' excentrisitas (e) =   6.6943800229e-3
    
    U1 = Atn((1# - F) * Tan(Radians(LintangKabah)))
    U2 = Atn((1# - F) * Tan(Radians(Lintang)))
    L = Radians(Bujur - BujurKabah)
   
    Lambda = L
    Lambda0 = 500#
    iterlimit = 0
    
  
    Do Until (Abs(Lambda0 - Lambda) < 0.000000000001 Or iterlimit > 100)
      iterlimit = iterlimit + 1
      Debug.Print Abs(Lambda0 - Lambda)
      Lambda0 = Lambda
      sSigma = Sqr((Cos(U2) * Sin(Lambda)) ^ 2 + (Cos(U1) * Sin(U2) - Sin(U1) * Cos(U2) * Cos(Lambda)) ^ 2)
      cSigma = Sin(U1) * Sin(U2) + Cos(U1) * Cos(U2) * Cos(Lambda)
      sigma = Atn2(cSigma, sSigma)
    
      sAlpha = Cos(U1) * Cos(U2) * Sin(Lambda) / sSigma
      c2Alpha = 1 - sAlpha ^ 2
      c2SigmaM = cSigma - 2 * Sin(U1) * Sin(U2) / c2Alpha
      c = F / 16# * c2Alpha * (4 + F * (4 - 3 * c2Alpha))
      Lambda = L + (1 - c) * F * sAlpha * (sigma + c * sSigma * (c2SigmaM + c * cSigma * (-1 + 2 * c2SigmaM ^ 2)))
    Loop
    
    Debug.Print Degrees(Lambda - L)
  
    up2 = c2Alpha * (ae ^ 2 - be ^ 2) / be ^ 2
    A = 1 + up2 / 16384 * (4096 + up2 * (-768 + up2 * (320 - 175 * up2)))
    B = up2 / 1024 * (256 + up2 * (-128 + up2 * (74 - 47 * up2)))
    
    dSigma = B * sSigma * (c2SigmaM + 0.25 * B * (cSigma * (-1 + 2 * c2SigmaM ^ 2) - 1# / 6# * B * c2SigmaM * (-3 + 4 * sSigma ^ 2) * (-3 + 4 * c2SigmaM ^ 2)))
    S = be * A * (sigma - dSigma)
    alpha1 = Atn2(Cos(U1) * Sin(U2) - Sin(U1) * Cos(U2) * Cos(Lambda), Cos(U2) * Sin(Lambda))
    alpha2 = Atn2(-Sin(U1) * Cos(U2) + Cos(U1) * Sin(U2) * Cos(Lambda), Cos(U1) * Sin(Lambda))
     
    ArahQiblaVincenty = -180# + Degrees(alpha2)
    ArahQiblaVincenty = Modulus(ArahQiblaVincenty, 360#)
End Function

Public Function JarakQiblaVincenty(ByVal Lintang As Double, ByVal Bujur As Double) As Double
' Fungsi JarakQiblaVincenty dimaksudkan untuk menghitung Jarak suatu Tempat ke ka'bah
' berdasarkan rumus Vincenty (Geodetic Approach).

' INPUT:
'    Lintang : koordinat lintang suatu tempat dalam koordinat geografik.
'    Bujur   : Koordinat Bujur suatu tempat dalam koordinat geografik
'
' OUTPUT:
'    Jarak ke Kabah dinyatakan dalam kilometers
'
' DATA EMBEDDED dalam fungsi ini:
'    Koordinat Ka'bah diperoleh dari GoogleEarth :
'      Lintang Ka'bah : 21° 25' 21.03"
'      Bujur Ka'bah   : 39° 49' 34.31"
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Arah ke Kiblat dari Jakarta dengan koordinat (6° 8' S 106° 45'E)
' dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur,
'    2. Jalankan fungsi JarakQiblaVincenty.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           x = JarakQiblaVincenty(Lintang, Bujur)
'    Hasil:
'         x = 7909.12506053699 km
'
' Referensi :
'   Vincenty, T. (April 1975). "Direct and Inverse Solutions of Geodesics
'   on the Ellipsoid with application of nested equations". Survey Review XXIII
'   (misprinted as XXII) (176): 88–93. http://www.ngs.noaa.gov/PUBS_LIB/inverse.pdf.
'   Retrieved 2009-07-11.
' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
'* 15 Juli 2010

    Dim S, A, B, c, E, F As Double
    Dim LintangTempatTerkoreksi As Double
    Dim LintangKabahTerkoreksi, c2SigmaM, dSigma As Double
    Dim sigma, sAlpha, cAlpha, c2Alpha, c2sigma, Bb As Double
    Dim ae, be, U1, U2, L, Lambda0, Lambda, sSigma, cSigma As Double
    Dim iterlimit As Integer
    Dim LintangKabah, BujurKabah, up2, alpha1, alpha2 As Double
    
' Koordinat Kabah
    LintangKabah = 21 + 25 / 60 + 21.03 / 3600
    BujurKabah = 39 + 49 / 60 + 34.31 / 3600
    
    F = 1# / 298.257223563
    ae = 6378137#
    be = 6356752.314
    
' Koreksi koordinat dari geografik (sistem bumi sebagai ellipsoid)
' ke koordinat geosentrik (sistem bumi sebagai bola). Koreksi
' dilakukan dengan menggunakan parameter WGS-84. Parameter ini biasa
' dipakai dalam pengukuran GPS (Global Positioning System)
' excentrisitas (e) =   6.6943800229e-3
    
    U1 = Atn((1# - F) * Tan(Radians(LintangKabah)))
    U2 = Atn((1# - F) * Tan(Radians(Lintang)))
    L = Radians(Bujur - BujurKabah)
   
    Lambda = L
    Lambda0 = 500#
    iterlimit = 0
    
  
    Do Until (Abs(Lambda0 - Lambda) < 0.000000000001 Or iterlimit > 100)
      iterlimit = iterlimit + 1
      Debug.Print Abs(Lambda0 - Lambda)
      Lambda0 = Lambda
      sSigma = Sqr((Cos(U2) * Sin(Lambda)) ^ 2 + (Cos(U1) * Sin(U2) - Sin(U1) * Cos(U2) * Cos(Lambda)) ^ 2)
      cSigma = Sin(U1) * Sin(U2) + Cos(U1) * Cos(U2) * Cos(Lambda)
      sigma = Atn2(cSigma, sSigma)
    
      sAlpha = Cos(U1) * Cos(U2) * Sin(Lambda) / sSigma
      c2Alpha = 1 - sAlpha ^ 2
      c2SigmaM = cSigma - 2 * Sin(U1) * Sin(U2) / c2Alpha
      c = F / 16# * c2Alpha * (4 + F * (4 - 3 * c2Alpha))
      Lambda = L + (1 - c) * F * sAlpha * (sigma + c * sSigma * (c2SigmaM + c * cSigma * (-1 + 2 * c2SigmaM ^ 2)))
    Loop
  
    up2 = c2Alpha * (ae ^ 2 - be ^ 2) / be ^ 2
    A = 1 + up2 / 16384 * (4096 + up2 * (-768 + up2 * (320 - 175 * up2)))
    B = up2 / 1024 * (256 + up2 * (-128 + up2 * (74 - 47 * up2)))
    
    dSigma = B * sSigma * (c2SigmaM + 0.25 * B * (cSigma * (-1 + 2 * c2SigmaM ^ 2) - 1# / 6# * B * c2SigmaM * (-3 + 4 * sSigma ^ 2) * (-3 + 4 * c2SigmaM ^ 2)))
    S = be * A * (sigma - dSigma)
    JarakQiblaVincenty = S / 1000#
End Function

Public Function ShadowRatioVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
                                          ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function ShadowRatioVSOP87 dimaksudkan untuk menghitung Perbandingan Panjang Bayang-bayang Matahari dan Tinggi Benda
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal (Waktu Universal)
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Rasio Bayang-bayang Matahari dalam %
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Rasio Bayang-bayang Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' di Jakarta dengan koordinat (6° 8' S 106° 45'E) dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ShadowRatioVSOP87.
'           L = -6. - 8./60.
'           B = 106. + 45./60.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ShadowRatioVSOP87(L, B, Y, M, D)
'    Hasil:
'         x = 0.466957782361196
'           = 46.70%
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985, hal. 12
'----------------------------
'* Demak, 18 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim hMatahari As Double
   
   hMatahari = ApparentSunAltitudeVSOP87(Lintang, Bujur, Year, Month, Day)
   If hMatahari < 0# Then
      ShadowRatioVSOP87 = 999#
   Else
      ShadowRatioVSOP87 = 1# / Tan(Radians(hMatahari))
   End If
End Function

Public Function ShadowDirectionVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
                                      ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function ShadowDirection dimaksudkan untuk menghitung Arah bayang-bayang Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Arah bayang-bayang Matahari dinyatakan dalam derajat diukur searah jaram jam dari arah Utara
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Azimuth Matahari di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tanggal 12 Februari 2010, jam 10:30:15 WIB dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ShadowDirectionVSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           L = -6. - 8./60.
'           B = 106. + 45./60.
'           x = ShadowDirectionVSOP87(L, B, Y, M, D)
'    Hasil:
'         x = 289.567213717889
'           = 289° 34' 1.97"
'
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'------------------------------ -------------------
     ShadowDirectionVSOP87 = Modulus(SunAzimuthVSOP87(Lintang, Bujur, Year, Month, Day) + 180#, 360#)
End Function

Public Function RashdulQibla(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                             ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Double, _
                             Optional ByVal Opsi As Integer = 1) As Double
' Function RashdulQibla dimaksudkan untuk menghitung saat bayang-bayang Matahari searah dengan arah Qiblat atau
' berlawanan dengan arah Qiblat.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Opsi    : Opsi perhitungan arah Qibla
'                1 : Vincenty
'                2 : Trigonometri Bola dengan koreksi Ellipsoid
'                3 : Murni Trigonometri Bola
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Rashul Qibla dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Rashdul Qibla di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tanggal 12 Februari 2010, jam 10:30:15 WIB dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi RashdulQibla.
'           L = -6. - 8./60.
'           B = 106. + 45./60.
'           TZ= 7.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = RashdulQibla(L, B, TZ, Y, M, D)
'    Hasil:
'         x = 10.9445741639097
'           = 10:56:40 WIB
'
'----------------------------
'*Singapura, 13 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Direction As Integer
   Dim delta, delta_, inv, RashdulQiblaOld, La, D, E, Qb, mDay, U, T, Az As Double
   
   La = Radians(Lintang)
   If Opsi = 1 Then Qb = ArahQiblaVincenty(Lintang, Bujur)
   If Opsi = 2 Then Qb = ArahQiblaWithEllipsoidCorrection(Lintang, Bujur)
   If Opsi = 3 Then Qb = ArahQibla(Lintang, Bujur)
   mDay = Day
   
   Do
     RashdulQiblaOld = RashdulQibla
     D = Radians(ApparentSunDeclinationVSOP87(Year, Month, mDay))
     E = EquationOfTime(Year, Month, mDay)
   
     If Qb > 180# Then
       U = Atn(1# / (Tan(Radians(360# - Qb)) * Sin(La)))
       T = Degrees(Acos(Tan(D) * Cos(U) / Tan(La)) + U) / 15#
       RashdulQibla = 12 - EquationOfTime(Year, Month, mDay) + T + kwd(TZ, Bujur)
     Else
       U = Atn(1# / (Tan(Radians(Qb)) * Sin(La)))
       T = Degrees(Acos(Tan(D) * Cos(U) / Tan(La)) + U) / 15#
       RashdulQibla = 12 - EquationOfTime(Year, Month, mDay) - T + kwd(TZ, Bujur)
     End If
     mDay = Int(Day) + (RashdulQibla - TZ) / 24#
   Loop Until Abs(RashdulQibla - RashdulQiblaOld) < (1# / 3600#)

   Direction = 0
   delta = (ShadowDirectionVSOP87(Lintang, Bujur, Year, Month, mDay) - Qb) * 3600#
   
   If Abs(delta) > 5# Then
      Direction = 1
      delta = (ShadowDirectionVSOP87(Lintang, Bujur, Year, Month, mDay) + 180# - Qb) * 3600#
   End If
   
   While Abs(delta) > 1#
      inv = 20# / 86400#
      delta_ = (ShadowDirectionVSOP87(Lintang, Bujur, Year, Month, mDay + inv) + 180# * Direction - Qb) * 3600#
      mDay = mDay + inv * delta / (delta - delta_)
      delta = (ShadowDirectionVSOP87(Lintang, Bujur, Year, Month, mDay) + 180# * Direction - Qb) * 3600#
   Wend
   RashdulQibla = Modulus(mDay * 24# + TZ, 24#)
   If ApparentSunAltitudeVSOP87(Lintang, Bujur, Year, Month, mDay) < 0# Then RashdulQibla = 999#
End Function

Public Function EquationOfTime(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function EquationOfTime dimaksudkan untuk menghitung Perata Waktu pada hari yang diinginkan.
' Equation of time atau perata adalah selisih antara waktu kulminasi matahari hakiki dengan waktu matahari ratarata.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
'
' OUTPUT
'    Perata Waktu dalam satuan jam
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Langkah-langkah  yang harus dipersiapkan untuk menghitung Perata Waktu adalah sebagai berikut:
'    1.  Mempersiapkan input berupa tahun, bulan, tanggal dengan sistem waktu Universal
'        Apabila kita menggunakan waktu lokal harus dikonversikan terlebih dahulu ke
'        waktu Universal dengan cara mengkoreksinya dengan Zona Waktu.
'    2.  Hitung Waktu Perata Waktu dengan fungsi EquationOfTime dengan cara sebagai berikut:
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'
'           E = EquationOfTime(Tahun, Bulan, Tanggal)
'
'    Hasil:
'           E = -0.236737399110455
'
' Catatan:
'    - Perata waktu berubah secara perlahan dana mempunyai periode satu tahun.
'      Untuk perhitungan jadwal shalat dengan ketelitian tinggi, perata waktu haruslah
'      dihitung pada saat waktu shalat tersebut.
'----------------------------
' Yogyakarta, 04 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Lambda, alpha, dPsi, Eps As Double

   Lambda = SunMeanLongitude(Year, Month, Day)
   alpha = ApparentSunRightAscensionVSOP87(Year, Month, Day)
   dPsi = NutationInLongitude(Year, Month, Day) / 3600#
   Eps = ObliquityOfEcliptic(Year, Month, Day)
   EquationOfTime = ((Lambda - 0.0057183 - alpha) + dPsi * Cos(Radians(Eps))) / 15#
   If (EquationOfTime > 23#) Then EquationOfTime = EquationOfTime - 24#
End Function


Public Function ShubuhTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                           ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                           Optional ByVal Alt As Double = -20#, Optional ByVal Ihtiyat As Double = 2#) As Double
' Function ShubuhTime adalah duplikat dari fungsi FajrTime
' Cara Pemakaiannya, lihat Fungsi FajrTime
'----------------------------
'* Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    ShubuhTime = FajrTime(Lintang, Bujur, TZ, Year, Month, Day, Alt, Ihtiyat)
End Function

Public Function FajrTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                         ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                         Optional ByVal Alt As Double = -20#, Optional ByVal Ihtiyat As Double = 2#) As Double
' Function FajrTime dimaksudkan untuk menghitung Waktu Fajr atau Shubuh berdasarkan opsi yang diinginkan.
' Opsi tersebut dinyatakan dalam bentuk posisi matahari di bawah ufuk dan tambahan ihtiyatnya.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Alt     : Posisi Matahari terhadap ufuk dalam derajat busur.
'                 Berikut adalah Opsi yang biasa dipakai
'                 - Kementerian Agama RI    : -20.
'                 - Ummul Qura, Makkah      : -18.5
'                 - UIS, Pakistan           : -18
'                 - ISNA, USA               : -15.
'                 - Muslim World League     : -18.
'                 - Egas Egypt              : -19.5
'                 - UOIF, France            : -12.
'                 - Custom                  : Pemakai bisa mendefinisikan sendiri sesuai dengan
'                                             kriteria yang diikutinya.
'                Jika Alt nilainya positif, maka yang dimaksudkan adalah waktu Fajar
'                (Alt) menit sebelum matahari terbit
'      ihtiyat : kehati-hatian dalam menit, untuk mengantisipasi pemberlakuan jadwal shalat untuk
'                daerah dengan radius tertentu dan kehati-hatian kemungkinan kesalahan perhitungan
'
' OUTPUT
'    FajrTime dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung waktu Fajr atau Shalat Shubuh di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Kriteria Fajr atau waktu Shalat Shubuh, misal matahari 20 derajat di bawah ufuk dan ihtiyat
'    3. Jalankan fungsi FajrTime.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           Alt = -20.
'           Ihtiyat = 2.
'           Shubuh = FajrTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt, ihtiyat)
'
'           Karena Alt = -20 (default) dan ihtiyat= 2 (defaut), maka dapat ditulis dengan singkat sbb:
'           Shubuh = FajrTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'           shubuh = 4.66150377132652
'                  = 04:39:41 WIB
'
' Catatan:
'    - Untuk perhitungan waktu Fajr berdasarkan selisih waktu dengan saat matahari terbit
'      dilakukan dengan parameter tekanan udara 1010 mbar dan suhu 27 derajat.
'----------------------------
'* Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim mDay, Terbit, E, h, hSun As Double

   FajrTime = 999#
   mDay = Day
   If Alt < 0# Then
     E = EquationOfTime(Year, Month, mDay)
     h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, mDay)
     If h <> 999# Then FajrTime = 12 - E - h + kwd(TZ, Bujur) + Ihtiyat / 60#
   Else
     hSun = SunAltitudeAtSunsetOrSunrise(Year, Month, mDay)
     Terbit = Sunrise(Lintang, Bujur, TZ, Year, Month, mDay, hSun)
     If Terbit <> 999# Then FajrTime = Terbit + Alt / 60#
   End If

   If FajrTime = 999# Then Exit Function
   mDay = Int(Day) + (FajrTime - TZ) / 24#
   If Alt < 0# Then
     E = EquationOfTime(Year, Month, mDay)
     h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, mDay)
     If h <> 999# Then FajrTime = 12 - E - h + kwd(TZ, Bujur) + Ihtiyat / 60#
   Else
     hSun = SunAltitudeAtSunsetOrSunrise(Year, Month, mDay)
     Terbit = Sunset(Lintang, Bujur, TZ, Year, Month, Int(Day) + (Terbit - TZ) / 24#, hSun)
     If Terbit <> 999# Then FajrTime = Terbit + Alt / 60#
   End If

End Function


Public Function DuhrTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                         ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                         Optional ByVal Ihtiyat As Double = 2#) As Double
' Function DuhrTime dimaksudkan untuk menghitung Waktu Duhr yakni saat matahari transit
' di garis bujur tempat yang akan dihitung waktu duhrnya. Tinggi Matahari saat waktu duhr
' bervariasi dari hari ke hari, sehingga sebagai opsi hanya perlu tambahan ihtiyatnya.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu Asharnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu Asharnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Ihtiyat : Adalah toleransi yang diberikan, terutama untuk mengakomoir wilayah-wilayah disekitarnya.
'                Untuk perhitungan yang tidak teliti, tambahan waktu ihtiyat ini dipakai juga untuk memastikan
'                bahwa waktu shalat betul-betul telah masuk.
'
' OUTPUT
'    Waktu Shalat Duhr dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung waktu Duhr di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Lama ihtiyat yang dikehendaki
'    3. Jalankan fungsi DuhrTime.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           Ihtiyat = 2.
'           Duhr = DuhrTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt, ihtiyat)
'
'           Karena ihtiyat= 2 (defaut), maka dapat ditulis dengan singkat sbb:
'           Duhr = DuhrTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'           Duhr = 12.15337991437
'                = 12:09:12 WIB
' Catatan:
'    - Untuk perhitungan waktu Shalat Duhr, tempat-tempat yang mempunyai koordinat bujur yang sama
'      akan serempak/sama waktu shalat duhrnya.
'----------------------------
'* Yogyakarta, 04 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E As Double

   E = EquationOfTime(Year, Month, Day)
   DuhrTime = 12 - E + kwd(TZ, Bujur) + Ihtiyat / 60#

   E = EquationOfTime(Year, Month, Int(Day) + (DuhrTime - TZ) / 24#)
   DuhrTime = 12 - E + kwd(TZ, Bujur) + Ihtiyat / 60#

End Function

Public Function AltitudeOfSunAtAshr(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                                    ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                                    Optional ByVal Rasio As Double = 1#, Optional ByVal IsShadowDuhr As Boolean = True, _
                                    Optional ByVal IsRefraction As Boolean = False) As Double
' Function AltitudeOfSunAtAshr dimaksudkan untuk menghitung tinggi Matahari saat Waktu Ashr.
' Opsi tersebut dinyatakan dalam bentuk tinggi matahari di atas ufuk dan tambahan ihtiyatnya.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang     : Koordinat Lintang tempat yang dihitung waktu Asharnya
'      Bujur       : Koordinat Bujur tempat yang dihitung waktu Asharnya
'      TZ          : Zona waktu (Time Zone)
'      Year        : tahun Gregorian/Julian (tahun Masehi)
'      Month       : bulan Masehi (1-12)
'      Day         : hari dalam decimal
'      Rasio       : Rasio Panjang bayang-bayang dibanding bendanya:
'                    - Syafii  : Rasio = 1
'                    - Hanafi  : Rasio = 2
'                    - Custom  : Rasio = Bebas terserah pemakai
'      IsShadowDuhr: Adalah opsi "Apakah Bayang-bayang saat AltitudeOfSunAtAshr ikut diperhitungkan".
'                    - Jika Ya maka IsDuhr = TRUE
'                    - Jika Tidak maka IsDuhr = FALSE
'
' OUTPUT
'    Tinggi Matahari saat Waktu Ashr dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Tinggi Matahari saat Waktu Ashr di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dengan madzab Syafii dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Mempersiapkan opsi perhitungan sebagai input
'    3. Jalankan fungsi AltitudeOfSunAtAshr.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           Rasio = 1
'           IsShadowDuhr = True
'
'           AltAshr = AltitudeOfSunAtAshr(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, 1., True, False)
'
'           Karena Rasio = 1 (default) dan IsDuhr= 2 (defaut), maka dapat ditulis dengan singkat sbb:
'           AltAshr = AltitudeOfSunAtAshr(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'           AltAshr = 41.3910677670231
'                   = 41° 23' 27.84"
'
'           AltAshr = AltitudeOfSunAtAshr(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal,1., True, True)
'           AltAshr = 41.3723694586778
'                   =  41° 22' 20.53"
'
' Catatan:
'    - Untuk perhitungan tinggi Matahari saat waktu Ashr  dilakukan
'      dengan memperhitungkan rasio, bayang-bayang sama panjang dengan bendanya (Syafii) atau
'      Bayang-bayang panjangnya dua kali benda (Hanafi). Kedua opsi ini dapat dilakukan dengan
'      memperhitungkan panjang bayangan saat duhr maupun tidak.
'----------------------------
'* Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim SunDec, A As Double
   
   SunDec = Radians(ApparentSunDeclinationVSOP87(Year, Month, Day))
   A = Asin(Cos(SunDec) * Cos(Radians(Lintang)) + Sin(SunDec) * Sin(Radians(Lintang)))
   If IsRefraction Then A = A - Radians(RefractionTrueAltitude(Degrees(A), 1010, 27#) / 60#)
   If IsShadowDuhr Then
       AltitudeOfSunAtAshr = Degrees(Atn(1# / (1# / Tan(A) + Rasio)))
   Else
       AltitudeOfSunAtAshr = Degrees(Atn(1# / Rasio))
   End If
   If IsRefraction Then AltitudeOfSunAtAshr = AltitudeOfSunAtAshr - RefractionApparentAltitude(AltitudeOfSunAtAshr, 1010, 27) / 60#
End Function

Public Function AshrTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                         ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                         Optional ByVal Alt As Double = 1#, Optional ByVal Ihtiyat As Double = 2#) As Double
' Function AshrTime dimaksudkan untuk menghitung Waktu Ashr berdasarkan opsi yang diinginkan.
' Opsi tersebut dinyatakan dalam bentuk tinggi matahari di atas ufuk dan tambahan ihtiyatnya.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu Asharnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu Asharnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Alt     : Posisi Matahari di atas ufuk pada saat waktu ashar setiap harinya berubah-ubah
'                Untuk menghitung Tinggi Matahari saat Ashar, dapat dilakukan dengan fungsi
'                Jika Alt=1 maka perhitungan akan dilakukan dengan madzab Syafii
'      Ihtiyat : Adalah toleransi yang diberikan, terutama untuk mengakomoir wilayah-wilayah disekitarnya.
'                Untuk perhitungan yang tidak teliti, tambahan waktu ihtiyat ini dipakai juga untuk memastikan
'                bahwa waktu shalat betul-betul telah masuk.
'
' OUTPUT
'    Waktu Shalat Ashr dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Waktu Ashr di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dengan madzab Syafii dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Mempersiapkan opsi perhitungan sebagai input
'    3. Hitung tinggi Matahari saat Ashr, lihat AltitudeOfSunAtAshr
'    4. Jalankan fungsi AltitudeOfSunAtAshr.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           Rasio = 1
'           IsShadowDuhr = True
'           Alt = AltitudeOfSunAtAshr(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, 1., True, False)
'           ihtiyat = 2.
'
'           Ashr = AshrTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt, ihtiyat)
'
'           Karena Alt = 1 (default) dan ihtiyat= 2 (defaut), maka dapat ditulis dengan singkat sbb:
'           Ashr = AshrTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'           Ashr = 15.4088426749099
'                = 15:24:32 WIB
'
'    Jika dengan koreksi Refraksi, maka
'
'           Alt = AltitudeOfSunAtAshr(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, 1., True, True)
'           Ashr = AshrTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt, ihtiyat)
'           Ashr = 15.4101286985861
'                = 15:24:36 WIB

' Catatan:
'    - Untuk perhitungan tinggi Matahari saat waktu Ashr dilakukan
'      dengan memperhitungkan opsi, bayang-bayang sama panjang dengan bendanya (Syafii) atau
'      Bayang-bayang panjangnya dua kali benda (Hanafi). Kedua opsi ini dapat dilakukan dengan
'      memperhitungkan panjang bayangan saat duhr maupun tidak.
'----------------------------
'* Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E, h, hSun As Double

   AshrTime = 999#
     
   If Alt = 1 Then Alt = AltitudeOfSunAtAshr(Lintang, Bujur, TZ, Year, Month, Day)

   E = EquationOfTime(Year, Month, Day)
   h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, Day)
   If h <> 999# Then AshrTime = 12 - E + h + kwd(TZ, Bujur) + Ihtiyat / 60#

   If AshrTime = 999# Then Exit Function
   E = EquationOfTime(Year, Month, Int(Day) + (AshrTime - TZ) / 24#)
   h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, Int(Day) + (AshrTime - TZ) / 24#)
   If h <> 999# Then AshrTime = 12 - E + h + kwd(TZ, Bujur) + Ihtiyat / 60#

End Function

Public Function MaghribTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                            ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                            Optional ByVal Alt As Double = -1#, Optional ByVal Ihtiyat As Double = 2#) As Double
' Function MaghribTime dimaksudkan untuk menghitung Waktu Maghrib berdasarkan opsi yang diinginkan.
' Opsi tersebut dinyatakan dalam bentuk posisi matahari di bawah ufuk dan tambahan ihtiyatnya.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Alt     : Posisi Matahari terhadap ufuk dalam derajat busur.
'                Untuk perhitungan yang tidak memerlukan ketelitian tinggi, maka Alt=1
'                Untuk perhitungan yang teliti, maka Alt harus dihitung berdasarkan koreksi
'                Semidiameter dan Refraksi. lihat fungsi SunAltitudeAtSunsetOrSunrise
'
' OUTPUT
'    Waktu Shalat Maghrib dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Waktu Ashr di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Mempersiapkan opsi perhitungan sebagai input
'    3. Hitung tinggi Matahari saat Terbenam, lihat SunAltitudeAtSunsetOrSunrise
'    4. Jalankan fungsi MaghribTime.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           p =  1010.
'           T = 27.
'           Alt = SunAltitudeAtSunsetOrSunrise(Tahun, Bulan, Tanggal, p, T, h) / 60#
'           ihtiyat = 2.
'
'           Maghrib = MaghribTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt, ihtiyat)
'
'           Karena ihtiyat= 2 (defaut), maka dapat ditulis dengan singkat sbb:
'           Maghrib = MaghribTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt)
'    Hasil:
'           Maghrib = 18.3090536548075
'                   = 18:18:33
' Catatan:
'    - Untuk perhitungan waktu Maghrib yang teliti, ketinggian Matahari di bawah ufuk
'      harus diperhitungkan koreksi semidiameter matahari dan refraksi. Sedangkan koreksi ketinggian tempat
'      (DIP) dapat diterapkan, namun penulis tidak menganjurkan untuk mengkoreksi DIB karena koreksi
'      ini hanya cocok untuk tempat tinggi yang tidak terhalang ke arah ufuk di laut, misalnya
'      menara yang menjulang tinggi.
'----------------------------
'* Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim mDay, Terbenam, E, h, hSun As Double
   
   Terbenam = Sunset(Lintang, Bujur, TZ, Year, Month, Day, Alt)
   MaghribTime = Terbenam
   If MaghribTime <> 999# Then MaghribTime = MaghribTime + Ihtiyat / 60#

End Function

Public Function IsyaTime(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                         ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                         Optional ByVal Alt As Double = -18#, Optional ByVal Ihtiyat As Double = 2#) As Double

' Function IsyaTime dimaksudkan untuk menghitung Waktu Isya berdasarkan opsi yang diinginkan.
' Opsi tersebut dinyatakan dalam bentuk posisi matahari di bawah ufuk dan tambahan ihtiyatnya.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Alt     : Posisi Matahari terhadap ufuk dalam derajat busur.
'                 Berikut adalah Opsi yang biasa dipakai
'                 - Kementerian Agama RI    : -18.
'                 - Ummul Qura, Makkah      : 90.
'                 - UIS, Pakistan           : -18
'                 - ISNA, USA               : -15.
'                 - Muslim World League     : -17.
'                 - Egas Egypt              : -17.5
'                 - UOIF, France            : -12.
'                 - Custom                  : Pemakai bisa mendefinisikan sendiri sesuai dengan
'                                             kriteria yang diikutinya.
'                Jika Alt nilainya positif, maka yang dimaksudkan adalah waktu Isya adalah
'                (Alt) menit setelah matahari terbenam
'      T       : Suhu di permukaan bumi dinyatakan dalam derajat Celcius
'
' OUTPUT
'    IsyaTime dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung waktu Shalat Isya di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Kriteria Fajr atau waktu Shalat Isya, misal matahari 20 derajat di bawah ufuk dan ihtiyat
'    3. Jalankan fungsi IsyaTime.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           Alt = -18.
'           Ihtiyat = 2.
'           Isya = IsyaTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt, ihtiyat)
'
'           Karena Alt = -18 (default) dan ihtiyat= 2 (defaut), maka dapat ditulis dengan singkat sbb:
'           Isya = IsyaTime(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'           Isya = 19.502247859682
'                = 19:30:08
' Catatan:
'    - Untuk perhitungan waktu Isya berdasarkan selisih waktu dengan saat matahari terbenam
'      dilakukan dengan parameter tekanan udara 1010 mbar dan suhu 27 derajat.
'----------------------------
'* Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim mDay, Terbenam, E, h, hSun As Double

   IsyaTime = 999#
   mDay = Day
   If Alt < 0# Then
     E = EquationOfTime(Year, Month, mDay)
     h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, mDay)
     If h <> 999# Then IsyaTime = 12 - E + h + kwd(TZ, Bujur) + Ihtiyat / 60#
   Else
     hSun = SunAltitudeAtSunsetOrSunrise(Year, Month, mDay)
     Terbenam = Sunrise(Lintang, Bujur, TZ, Year, Month, mDay, hSun)
     If Terbenam <> 999# Then IsyaTime = Terbenam + Alt / 60#
   End If

   If IsyaTime = 999# Then Exit Function
   mDay = Int(Day) + (IsyaTime - TZ) / 24#
   If Alt < 0# Then
     E = EquationOfTime(Year, Month, mDay)
     h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, mDay)
     If h <> 999# Then IsyaTime = 12 - E + h + kwd(TZ, Bujur) + Ihtiyat / 60#
   Else
     hSun = SunAltitudeAtSunsetOrSunrise(Year, Month, mDay)
     Terbenam = Sunset(Lintang, Bujur, TZ, Year, Month, Int(Day) + (Terbenam - TZ) / 24#, hSun)
     If Terbenam <> 999# Then IsyaTime = Terbenam + Alt / 60#
   End If

End Function

Public Function SunMeanLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunMeanLongitude berfungsi untuk menghitung rata-rata bujur Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Rata-rata bujur matahari dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Astronomical Algorithm Example 24.a - Calculate the Sun's position on 1992 October 13
'                                       at 0h TD = JDE 2448 908.5.
'           Y = 1992
'           M = 10
'           D = 13
'           dT = DeltaT(Y, M, D)/86400.
'           x = SunMeanLongitude(Y, M, D-dT)
'    Hasil:
'         x = 201.80719336624
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 152
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim JDE, T, tau As Double
   
   JDE = JulianDatum(Year, Month, Day) + DeltaT(Year, Month, Day) / 86400#
   T = (JDE - 2451545#) / 36525#
   tau = T / 10#
   SunMeanLongitude = Modulus(280.4664567 + 360007.6982779 * tau + 0.03032028 * tau ^ 2 + _
                               tau ^ 3 / 49931 - tau ^ 4 / 15299 - tau ^ 5 / 1988000, 360#)
End Function

Public Function SunMeanAnomaly(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunMeanAnomaly berfungsi untuk menghitung Rata-rata anomaly Matahari.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Rata-rata bujur matahari dalam derajat
'
' CATATAN:
'    Dalam Buku Astronomical Algorithm disimbolkan dengan M
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung rata-rata Anomaly Matahari pada 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunMeanAnomaly(Y, M, D)
'    Hasil:
'         x = 3998.97401450417
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 308
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   SunMeanAnomaly = 357.5291092 + 35999.0502909 * T - 0.0001536 * T ^ 2 + T ^ 3 / 24490000
End Function

Public Function SunLongitudeLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunLongitudeLowAccuracy berfungsi untuk menghitung Sun's True Longitude (Bujur Ekliptik Matahari)
' dengan akurasi rendah. Perhitungan dilakukan dengan rumus yang sederhana.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Bujur Matahari dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Bujur Ekliptika Matahari pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunLongitudeLowAccuracy(Y, M, D)
'    Hasil:
'         x = 323.309014095957
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 152
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T, M, c, Lo As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   M = Radians(357.5291 + 35999.0503 * T - 0.0001559 * T ^ 2 - 0.00000048 * T ^ 3)
   c = (1.9146 - 0.004817 * T - 0.000014 * T ^ 2) * Sin(M) + (0.019993 - 0.000101 * T) * Sin(2 * M) + 0.00029 * Sin(3 * M)
   Lo = Modulus(280.46645 + 36000.76983 * T + 0.0003032 * T ^ 2, 360#)
   SunLongitudeLowAccuracy = Modulus(Lo + c, 360#)
   
End Function

Public Function SunLatitudeLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunLatitudeLowAccuracy berfungsi untuk menghitung Sun's True Latitude (Lintang Ekliptik Matahari)
' dengan akurasi rendah. Perhitungan dilakukan dengan rumus yang sederhana.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'    Lintang Matahari nilainya 0 untuk kapan-pun
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Bujur Ekliptika Matahari pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunLatitudeLowAccuracy(Y, M, D)
'    Hasil:
'         x = 0.
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, bab 24
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Ra*ya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   SunLatitudeLowAccuracy = 0#
   
End Function

Public Function DistanceToTheSunLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  The Sun 's radius vector, or the distance from the Earth to the Sun, expressed in astronomical units,
'  is given by R = 1.000001018 (1 - e^2) / (1 + e cos v)
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'    Jarak bumi ke Matahari dalam Astronomical Unit
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Bujur Ekliptika Matahari pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = DistanceToTheSunLowAccuracy(Y, M, D)
'    Hasil:
'         x = 0.987127076021332
'
'  Referensi
'    Astronomical Algorithm hal. 152
' -------------------------------------------------------
'*  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E, M, c, V As Double
   
   E = EccentricityOfEarthOrbit(Year, Month, Day)
   M = Radians(SunMeanAnomaly(Year, Month, Day))
   c = Radians(SunsEqCenter(Year, Month, Day))
   V = M + c
   DistanceToTheSunLowAccuracy = 1.000001018 * (1 - E ^ 2) / (1 + E * Cos(V))
End Function

Public Function ApparentSunLongitudeLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi ApparentSunLongitudeLowAccuracy dimaksudkan untuk menghitung bujur Matahari nampak
' (Apparent Sun Longitude) dengan akurasi rendah.
' Bujur Matahari Nampak adalah sama dengan Bujur ekliptik Matahari yang dikoreksi dengan
' Aberasi dan Nutasi
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Apparent Sun Longitude dinyatakan dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Apparent Sun Longitude Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentSunLongitudeLowAccuracy.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentSunLongitudeLowAccuracy(Y, M, D)
'    Hasil:
'         x = 323.303183547447
'
' ---------------------------------------------
'  Programmer : Dr.-Ing. Khafid
'  Badan Koordinasi Survei dan Pemetaan Nasional
'* 17 Juli 2010

   ApparentSunLongitudeLowAccuracy = SunLongitudeLowAccuracy(Year, Month, Day) _
                                   + SunAberration(Year, Month, Day) / 3600# _
                                   + NutationInLongitude(Year, Month, Day) / 3600#
End Function

Public Function SunLongitudeVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunLongitudeVSOP87 dimaksudkan untuk menghitung Buujur Matahari ekliptic
' (Ecliptic Sun Longitude) dengan metode VSOP87.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Bujur Matahari dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Bujur Matahari Ekliptik Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunLongitudeVSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunLongitudeVSOP87(Y, M, D)
'    Hasil:
'         x = 323.303312808408
'           = 323° 18' 11.926"
'
' Referensi VSOP87
' ---------------------------------------------
'  Programmer : Dr.-Ing. Khafid
'  Badan Koordinasi Survei dan Pemetaan Nasional
'* 18 Juni 2011
'
   
   Dim T, tau, Theta, Beta, lambda1, dBeta, dTheta As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   tau = T / 10#
  
' Perhitungan Bujur Matahari dengan menggunakan rumus VSOP87
   Theta = L0(tau) + L1(tau) * tau + L2(tau) * tau ^ 2 + L3(tau) * tau ^ 3 + L4(tau) * tau ^ 4 + L5(tau) * tau ^ 5
   Beta = B0(tau) + B1(tau) * tau + B2(tau) * tau ^ 2 + B3(tau) * tau ^ 3 + B4(tau) * tau ^ 4
   Beta = -(Degrees(Beta))
   Theta = 180# + Degrees(Theta)
   
' Mencari sudut Theta di antara 0 dan 360 derajat
   Theta = Modulus(Theta, 360#)
   
' Menghitung koreksi Theta (Bujur Matahari)
' Perhitungan koreksi ini mengharuskan adanya perhitungan Lintang Matahari (Beta)
   lambda1 = Theta - 1.397 * T - 0.00031 * T ^ 2
   dBeta = 0.03916 * (Cos(Radians(lambda1)) - Sin(Radians(lambda1))) / 3600#
   Beta = Beta + dBeta
   dTheta = (-0.09033 + 0.03916 * (Cos(Radians(lambda1)) + Sin(Radians(lambda1))) * Tan(Beta)) / 3600#
   SunLongitudeVSOP87 = Modulus(Theta + dTheta, 360#)
End Function

Public Function SunLatitudeVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunLatitudeVSOP87 dimaksudkan untuk menghitung Lintang Matahari ekliptic
' (Ecliptic Sun Latitude) dengan metode VSOP87.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Lintang Matahari Ekliptik dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Lintang Ekliptik Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunLatitudeVSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunLatitudeVSOP87(Y, M, D)
'    Hasil:
'         x = -0.0000134540652726544
'
' ---------------------------------------------
'  Programmer : Dr.-Ing. Khafid
'  Badan Koordinasi Survei dan Pemetaan Nasional
'* 18 Juni 2011
'
   
   Dim DT, JD, T, tau As Double
   Dim Theta, Beta, lambda1, dBeta As Double
   
' Mengkonversikan Tanggal/Bulan/Tahun ke dalam tanggal julian (Julian Date) dan
' Julian Century (tau) serta Julian Millenea (tau) dengan epoch J2000 Januari 1.5
   DT = DeltaT(Year, Month, Day) / 86400#
   JD = JulianDatum(Year, Month, Day + DT)
   T = (JD - 2451545#) / 36525#
   tau = T / 10#
  
' Perhitungan Lintang Matahari dengan menggunakan rumus VSOP87
   Theta = L0(tau) + L1(tau) * tau + L2(tau) * tau ^ 2 + L3(tau) * tau ^ 3 + L4(tau) * tau ^ 4 + L5(tau) * tau ^ 5
   Beta = B0(tau) + B1(tau) * tau + B2(tau) * tau ^ 2 + B3(tau) * tau ^ 3 + B4(tau) * tau ^ 4
   Beta = -(Degrees(Beta))
   Theta = 180# + Degrees(Theta)
   
' Mencari sudut Theta di antara 0 dan 360 derajat
   Theta = Modulus(Theta, 360#)
   lambda1 = Theta - 1.397 * T - 0.00031 * T ^ 2
   
   dBeta = 0.03916 * (Cos(Radians(lambda1)) - Sin(Radians(lambda1))) / 3600#
   SunLatitudeVSOP87 = Beta + dBeta
End Function

Public Function DistanceToTheSunVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function DistanceToTheSunVSOP87 dimaksudkan untuk menghitung Jarak dari Bumi ke Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 13:30:15 dalam waktu Universal
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
' OUTPUT:
'    DistanceToTheSunVSOP87 dalam astronomical unit
'    i AU = 149597870.691 km
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Jarak dari Bumi ke Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Hitung Jarak Matahari ke bumi dengan fungsi DistanceToTheSunVSOP87
'       dengan cara sebagai berikut:
'           Y = 2010
'           M = 2
'           D   = 12 + (10+30/60+15/3600)/24 - 7
'           R = DistanceToTheSunVSOP87(Y, M, D)
'
'    Hasil:
'           R = 0.985967705399629
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 24
'----------------------------
'*Cibinong, 14 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim tau As Double
   
   tau = JulianEphemerisCentury(Year, Month, Day) / 10#
   DistanceToTheSunVSOP87 = R0(tau) + R1(tau) * tau + R2(tau) * tau ^ 2 + R3(tau) * tau ^ 3 + R4(tau) * tau ^ 4 + R5(tau) * tau ^ 5
End Function

Public Function ApparentSunLongitudeVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi ApparentSunLongitudeVSOP87 dimaksudkan untuk menghitung bujur Matahari nampak
' (Apparent Sun Longitude) dengan metode VSOP87.
' Bujur Matahari Nampak adalah sama dengan Bujur ekliptik Matahari yang dikoreksi dengan
' Aberasi dan Nutasi
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Apparent Sun Longitude dinyatakan dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Apparent Sun Longitude Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentSunLongitudeVSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentSunLongitudeVSOP87(Y, M, D)
'    Hasil:
'         x = 323.297482259898
'
' ---------------------------------------------
'  Programmer : Dr.-Ing. Khafid
'  Badan Koordinasi Survei dan Pemetaan Nasional
'* 17 Juli 2010

   ApparentSunLongitudeVSOP87 = SunLongitudeVSOP87(Year, Month, Day) _
                              + SunAberration(Year, Month, Day) / 3600# _
                              + NutationInLongitude(Year, Month, Day) / 3600#
End Function

Public Function SunDeclinationLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunDeclination dimaksudkan untuk menghitung Deklinasi Matahari dengan ketelitian rendah.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Deklinasi Matahari dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Deklinasi Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunDeclinationLowAccuracy.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunDeclinationLowAccuracy(Y, M, D)
'    Hasil:
'         x = -13.7489164095265
'           = -13° 44' 56.10"
'
' Referensi
'    Astronomical Algorithm
' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
'*18 Juni 2011
'
   Dim Lambda, Beta, Eps, sinDelta As Double
   
   Lambda = Radians(SunLongitudeLowAccuracy(Year, Month, Day))
   Beta = Radians(SunLatitudeLowAccuracy(Year, Month, Day))
   Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
   sinDelta = Sin(Beta) * Cos(Eps) + Cos(Beta) * Sin(Eps) * Sin(Lambda)
   SunDeclinationLowAccuracy = Degrees(Asin(sinDelta))
End Function

Public Function SunDeclinationVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunDeclination dimaksudkan untuk menghitung Deklinasi Matahari dengan ketelitian tinggi menggunakan
' metode VSOP87.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Deklinasi Matahari dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Deklinasi Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunDeclinationVSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunDeclinationVSOP87(Y, M, D)
'    Hasil:
'         x = -13.7508011821537
'           = -13° 45' 2.88"
'
' Referensi
'    Astronomical Algorithm
' ---------------------------------------------
'* Programmer : Dr.-Ing. Khafid
'  Badan Koordinasi Survei dan Pemetaan Nasional
'  18 Juni 2011
'
   Dim Lambda, Beta, Eps, sinDelta As Double
   
   Lambda = Radians(SunLongitudeVSOP87(Year, Month, Day))
   Beta = Radians(SunLatitudeVSOP87(Year, Month, Day))
   Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
   sinDelta = Sin(Beta) * Cos(Eps) + Cos(Beta) * Sin(Eps) * Sin(Lambda)
   SunDeclinationVSOP87 = Degrees(Asin(sinDelta))
End Function

Public Function ApparentSunDeclinationLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi ApparentSunDeclinationLowAccuracy dimaksudkan untuk menghitung Deklinasi Matahari Nampak dengan ketelitian rendah.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Deklinasi Matahari Nampak dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Deklinasi Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentSunDeclinationLowAccuracy.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentSunDeclinationLowAccuracy(Y, M, D)
'    Hasil:
'         x = -13.7508309164815
'           = -13° 45' 2.99"
'
' Referensi
'    Astronomical Algorithm
' ---------------------------------------------
'  Programmer : Dr.-Ing. Khafid
'  Badan Koordinasi Survei dan Pemetaan Nasional
'* 18 Juni 2011
'
   Dim Lambda, Beta, Eps, sinDelta As Double
   
   Lambda = Radians(ApparentSunLongitudeLowAccuracy(Year, Month, Day))
   Beta = Radians(SunLatitudeLowAccuracy(Year, Month, Day))
   Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
   sinDelta = Sin(Beta) * Cos(Eps) + Cos(Beta) * Sin(Eps) * Sin(Lambda)
   ApparentSunDeclinationLowAccuracy = Degrees(Asin(sinDelta))
End Function

Public Function ApparentSunDeclinationVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi ApparentSunDeclinationVSOP87 dimaksudkan untuk menghitung Deklinasi Matahari Nampak dengan
' ketelitian tinggi menggunakan metode VSOP87.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Deklinasi Matahari Nampak dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Deklinasi Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentSunDeclinationVSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentSunDeclinationVSOP87(Y, M, D)
'    Hasil:
'         x = -13.7527155625426
'           = -13° 45' 9.78
' Referensi
'    Astronomical Algorithm
' ---------------------------------------------
'  Programmer : Dr.-Ing. Khafid
'  Badan Koordinasi Survei dan Pemetaan Nasional
'* 18 Juni 2011
'
   Dim Lambda, Beta, Eps, sinDelta As Double
   
   Lambda = Radians(ApparentSunLongitudeVSOP87(Year, Month, Day))
   Beta = Radians(SunLatitudeVSOP87(Year, Month, Day))
   Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
   sinDelta = Sin(Beta) * Cos(Eps) + Cos(Beta) * Sin(Eps) * Sin(Lambda)
   ApparentSunDeclinationVSOP87 = Degrees(Asin(sinDelta))
End Function

Public Function SunRightAscensionLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunRightAscention dimaksudkan untuk menghitung askensio rekta Matahari dengan
' rumus akurasi rendah
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Askensio Rekta Matahari ( Sun Right Ascension) bujur derajat
'
' CATATAN:
'    Untuk mengkonversikan dalam satuan jam, maka nilai dalam Busur derajat tersebut
'    dibagi dengan 15.
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Askensio Rekta  Matahari terlihat Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunRightAscensionLowAccuracy.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunRightAscensionLowAccuracy(Y, M, D)
'    Hasil:
'         x = 325.641627237425
'           = 21:42:33.99
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 12. hal. 89
'------===========================----------------------
'*Yogyakarta, 04 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim sinAlpha, cosAlpha, Eps, L, B As Double
  
  B = Radians(SunLatitudeLowAccuracy(Year, Month, Day))
  L = Radians(SunLongitudeLowAccuracy(Year, Month, Day))
  Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
  
  sinAlpha = -Sin(B) * Sin(Eps) + Cos(B) * Cos(Eps) * Sin(L)
  cosAlpha = Cos(B) * Cos(L)
  SunRightAscensionLowAccuracy = Modulus(Degrees(Atn2(cosAlpha, sinAlpha)), 360#)
End Function

Public Function SunRightAscensionVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunRightAscention dimaksudkan untuk menghitung askensio rekta Matahari Nampak
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Askensio Rekta Matahari terlihat (Apparent Sun Right Ascension) bujur derajat
'
' CATATAN:
'    Untuk mengkonversikan dalam satuan jam, maka nilai dalam Busur derajat tersebut
'    dibagi dengan 15.
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Askensio Rekta  Matahari terlihat Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunRightAscensionVSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunRightAscensionVSOP87(Y, M, D)
'    Hasil:
'         x = 325.636087738009
'           = 21:42:32.66
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 12. hal. 89
'------===========================----------------------
'*Yogyakarta, 04 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim sinAlpha, cosAlpha, Eps, L, B As Double
  
  B = Radians(SunLatitudeVSOP87(Year, Month, Day))
  L = Radians(SunLongitudeVSOP87(Year, Month, Day))
  Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
  
  sinAlpha = -Sin(B) * Sin(Eps) + Cos(B) * Cos(Eps) * Sin(L)
  cosAlpha = Cos(B) * Cos(L)
  SunRightAscensionVSOP87 = Modulus(Degrees(Atn2(cosAlpha, sinAlpha)), 360#)
End Function

Public Function ApparentSunRightAscensionLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi ApparentSunRightAscention dimaksudkan untuk menghitung askensio rekta Matahari terlihat dengan
' rumus akurasi rendah
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Askensio Rekta Matahari terlihat (Apparent Sun Right Ascension) bujur derajat
'
' CATATAN:
'    Untuk mengkonversikan dalam satuan jam, maka nilai dalam Busur derajat tersebut
'    dibagi dengan 15.
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Askensio Rekta  Matahari terlihat Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentSunRightAscensionLowAccuracy.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentSunRightAscensionLowAccuracy(Y, M, D)
'    Hasil:
'         x = 325.63595749305
'           = 21:42:32.63
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 12. hal. 89
'------===========================----------------------
'*Yogyakarta, 04 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim sinAlpha, cosAlpha, Eps, L, B As Double
  
  B = Radians(SunLatitudeLowAccuracy(Year, Month, Day))
  L = Radians(ApparentSunLongitudeLowAccuracy(Year, Month, Day))
  Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
  
  sinAlpha = -Sin(B) * Sin(Eps) + Cos(B) * Cos(Eps) * Sin(L)
  cosAlpha = Cos(B) * Cos(L)
  ApparentSunRightAscensionLowAccuracy = Modulus(Degrees(Atn2(cosAlpha, sinAlpha)), 360#)
End Function


Public Function ApparentSunRightAscensionVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi ApparentSunRightAscention dimaksudkan untuk menghitung askensio rekta Matahari terlihat
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Askensio Rekta Matahari terlihat (Apparent Sun Right Ascension) bujur derajat
'
' CATATAN:
'    Untuk mengkonversikan dalam satuan jam, maka nilai dalam Busur derajat tersebut
'    dibagi dengan 15.
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Askensio Rekta  Matahari terlihat Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentSunRightAscension.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentSunRightAscension(Y, M, D)
'    Hasil:
'         x = 325.630417902699
'           = 21:42:31.30
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 12. hal. 89
'------===========================----------------------
'*Yogyakarta, 04 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim sinAlpha, cosAlpha, Eps, L, B As Double
  
  B = Radians(SunLatitudeVSOP87(Year, Month, Day))
  L = Radians(ApparentSunLongitudeVSOP87(Year, Month, Day))
  Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
  
  sinAlpha = -Sin(B) * Sin(Eps) + Cos(B) * Cos(Eps) * Sin(L)
  cosAlpha = Cos(B) * Cos(L)
  ApparentSunRightAscensionVSOP87 = Modulus(Degrees(Atn2(cosAlpha, sinAlpha)), 360#)
End Function

Public Function SunAltitudeVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
                                  ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function SunAltitudeVSOP87 dimaksudkan untuk menghitung Tinggi Matahari di atas ufuk dengan rumus-rumus
' metode VSOP87.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal (Waktu Universal)
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Ketinggian Matahari dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Tinggi Matahari di atas ufuk Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' di Jakarta dengan koordinat (6° 8' S 106° 45'E) dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunAltitudeVSOP87.
'           L = -6. - 8./60.
'           B = 106. + 45./60.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunAltitudeVSOP87(L, B, Y, M, D)
'    Hasil:
'         x = 64.9620859194381
'           = 64 ° 57 ' 43.51"
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985, hal. 12
'----------------------------
'*Demak, 18 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim La, h, DEC As Double
   
   h = Radians(HourAngleOfTheSun(Bujur, Year, Month, Day) * 15#)
   DEC = Radians(ApparentSunDeclinationVSOP87(Year, Month, Day))
   La = Radians(Lintang)
   SunAltitudeVSOP87 = Cos(h) * Cos(DEC) * Cos(La) + Sin(DEC) * Sin(La)
   SunAltitudeVSOP87 = Degrees(Asin(SunAltitudeVSOP87))
End Function

Public Function ApparentSunAltitudeVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
                                          ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function ApparentSunAltitudeVSOP87 dimaksudkan untuk menghitung Tinggi Matahari Nampak yakni
' Tinggi Matahari sebenarnya dikoreksi dengan pengaruh refraksi
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal (Waktu Universal)
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Ketinggian Matahari dalam derajat yang terkoreksi dengan Refraksi
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Tinggi Matahari di atas ufuk terkoreksi dengan refraksi
' Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' di Jakarta dengan koordinat (6° 8' S 106° 45'E) dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentSunAltitudeVSOP87.
'           L = -6. - 8./60.
'           B = 106. + 45./60.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentSunAltitudeVSOP87(L, B, Y, M, D)
'    Hasil:
'         x = 64.9694112155221
'           = 64° 58' 9.88"
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985, hal. 12
'----------------------------
'* Demak, 18 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim h, R As Double
   
   h = SunAltitudeVSOP87(Lintang, Bujur, Year, Month, Day)
   R = RefractionTrueAltitude(0, 1010, 27#)
   ApparentSunAltitudeVSOP87 = h + RefractionTrueAltitude(h, 1010, 27#) / 60#
End Function

Public Function SunAzimuthVSOP87(ByVal Lintang As Double, ByVal Bujur As Double, _
                                 ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function SunAzimut dimaksudkan untuk menghitung Azimuth Matahari dengan rumus
' perhitungan akurasi rendah ketika menghitung Lintang dan Bujur Ekliptik Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Azimuth Matahari dinyatakan dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Azimuth Matahari di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tanggal 12 Februari 2010, jam 10:30:15 WIB dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunAzimuthVSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           L = -6. - 8./60.
'           B = 106. + 45./60.
'           x = SunAzimuthVSOP87(L,B, Y, M, D)
'    Hasil:
'         x = 109.567213717889
'           = 109° 34' 1.97"
'
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim h, DEC, A, cosA As Double
   
   h = Radians(HourAngleOfTheSun(Bujur, Year, Month, Day) * 15#)
   DEC = Radians(ApparentSunDeclinationVSOP87(Year, Month, Day))
   cosA = Cos(h) * Sin(Radians(Lintang)) - Tan(DEC) * Cos(Radians(Lintang))
   SunAzimuthVSOP87 = Modulus(Degrees(Atn2(cosA, Sin(h))) + 180#, 360#)
End Function

Public Function SunAltitudeAtSunsetOrSunrise(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                                             Optional ByVal P As Double = 1010, _
                                             Optional ByVal T As Double = 27, _
                                             Optional ByVal h As Double = 0) As Double
' Function SunAltitudeAtSunsetOrSunrise dimaksudkan untuk menghitung ketinggian matahari saat terbenam atau terbit
' Perhitungan dilakukan dengan menggunakan koreksi Semidiameter matahari dan pengaruh refraksi
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      p       : Tekanan Udara di permukaan bumi dinyatakan dalam millibar
'      T       : Suhu di permukaan bumi dinyatakan dalam derajat Celcius
'      h       : Tinggi Tempat terhadap muka laut dalam meter
'
' OUTPUT
'    SunAltitudeAtSunsetOrSunrise (Tinggi Matahari saat terbenam atau terbit) dinyatakan dalam menit busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Langkah-langkah  yang harus dipersiapkan untuk menghitung Tinggi Matahari pada saat terbenam
' adalah sebagai berikut:
'    1.  Mempersiapkan input berupa tahun, bulan, tanggal, tekanan udara dan suhu udara
'    2.  Hitung tinggi Matahari saat terbenam dengan fungsi SunAltitudeAtSunsetOrSunrise
'        dengan cara sebagai berikut:
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           p =  1010.
'           T = 27.
'           h = 0.
'           Alt = SunAltitudeAtSunsetOrSunrise(Tahun, Bulan, Tanggal, p, T, h)
'
'    Hasil:
'           Alt = -48.7256313987265
'               = -48' 43.54"
' Catatan:
'   - Ketinggian Matahari saat terbenam atau terbit tidak nol, melainkan berada dibawah ufuk akibat koreksi
'     Semidiameter matahari (setengah besar piringan matahari) dan pengaruh refraksi udara.
'   - Dalam perhitungan di sini diasumsikan semidiameter matahari tidak berubah, antara saat matahari terbit
'     dan saat matahari terbenam. Perbedaan atau perubahan ini tidak lebih dari 0.01 detik busur, sehingga tidak
'     ada pengaruhnya pada perhitungan waktu matahari terbit dan terbenam.
'----------------------------
'* Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  SunAltitudeAtSunsetOrSunrise = -(SunSemiDiameterVSOP87(Year, Month, Int(Day) + 0.5) + RefractionApparentAltitude(0, P, T) + KoreksiDIP(h))
End Function

Public Function Sunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                       ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                       Optional ByVal Alt As Double = -1#) As Double
' Function Sunset dimaksudkan untuk menghitung waktu Terbenam Matahari
' Matahari dinyatakan tepat Terbenam apabila piringan atas matahari berada persis di garis ufuk
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Alt     : Posisi Matahari terhadap ufuk dalam derajat busur.
'                Untuk perhitungan yang tidak memerlukan ketelitian tinggi, maka Alt=1
'                Untuk perhitungan yang teliti, maka Alt harus dihitung berdasarkan koreksi
'                Semidiameter dan Refraksi. lihat fungsi SunAltitudeAtSunsetOrSunrise
'
' OUTPUT
'    Waktu Matahari Terbenam dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Waktu Terbenam Matahari di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Mempersiapkan opsi perhitungan sebagai input
'    3. Hitung tinggi Matahari saat terbenam, lihat SunAltitudeAtSunsetOrSunrise
'    4. Jalankan fungsi Sunset.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           p =  1010.
'           T = 27.
'           Alt = SunAltitudeAtSunsetOrSunrise(Tahun, Bulan, Tanggal, p, T, h) / 60#
'
'           Terbenam = Sunset(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt)
'
'    Hasil:
'           Terbenam = 18.2757203214741
'                    = 18:16:33 WIB
'
' Catatan:
'   - Ketinggian Matahari saat terbenam atau terbit tidak nol, melainkan berada dibawah ufuk akibat koreksi
'     Semidiameter matahari (setengah besar piringan matahari) dan pengaruh refraksi udara.
'   - Dalam perhitungan di sini diasumsikan semidiameter matahari tidak berubah, antara saat matahari terbit
'     dan saat matahari terbenam. Perbedaan atau perubahan ini tidak lebih dari 0.01 detik busur, sehingga tidak
'     ada pengaruhnya pada perhitungan waktu matahari terbit dan terbenam.
'----------------------------
'* Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E, h, mDay As Double
   
   E = EquationOfTime(Year, Month, Day)
   h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, Day)
   Sunset = 999#
   If h <> 999# Then Sunset = 12 - E + h + kwd(TZ, Bujur)
   If Sunset <> 999# Then
     E = EquationOfTime(Year, Month, Int(Day) + (Sunset - TZ) / 24#)
     h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, Int(Day) + (Sunset - TZ) / 24#)
     Sunset = 12 - E + h + kwd(TZ, Bujur)
   End If
  
End Function

Public Function Sunrise(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                        Optional ByVal Alt As Double = -1#) As Double
' Function Sunrise dimaksudkan untuk menghitung waktu terbit matahari
' Matahari dinyatakan tepat Terbit apabila piringan atas matahari berada persis di garis ufuk
'
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Alt     : Posisi Matahari terhadap ufuk dalam derajat busur.
'                Untuk perhitungan yang tidak memerlukan ketelitian tinggi, maka Alt=1
'                Untuk perhitungan yang teliti, maka Alt harus dihitung berdasarkan koreksi
'                Semidiameter dan Refraksi. lihat fungsi SunAltitudeAtSunsetOrSunrise
'
' OUTPUT
'    Waktu Matahari Terbit dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Waktu Terbenam Matahari di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010  dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Mempersiapkan opsi perhitungan sebagai input
'    3. Hitung tinggi Matahari saat Terbit, lihat SunAltitudeAtSunsetOrSunrise
'    4. Jalankan fungsi Sunrise.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           p =  1010.
'           T = 27.
'           Alt = SunAltitudeAtSunsetOrSunrise(Tahun, Bulan, Tanggal, p, T, h) / 60#
'
'           Terbit = Sunrise(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt)
'
'    Hasil:
'           Terbit = 5.9630226610661
'                  = 05:57:47 WIB
'
' Catatan:
'   - Ketinggian Matahari saat terbenam atau terbit tidak nol, melainkan berada dibawah ufuk akibat koreksi
'     Semidiameter matahari (setengah besar piringan matahari) dan pengaruh refraksi udara.
'   - Dalam perhitungan di sini diasumsikan semidiameter matahari tidak berubah, antara saat matahari terbit
'     dan saat matahari terbenam. Perbedaan atau perubahan ini tidak lebih dari 0.01 detik busur, sehingga tidak
'     ada pengaruhnya pada perhitungan waktu matahari terbit dan terbenam.
'----------------------------
'* Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E, h As Double
   
   E = EquationOfTime(Year, Month, Day)
   h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, Day)
   Sunrise = 999#
   If h <> 999# Then Sunrise = 12 - E - h + kwd(TZ, Bujur)
   If Sunrise <> 999# Then
     E = EquationOfTime(Year, Month, Int(Day) + (Sunrise - TZ) / 24#)
     h = SunHourAngleVSOP87(Lintang, Alt, Year, Month, Int(Day) + (Sunrise - TZ) / 24#)
     Sunrise = 12 - E - h + kwd(TZ, Bujur)
   End If
End Function

Public Function SunSemiDiameterLowAccuracy(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function SunSemiDiameterVSOP87 dimaksudkan untuk menghitung Semidiameter Matahari dengan
' rumus sederhana
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Satuan Waktu adalah UT (Universal Time)
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24)
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Semidiameter Matahari dalam menit busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Semidiameter Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunSemiDiameterLowAccuracy.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           sD = SunSemiDiameterLowAccuracy(Y, M, D)
'    Hasil:
'         sD = 16.2024056697921
'            = 16' 12.14"
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 53 hal 359
' Catatan:
'     Perhitungan didasarkan dengan metode VSOP87 untuk menghitung jarak Bumi-Matahari
'----------------------------
'*Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim R As Double
     
' Mengkonversikan Tanggal/Bulan/Tahun ke dalam tanggal julian (Julian Date) dan
' Julian Century (tau) serta Julian Millenea (tau) dengan epoch J2000 Januari 1.5
' dalam sistem waktu Dynamical Time.
  R = DistanceToTheSunLowAccuracy(Year, Month, Day)
  SunSemiDiameterLowAccuracy = (959.63 / R) / 60#
End Function

Public Function SunSemiDiameterVSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function SunSemiDiameterVSOP87 dimaksudkan untuk menghitung Semidiameter Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Satuan Waktu adalah UT (Universal Time)
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24)
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Semidiameter Matahari dalam menit busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Semidiameter Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunSemiDiameterVSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           sD = SunSemiDiameterVSOP87(Y, M, D)
'    Hasil:
'         sD = 16.2019852588471
'            = 16' 12.12"
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 53 hal 359
' Catatan:
'     Perhitungan didasarkan dengan metode VSOP87 untuk menghitung jarak Bumi-Matahari
'----------------------------
'*Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim R As Double
     
' Mengkonversikan Tanggal/Bulan/Tahun ke dalam tanggal julian (Julian Date) dan
' Julian Century (tau) serta Julian Millenea (tau) dengan epoch J2000 Januari 1.5
' dalam sistem waktu Dynamical Time.
  R = DistanceToTheSunVSOP87(Year, Month, Day)
  SunSemiDiameterVSOP87 = (959.63 / R) / 60#
End Function

Public Function SunsEqCenter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Then find the Sun's equation of center C as follows :
'  C = + (l°.914600 - 0°.004817 * T - 0°.000014T2) sin M
'      + (0°.019993 - 0°.000 101 * T) sin 2M
'      + 0°.000 290 sin 3M
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'    SunsEqCenter dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Sun's equation of center pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunsEqCenter(Y, M, D)
'    Hasil:
'         x = 1.22371622104403
'
'  Referensi
'    Astronomical Algorithm hal. 152
' -------------------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T, M As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   M = Radians(SunMeanAnomaly(Year, Month, Day))
   SunsEqCenter = (1.9146 - 0.004817 * T - 0.000014 * T ^ 2) * Sin(M) _
                  + (0.019993 - 0.000101 * T) * Sin(2 * M) _
                  + 0.00029 * Sin(3 * M)
End Function

Public Function EccentricityOfEarthOrbit(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  The eccentricity of the Earth's orbit is
'  e = 0.016 708 617 - 0.000 042 037 T - 0.000 000 1236 T^2  (24.4)
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'    Eksentrisitas tidak pakai satuan ukur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Ekentrisitas dari orbit Bumi saat mengelilingi Matahari pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = EccentricityOfEarthOrbit(Y, M, D)
'    Hasil:
'         x = 0.0167043635281465
'
'  Referensi
'    Astronomical Algorithm hal. 151
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   EccentricityOfEarthOrbit = 0.016708617 - 0.000042037 * T - 0.0000001236 * T ^ 2
End Function

Public Function HourAngleOfTheSun(ByVal Bujur As Double, ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function HourAngleOfTheSun dimaksudkan untuk menghitung Sudut Jam
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal (Waktu Universal)
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Sudut Jam Matahari dalam jam
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Sudut Jam Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi HourAngleOfTheSun.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = HourAngleOfTheSun(B, Y, M, D)
'    Hasil:
'         x = 22.4026470076929
'           = 22:24:10
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985, hal. 12
'----------------------------
'*Demak, 18 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   HourAngleOfTheSun = Modulus(ApparentLocalSiderealTime(Bujur, Year, Month, Day) - ApparentSunRightAscensionVSOP87(Year, Month, Day) / 15#, 24#)
End Function

Public Function SunHourAngleVSOP87(ByVal Lintang As Double, ByVal Alt As Double, ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SunHourAngleVSOP87 dimaksudkan untuk menghitung Sudut Jam Matahari.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Lintang : Koordinat Lintang Tempat yang akan dihitung
'      - ALt     : Tinggi Matahari
'      - Year    : tahun Gregorian/Julian (tahun Masehi)
'      - Month   : bulan Masehi (1-12)
'      - Day     : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Sudut Jam Matahari dalam jam
'
' CATATAN:
'   Perhitungan Sudut Jam Matahari didasarkan pada Deklinasi Matahari Nampak.
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Sudut Jam Matahari saat ketinggian 15 di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang tempat, Zona Waktu, Tahun, Bulan dan Tanggal.
'    2. Mempersiapkan tinggi Matahari
'    3. Jalankan fungsi SunHourAngleVSOP87.
'           Lintang = -6 - 8/60.
'           Tz = 7.
'           Alt = 15.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunHourAngleVSOP87(Lintang, Alt, Y, M, D)
'    Hasil:
'         x = 5.06757136603873
'           = 05:04:03
' Referensi
'    Astronomical Algorithm
' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
'* 18 Juni 2011
      
   Dim delta, cosh As Double
   
   On Error GoTo Keluar
   delta = Radians(ApparentSunDeclinationVSOP87(Year, Month, Day))
   cosh = (Sin(Radians(Alt)) - Sin(delta) * Sin(Radians(Lintang))) / (Cos(delta) * Cos(Radians(Lintang)))
   SunHourAngleVSOP87 = Degrees(Acos(cosh)) * 24# / 360#
   Exit Function
Keluar:
   SunHourAngleVSOP87 = 999#
End Function

Public Function MoonMeanLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonMeanLongitude berfungsi untuk menghitung Rata-rata anomaly Bulan.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Rata-rata bujur Bulan dalam derajat
'
' CATATAN:
'    Dalam Buku Astronomical Algorithm disimbolkan dengan L'
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung rata-rata Bujur Bulan pada 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonMeanLongitude(Y, M, D)
'    Hasil:
'         x = 48900.4472571836
'           = 300° 26' 50.13"
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 308
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   MoonMeanLongitude = 218.3164591 + 481267.88134236 * T - 0.0013268 * T ^ 2 + T ^ 3 / 538841# - T ^ 4 / 65194000#
End Function

Public Function Ijtimak(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function Ijtimak adalah duplikat dari fungsi NewMoon
' ------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Ijtimak = NewMoon(Year, Month, Day)
End Function

Public Function Konjungsi(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function Konjungsi adalah duplikat dari fungsi NewMoon
' ------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Konjungsi = NewMoon(Year, Month, Day)
End Function

Public Function NewMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function NewMoon dipakai untuk menghitung saat terjadinya konjungsi.
'  NewMoon  adalah dinyatakan dalam Julian Day, oleh karena itu diperlukan konversi
'  ke satuan waktu sehari-hari. lihat fungsi caldat.
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'  OUTPUT:
'    Saat terjadinya Newmoon atau Konjungsi atau Ijtima dalam waktu universal (Universal Time)
'
'  Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      x = NewMoon(Y, M, D)
'      x = 2443192.65062192
'        = Jum'at, 18 Februari 1977 03:36:54 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab. 47
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, M As Integer
   Dim D, JDE, JDEo, CP, c As Double
   
   Y = Year
   M = Month
   D = Day
   JDE = JulianDatum(Y, M, D)
   Do
      JDEo = nilai_JDE(Y, M, D, 0#)
      CP = KoreksiNewMoon(Y, M, D)
      c = Koreksi_PlanetaryArguments(Y, M, D, 0#)
      NewMoon = JDEo + CP + c
      Y = JD2Year(NewMoon)
      M = JD2Month(NewMoon)
      D = JD2Date(NewMoon) + JD2Time(NewMoon) + 29.2
   Loop While NewMoon < JDE

   Do Until (NewMoon - JDE) < 30#
      JDEo = nilai_JDE(Y, M, D, 0#)
      CP = KoreksiNewMoon(Y, M, D)
      c = Koreksi_PlanetaryArguments(Y, M, D, 0#)
      NewMoon = JDEo + CP + c
      Y = JD2Year(NewMoon)
      M = JD2Month(NewMoon)
      D = JD2Date(NewMoon) + JD2Time(NewMoon) - 29.2
   Loop
   NewMoon = NewMoon - DeltaT(Year, Month, Day) / 86400#
End Function

Public Function FirstQuarter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function FirstQuarter dipakai untuk menghitung saat terjadinya Fase Bulan First Quarter.
'  FirstQuarter  dinyatakan dalam Julian Day, oleh karena itu diperlukan konversi
'  ke satuan waktu sehari-hari. lihat fungsi caldat.
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'  OUTPUT:
'    Saat terjadinya First Quarter atau Seperempat Pertama dalam waktu universal (Universal Time)
'
'  Contoh pemakaian fungsi:
'   Astronomical Algorithm Example analogously 4-7.a — Calculate the instant of the First Quarter which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      x = FirstQuarter(Y, M, D)
'      x = 2443200.61907792
'        = Sabtu, 26 Februari 1977 02:51:28 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab. 47
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, M As Integer
   Dim D, JDE, JDEo, CP, w, c As Double
   
   Y = Year
   M = Month
   D = Day
   JDE = JulianDatum(Y, M, D)
   Do
      JDEo = nilai_JDE(Y, M, D, 0.25)
      CP = KoreksiFirstQuarter(Y, M, D)
      w = Koreksi_W(Y, M, D, 0.25)
      c = Koreksi_PlanetaryArguments(Y, M, D, 0.25)
      FirstQuarter = JDEo + w + CP + c
      Y = JD2Year(FirstQuarter)
      M = JD2Month(FirstQuarter)
      D = JD2Date(FirstQuarter) + JD2Time(FirstQuarter) + 29.2
   Loop While FirstQuarter < JDE

   Do Until (FirstQuarter - JDE) < 30#
      JDEo = nilai_JDE(Y, M, D, 0.25)
      CP = KoreksiFirstQuarter(Y, M, D)
      w = Koreksi_W(Y, M, D, 0.25)
      c = Koreksi_PlanetaryArguments(Y, M, D, 0.25)
      FirstQuarter = JDEo + w + CP + c
      Y = JD2Year(FirstQuarter)
      M = JD2Month(FirstQuarter)
      D = JD2Date(FirstQuarter) + JD2Time(FirstQuarter) - 29.2
   Loop
   FirstQuarter = FirstQuarter - DeltaT(Year, Month, Day) / 86400#
End Function

Public Function FullMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function FullMoon dipakai untuk menghitung saat terjadinya Bulan Purnama (Full Moon).
'  FullMoon  dinyatakan  dalam Julian Day, oleh karena itu diperlukan konversi
'  ke satuan waktu sehari-hari. lihat fungsi caldat.
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'  OUTPUT:
'    Saat terjadinya Full Moon atau Bulan Purnama dalam waktu universal (Universal Time)
'
'  Contoh pemakaian fungsi:
'   Astronomical Algorithm analogously Example 4-7.a — Calculate the instant of the Full Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      x = FullMoon(Y, M, D)
'      x = 2443208.21750627
'        = Sabtu, 05 Maret 1977 17:13:13 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab. 47
' ------------------------------------------
'  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, M As Integer
   Dim JDE, JDEo, CP, c, D As Double
   
   Y = Year
   M = Month
   D = Day
   JDE = JulianDatum(Y, M, D)
   Do
      JDEo = nilai_JDE(Y, M, D, 0.5)
      CP = KoreksiFullMoon(Y, M, D)
      c = Koreksi_PlanetaryArguments(Y, M, D, 0.5)
      FullMoon = JDEo + CP + c
      Y = JD2Year(FullMoon)
      M = JD2Month(FullMoon)
      D = JD2Date(FullMoon) + JD2Time(FullMoon) + 29.2
   Loop While FullMoon < JDE

   Do Until (FullMoon - JDE) < 30#
      JDEo = nilai_JDE(Y, M, D, 0.5)
      CP = KoreksiFullMoon(Y, M, D)
      c = Koreksi_PlanetaryArguments(Y, M, D, 0.5)
      FullMoon = JDEo + CP + c
      Y = JD2Year(FullMoon)
      M = JD2Month(FullMoon)
      D = JD2Date(FullMoon) + JD2Time(FullMoon) - 29.2
   Loop
   FullMoon = FullMoon - DeltaT(Year, Month, Day) / 86400#
End Function

Public Function LastQuarter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function LastQuarter dipakai untuk menghitung saat terjadinya Fase Bulan Last Quarter.
'  LastQuarter  dinyatakan dalam Julian Day, oleh karena itu diperlukan konversi
'  ke satuan waktu sehari-hari. lihat fungsi caldat.
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'  OUTPUT:
'    Saat terjadinya Last Quarter atau Seperempat Terakhir dalam waktu universal (Universal Time)
'
'  Contoh pemakaian fungsi:
'   Astronomical Algorithm Example analogously 4-7.a — Calculate the instant of the Last Quarter which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      x = LastQuarter(Y, M, D)
'      x = 2443214.98275591
'        = Sabtu, 12 Maret 1977 11:35:10 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab. 47
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, M As Integer
   Dim D, JDE, JDEo, CP, w, c As Double
   
   Y = Year
   M = Month
   D = Day
   JDE = JulianDatum(Y, M, D)
   Do
      JDEo = nilai_JDE(Y, M, D, 0.75)
      CP = KoreksiLastQuarter(Y, M, D)
      w = Koreksi_W(Y, M, D, 0.75)
      c = Koreksi_PlanetaryArguments(Y, M, D, 0.75)
      LastQuarter = JDEo - w + CP + c
      Y = JD2Year(LastQuarter)
      M = JD2Month(LastQuarter)
      D = JD2Date(LastQuarter) + JD2Time(LastQuarter) + 29.2
   Loop While LastQuarter < JDE

   Do Until (LastQuarter - JDE) < 30#
      JDEo = nilai_JDE(Y, M, D, 0.75)
      CP = KoreksiLastQuarter(Y, M, D)
      w = Koreksi_W(Y, M, D, 0.75)
      c = Koreksi_PlanetaryArguments(Y, M, D, 0.75)
      Y = JD2Year(LastQuarter)
      M = JD2Month(LastQuarter)
      D = JD2Date(LastQuarter) + JD2Time(LastQuarter) - 29.2
   Loop
   LastQuarter = LastQuarter - DeltaT(Year, Month, Day) / 86400#
End Function

Public Function MoonLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonLongitude berfungsi untuk menghitung Moon's Longitude (Bujur Ekliptik Bulan)
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'    Bujur Ekliptik Bulan dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Bujur Ekliptik Bulan pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonLongitude(Y, M, D)
'    Hasil:
'         x = 302.012884124058
'           = 302° 00' 46.38"
' Referensi:
'    Jean Meeus, Astronomical Algorithm, bab 45
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    MoonLongitude = Modulus(MoonMeanLongitude(Year, Month, Day) + SigmaL(Year, Month, Day) / 1000000, 360#)
End Function

Public Function ApparentMoonLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi ApparentMoonLongitude berfungsi untuk menghitung Apparent Moon's Longitude (Bujur Ekliptik Bulan Nampak)
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'    Bujur Ekliptik Bulan Nampak dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Bujur Ekliptik Bulan Nampak pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentMoonLongitude(Y, M, D)
'    Hasil:
'         x = 302.017805418309
'           = 302° 01' 4.10"
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, bab 45
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    ApparentMoonLongitude = Modulus(MoonLongitude(Year, Month, Day) + NutationInLongitude(Year, Month, Day), 360#)
End Function

Public Function MoonLatitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonLatitude berfungsi untuk menghitung Moon's Latitude (Lintang Ekliptik Bulan)
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'    Lintang Ekliptik Bulan dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Lintang Ekliptik Bulan pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonLatitude(Y, M, D)
'    Hasil:
'         x = 1.00647050688368
'           = 1° 00' 23.29"
' Referensi:
'    Jean Meeus, Astronomical Algorithm, bab 45
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    MoonLatitude = SigmaB(Year, Month, Day) / 1000000#
End Function

Public Function DistanceToTheMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi DistanceToTheMoon dimaksudkan untuk menghitung jarak Bulan dari Bumi.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Jarak Bulan dari Bumi dalam km
'
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung jarak Bulan dari Bumi pada 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = DistanceToTheMoon(Y, M, D)
'    Hasil:
'         x = 406234.047137886 km
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 308
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    DistanceToTheMoon = 385000.56 + SigmaR(Year, Month, Day) / 1000#
End Function

Public Function MoonMeanElongation(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonMeanElongation berfungsi untuk menghitung Rata-rata Elongasi Bulan.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    MoonMeanElongation dalam derajat
'
' CATATAN:
'    Dalam Buku Astronomical Algorithm disimbolkan dengan D
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung MoonMeanElongation 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Yr = 2010
'           Mn = 2
'           Dy = 12.14600694444440
'           x = MeanElongationOfTheMoon(Yr, Mn, Dy)
'    Hasil:
'         x = 45338.3621546409
'           = 338° 21' 43.76"
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 308
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   MoonMeanElongation = 297.8502042 + 445267.1115168 * T - 0.00163 * T ^ 2 + T ^ 3 / 545868 - T ^ 4 / 113065000
End Function

Public Function MoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonDeclination dimaksudkan untuk menghitung deklinasi Bulan
'
' INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Deklinasi Bulan dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Deklinasi Bulan Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapka
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi MoonDeclination.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonDeclination(Y, M, D)
'    Hasil:
'         x = -18.7302610333087
'           = -18° 43' 48.94"
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 12. hal. 89
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim B, L, Eps, sinDEC As Double

' Perhitungan deklinasi mendasarkan pada rumus transformasi koordinat,
' yaitu dari koordinat ekliptik ke koordinat ekuator.
' Untuk itu terlebih dahulu harus dihitung koordinat Bujur, Lintang ekliptik Bulan.
' Epssilon (Eps) adalah (Obliquity of the Ecliptic) yakni sudut antara
' celestial ekuator dan celestial ekliptik.
  B = Radians(MoonLatitude(Year, Month, Day))
  L = Radians(MoonLongitude(Year, Month, Day))
  Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
  
  sinDEC = Sin(B) * Cos(Eps) + Cos(B) * Sin(Eps) * Sin(L)
  MoonDeclination = Degrees(Asin(sinDEC))
End Function

Public Function ApparentMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonDeclination dimaksudkan untuk menghitung deklinasi Bulan Nampak
'
' INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Deklinasi Bulan Nampak dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Deklinasi Bulan Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapka
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentMoonDeclination.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentMoonDeclination(Y, M, D)
'    Hasil:
'         x = -18.7291653862793
'           = -18° 43' 45.00"
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 12. hal. 89
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim B, L, Eps, sinDEC As Double

' Perhitungan deklinasi mendasarkan pada rumus transformasi koordinat,
' yaitu dari koordinat ekliptik ke koordinat ekuator.
' Untuk itu terlebih dahulu harus dihitung koordinat Bujur, Lintang ekliptik Bulan.
' Epssilon (Eps) adalah (Obliquity of the Ecliptic) yakni sudut antara
' celestial ekuator dan celestial ekliptik.
  B = Radians(MoonLatitude(Year, Month, Day))
  L = Radians(ApparentMoonLongitude(Year, Month, Day))
  Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
  
  sinDEC = Sin(B) * Cos(Eps) + Cos(B) * Sin(Eps) * Sin(L)
  ApparentMoonDeclination = Degrees(Asin(sinDEC))
End Function

Public Function MoonRightAscention(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonRightAscention dimaksudkan untuk menghitung askensio rekta Bulan
'
' INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Askensio Rekta Bulan dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Askensio Rekta  Bulan Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapka
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi MoonRightAscention.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonRightAscention(Y, M, D)
'    Hasil:
'         x = 304.032820971634
'           = 304° 01' 58.16"
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 12. hal. 89
'------===========================----------------------
'*Cibinong, 21 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Eps, DEC, sinAlpha, cosAlpha, B, L As Double
  
  DEC = Radians(MoonDeclination(Year, Month, Day))
  B = Radians(MoonLatitude(Year, Month, Day))
  L = Radians(MoonLongitude(Year, Month, Day))
  Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
  
  sinAlpha = -Sin(B) * Sin(Eps) + Cos(B) * Cos(Eps) * Sin(L)
  sinAlpha = sinAlpha / Cos(DEC)
  cosAlpha = Cos(B) * Cos(L) / Cos(DEC)
  MoonRightAscention = Degrees(Atn2(cosAlpha, sinAlpha))
End Function

Public Function ApparentMoonRightAscention(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi ApparentMoonRightAscention dimaksudkan untuk menghitung askensio rekta Bulan Nampak
'
' INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Askensio Rekta Bulan Nampak dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Askensio Rekta  Bulan Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapka
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentMoonRightAscention.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentMoonRightAscention(Y, M, D)
'    Hasil:
'         x = 304.037886208266
'           = 304° 02' 16.39"
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 12. hal. 89
'------===========================----------------------
'*Cibinong, 21 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Eps, DEC, sinAlpha, cosAlpha, B, L As Double
  
  DEC = Radians(ApparentMoonDeclination(Year, Month, Day))
  B = Radians(MoonLatitude(Year, Month, Day))
  L = Radians(ApparentMoonLongitude(Year, Month, Day))
  Eps = Radians(ObliquityOfEcliptic(Year, Month, Day))
  
  sinAlpha = -Sin(B) * Sin(Eps) + Cos(B) * Cos(Eps) * Sin(L)
  sinAlpha = sinAlpha / Cos(DEC)
  cosAlpha = Cos(B) * Cos(L) / Cos(DEC)
  ApparentMoonRightAscention = Degrees(Atn2(cosAlpha, sinAlpha))
End Function

Public Function MoonElongation(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function MoonElongation dimaksudkan untuk menghitung Jarak Busur Bulan-Matahari Nampak
'
' INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Jarak Elongasi Bulan Matahari dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Elongasi Bulan Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi MoonElongation.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonElongation(Y, M, D)
'    Hasil:
'         x = 21.302362445509
'           = 21° 18' 8.50"
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 46 hal 315
'----------------------------
'* Cibinong, 14 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
   Dim d0, d1, a0, A1, cosELO As Double
   
   d0 = Radians(ApparentSunDeclinationVSOP87(Year, Month, Day))
   d1 = Radians(ApparentMoonDeclination(Year, Month, Day))
   a0 = Radians(ApparentSunRightAscensionVSOP87(Year, Month, Day))
   A1 = Radians(ApparentMoonRightAscention(Year, Month, Day))
   cosELO = Sin(d0) * Sin(d1) + Cos(d0) * Cos(d1) * Cos(a0 - A1)
   MoonElongation = Degrees(Acos(cosELO))
End Function

Public Function MoonElongationAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                        Optional ByVal Alt As Double = -1#) As Double
' Function MoonElongationAtSunset dimaksudkan untuk menghitung Jarak Elongasi Bulan dan Matahari
' pada saat Terbenam Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Alt     : Posisi Matahari terhadap ufuk dalam derajat busur.
'                Untuk perhitungan yang tidak memerlukan ketelitian tinggi, maka Alt=1
'                Untuk perhitungan yang teliti, maka Alt harus dihitung berdasarkan koreksi
'                Semidiameter dan Refraksi. lihat fungsi SunAltitudeAtSunsetOrSunrise
'
' OUTPUT
'    Jarak Elongasi Bulan dan Matahari dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Jarak Elonasi Bulan-Matahari di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 19 Juli 2012 pada saat Matahari terbenam dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Mempersiapkan opsi perhitungan sebagai input
'    3. Hitung tinggi Matahari saat terbenam, lihat SunAltitudeAtSunsetOrSunrise
'    4. Jalankan fungsi MoonElongationAtSunset.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2012
'           Bulan = 7
'           Tanggal = 19
'           p =  1010.
'           T = 27.
'           Alt = SunAltitudeAtSunsetOrSunrise(Tahun, Bulan, Tanggal, p, T, h) / 60#
'
'           x = MoonElongationAtSunset(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt)
'
'    Hasil:
'           x = 5.32094539895012
'           x = 5° 19' 15.40"
'
'----------------------------
'* Cibinong, 09 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Terbenam As Double
   
   Terbenam = Sunset(Lintang, Bujur, TZ, Year, Month, Day, Alt)
   MoonElongationAtSunset = MoonElongation(Year, Month, Int(Day) + (Terbenam - TZ) / 24#)
End Function


Public Function MoonsAge(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function MoonsAge dimaksudkan untuk menghitung umur bulan saat Terbenam Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT
'    Umur Bulan dalam hari
'
' CATATAN:
'   Umur Bulan yang dihitung sejak terjadinya konjungsi terdekat.
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Umur Bulan di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 jam 10:30:15 WIB dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan Tahun, Bulan, Tanggal dan jam.
'    2. Jalankan fungsi MoonsAge.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12.14600694444440
'
'           UmurBulan = MoonsAge(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt)
'
'    Hasil:
'           UmurBulan = 27.8463892485015 hari
'                     = 27 Hari 20 jam 18 menit 48 detik
'
'----------------------------
'* Cibinong, 09 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, D, M, JD, JDI As Double
   
   JD = JulianDatum(Year, Month, Day)
   JDI = NewMoon(Year, Month, Day)
   If JDI > JD Then
      Y = JD2Year(JDI)
      M = JD2Month(JDI)
      D = JD2Date(JDI) + JD2Time(JDI) - 30#
      JDI = NewMoon(Y, M, D)
   End If
   MoonsAge = JD - JDI
End Function

Public Function MoonsAgeAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                        Optional ByVal Alt As Double = -1#) As Double
' Function MoonsAgeAtSunset dimaksudkan untuk menghitung umur bulan saat Terbenam Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Alt     : Posisi Matahari terhadap ufuk dalam derajat busur.
'                Untuk perhitungan yang tidak memerlukan ketelitian tinggi, maka Alt=1
'                Untuk perhitungan yang teliti, maka Alt harus dihitung berdasarkan koreksi
'                Semidiameter dan Refraksi. lihat fungsi SunAltitudeAtSunsetOrSunrise
'
' OUTPUT
'    Umur Bulan dalam hari
'
' CATATAN:
'   Umur Bulan yang dihitung sejak terjadinya konjungsi terdekat sebelum Matahari terbenam
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Umur Bulan di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 19 Juli 2012 pada saat Matahari Terbenam dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Mempersiapkan opsi perhitungan sebagai input
'    3. Hitung tinggi Matahari saat terbenam, lihat SunAltitudeAtSunsetOrSunrise
'    4. Jalankan fungsi MoonsAgeAtSunset.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2012
'           Bulan = 7
'           Tanggal = 19
'           p =  1010.
'           T = 27.
'           Alt = SunAltitudeAtSunsetOrSunrise(Tahun, Bulan, Tanggal, p, T, h) / 60#
'
'           UmurBulan = MoonsAgeAtSunset(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt)
'
'    Hasil:
'           UmurBulan = hari 0.270443771034479
'                     = 6 jam 29 menit 26 detik
'
'----------------------------
'* Cibinong, 09 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, D, M, Terbenam, JDTerbenam, JDIjtimak As Double
   
   Terbenam = Sunset(Lintang, Bujur, TZ, Year, Month, Day, Alt)
   JDTerbenam = JulianDatum(Year, Month, Int(Day) + (Terbenam - TZ) / 24#)
   JDIjtimak = NewMoon(Year, Month, Day)
   If JDIjtimak > JDTerbenam Then
      Y = JD2Year(JDIjtimak)
      M = JD2Month(JDIjtimak)
      D = JD2Date(JDIjtimak) + JD2Time(JDIjtimak) - 30#
      JDIjtimak = NewMoon(Y, M, D)
   End If
   MoonsAgeAtSunset = JDTerbenam - JDIjtimak
End Function

Public Function MoonIllumination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function MoonIllumination dimaksudkan untuk menghitung Illuminasi (kecerlangan) bulan
'
' INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Illuminasi Bulan (dalam procent), jika pada saat FullMoon mendekati 100% atau mendekati 1
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Illuminasi Bulan Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi MoonIllumination.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonIllumination(Y, M, D)
'    Hasil:
'         x = 0.0343441032096664
'           = 3.43%
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 46 hal 315
'----------------------------
'*Cibinong, 14 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
    Dim Elo, R, D, i, k As Double
   
    Elo = Radians(MoonElongation(Year, Month, Day))
    R = DistanceToTheSunVSOP87(Year, Month, Day) * 149597870.691
    D = DistanceToTheMoon(Year, Month, Day)
    i = Atn2(D - R * Cos(Elo), R * Sin(Elo))
    k = (1 + Cos(i)) / 2#
    MoonIllumination = k
End Function

Public Function MoonIlluminationAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, _
                        Optional ByVal Alt As Double = -1#) As Double
' Function MoonIlluminationAtSunset dimaksudkan untuk menghitung Illuminasi (kecerlangan) bulan
' pada saat Terbenam Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'      Alt     : Posisi Matahari terhadap ufuk dalam derajat busur.
'                Untuk perhitungan yang tidak memerlukan ketelitian tinggi, maka Alt=1
'                Untuk perhitungan yang teliti, maka Alt harus dihitung berdasarkan koreksi
'                Semidiameter dan Refraksi. lihat fungsi SunAltitudeAtSunsetOrSunrise
'
' OUTPUT
'    Kecerlangan bulan dalam hari
'
' CATATAN:
'    Illuminasi Bulan (dalam procent), jika pada saat FullMoon mendekati 100% atau mendekati 1
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Kecerlangan Bulan di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 19 Juli 2012 pada saat Matahari Terbenam dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Mempersiapkan opsi perhitungan sebagai input
'    3. Hitung tinggi Matahari saat terbenam, lihat SunAltitudeAtSunsetOrSunrise
'    4. Jalankan fungsi MoonIlluminationAtSunset.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2012
'           Bulan = 7
'           Tanggal = 19
'           p =  1010.
'           T = 27.
'           Alt = SunAltitudeAtSunsetOrSunrise(Tahun, Bulan, Tanggal, p, T, h) / 60#
'
'           x = MoonIlluminationAtSunset(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal, Alt)
'
'    Hasil:
'           x = 0.00216568025802738
'             = 0.22%
'
'----------------------------
'* Cibinong, 09 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Terbenam As Double
   
   Terbenam = Sunset(Lintang, Bujur, TZ, Year, Month, Day, Alt)
   MoonIlluminationAtSunset = MoonIllumination(Year, Month, Int(Day) + (Terbenam - TZ) / 24#)
End Function

Public Function CrescentWidth(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function CrescentWidth dimaksudkan untuk menghitung Lebar Hilal pada saat tertentu
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' OUTPUT
'    Lebar hilal dalam menit busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Lebar Hilal pada tangal 19 Juli 2012 jam 10:30:15 WIB dilakukan
' dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data Tahun, Bulan dan Tanggal.
'    2. Jalankan fungsi CrescentWidth.
'           Tahun = 2012
'           Bulan = 7
'           Tanggal = 19 + (10. + 30./60. + 15./3600.)/24. - 7./24
'           x = CrescentWidth(Tahun, Bulan, Tanggal)
'    Hasil:
'         x = 0.0387762406235933
'           = 0' 2.33"
'
'----------------------------
'*Pontianak, 19 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
 '-------------------------------------------------
   
   Dim mSD, illum, MTerbenam, Alt As Double
   mSD = MoonSemidiameter(Year, Month, Day)
   illum = MoonIllumination(Year, Month, Day)
   CrescentWidth = illum * 2# * mSD
End Function

Public Function CrescentWidthAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function CrescentWidthAtSunset dimaksudkan untuk menghitung Lebar Hilal pada saat matahari terbenam
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' OUTPUT
'    Lebar hilal dalam menit busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Lebar Hilal pada saat matahari terbenam di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 19 Juli 2012 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Jalankan fungsi CrescentWidthAtSunset.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2012
'           Bulan = 7
'           Tanggal = 19
'           x = CrescentWidthAtSunset(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'         x = 0.0661222387145634
'           = 0' 3.97"
'
'----------------------------
'*Pontianak, 19 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
 '-------------------------------------------------
   
   Dim mSD, illum, MTerbenam, Alt As Double
   
   Alt = SunAltitudeAtSunsetOrSunrise(Year, Month, Day, 1010, 27#, 0#) / 60#
   MTerbenam = Sunset(Lintang, Bujur, TZ, Year, Month, Day, Alt)
   mSD = MoonSemidiameter(Year, Month, Int(Day) + (MTerbenam - TZ) / 24#)
   illum = MoonIllumination(Year, Month, Int(Day) + (MTerbenam - TZ) / 24#)
   CrescentWidthAtSunset = illum * 2# * mSD
End Function


Public Function AzimuthDifferenceAtSunset(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function AzimuthDifferenceAtSunset dimaksudkan untuk menghitung Beda Azimuth Bulan dan Matahari
' pada saat matahari terbenam
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' OUTPUT
'    Beda Azimuth Bulan dan Matahari saat Matahari terbenam dinyatakan dalam derajat
'
' CATATAN:
'    - Jika Bulan berada di sebelah utara Matahari maka beda azimuthnya positif
'    - Sedangkan beda azimuth negatif apabila Bulan berada di selatan Matahari
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Beda Azimuth antara Bulan dan Matahari pada saat matahari terbenam
' di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 19 Juli 2012 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Jalankan fungsi AzimuthDifferenceAtSunset.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2012
'           Bulan = 7
'           Tanggal = 19
'           x = AzimuthDifferenceAtSunset(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'         x = -4.52041368787786
'           = -4° 31' 13.49"
'
'----------------------------
'*Pontianak, 19 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
 '-------------------------------------------------
   
   Dim MTerbenam, Alt, Maz, Baz As Double
   
   Alt = SunAltitudeAtSunsetOrSunrise(Year, Month, Day, 1010, 27#, 0#) / 60#
   MTerbenam = Sunset(Lintang, Bujur, TZ, Year, Month, Day, Alt)
   Maz = SunAzimuthVSOP87(Lintang, Bujur, Year, Month, Int(Day) + (MTerbenam - TZ) / 24#)
   Baz = MoonAzimuth(Lintang, Bujur, Year, Month, Int(Day) + (MTerbenam - TZ) / 24#)
   AzimuthDifferenceAtSunset = Baz - Maz
End Function

Public Function AzimuthDifference(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function AzimuthDifference dimaksudkan untuk menghitung Beda Azimuth Bulan dan Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT
'    Beda Azimuth Bulan dan Matahari dinyatakan dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Beda Azimuth antara Bulan dan Matahari
' di Jakarta dengan koordinat (6° 8' S 106° 45'E) pada tangal 1 Februari 2010 jam 10:30:15 WIB
' dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Jalankan fungsi AzimuthDifference.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12.14600694444440
'           x = AzimuthDifference(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'         x = 59.0988929639763
'           = 59° 05' 56.01"
'
'----------------------------
'*Pontianak, 19 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
 '-------------------------------------------------
   
   Dim Maz, Baz As Double
   
   Maz = SunAzimuthVSOP87(Lintang, Bujur, Year, Month, Day)
   Baz = MoonAzimuth(Lintang, Bujur, Year, Month, Day)
   AzimuthDifference = Baz - Maz
End Function

Public Function MoonSet(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function MoonSet dimaksudkan untuk menghitung Waktu Bulan terbenam
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' OUTPUT
'    Waktu Bulan Terbenam dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Waktu Terbenam Bulan di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Jalankan fungsi MoonSet.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           x = MoonSet(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'         x = 16.999439612031
'           = 16:59:58 WIB
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 14 hal 98
'----------------------------
'*Singapura, 13 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim m0, h, h0, hh, M, DEC, RA As Double
  Dim DT, JDE, A1, A2, B1, B2, c1, c2 As Double
  Dim D, hset, La, GMST, dm, n, Theta As Double
  
  La = Radians(Lintang)
  
  DT = DeltaT(Year, Month, Day) / 86400#
  hset = 0.7275 * HorizontalMoonParallax(Year, Month, Day) - RefractionApparentAltitude(0#, 1010#, 27#) / 60#

  JDE = JulianEphemerisDay(Year, Month, Day)
  GMST = GreenwichMeanSiderealTime(Year, Month, Day) * 15#
  
  DEC = ApparentMoonDeclination(Year, Month, Day)
  RA = ApparentMoonRightAscention(Year, Month, Day)
  
  h0 = Acos((Sin(Radians(hset)) - Sin(La) * Sin(Radians(DEC))) / (Cos(La) * Cos(Radians(DEC))))
  
  A1 = DEC - ApparentMoonDeclination(Year, Month, Day - 1#)
  B1 = ApparentMoonDeclination(Year, Month, Day + 1#) - DEC
  c1 = B1 - A1
  A2 = RA - ApparentMoonRightAscention(Year, Month, Day - 1#)
  B2 = ApparentMoonRightAscention(Year, Month, Day + 1#) - RA
  c2 = B2 - A2
  
  m0 = Modulus((RA - Bujur - GMST) / 360#, 1)
  M = m0 + Degrees(h0) / 360#
  
  dm = 1#
  While dm > 0.0000001
    Theta = GMST + 360.985647 * M
    n = M + DT / 86400#
    D = Radians(DEC + 0.5 * n * (A1 + B1 + n * c1))
    h = Radians(Theta + Bujur - (RA + 0.5 * n * (A2 + B2 + n * c2)))
    hh = Degrees(Asin(Sin(La) * Sin(D) + Cos(La) * Cos(D) * Cos(h)))
    dm = (hh - hset) / (360 * Cos(La) * Cos(D) * Sin(h))
    M = M + dm
  Wend
  
  MoonSet = (JDE + M + TZ / 24# - DT) - JulianDatum(Year, Month, Int(Day))
  MoonSet = MoonSet * 24#
End Function

Public Function Moonrise(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function Moonrise dimaksudkan untuk menghitung Waktu Bulan terbit
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' OUTPUT
'    Waktu Bulan Terbit dalam waktu lokal
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Waktu Terbit Bulan di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Jalankan fungsi Moonrise.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           x = moonrise(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'         x = 4.35744645819068
'           = 04:21:27 WIB
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 14 hal 98
'----------------------------
'*Singapura, 13 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim m0, h, h0, hh, M, DEC, RA, DT, JDE, A1, A2, B1, B2, c1, c2 As Double
  Dim D, hrise, La, GMST, dm, n, Theta, S As Double
  
  La = Radians(Lintang)
  
  DT = DeltaT(Year, Month, Day) / 86400#
  hrise = 0.7275 * HorizontalMoonParallax(Year, Month, Day) - RefractionApparentAltitude(0#, 1010#, 27#) / 60#

  JDE = JulianEphemerisDay(Year, Month, Day)
  GMST = GreenwichMeanSiderealTime(Year, Month, Day) * 15#
  
  DEC = ApparentMoonDeclination(Year, Month, Day)
  RA = ApparentMoonRightAscention(Year, Month, Day)
  
  h0 = Acos((Sin(Radians(hrise)) - Sin(La) * Sin(Radians(DEC))) / (Cos(La) * Cos(Radians(DEC))))
  
  A1 = DEC - ApparentMoonDeclination(Year, Month, Day - 1#)
  B1 = ApparentMoonDeclination(Year, Month, Day + 1#) - DEC
  c1 = B1 - A1
  A2 = RA - ApparentMoonRightAscention(Year, Month, Day - 1#)
  B2 = ApparentMoonRightAscention(Year, Month, Day + 1#) - RA
  c2 = B2 - A2
  
  m0 = Modulus((RA - Bujur - GMST) / 360#, 1)
  M = m0 - Degrees(h0) / 360#
  
  dm = 1#
  While dm > 0.0000001
    Theta = GMST + 360.985647 * M
    n = M + DT / 86400#
    D = Radians(DEC + 0.5 * n * (A1 + B1 + n * c1))
    h = Radians(Theta + Bujur - (RA + 0.5 * n * (A2 + B2 + n * c2)))
    hh = Degrees(Asin(Sin(La) * Sin(D) + Cos(La) * Cos(D) * Cos(h)))
    dm = (hh - hrise) / (360 * Cos(La) * Cos(D) * Sin(h))
    M = M + dm
  Wend
  
  Moonrise = (JDE + M + TZ / 24# - DT) - JulianDatum(Year, Month, Int(Day))
  Moonrise = Moonrise * 24#
End Function

Public Function MoonTransit(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function MoonTransit dimaksudkan untuk menghitung Waktu Bulan saat transit
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' OUTPUT
'    Waktu Bulan saat transit dinyatan dalam Julian Day
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Waktu Terbenam MataBulan di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 12 Februari 2010 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Jalankan fungsi MoonTransit.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2010
'           Bulan = 2
'           Tanggal = 12
'           x = MoonTransit(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'         x = 2455239.94525875
'           = Jum'at, 12 Februari 2010 10:41:10 WIB
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 14 hal 98
'----------------------------
'*Singapura, 13 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim RA, DT, JDE, Theta, GMST, A, B, c, dm, n, h, M As Double
  
  JDE = JulianEphemerisDay(Year, Month, Day)
  GMST = GreenwichMeanSiderealTime(Year, Month, Day) * 15#
  
  DT = DeltaT(Year, Month, Day) / 86400#
  RA = ApparentMoonRightAscention(Year, Month, Day)
  A = RA - ApparentMoonRightAscention(Year, Month, Day - 1#)
  B = ApparentMoonRightAscention(Year, Month, Day + 1#) - RA
  c = B - A
  
  M = Modulus((RA - Bujur - GMST) / 360#, 1)
  
  dm = 1#
  While dm > 0.0000001
    Theta = GMST + 360.985647 * M
    n = M + DT / 86400#
    h = Theta + Bujur - (RA + 0.5 * n * (A + B + n * c))
    dm = -h / 360#
    M = M + dm
  Wend
  
  MoonTransit = JDE + M + TZ / 24# - DT
End Function


Public Function MoonAltitude(ByVal Lintang As Double, ByVal Bujur As Double, _
                             ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function MoonAltitude dimaksudkan untuk menghitung Tinggi Bulan di atas ufuk
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Ketinggian Bulan dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Tinggi Bulan di atas ufuk Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' di Jakarta dengan koordinat (6° 8' S 106° 45'E) dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi MoonAltitude.
'           L = -6. - 8./60.
'           B = 106. + 45./60.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonAltitude(L, B, Y, M, D)
'    Hasil:
'         x = 77.1429299018723
'           = 77° 08' 34.55"
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985, hal. 12
'----------------------------
'*Demak, 18 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim La, h, DEC As Double
   
   La = Radians(Lintang)
   h = Radians(HourAngleOfTheMoon(Bujur, Year, Month, Day) * 15#)
   DEC = Radians(ApparentMoonDeclination(Year, Month, Day))
   MoonAltitude = Cos(h) * Cos(DEC) * Cos(La) + Sin(DEC) * Sin(La)
   MoonAltitude = Degrees(Asin(MoonAltitude))
End Function

Public Function ApparentMoonAltitude(ByVal Lintang As Double, ByVal Bujur As Double, _
                             ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function ApparentMoonAltitude dimaksudkan untuk menghitung Tinggi Bulan Nampak di atas ufuk
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Ketinggian Bulan Nampak dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Tinggi Bulan di atas ufuk Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' di Jakarta dengan koordinat (6° 8' S 106° 45'E) dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ApparentMoonAltitude.
'           L = -6. - 8./60.
'           B = 106. + 45./60.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentMoonAltitude(L, B, Y, M, D)
'    Hasil:
'         x = 64.1017246735075
'           = 64° 06' 06.209"
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985, hal. 12
'----------------------------
'*Demak, 18 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim h As Double
   
   h = MoonAltitude(Lintang, Bujur, Year, Month, Day)
   ApparentMoonAltitude = h + RefractionTrueAltitude(h, 1010, 27#) / 60#
End Function

Public Function MoonAzimuth(ByVal Lintang As Double, ByVal Bujur As Double, _
                            ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function MoonAzimuth dimaksudkan untuk menghitung Azimuth Bulan
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Azimuth Bulan dinyatakan dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Azimut Bulan Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' di Jakarta dengan koordinat (6° 8' S 106° 45'E) dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi MoonAzimuth.
'           L = -6. - 8./60.
'           B = 106. + 45./60.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonAzimuth(L, B, Y, M, D)
'    Hasil:
'         x = 168.666024929864
'           = 168° 39' 57.69"
'
'----------------------------
'*Cibinong, 18 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim h, DEC, cosA As Double
  
   h = Radians(HourAngleOfTheMoon(Bujur, Year, Month, Day) * 15#)
   DEC = Radians(ApparentMoonDeclination(Year, Month, Day))
   cosA = Cos(h) * Sin(Radians(Lintang)) - Tan(DEC) * Cos(Radians(Lintang))
   MoonAzimuth = Modulus(Degrees(Atn2(cosA, Sin(h))) + 180#, 360#)
End Function

Public Function HilalDuration(ByVal Lintang As Double, ByVal Bujur As Double, ByVal TZ As Double, _
                        ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function HilalDuration dimaksudkan untuk menghitung Lama hilal sejak matahari terbenam
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      Lintang : Koordinat Lintang tempat yang dihitung waktu fajrnya
'      Bujur   : Koordinat Bujur tempat yang dihitung waktu fajrnya
'      TZ      : Zona waktu (Time Zone)
'      Year    : tahun Gregorian/Julian (tahun Masehi)
'      Month   : bulan Masehi (1-12)
'      Day     : hari dalam decimal
'
' OUTPUT
'    Lama Hilal dalam menit
'
'CATATAN:
'   - Fungsi ini untuk menghitung selisih waktu antara Bulan terbenam dan Matahari terbenam
'   - Selisih waktu positif, jika Matahari terbenam terlebih dahulu yang disusul Bulan terbenam.
'     Saat bulan terbenam lebih dulu, maka lama hilal negatif
'   - Fungsi ini hanya menghitung apabila nilai absolut lama hilal tidak lebih dari 100 menit
'     Jika lama hilal lebih dari 100 menit maka Fungsi ini akan menghasilkan nila 999.
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Lama Hilal di Jakarta dengan koordinat (6° 8' S 106° 45'E)
' pada tangal 19 Juli 2012 dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mempersiapkan data lintang dan bujur, zona waktu, Tahun, Bulan dan Tanggal.
'    2. Jalankan fungsi HilalDuration.
'           Lintang = -6 - 8/60.
'           Bujur = 106 + 45./60.
'           Tz = 7.
'           Tahun = 2012
'           Bulan = 7
'           Tanggal = 19
'           x = HilalDuration(Lintang, Bujur, TZ, Tahun, Bulan, Tanggal)
'    Hasil:
'         x = 8.09243147289862
'           = 8 menit 6 detik
'
'----------------------------
'*Pontianak, 19 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   
   Dim Alt, MTerbenam, BTerbenam As Double
   
   Alt = SunAltitudeAtSunsetOrSunrise(Year, Month, Day, 1010, 27#, 0#) / 60#
   MTerbenam = Sunset(Lintang, Bujur, TZ, Year, Month, Day, Alt)
   BTerbenam = MoonSet(Lintang, Bujur, TZ, Year, Month, Day)
   HilalDuration = (BTerbenam - MTerbenam) * 60#
   If Abs(HilalDuration) > 100# Then HilalDuration = 999#
End Function

Public Function MoonSemidiameter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function MoonSemidiameter dimaksudkan untuk menghitung Semidiameter Bulan
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal (Waktu Universal)
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Semidiameter Bulan dalam menit busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Semidiameter Bulan Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapka
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi MoonSemidiameter.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           sD = MoonSemidiameter(Y, M, D)
'    Hasil:
'         sD = 14.7071785557126
'            = 0° 14' 42.43"
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 53 hal 359
'----------------------------
'*Cibinong, 14 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim R As Double
     
  R = DistanceToTheMoon(Year, Month, Day)
  MoonSemidiameter = (358473400 / R) / 60#
End Function

Public Function MoonMeanAnomaly(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonMeanLongitude berfungsi untuk menghitung rata-rata Anomali Matahari
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Rata-rata Anomali matahari dalam derajat
'
' CATATAN:
'    Dalam Buku Astronomical Algorithm disimbolkan dengan M'
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung rata-rata Bujur Bulan pada 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonMeanAnomaly(Y, M, D)
'    Hasil:
'         x = 48405.4976442627
'           = 165° 29' 51.52"
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 132
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   T = JulianEphemerisCentury(Year, Month, Day)
   MoonMeanAnomaly = 134.9634114 + 477198.8676313 * T + 0.008997 * T ^ 2 + T ^ 3 / 69699 - T ^ 4 / 14712000
End Function


Public Function JDEofMaxNorthMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function JDEofMaxNorthMoonDeclination dipakai untuk menghitung saat terjadinya maximum deklinasi Bulan di belahan Utara.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'      Saat Bulan mencapai deklinasi maxmimum di belahan Utara dalam Julian Day.
'
' CATATAN:
'      Untuk mengkonversikan Julian Day dalam waktu sehari-hari, silahkan lihat fungsi Caldat
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 50.a — Greatest northern declination of the Moon in December 1988#
'      maka :
'      Y = 1988
'      M = 12
'      D = 15
'      x = JDEofMaxNorthMoonDeclination(Y, M, D)
'      x = 2447518.33399884
'        = Kamis, 22 Desember 1988 20:00:57 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab 50.
' ------------------------------------------
'* Cibinong, 23 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Tahun, Bulan As Integer
  Dim k, T, D, M, M1, F, E, P, JDE, Tanggal As Double
  
   k = value_K(Year, Month, Day)
   T = value_T(Year, Month, Day)
   D = Radians(152.2029 + 333.0705546 * k - 0.0004025 * T ^ 2 + 0.00000011 * T ^ 3)
   M = Radians(14.8591 + 26.9281592 * k - 0.0000544 * T ^ 2 - 0.0000001 * T ^ 3)
   M1 = Radians(4.6881 + 356.9562795 * k + 0.0103126 * T ^ 2 + 0.00001251 * T ^ 3)
   F = Radians(325.8867 + 1.4467806 * k - 0.0020708 * T ^ 2 - 0.00000215 * T ^ 3)
   E = nilai_E(Year, Month, Day)
   
   P = 0.8975 * Cos(F)
   P = P - 0.4726 * Sin(M1)
   P = P - 0.103 * Sin(2 * F)
   P = P - 0.0976 * Sin(2 * D - M1)
   P = P - 0.0462 * Cos(M1 - F)
   P = P - 0.0461 * Cos(M1 + F)
   P = P - 0.0438 * Sin(2 * D)
   P = P + 0.0162 * E * Sin(M)
   P = P - 0.0157 * Cos(3 * F)
   P = P + 0.0145 * Sin(M1 + 2 * F)
   P = P + 0.0136 * Cos(2 * D - F)
   P = P - 0.0095 * Cos(2 * D - M1 - F)
   P = P - 0.0091 * Cos(2 * D - M1 + F)
   P = P - 0.0089 * Cos(2 * D + F)
   P = P + 0.0075 * Sin(2 * M1)
   P = P - 0.0068 * Sin(M1 - 2 * F)
   P = P + 0.0061 * Cos(2 * M1 - F)
   P = P - 0.0047 * Sin(M1 + 3 * F)
   P = P - 0.0043 * E * Sin(2 * D - M - M1)
   P = P - 0.004 * Cos(M1 - 2 * F)
   P = P - 0.0037 * Sin(2 * D - 2 * M1)
   P = P + 0.0031 * Sin(F)
   P = P + 0.003 * Sin(2 * D + M1)
   P = P - 0.0029 * Cos(M1 + 2 * F)
   P = P - 0.0029 * E * Sin(2 * D - M)
   P = P - 0.0027 * Sin(M1 + F)
   P = P + 0.0024 * E * Sin(M - M1)
   P = P - 0.0021 * Sin(M1 - 3 * F)
   P = P + 0.0019 * Sin(2 * M1 + F)
   P = P + 0.0018 * Cos(2 * D - 2 * M1 - F)
   P = P + 0.0018 * Sin(3 * F)
   P = P + 0.0017 * Cos(M1 + 3 * F)
   P = P + 0.0017 * Cos(2 * M1)
   P = P - 0.0014 * Cos(2 * D - M1)
   P = P + 0.0013 * Cos(2 * D + M1 + F)
   P = P + 0.0013 * Cos(M1)
   P = P + 0.0012 * Sin(3 * M1 + F)
   P = P + 0.0011 * Sin(2 * D - M1 + F)
   P = P - 0.0011 * Cos(2 * D - 2 * M1)
   P = P + 0.001 * Cos(D + F)
   P = P + 0.001 * E * Sin(M + M1)
   P = P - 0.0009 * Sin(2 * D - 2 * F)
   P = P + 0.0007 * Cos(2 * M1 + F)
   P = P - 0.0007 * Cos(3 * M1 + F)
   JDE = 2451562.5897 + 27.321582241 * k + 0.000100695 * T ^ 2 - 0.000000141 * T ^ 3 + P
   Tahun = JD2Year(JDE)
   Bulan = JD2Month(JDE)
   Tanggal = JD2Date(JDE) + JD2Time(JDE)
   JDEofMaxNorthMoonDeclination = JDE - DeltaT(Tahun, Bulan, Tanggal) / 86400#
End Function

Public Function JDEofMaxSouthMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function JDEofMaxSouthMoonDeclination dipakai untuk menghitung saat terjadinya maximum deklinasi Bulan di belahan Selatan.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'      Saat Bulan mencapai deklinasi maxmimum di belahan selatan dalam Julian Day.
'
' CATATAN:
'      Untuk mengkonversikan Julian Day dalam waktu sehari-hari, silahkan lihat fungsi Caldat
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 50.b — If we calculate the maximum southern declination
'   for k = +659, we obtain JDE = 2469 553.0834, which  corresponds to 2049 April 21, at 14h TD, and
'   d = 22.1384, so the greatest southern declination is -22° 08'.
'
'      maka :
'      Y = 2049
'      M = 4
'      D = 15
'      x = JDEofMaxSouthMoonDeclination(Y, M, D)
'      x = 2469553.08236493
'        = Rabu, 21 April 2049 13:58:36 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab 50.
' ------------------------------------------
'* Cibinong, 23 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim Tahun, Bulan As Integer
  Dim k, T, D, M, M1, F, E, P, JDE, Tanggal As Double
  
   k = value_K(Year, Month, Day)
   T = value_T(Year, Month, Day)
   D = Radians(345.6676 + 333.0705546 * k - 0.0004025 * T ^ 2 + 0.00000011 * T ^ 3)
   M = Radians(1.3951 + 26.9281592 * k - 0.0000544 * T ^ 2 - 0.0000001 * T ^ 3)
   M1 = Radians(186.21 + 356.9562795 * k + 0.0103126 * T ^ 2 + 0.00001251 * T ^ 3)
   F = Radians(145.1633 + 1.4467806 * k - 0.0020708 * T ^ 2 - 0.00000215 * T ^ 3)
   E = nilai_E(Year, Month, Day)
   
   P = -0.8975 * Cos(F) - 0.4726 * Sin(M1)
   P = P - 0.103 * Sin(2 * F)
   P = P - 0.0976 * Sin(2 * D - M1)
   P = P + 0.0541 * Cos(M1 - F)
   P = P + 0.0516 * Cos(M1 + F)
   P = P - 0.0438 * Sin(2 * D)
   P = P + 0.0112 * Sin(M) * E
   P = P + 0.0157 * Cos(3 * F)
   P = P + 0.0023 * Sin(M1 + 2 * F)
   P = P - 0.0136 * Cos(2 * D - F)
   P = P + 0.011 * Cos(2 * D - M1 - F)
   P = P + 0.0091 * Cos(2 * D - M1 + F)
   P = P + 0.0089 * Cos(2 * D + F)
   P = P + 0.0075 * Sin(2 * M1)
   P = P - 0.003 * Sin(M1 - 2 * F)
   P = P - 0.0061 * Cos(2 * M1 - F)
   P = P - 0.0047 * Sin(M1 + 3 * F)
   P = P - 0.0043 * Sin(2 * D - M - M1) * E
   P = P + 0.004 * Cos(M1 - 2 * F)
   P = P - 0.0037 * Sin(2 * D - 2 * M1)
   P = P - 0.0031 * Sin(F)
   P = P + 0.003 * Sin(2 * D + M1)
   P = P + 0.0029 * Cos(M1 + 2 * F)
   P = P - 0.0029 * Sin(2 * D - M) * E
   P = P - 0.0027 * Sin(M1 + F)
   P = P + 0.0024 * Sin(M - M1) * E
   P = P - 0.0021 * Sin(M1 - 3 * F)
   P = P - 0.0019 * Sin(2 * M1 + F)
   P = P - 0.0006 * Cos(2 * D - 2 * M1 - F)
   P = P - 0.0018 * Sin(3 * F)
   P = P - 0.0017 * Cos(M1 + 3 * F)
   P = P + 0.0017 * Cos(2 * M1)
   P = P + 0.0014 * Cos(2 * D - M1)
   P = P - 0.0013 * Cos(2 * D + M1 + F)
   P = P - 0.0013 * Cos(M1)
   P = P + 0.0012 * Sin(3 * M1 + F)
   P = P + 0.0011 * Sin(2 * D - M1 + F)
   P = P + 0.0011 * Cos(2 * D - 2 * M1)
   P = P + 0.001 * Cos(D + F)
   P = P + 0.001 * Sin(M + M1) * E
   P = P - 0.0009 * Sin(2 * D - 2 * F)
   P = P - 0.0007 * Cos(2 * M1 + F)
   P = P - 0.0007 * Cos(3 * M1 + F)

   JDE = 2451548.9289 + 27.321582241 * k + 0.000100695 * T ^ 2 - 0.000000141 * T ^ 3 + P
   Tahun = JD2Year(JDE)
   Bulan = JD2Month(JDE)
   Tanggal = JD2Date(JDE) + JD2Time(JDE)
   JDEofMaxSouthMoonDeclination = JDE - DeltaT(Tahun, Bulan, Tanggal) / 86400#
End Function

Public Function MaxNorthMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function MaxNorthMoonDeclination dipakai untuk menghitung maximum deklinasi Bulan di belahan Utara.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'      Saat Bulan mencapai deklinasi maxmimum di belahan Utara dalam derajat.
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 50.a — Greatest northern declination of the Moon in December 1988#
'      maka :
'      Y = 1988
'      M = 12
'      D = 15
'      x = MaxNorthMoonDeclination(Y, M, D)
'      x = 28.1562340347414
'        = 28° 09' 22.44"
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab 50.
' ------------------------------------------
'* Cibinong, 23 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim k, T, D, M, M1, F, E, P As Double
  
   k = value_K(Year, Month, Day)
   T = value_T(Year, Month, Day)
   D = Radians(152.2029 + 333.0705546 * k - 0.0004025 * T ^ 2 + 0.00000011 * T ^ 3)
   M = Radians(14.8591 + 26.9281592 * k - 0.0000544 * T ^ 2 - 0.0000001 * T ^ 3)
   M1 = Radians(4.6881 + 356.9562795 * k + 0.0103126 * T ^ 2 + 0.00001251 * T ^ 3)
   F = Radians(325.8867 + 1.4467806 * k - 0.0020708 * T ^ 2 - 0.00000215 * T ^ 3)
   E = nilai_E(Year, Month, Day)
  
   P = 5.1093 * Sin(F)
   P = P + 0.2658 * Cos(2 * F)
   P = P + 0.1448 * Sin(2 * D - F)
   P = P - 0.0322 * Sin(3 * F)
   P = P + 0.0133 * Cos(2 * D - 2 * F)
   P = P + 0.0125 * Cos(2 * D)
   P = P - 0.0124 * Sin(M1 - F)
   P = P - 0.0101 * Sin(M1 + 2 * F)
   P = P + 0.0097 * Cos(F)
   P = P - 0.0087 * E * Sin(2 * D + M - F)
   P = P + 0.0074 * Sin(M1 + 3 * F)
   P = P + 0.0067 * Sin(D + F)
   P = P + 0.0063 * Sin(M1 - 2 * F)
   P = P + 0.006 * E * Sin(2 * D - M - F)
   P = P - 0.0057 * Sin(2 * D - M1 - F)
   P = P - 0.0056 * Cos(M1 + F)
   P = P + 0.0052 * Cos(M1 + 2 * F)
   P = P + 0.0041 * Cos(2 * M1 + F)
   P = P - 0.004 * Cos(M1 - 3 * F)
   P = P + 0.0038 * Cos(2 * M1 - F)
   P = P - 0.0034 * Cos(M1 - 2 * F)
   P = P - 0.0029 * Sin(2 * M1)
   P = P + 0.0029 * Sin(3 * M1 + F)
   P = P - 0.0028 * E * Cos(2 * D + M - F)
   P = P - 0.0028 * Cos(M1 - F)
   P = P - 0.0023 * Cos(3 * F)
   P = P - 0.0021 * Sin(2 * D + F)
   P = P + 0.0019 * Cos(M1 + 3 * F)
   P = P + 0.0018 * Cos(D + F)
   P = P + 0.0017 * Sin(2 * M1 - F)
   P = P + 0.0015 * Cos(3 * M1 + F)
   P = P + 0.0014 * Cos(2 * D + 2 * M1 + F)
   P = P - 0.0012 * Sin(2 * D - 2 * M1 - F)
   P = P - 0.0012 * Cos(2 * M1)
   P = P - 0.001 * Cos(M1)
   P = P - 0.001 * Sin(2 * F)
   P = P + 0.0006 * Sin(M1 + F)
   MaxNorthMoonDeclination = 23.6961 - 0.013004 * T + P
End Function

Public Function MaxSouthMoonDeclination(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function MaxSouthMoonDeclination dipakai untuk menghitung maximum deklinasi Bulan di belahan Selatan.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' OUTPUT:
'      Saat Bulan mencapai deklinasi maxmimum di belahan Selatan dalam derajat.
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 50.b — If we calculate the maximum southern declination
'   for k = +659, we obtain JDE = 2469 553.0834, which  corresponds to 2049 April 21, at 14h TD, and
'   d = 22.1384, so the greatest southern declination is -22° 08'.
'
'      maka :
'      Y = 2049
'      M = 4
'      D = 15
'      x = MaxSouthMoonDeclination(Y, M, D)
'      x = 22.1383613967941
'        = 22° 08' 18.10"
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab 50.
' ------------------------------------------
'* Cibinong, 23 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim k, T, D, M, M1, F, E, P As Double
  
   k = value_K(Year, Month, Day)
   T = value_T(Year, Month, Day)
   D = Radians(345.6676 + 333.0705546 * k - 0.0004025 * T ^ 2 + 0.00000011 * T ^ 3)
   M = Radians(1.3951 + 26.9281592 * k - 0.0000544 * T ^ 2 - 0.0000001 * T ^ 3)
   M1 = Radians(186.21 + 356.9562795 * k + 0.0103126 * T ^ 2 + 0.00001251 * T ^ 3)
   F = Radians(145.1633 + 1.4467806 * k - 0.0020708 * T ^ 2 - 0.00000215 * T ^ 3)
   E = nilai_E(Year, Month, Day)
  
   P = -5.1093 * Sin(F)
   P = P + 0.2658 * Cos(2 * F)
   P = P - 0.1448 * Sin(2 * D - F)
   P = P + 0.0322 * Sin(3 * F)
   P = P + 0.0133 * Cos(2 * D - 2 * F)
   P = P + 0.0125 * Cos(2 * D)
   P = P - 0.0015 * Sin(M1 - F)
   P = P + 0.0101 * Sin(M1 + 2 * F)
   P = P - 0.0097 * Cos(F)
   P = P + 0.0087 * E * Sin(2 * D + M - F)
   P = P + 0.0074 * Sin(M1 + 3 * F)
   P = P + 0.0067 * Sin(D + F)
   P = P - 0.0063 * Sin(M1 - 2 * F)
   P = P - 0.006 * E * Sin(2 * D - M - F)
   P = P + 0.0057 * Sin(2 * D - M1 - F)
   P = P - 0.0056 * Cos(M1 + F)
   P = P - 0.0052 * Cos(M1 + 2 * F)
   P = P - 0.0041 * Cos(2 * M1 + F)
   P = P - 0.004 * Cos(M1 - 3 * F)
   P = P - 0.0038 * Cos(2 * M1 - F)
   P = P + 0.0034 * Cos(M1 - 2 * F)
   P = P - 0.0029 * Sin(2 * M1)
   P = P + 0.0029 * Sin(3 * M1 + F)
   P = P + 0.0028 * E * Cos(2 * D + M - F)
   P = P - 0.0028 * Cos(M1 - F)
   P = P + 0.0023 * Cos(3 * F)
   P = P + 0.0021 * Sin(2 * D + F)
   P = P + 0.0019 * Cos(M1 + 3 * F)
   P = P + 0.0018 * Cos(D + F)
   P = P - 0.0017 * Sin(2 * M1 - F)
   P = P + 0.0015 * Cos(3 * M1 + F)
   P = P + 0.0014 * Cos(2 * D + 2 * M1 + F)
   P = P + 0.0012 * Sin(2 * D - 2 * M1 - F)
   P = P - 0.0012 * Cos(2 * M1)
   P = P + 0.001 * Cos(M1)
   P = P - 0.001 * Sin(2 * F)
   P = P + 0.0037 * Sin(M1 + F)
   
   MaxSouthMoonDeclination = 23.6961 - 0.013004 * T + P
End Function

Public Function MoonArgumentOfLatitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonArgumentOfLatitude berfungsi untuk menghitung Agumen Lintang ulan
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Argumen Lintang Bulan dalam derajat
'
' CATATAN:
'    Dalam Buku Astronomical Algorithm disimbolkan dengan F
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung rata-rata Bujur Bulan pada 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonArgumentOfLatitude(Y, M, D)
'    Hasil:
'         x = 48971.0483251973
'           = 11° 02' 53.97"
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 308
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   MoonArgumentOfLatitude = 93.2720993 + 483202.0175273 * T - 0.0034029 * T ^ 2 - T ^ 3 / 3526000# + T ^ 4 / 863310000#
End Function

Public Function SigmaB(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SigmaB berfungsi untuk menghitung bagian dari Lintang Bulan
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    SigmaB dalam seper sejuta derajat
'
' CATATAN:
'    Dalam Buku Astronomical Algorithm disimbolkan dengan Sigma B
'    Lihat MoonLatitude
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Sigma B pada 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SigmaB(Y, M, D)
'    Hasil:
'         x = 1006470.50688368
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 311-312
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    Dim L1, D, M, M1, F, A1, A2, A3, E, sB As Double
   
    L1 = Radians(MoonMeanLongitude(Year, Month, Day))
    D = Radians(MoonMeanElongation(Year, Month, Day))
    M = Radians(SunMeanAnomaly(Year, Month, Day))
    M1 = Radians(MoonMeanAnomaly(Year, Month, Day))
    F = Radians(MoonArgumentOfLatitude(Year, Month, Day))
    A1 = Radians(Argument_A1(Year, Month, Day))
    A2 = Radians(Argument_A2(Year, Month, Day))
    A3 = Radians(Argument_A3(Year, Month, Day))
    E = nilai_E(Year, Month, Day)
   
    sB = 0#
    sB = sB + 5128122 * Sin(F)
    sB = sB + 280602 * Sin(M1 + F)
    sB = sB + 277693 * Sin(M1 - F)
    sB = sB + 173237 * Sin(2 * D - F)
    sB = sB + 55413 * Sin(2 * D - M1 + F)
    sB = sB + 46271 * Sin(2 * D - M1 - F)
    sB = sB + 32573 * Sin(2 * D + F)
    sB = sB + 17198 * Sin(2 * M1 + F)
    sB = sB + 9266 * Sin(2 * D + M1 - F)
    sB = sB + 8822 * Sin(2 * M1 - F)
    sB = sB + 8216 * E * Sin(2 * D - M - F)
    sB = sB + 4324 * Sin(2 * D - 2 * M1 - F)
    sB = sB + 4200 * Sin(2 * D + M1 + F)
    sB = sB - 3359 * E * Sin(2 * D + M - F)
    sB = sB + 2463 * E * Sin(2 * D - M - M1 + F)
    sB = sB + 2211 * E * Sin(2 * D - M + F)
    sB = sB + 2065 * E * Sin(2 * D - M - M1 - F)
    sB = sB - 1870 * E * Sin(M - M1 - F)
    sB = sB + 1828 * Sin(4 * D - M1 - F)
    sB = sB - 1794 * E * Sin(M + F)
    sB = sB - 1749 * Sin(3 * F)
    sB = sB - 1565 * E * Sin(M - M1 + F)
    sB = sB - 1491 * Sin(D + F)
    sB = sB - 1475 * E * Sin(M + M1 + F)
    sB = sB - 1410 * E * Sin(M + M1 - F)
    sB = sB - 1344 * E * Sin(M - F)
    sB = sB - 1335 * Sin(D - F)
    sB = sB + 1107 * Sin(3 * M1 + F)
    sB = sB + 1021 * Sin(4 * D - F)
    sB = sB + 833 * Sin(4 * D - M1 + F)
    sB = sB + 777 * Sin(M1 - 3 * F)
    sB = sB + 671 * Sin(4 * D - 2 * M1 + F)
    sB = sB + 607 * Sin(2 * D - 3 * F)
    sB = sB + 596 * Sin(2 * D + 2 * M1 - F)
    sB = sB + 491 * E * Sin(2 * D - M + M1 - F)
    sB = sB - 451 * Sin(2 * D - 2 * M1 + F)
    sB = sB + 439 * Sin(3 * M1 - F)
    sB = sB + 422 * Sin(2 * D + 2 * M1 + F)
    sB = sB + 421 * Sin(2 * D - 3 * M1 - F)
    sB = sB - 366 * E * Sin(2 * D + M - M1 + F)
    sB = sB - 351 * E * Sin(2 * D + M + F)
    sB = sB + 331 * Sin(4 * D + F)
    sB = sB + 315 * E * Sin(2 * D - M + M1 + F)
    sB = sB + 302 * E ^ 2 * Sin(2 * D - 2 * M - F)
    sB = sB - 283 * Sin(M1 + 3 * F)
    sB = sB - 229 * E * Sin(2 * D + M + M1 - F)
    sB = sB + 223 * E * Sin(D + M - F)
    sB = sB + 223 * E * Sin(D + M + F)
    sB = sB - 220 * E * Sin(M - 2 * M1 - F)
    sB = sB - 220 * E * Sin(2 * D + M - M1 - F)
    sB = sB - 185 * Sin(D + M1 + F)
    sB = sB + 181 * E * Sin(2 * D - M - 2 * M1 - F)
    sB = sB - 177 * E * Sin(M + 2 * M1 + F)
    sB = sB + 176 * Sin(4 * D - 2 * M1 - F)
    sB = sB + 166 * E * Sin(4 * D - M - M1 - F)
    sB = sB - 164 * Sin(D + M1 - F)
    sB = sB + 132 * Sin(4 * D + M1 - F)
    sB = sB - 119 * Sin(D - M1 - F)
    sB = sB + 115 * E * Sin(4 * D - M - F)
    sB = sB + 107 * E ^ 2 * Sin(2 * D - 2 * M + F)
    sB = sB - 2235 * Sin(L1) + 382 * Sin(A3) + 175 * Sin(A1 - F) + 175 * Sin(A1 + F) + 127 * Sin(L1 - M1) - 115 * Sin(L1 + M1)
    SigmaB = sB
End Function

Public Function SigmaL(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SigmaL berfungsi untuk menghitung bagian dari komponen Bujur Bulan.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    SigmaL dalam derajat
'
' CATATAN:
'    Dalam Buku Astronomical Algorithm disimbolkan dengan sigma L
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung sigma L pada 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SigmaL(Y, M, D)
'    Hasil:
'         x = 1565626.94050891
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 310
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    Dim L1, D, M, M1, F, A1, A2, A3, E, sL As Double
    
    L1 = Radians(MoonMeanLongitude(Year, Month, Day))
    D = Radians(MoonMeanElongation(Year, Month, Day))
    M = Radians(SunMeanAnomaly(Year, Month, Day))
    M1 = Radians(MoonMeanAnomaly(Year, Month, Day))
    F = Radians(MoonArgumentOfLatitude(Year, Month, Day))
    A1 = Radians(Argument_A1(Year, Month, Day))
    A2 = Radians(Argument_A2(Year, Month, Day))
    A3 = Radians(Argument_A3(Year, Month, Day))
    E = nilai_E(Year, Month, Day)

    
    sL = 0#
    sL = sL + 6288774 * Sin(M1)
    sL = sL + 1274027 * Sin(2 * D - M1)
    sL = sL + 658314 * Sin(2 * D)
    sL = sL + 213618 * Sin(2 * M1)
    sL = sL - 185116 * E * Sin(M)
    sL = sL - 114332 * Sin(2 * F)
    sL = sL + 58793 * Sin(2 * D - 2 * M1)
    sL = sL + 57066 * E * Sin(2 * D - M - M1)
    sL = sL + 53322 * Sin(2 * D + M1)
    sL = sL + 45758 * E * Sin(2 * D - M)
    sL = sL - 40923 * E * Sin(M - M1)
    sL = sL - 34720 * Sin(D)
    sL = sL - 30383 * E * Sin(M + M1)
    sL = sL + 15327 * Sin(2 * D - 2 * F)
    sL = sL - 12528 * Sin(M1 + 2 * F)
    sL = sL + 10980 * Sin(M1 - 2 * F)
    sL = sL + 10675 * Sin(4 * D - M1)
    sL = sL + 10034 * Sin(3 * M1)
    sL = sL + 8548 * Sin(4 * D - 2 * M1)
    sL = sL - 7888 * E * Sin(2 * D + M - M1)
    sL = sL - 6766 * E * Sin(2 * D + M)
    sL = sL - 5163 * Sin(D - M1)
    sL = sL + 4987 * E * Sin(D + M)
    sL = sL + 4036 * E * Sin(2 * D - M + M1)
    sL = sL + 3994 * E * Sin(2 * D + 2 * M1)
    sL = sL + 3861 * Sin(4 * D)
    sL = sL + 3665 * Sin(2 * D - 3 * M1)
    sL = sL - 2689 * E * Sin(M - 2 * M1)
    sL = sL - 2602 * Sin(2 * D - M1 + 2 * F)
    sL = sL + 2390 * E * Sin(2 * D - M - 2 * M1)
    sL = sL - 2348 * Sin(D + M1)
    sL = sL + 2236 * E ^ 2 * Sin(2 * D - 2 * M)
    sL = sL - 2120 * E * Sin(M + 2 * M1)
    sL = sL - 2069 * E ^ 2 * Sin(2 * M)
    sL = sL + 2048 * E ^ 2 * Sin(2 * D - 2 * M - M1)
    sL = sL - 1773 * Sin(2 * D + M1 - 2 * F)
    sL = sL - 1595 * Sin(2 * D + 2 * F)
    sL = sL + 1215 * E * Sin(4 * D - M - M1)
    sL = sL - 1110 * Sin(2 * M1 + 2 * F)
    sL = sL - 892 * Sin(3 * D - M1)
    sL = sL - 810 * E * Sin(2 * D + M + M1)
    sL = sL + 759 * E * Sin(4 * D - M - 2 * M1)
    sL = sL - 713 * E ^ 2 * Sin(2 * M - M1)
    sL = sL - 700 * E ^ 2 * Sin(2 * D + 2 * M - M1)
    sL = sL + 691 * E * Sin(2 * D + M - 2 * M1)
    sL = sL + 596 * E * Sin(2 * D - M - 2 * F)
    sL = sL + 549 * Sin(4 * D + M1)
    sL = sL + 537 * Sin(4 * M1)
    sL = sL + 520 * E * Sin(4 * D - M)
    sL = sL - 487 * Sin(D - 2 * M1)
    sL = sL - 399 * E * Sin(2 * D + M - 2 * F)
    sL = sL - 381 * Sin(2 * M1 - 2 * F)
    sL = sL + 351 * E * Sin(D + M + M1)
    sL = sL - 340 * Sin(3 * D - 2 * M1)
    sL = sL + 330 * Sin(4 * D - 3 * M1)
    sL = sL + 327 * E * Sin(2 * D - M + 2 * M1)
    sL = sL - 323 * E ^ 2 * Sin(2 * M + M1)
    sL = sL + 299 * E * Sin(D + M - M1)
    sL = sL + 294 * Sin(2 * D + 3 * M1)
    SigmaL = sL + 3958 * Sin(A1) + 1962 * Sin(L1 - F) + 318 * Sin(A2)
End Function


Public Function SigmaR(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi SigmaR berfungsi untuk menghitung bagian dari komponen Jarak ke Bulan.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    SigmaR dalam derajat
'
' CATATAN:
'    Dalam Buku Astronomical Algorithm disimbolkan dengan sigma r
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung sigma r pada 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SigmaR(Y, M, D)
'    Hasil:
'         x = 21233487.1378864
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 310
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    Dim L1, D, M, M1, F, A1, A2, A3, E, mD As Double
    
    L1 = Radians(MoonMeanLongitude(Year, Month, Day))
    D = Radians(MoonMeanElongation(Year, Month, Day))
    M = Radians(SunMeanAnomaly(Year, Month, Day))
    M1 = Radians(MoonMeanAnomaly(Year, Month, Day))
    F = Radians(MoonArgumentOfLatitude(Year, Month, Day))
    A1 = Radians(Argument_A1(Year, Month, Day))
    A2 = Radians(Argument_A2(Year, Month, Day))
    A3 = Radians(Argument_A3(Year, Month, Day))
    E = nilai_E(Year, Month, Day)
    
    mD = 0#
    mD = mD - 20905355 * Cos(M1)
    mD = mD - 3699111 * Cos(2 * D - M1)
    mD = mD - 2955968 * Cos(2 * D)
    mD = mD - 569925 * Cos(2 * M1)
    mD = mD + 48888 * E * Cos(M)
    mD = mD - 3149 * Cos(2 * F)
    mD = mD + 246158 * Cos(2 * D - 2 * M1)
    mD = mD - 152138 * E * Cos(2 * D - M - M1)
    mD = mD - 170733 * Cos(2 * D + M1)
    mD = mD - 204586 * E * Cos(2 * D - M)
    mD = mD - 129620 * E * Cos(M - M1)
    mD = mD + 108743 * Cos(D)
    mD = mD + 104755 * E * Cos(M + M1)
    mD = mD + 10321 * Cos(2 * D - 2 * F)
    mD = mD + 79661 * Cos(M1 - 2 * F)
    mD = mD - 34782 * Cos(4 * D - M1)
    mD = mD - 23210 * Cos(3 * M1)
    mD = mD - 21636 * Cos(4 * D - 2 * M1)
    mD = mD + 24208 * E * Cos(2 * D + M - M1)
    mD = mD + 30824 * E * Cos(2 * D + M)
    mD = mD - 8379 * Cos(D - M1)
    mD = mD - 16675 * E * Cos(D + M)
    mD = mD - 12831 * E * Cos(2 * D - M + M1)
    mD = mD - 10445 * Cos(2 * D + 2 * M1)
    mD = mD - 11650 * Cos(4 * D)
    mD = mD + 14403 * Cos(2 * D - 3 * M1)
    mD = mD - 7003 * E * Cos(M - 2 * M1)
    mD = mD + 10056 * E * Cos(2 * D - M - 2 * M1)
    mD = mD + 6322 * Cos(D + M1)
    mD = mD - 9884 * E ^ 2 * Cos(2 * D - 2 * M)
    mD = mD + 5751 * E * Cos(M + 2 * M1)
    mD = mD - 4950 * E ^ 2 * Cos(2 * D - 2 * M - M1)
    mD = mD + 4130 * Cos(2 * D + M1 - 2 * F)
    mD = mD - 3958 * E * Cos(4 * D - M - M1)
    mD = mD + 3258 * Cos(3 * D - M1)
    mD = mD + 2616 * E * Cos(2 * D + M + M1)
    mD = mD - 1897 * E * Cos(4 * D - M - 2 * M1)
    mD = mD - 2117 * E ^ 2 * Cos(2 * M - M1)
    mD = mD + 2354 * E ^ 2 * Cos(2 * D + 2 * M - M1)
    mD = mD - 1423 * Cos(4 * D + M1)
    mD = mD - 1117 * Cos(4 * M1)
    mD = mD - 1571 * E * Cos(4 * D - M)
    mD = mD - 1739 * Cos(D - 2 * M1)
    mD = mD - 4421 * Cos(2 * M1 - 2 * F)
    mD = mD + 1165 * E ^ 2 * Cos(2 * M + M1)
    mD = mD + 8752 * Cos(2 * D - M1 - 2 * F)
    SigmaR = mD
End Function

 
Public Function HourAngleOfTheMoon(ByVal Bujur As Double, ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function HourAngleOfTheMoon dimaksudkan untuk menghitung Moon Hour Angle
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal (Waktu Universal)
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Hour Angle Bulan dalam jam
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Hour Angle (Sudut Jam) Bulan di atas ufuk Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' di Jakarta dengan koordinat (6° 8' S 106° 45'E) dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Mempersiapkan Bujur tempat dalam derajat
'    2. Jalankan fungsi HourAngleOfTheMoon.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           B = 106. + 45./60.
'           x = HourAngleOfTheMoon(Y, M, D)
'    Hasil:
'         x = 5.74337121760952
'           = 05:44:36
'
' Referensi :
'    Kartunen, H. et.al., Fundamental Astronomy, ISBN 3-540-572o34 2nd Edition
'    Springer-Verlag Berlin Heidelberg New York, 1985, hal. 12
'----------------------------
'*Demak, 18 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   HourAngleOfTheMoon = Modulus(ApparentLocalSiderealTime(Bujur, Year, Month, Day) - ApparentMoonRightAscention(Year, Month, Day) / 15#, 24#)
End Function

Public Function MoonLongitudeELP2000(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonLongitudeELP2000 berfungsi untuk menghitung  Moon's Longitude (Bujur Ekliptik Bulan)
' dengan metode ELP-2000/82
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Bujur Ekliptik Bulan dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Bujur Ekliptik Bulan pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonLongitudeELP2000(Y, M, D)
'    Hasil:
'         x = 302.01367236314
'           = 302° 00' 49.22"
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, bab 45
'    M. Chapront-Touze and J. Chapront, "The lunar ephemeris ELP 2000",
'       Astronomy and Astrophysics, Vol. 124, pages 50-62 A983). —
' -------------------------------------------------------
'* Cibinong, 20 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------

   Dim T, JD, Pa, Lo, R(3) As Double
   
   JD = JulianEphemerisDay(Year, Month, Day)
   T = JulianEphemerisCentury(Year, Month, Day)
   
   Call ELP82B(JD, R)
   Lo = Atn2(R(1), R(2))
   
   ' convert to inertial mean ecliptic of date and equinox of date.
   Pa = 5029.0966 * T + 1.112 * T ^ 2 + 0.000077 * T ^ 3 - 0.00002353 * T ^ 4
   Pa = Radians(Pa / 3600#)
   
   MoonLongitudeELP2000 = Degrees(Lo + Pa)
End Function

Public Function ApparentMoonLongitudeELP2000(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi ApparentMoonLongitudeELP2000 berfungsi untuk menghitung Apparent Moon's Longitude (Bujur Ekliptik Bulan Nampak)
' dengan metode ELP-2000/82
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Bujur Ekliptik Bulan Nampak dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Bujur Ekliptik Bulan Nampak pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ApparentMoonLongitudeELP2000(Y, M, D)
'    Hasil:
'         x = 302.018593657391
'           = 302° 01' 6.94"
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, bab 45
'    M. Chapront-Touze and J. Chapront, "The lunar ephemeris ELP 2000",
'       Astronomy and Astrophysics, Vol. 124, pages 50-62 A983). —
' -------------------------------------------------------
'* Cibinong, 20 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    ApparentMoonLongitudeELP2000 = Modulus(MoonLongitudeELP2000(Year, Month, Day) + NutationInLongitude(Year, Month, Day), 360#)
End Function

Public Function MoonLatitudeELP2000(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi MoonLatitudeELP2000 dimaksudkan untuk menghitung Lintang Bulan ekliptic
' (Ecliptic Moon Latitude) dengan metode ELP2000/82.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Lintang Ekliptik Bulan Nampak dalam derajat busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Lintang Ekliptik Bulan pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MoonLatitudeELP2000(Y, M, D)
'    Hasil:
'         x = 1.00758267212283
'           = 1° 00' 27.30"
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, bab 45
'    M. Chapront-Touze and J. Chapront, "The lunar ephemeris ELP 2000",
'       Astronomy and Astrophysics, Vol. 124, pages 50-62 A983). —
' -------------------------------------------------------
'* Cibinong, 20 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim JD, R(3) As Double
   
   JD = JulianEphemerisDay(Year, Month, Day)
   Call ELP82B(JD, R)
   
   MoonLatitudeELP2000 = Degrees(Asin(R(3) / Sqr(R(1) * R(1) + R(2) * R(2) + R(3) * R(3))))
End Function

Public Function DistanceToTheMoonELP2000(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi DistanceToTheMoonELP2000 dimaksudkan untuk menghitung jarak Bumi ke Bulan
' (Earth-Moon Distance) dengan metode ELP2000.

' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Jarak Bumi ke Bulan dinyatakan dalam km
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Jarak Bumi ke Bulan pada tanggal 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = DistanceToTheMoonELP2000(Y, M, D)
'    Hasil:
'         x = 406236.835402202 km
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, bab 45
'    M. Chapront-Touze and J. Chapront, "The lunar ephemeris ELP 2000",
'       Astronomy and Astrophysics, Vol. 124, pages 50-62 A983). —
' -------------------------------------------------------
'* Cibinong, 20 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim JDE, R(3) As Double
   
   JDE = JulianEphemerisDay(Year, Month, Day)
   Call ELP82B(JDE, R)
   DistanceToTheMoonELP2000 = Sqr(R(1) ^ 2 + R(2) ^ 2 + R(3) ^ 2)
End Function

Public Function IjtimakELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function IjtimakELP2000_VSOP87 adalah duplikat dari fungsi NewMoonELP2000_VSOP87
' ------------------------------------------
'* Cibinong, 21 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   IjtimakELP2000_VSOP87 = NewMoonELP2000_VSOP87(Year, Month, Day)
End Function

Public Function KonjungsiELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function KonjungsiELP2000_VSOP87 adalah duplikat dari fungsi NewMoonELP2000_VSOP87
' ------------------------------------------
'* Cibinong, 21 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   KonjungsiELP2000_VSOP87 = NewMoonELP2000_VSOP87(Year, Month, Day)
End Function

Public Function NewMoonELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function NewMoonELP2000_VSOP87 dipakai untuk menghitung saat terjadinya konjungsi dengan metode ELP 2000/82
'  dan VSOP87.
'
'  INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'  OUTPUT:
'    Saat terjadinya NewMoon (bulan baru) atau Konjungsi atau Ijtima dalam waktu universal (Universal Time)
'
'  CATATAN:
'    - Perhitungan dilakukan dengan mencari saat Bujur Bulan Nampak dan Bujur Matahari Nampak sama (atau dengan
'      kata lain berhimpit atau membentuk sudut nol derajat)
'    - Hasil akhir dari fungsi ini dinyatakan dalam Julian Day, oleh karena itu diperlukan konversi
'      ke satuan waktu sehari-hari. lihat fungsi caldat.
'
'  Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      x = NewMoonELP2000_VSOP87(Y, M, D)
'      x = 2443192.65038556
'        = Jum'at, 18 Februari 1977 03:36:33 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab. 47
'    M. Chapront-Touze and J. Chapront, "The lunar ephemeris ELP 2000",
'       Astronomy and Astrophysics, Vol. 124, pages 50-62 A983). —
' ------------------------------------------
'* Cibinong, 21 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, M As Integer
   Dim inv, D, JDE, dB, dB_ As Double
   
   JDE = NewMoon(Year, Month, Day)
   Y = JD2Year(JDE)
   M = JD2Month(JDE)
   D = JD2Date(JDE) + JD2Time(JDE)
   dB = ApparentSunLongitudeVSOP87(Y, M, D) - ApparentMoonLongitudeELP2000(Y, M, D)
   inv = 20# * dB / (86400# * Abs(dB))
   dB_ = ApparentSunLongitudeVSOP87(Y, M, D + inv) - ApparentMoonLongitudeELP2000(Y, M, D + inv)
   D = D + inv * dB / (dB - dB_)
   NewMoonELP2000_VSOP87 = JulianDatum(Y, M, D)
End Function

Public Function FirstQuarterELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function FirstQuarterELP2000_VSOP87 dipakai untuk menghitung saat terjadinya seperempat pertama (First Quarter)
'  dengan metode ELP 2000/82 dan VSOP87.
'
'  INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'  OUTPUT:
'    Saat terjadinya First Quarter (seperempat pertama) dalam waktu universal (Universal Time)
'
'  CATATAN:
'    - Metode VSOP87 untuk menghitung Bujur Ekliptik Matahari, sedangkan metode ELP 2000/82 untuk menghitung
'      Bujur Ekliptik Bulan.
'    - Perhitungan dilakukan dengan mencari saat Bujur Bulan Nampak dan Bujur Matahari Nampak membentuk
'      sudut 270 derajat.
'    - Hasil akhir dari fungsi ini dinyatakan dalam Julian Day, oleh karena itu diperlukan konversi
'      ke satuan waktu sehari-hari. lihat fungsi caldat.
'
'  Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the First Quarter which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      x = FirstQuarterELP2000_VSOP87(Y, M, D)
'      x = 2443200.61781809
'        = Sabtu, 26 Februari 1977 02:49:39 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab. 47
'    M. Chapront-Touze and J. Chapront, "The lunar ephemeris ELP 2000",
'       Astronomy and Astrophysics, Vol. 124, pages 50-62 A983). —
' ------------------------------------------
'* Cibinong, 21 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, M As Integer
   Dim inv, D, JDE, dB, dB_ As Double
   
   JDE = FirstQuarter(Year, Month, Day)
   Y = JD2Year(JDE)
   M = JD2Month(JDE)
   D = JD2Date(JDE) + JD2Time(JDE)
   dB = ApparentSunLongitudeVSOP87(Y, M, D) - ApparentMoonLongitudeELP2000(Y, M, D) - 270#
   inv = 20# * dB / (86400# * Abs(dB))
   dB_ = ApparentSunLongitudeVSOP87(Y, M, D + inv) - ApparentMoonLongitudeELP2000(Y, M, D + inv) - 270#
   D = D + inv * dB / (dB - dB_)
   FirstQuarterELP2000_VSOP87 = JulianDatum(Y, M, D)
End Function

Public Function FullMoonELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function FullMoonELP2000_VSOP87 dipakai untuk menghitung saat terjadinya bulan purnama (Full Moon)
'  dengan metode ELP 2000/82 dan VSOP87.
'
'  INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'  OUTPUT:
'    Saat terjadinya Full Moon (bulan purnama) dalam waktu universal (Universal Time)
'
'  CATATAN:
'    - Metode VSOP87 untuk menghitung Bujur Ekliptik Matahari, sedangkan metode ELP 2000/82 untuk menghitung
'      Bujur Ekliptik Bulan.
'    - Perhitungan dilakukan dengan mencari saat Bujur Bulan Nampak dan Bujur Matahari Nampak membentuk
'      sudut 180 derajat.
'    - Hasil akhir dari fungsi ini dinyatakan dalam Julian Day, oleh karena itu diperlukan konversi
'      ke satuan waktu sehari-hari. lihat fungsi caldat.
'
'  Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the Full Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      x = FullMoonELP2000_VSOP87(Y, M, D)
'      x = 2443208.21731775
'        = Sabtu, 05 Maret 1977 17:12:56 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab. 47
'    M. Chapront-Touze and J. Chapront, "The lunar ephemeris ELP 2000",
'       Astronomy and Astrophysics, Vol. 124, pages 50-62 A983). —
' ------------------------------------------
'* Cibinong, 21 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, M As Integer
   Dim inv, D, JDE, dB, dB_ As Double
   
   JDE = FullMoon(Year, Month, Day)
   Y = JD2Year(JDE)
   M = JD2Month(JDE)
   D = JD2Date(JDE) + JD2Time(JDE)
   dB = ApparentSunLongitudeVSOP87(Y, M, D) - ApparentMoonLongitudeELP2000(Y, M, D) - 180#
   inv = 20# * dB / (86400# * Abs(dB))
   dB_ = ApparentSunLongitudeVSOP87(Y, M, D + inv) - ApparentMoonLongitudeELP2000(Y, M, D + inv) - 180#
   D = D + inv * dB / (dB - dB_)
   FullMoonELP2000_VSOP87 = JulianDatum(Y, M, D)
End Function

Public Function LastQuarterELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function LastQuarterELP2000_VSOP87 dipakai untuk menghitung saat terjadinya seperempat terakhir (Last Quarter)
'  dengan metode ELP 2000/82 dan VSOP87.
'
'  INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'  OUTPUT:
'    Saat terjadinya Last Quarter (seperempat terakhir) dalam waktu universal (Universal Time)
'
'  CATATAN:
'    - Metode VSOP87 untuk menghitung Bujur Ekliptik Matahari, sedangkan metode ELP 2000/82 untuk menghitung
'      Bujur Ekliptik Bulan.
'    - Perhitungan dilakukan dengan mencari saat Bujur Bulan Nampak dan Bujur Matahari Nampak membentuk
'      sudut 90 derajat.
'    - Hasil akhir dari fungsi ini dinyatakan dalam Julian Day, oleh karena itu diperlukan konversi
'      ke satuan waktu sehari-hari. lihat fungsi caldat.
'
'  Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the Last Quarter which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      x = LastQuarterELP2000_VSOP87(Y, M, D)
'      x = 2443214.98215079
'        = Sabtu, 12 Maret 1977 11:34:18 UT
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab. 47
'    M. Chapront-Touze and J. Chapront, "The lunar ephemeris ELP 2000",
'       Astronomy and Astrophysics, Vol. 124, pages 50-62 A983). —
' ------------------------------------------
'* Cibinong, 21 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, M As Integer
   Dim inv, D, JDE, dB, dB_ As Double
   
   JDE = LastQuarter(Year, Month, Day)
   Y = JD2Year(JDE)
   M = JD2Month(JDE)
   D = JD2Date(JDE) + JD2Time(JDE)
   dB = ApparentSunLongitudeVSOP87(Y, M, D) - ApparentMoonLongitudeELP2000(Y, M, D) - 90#
   inv = 20# * dB / (86400# * Abs(dB))
   dB_ = ApparentSunLongitudeVSOP87(Y, M, D + inv) - ApparentMoonLongitudeELP2000(Y, M, D + inv) - 90#
   D = D + inv * dB / (dB - dB_)
   LastQuarterELP2000_VSOP87 = JulianDatum(Y, M, D)
End Function


Public Function NewMoonToNewMoonELP2000_VSOP87(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function NewMoonToNewMoonELP2000_VSOP87 dipakai untuk menghitung umur bulan dari konjungsi ke konjugsi berikutnya.
'  dengan metode ELP 2000/82 dan VSOP87.
'
'  INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  OUTPUT:
'    Waktu dari NewMoon (sebelum tanggal pada input) ke newmoon berikutnya.
'    Waktu dari NewMoon ke NewMoon berikutnya ini dinyatakan dalam hari
'
'  Contoh pemakaian fungsi:
'   Untuk menghitung Waktu NewMoon ke NewMoon berikutnya Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
'   dilakukan dengan Langkah-langkah sebagai berikut:
'     1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'        Waktu Universal = Waktu Lokal - Zona Waktu
'        Dalam hal ini WIB berarti Zona Waktunya 7
'     2. Jalankan fungsi NewMoonToNewMoonELP2000_VSOP87.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = NewMoonToNewMoonELP2000_VSOP87(Y, M, D)
'    Hasil:
'           x = 29.8193850787357 hari
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, Bab. 47
'    M. Chapront-Touze and J. Chapront, "The lunar ephemeris ELP 2000",
'       Astronomy and Astrophysics, Vol. 124, pages 50-62 A983). —
' ------------------------------------------
'* Cibinong, 22 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Y, M As Integer
   Dim D, NewMoonAfter, NewMoonBefore As Double
   
   NewMoonAfter = NewMoonELP2000_VSOP87(Year, Month, Day)
   Y = JD2Year(NewMoonAfter - 30#)
   M = JD2Month(NewMoonAfter - 30#)
   D = JD2Date(NewMoonAfter - 30#) + JD2Time(NewMoonAfter - 30#)
   NewMoonBefore = NewMoonELP2000_VSOP87(Y, M, D)
   NewMoonToNewMoonELP2000_VSOP87 = NewMoonAfter - NewMoonBefore
End Function

Public Function dLatitude(ByVal Lintang As Double)
' Function dLatitude dimaksudkan untuk menghitung koreksi koordinat geografik bumi
' ke koordinat geosentrik. Hal ini diperlukan karena sistem perhitungan
' dilakukan dalam rumusan segitiga bola.
' ---------------------------------------------------------------------------------------------
' The latitude is usually supposed to mean the geographical latitude, which is the angle between
' the plumb line and the equatorial plane. Because the Earth is rotating, it is slightly flattened.
' The shape defined by the surface of the oceans, is very close to an ablate spheroid, te short axis,
' which coincides with the rotation axis. In 1979, the International Astronomical Union (IAU)
' adopted the following dimensions for the meredional ellipse:

'     equatorial radius  :     a = 6378140 m
'     polar radius       :     b = 6356755 m
'     flattening         :     f = (a - b) / a
'                                = 1 / 298.257
'
' The angle between the equator and the normal to the ellipsoid approximating the true Earth
' is called the geodetic latitude. Because the surface of a liquid (like an ocean) is perpendicular
' the geodetic and geographical latitudes are practically the same.
'
' The new geodetic reference (World Geodetic System 1984), the earth parameters are:
'      a = 6378137 m
'      b = 6356752 m
'      f = 1 / 298.257223563
'
' INPUT:
'      - Lintang : Koordinat Lintang geografik suatu tempat di permukaan bumi
'
' OUTPUT:
'    dLatitude dalam derajat
'
' CONTOH:
' Koordinat Geografik Masjid Istiqlal adalah: Bujur = 106.830928° dan Lintang =  -6.169788°
' Sehingga koordinat Lintang Geosentriknya
'                  LintangGeosentrik = Lintang + dLatitude(Lintang)
'                  -6.12879988070087 = -6.169788 + 0.0409881192991284
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm Bab 10, hal. 77-78
'----------------------------
' Cibinong, 14 Juni 2011
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
 
   Dim A, B As Double
   A = 6378137
   B = 6356752
   dLatitude = Degrees(Atn((B ^ 2 / A ^ 2) * Tan(Radians(Lintang)))) - Lintang
End Function

Public Function SunAberration(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function SunAberration dimaksudkan untuk menghitung koreksi Aberasi
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    Aberasi Matahari dalam detik busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Untuk menghitung Aberasi Matahari Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' dilakukan dengan Langkah-langkah sebagai berikut:
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi SunAberration.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = SunAberration(Y, M, D)
'    Hasil:
'         x = -20.7748666716828
'
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm hal. 155
'     Menurut Jean Meeus ketelitian perhitunan aberasi adalah 0.001"
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim dL, R, tau As Double
   
   tau = JulianEphemerisCentury(Year, Month, Day) / 10#
   
   dL = 3548.193
   dL = dL + 118.568 * Sin(Radians(87.5287 + 359993.7286 * tau))
   dL = dL + 2.476 * Sin(Radians(85.0561 + 719987.4571 * tau))
   dL = dL + 1.376 * Sin(Radians(27.8502 + 4452671.1152 * tau))
   dL = dL + 0.119 * Sin(Radians(73.1375 + 450368.8564 * tau))
   dL = dL + 0.114 * Sin(Radians(337.2264 + 329644.6718 * tau))
   dL = dL + 0.086 * Sin(Radians(222.54 + 659289.3436 * tau))
   dL = dL + 0.078 * Sin(Radians(162.8136 + 9224659.7915 * tau))
   dL = dL + 0.054 * Sin(Radians(82.5823 + 1079981.1857 * tau))
   dL = dL + 0.052 * Sin(Radians(171.5189 + 225184.4282 * tau))
   dL = dL + 0.034 * Sin(Radians(30.3214 + 4092677.3866 * tau))
   dL = dL + 0.033 * Sin(Radians(119.8105 + 337181.4711 * tau))
   dL = dL + 0.023 * Sin(Radians(247.5418 + 299295.6151 * tau))
   dL = dL + 0.023 * Sin(Radians(325.1526 + 315559.556 * tau))
   dL = dL + 0.021 * Sin(Radians(155.1241 + 675553.2846 * tau))
   dL = dL + 7.311 * Sin(Radians(333.4515 + 359993.7286 * tau))
   dL = dL + 0.305 * Sin(Radians(330.9814 + 719987.4571 * tau))
   dL = dL + 0.01 * Sin(Radians(328.517 + 1079981.1857 * tau))
   dL = dL + 0.309 * Sin(Radians(241.4518 + 359993.7286 * tau))
   dL = dL + 0.021 * Sin(Radians(205.0482 + 719987.4571 * tau))
   dL = dL + 0.004 * Sin(Radians(297.861 + 4452671.1152 * tau))
   dL = dL + 0.01 * Sin(Radians(154.7066 + 359993.7286 * tau))
   R = DistanceToTheSunVSOP87(Year, Month, Day)
   SunAberration = -0.005775518 * R * dL
End Function

Public Function NutationInLongitude(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function Nutation dimaksudkan untuk menghitung koreksi Nutasi pengaruhnya pada koordinat Bujur Ecliptic
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'     NutationInLongitude dalam satuan derajat busur
'
' CATATAN:
'    Koreksi ini diperlukan untuk menghitung Bujur Bulan Nampak (Apparent Moon Longitude) dan
'    Bujur Matahari Nampak (Apparent Sun Longitude)
'
' Contoh pemakaian fungsi:
'     Hitung Koreksi Nutation in Longitude tanggal 12 Februari 2010, jam 10:30:15 WIB
'
'     maka :
'        Y = 2010
'        M = 2
'        D = 12.14600694444440
'        x = NutationInLongitude(Y,M,D)
'        x =  0.00492129425061738 derajat
'          = 17.7166593022226 detik
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 21.
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T, D, M, M1, F, Omega, dPsi As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   D = Radians(MoonMeanElongation(Year, Month, Day))
   M = Radians(SunMeanAnomaly(Year, Month, Day))
   M1 = Radians(MoonMeanAnomaly(Year, Month, Day))
   F = Radians(MoonArgumentOfLatitude(Year, Month, Day))
   Omega = Radians(nilai_Omega(Year, Month, Day))
   
   dPsi = 0#
   dPsi = dPsi + (-171996# + -174.2 * T) * Sin(0# * D + 0# * M + 0# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (-13187# + -1.6 * T) * Sin(-2# * D + 0# * M + 0# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-2274# + -0.2 * T) * Sin(0# * D + 0# * M + 0# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (2062# + 0.2 * T) * Sin(0# * D + 0# * M + 0# * M1 + 0# * F + 2# * Omega)
   dPsi = dPsi + (1426# + -3.4 * T) * Sin(0# * D + 1# * M + 0# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (712# + 0.1 * T) * Sin(0# * D + 0# * M + 1# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (-517# + 1.2 * T) * Sin(-2# * D + 1# * M + 0# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-386# + -0.4 * T) * Sin(0# * D + 0# * M + 0# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (217# + -0.5 * T) * Sin(-2# * D + -1# * M + 0# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (129# + 0.1 * T) * Sin(-2# * D + 0# * M + 0# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (63# + 0.1 * T) * Sin(0# * D + 0# * M + 1# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (-58# + -0.1 * T) * Sin(0# * D + 0# * M + -1# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (17# + -0.1 * T) * Sin(0# * D + 2# * M + 0# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (-16# + 0.1 * T) * Sin(-2# * D + 2# * M + 0# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-301# + 0# * T) * Sin(0# * D + 0# * M + 1# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-158# + 0# * T) * Sin(-2# * D + 0# * M + 1# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (123# + 0# * T) * Sin(0# * D + 0# * M + -1# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (63# + 0# * T) * Sin(2# * D + 0# * M + 0# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (-59# + 0# * T) * Sin(2# * D + 0# * M + -1# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-51# + 0# * T) * Sin(0# * D + 0# * M + 1# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (48# + 0# * T) * Sin(-2# * D + 0# * M + 2# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (46# + 0# * T) * Sin(0# * D + 0# * M + -2# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (-38# + 0# * T) * Sin(2# * D + 0# * M + 0# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-31# + 0# * T) * Sin(0# * D + 0# * M + 2# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (29# + 0# * T) * Sin(0# * D + 0# * M + 2# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (29# + 0# * T) * Sin(-2# * D + 0# * M + 1# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (26# + 0# * T) * Sin(0# * D + 0# * M + 0# * M1 + 2# * F + 0# * Omega)
   dPsi = dPsi + (-22# + 0# * T) * Sin(-2# * D + 0# * M + 0# * M1 + 2# * F + 0# * Omega)
   dPsi = dPsi + (21# + 0# * T) * Sin(0# * D + 0# * M + -1# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (16# + 0# * T) * Sin(2# * D + 0# * M + -1# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (-15# + 0# * T) * Sin(0# * D + 1# * M + 0# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (-13# + 0# * T) * Sin(-2# * D + 0# * M + 1# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (-12# + 0# * T) * Sin(0# * D + -1# * M + 0# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (11# + 0# * T) * Sin(0# * D + 0# * M + 2# * M1 + -2# * F + 0# * Omega)
   dPsi = dPsi + (-10# + 0# * T) * Sin(2# * D + 0# * M + -1# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (-8# + 0# * T) * Sin(2# * D + 0# * M + 1# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (7# + 0# * T) * Sin(0# * D + 1# * M + 0# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-7# + 0# * T) * Sin(-2# * D + 1# * M + 1# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (-7# + 0# * T) * Sin(0# * D + -1# * M + 0# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-7# + 0# * T) * Sin(2# * D + 0# * M + 0# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (6# + 0# * T) * Sin(2# * D + 0# * M + 1# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (6# + 0# * T) * Sin(-2# * D + 0# * M + 2# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (6# + 0# * T) * Sin(-2# * D + 0# * M + 1# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (-6# + 0# * T) * Sin(2# * D + 0# * M + -2# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (-6# + 0# * T) * Sin(2# * D + 0# * M + 0# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (5# + 0# * T) * Sin(0# * D + -1# * M + 1# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (-5# + 0# * T) * Sin(-2# * D + -1# * M + 0# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (-5# + 0# * T) * Sin(-2# * D + 0# * M + 0# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (-5# + 0# * T) * Sin(0# * D + 0# * M + 2# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (4# + 0# * T) * Sin(-2# * D + 0# * M + 2# * M1 + 0# * F + 1# * Omega)
   dPsi = dPsi + (4# + 0# * T) * Sin(-2# * D + 1# * M + 0# * M1 + 2# * F + 1# * Omega)
   dPsi = dPsi + (4# + 0# * T) * Sin(0# * D + 0# * M + 1# * M1 + -2# * F + 0# * Omega)
   dPsi = dPsi + (-4# + 0# * T) * Sin(-1# * D + 0# * M + 1# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (-4# + 0# * T) * Sin(-2# * D + 1# * M + 0# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (-4# + 0# * T) * Sin(1# * D + 0# * M + 0# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (3# + 0# * T) * Sin(0# * D + 0# * M + 1# * M1 + 2# * F + 0# * Omega)
   dPsi = dPsi + (-3# + 0# * T) * Sin(0# * D + 0# * M + -2# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-3# + 0# * T) * Sin(-1# * D + -1# * M + 1# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (-3# + 0# * T) * Sin(0# * D + 1# * M + 1# * M1 + 0# * F + 0# * Omega)
   dPsi = dPsi + (-3# + 0# * T) * Sin(0# * D + -1# * M + 1# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-3# + 0# * T) * Sin(2# * D + -1# * M + -1# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-3# + 0# * T) * Sin(0# * D + 0# * M + 3# * M1 + 2# * F + 2# * Omega)
   dPsi = dPsi + (-3# + 0# * T) * Sin(2# * D + -1# * M + 0# * M1 + 2# * F + 2# * Omega)
   NutationInLongitude = dPsi / (3600# * 10000#)
End Function

Public Function NutationInObliquity(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function NutationInObliquity dimaksudkan untuk menghitung pengaruh Nutasi pada Epsilon (Obliquity of the Ecliptic)
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'   NutationInObliquity (Delta Epsilon) dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Nutation in Obliquity of the Ecliptic (Delta Epsilon) Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi NutationInObliquity.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = NutationInObliquity(Y, M, D)
'    Hasil:
'           x = 0.00090459553607744 derajat
'             = 3.25654392987878 detik
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 21. hal. 135
'----------------------------
' Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 21. hal. 131
'     Menurut Jean Meeus : The accuracy of this expression is estimated
'     at 0".01 after 1000 years (that is between A.D. 1000 and 3000)
'     and a few second after 10000 years
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T, D, M, M1, F, Omega, dEpsilon As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   D = Radians(MoonMeanElongation(Year, Month, Day))
   M = Radians(SunMeanAnomaly(Year, Month, Day))
   M1 = Radians(MoonMeanAnomaly(Year, Month, Day))
   F = Radians(MoonArgumentOfLatitude(Year, Month, Day))
   Omega = Radians(nilai_Omega(Year, Month, Day))
   
   dEpsilon = 0#
   dEpsilon = dEpsilon + (92025# + 8.9 * T) * Cos(0# * D + 0# * M + 0# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (5736# + -3.1 * T) * Cos(-2# * D + 0# * M + 0# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (977# + -0.5 * T) * Cos(0# * D + 0# * M + 0# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (-895# + 0.5 * T) * Cos(0# * D + 0# * M + 0# * M1 + 0# * F + 2# * Omega)
   dEpsilon = dEpsilon + (54# + -0.1 * T) * Cos(0# * D + 1# * M + 0# * M1 + 0# * F + 0# * Omega)
   dEpsilon = dEpsilon + (224# + -0.6 * T) * Cos(-2# * D + 1# * M + 0# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (129# + -0.1 * T) * Cos(0# * D + 0# * M + 1# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (-95# + 0.3 * T) * Cos(-2# * D + -1# * M + 0# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (-7# + 0# * T) * Cos(0# * D + 0# * M + 1# * M1 + 0# * F + 0# * Omega)
   dEpsilon = dEpsilon + (200# + 0# * T) * Cos(0# * D + 0# * M + 0# * M1 + 2# * F + 1# * Omega)
   dEpsilon = dEpsilon + (-70# + 0# * T) * Cos(-2# * D + 0# * M + 0# * M1 + 2# * F + 1# * Omega)
   dEpsilon = dEpsilon + (-53# + 0# * T) * Cos(0# * D + 0# * M + -1# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (-33# + 0# * T) * Cos(0# * D + 0# * M + 1# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (26# + 0# * T) * Cos(2# * D + 0# * M + -1# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (32# + 0# * T) * Cos(0# * D + 0# * M + -1# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (27# + 0# * T) * Cos(0# * D + 0# * M + 1# * M1 + 2# * F + 1# * Omega)
   dEpsilon = dEpsilon + (-24# + 0# * T) * Cos(0# * D + 0# * M + -2# * M1 + 2# * F + 1# * Omega)
   dEpsilon = dEpsilon + (16# + 0# * T) * Cos(2# * D + 0# * M + 0# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (13# + 0# * T) * Cos(0# * D + 0# * M + 2# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (-12# + 0# * T) * Cos(-2# * D + 0# * M + 1# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (-10# + 0# * T) * Cos(0# * D + 0# * M + -1# * M1 + 2# * F + 1# * Omega)
   dEpsilon = dEpsilon + (-8# + 0# * T) * Cos(2# * D + 0# * M + -1# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (7# + 0# * T) * Cos(-2# * D + 2# * M + 0# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (9# + 0# * T) * Cos(0# * D + 1# * M + 0# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (7# + 0# * T) * Cos(-2# * D + 0# * M + 1# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (6# + 0# * T) * Cos(0# * D + -1# * M + 0# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (5# + 0# * T) * Cos(2# * D + 0# * M + -1# * M1 + 2# * F + 1# * Omega)
   dEpsilon = dEpsilon + (3# + 0# * T) * Cos(2# * D + 0# * M + 1# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (-3# + 0# * T) * Cos(0# * D + 1# * M + 0# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (3# + 0# * T) * Cos(0# * D + -1# * M + 0# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (3# + 0# * T) * Cos(2# * D + 0# * M + 0# * M1 + 2# * F + 1# * Omega)
   dEpsilon = dEpsilon + (-3# + 0# * T) * Cos(-2# * D + 0# * M + 2# * M1 + 2# * F + 2# * Omega)
   dEpsilon = dEpsilon + (-3# + 0# * T) * Cos(-2# * D + 0# * M + 1# * M1 + 2# * F + 1# * Omega)
   dEpsilon = dEpsilon + (3# + 0# * T) * Cos(2# * D + 0# * M + -2# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (3# + 0# * T) * Cos(2# * D + 0# * M + 0# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (3# + 0# * T) * Cos(-2# * D + -1# * M + 0# * M1 + 2# * F + 1# * Omega)
   dEpsilon = dEpsilon + (3# + 0# * T) * Cos(-2# * D + 0# * M + 0# * M1 + 0# * F + 1# * Omega)
   dEpsilon = dEpsilon + (3# + 0# * T) * Cos(0# * D + 0# * M + 2# * M1 + 2# * F + 1# * Omega)
   NutationInObliquity = dEpsilon / (3600# * 10000#)
End Function

Public Function HorizontalMoonParallax(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Function HorizontalMoonParallax dimaksudkan untuk menghitung koreksi horizontal parallax pada kenampakan bulan
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'     Parallax Bulan dalam satuan derajat busur
'
' CATATAN:
'    Koreksi ini diperlukan dalm menghitung tinggi benda-benda langit termasuk bumi dan bulan
'
' Contoh pemakaian fungsi:
'     Hitung Koreksi Nutation in Longitude tanggal 12 Februari 2010, jam 10:30:15 WIB
'
'     maka :
'        Y = 2010
'        M = 2
'        D = 12.14600694444440
'        x = HorizontalMoonParallax(Y,M,D)
'        x =
'            0.89961930924678
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 21.
'----------------------------
' Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    HorizontalMoonParallax = Degrees(Asin(6378.14 / DistanceToTheMoon(Year, Month, Day)))
End Function

Public Function RefractionApparentAltitude(ByVal h0 As Double, ByVal P As Double, ByVal T As Double)
' Function RefractionApparentAltitude dimaksudkan untuk menghitung koreksi Refraksi jika ketinggian
' benda langit Nampak diketahui. Jika tinggi benda langit Nampak dikoreksi refraksi maka akan
' menjadi Tinggi Benda Langit Sebenarnya.
'
' INPUT
'     h0 : Apparent Altitude (Ketinggian Benda Langit dengan koreksi faktor Aberasi dan Nutasi)
'          dinyatakan dalam derajat
'     P  : Tekanan Udara di permukaan bumi dinyatakan dalam millibar
'     T  : Suhu di permukaan bumi dinyatakan dalam derajat Celcius
'
' OUTPUT
'    Refraksi dinyatakan dalam menit busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Refraction  Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Untuk menghitung koreksi refraksi saat bulan atau matahari terbenam
'       maka ketinggian benda langit 0 derajat, tekanan udara 1010 milibar dan
'       suhu udara 27 derajat celsius.
'    2. Jalankan fungsi RefractionApparentAltitude.
'           h0 = 0
'           p = 1010
'           T = 27
'           Ref = RefractionApparentAltitude(h0,p,T)
'    Hasil:
'         Ref = 32.5247368203474
'             = 32' 31.48"
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 15. hal. 102
'     Menurut Jean Meeus kesalahan maximal perhitunan refraksi adalah 0.9"
' Catatan:
'   -  Rata-rata tekanan udara di Indonesia sekitar 1010 mbar
'   -  Rata-rata suhu udara di daerah pantai Indonesia pada saat matahari terbenam 28 derajat celsius
'   -  Untuk perhitungan lebih teliti perlu dilakukan pengukuran tekanan udara dan suhu di tempat pengamatan
'----------------------------
'*Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim R, dR1, dR2 As Double

   R = 1# / Tan(Radians(h0 + 7.31 / (h0 + 4.4))) + 0.0013515
   dR1 = -0.06 * Sin(Radians(14.7 * R / 60# + 13))
   dR2 = (P / 1010) * (283 / (273 + T))
   RefractionApparentAltitude = (R + dR1 / 60#) * dR2
End Function

Public Function RefractionTrueAltitude(ByVal h As Double, ByVal P As Double, ByVal T As Double)
' Function RefractionTrueAltitude dimaksudkan untuk menghitung koreksi Refraksi jika Ketinggian
' sebenarnya diketahui.
'
' CATATAN:
'   Jika Tinggi Benda Langit sebenarnya dikoreksi dengan Refraksi ini, maka Tinggi Benda Langit tersebut
'   menjadi Tinggi Benda Langit Nampak
'
' INPUT
'     h  : True Altitude (Ketinggian sebelum dikoreksi dengan Refraksi)
'     P  : Tekanan Udara di permukaan bumi dinyatakan dalam millibar
'     T  : Suhu di permukaan bumi dinyatakan dalam derajat Celcius
'
' OUTPUT
'    Refraksi dinyatakan dalam menit busur
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Refraction  Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Untuk menghitung koreksi refraksi saat bulan atau matahari terbenam
'       maka ketinggian benda langit 0 derajat, tekanan udara 1010 milibar dan
'       suhu udara 27 derajat celsius.
'    2. Jalankan fungsi RefractionTrueAltitude.
'           h = 0
'           p = 1010
'           T = 27
'           Ref = RefractionTrueAltitude(h0,p,T)
'    Hasil:
'         Ref = 26.8053658738998
'             = 26' 48.32"
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 15. hal. 102
'     Menurut Jean Meeus kesalahan maximal perhitunan refraksi adalah 0.9"
' Catatan:
'   -  Rata-rata tekanan udara di Indonesia sekitar 1010 mbar
'   -  Rata-rata suhu udara di daerah pantai Indonesia pada saat matahari terbenam 28 derajat celsius
'   -  Untuk perhitungan lebih teliti perlu dilakukan pengukuran tekanan udara dan suhu di tempat pengamatan
'----------------------------
'*Yogyakarta, 03 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim R, dR As Double

   R = 1# / Tan(Radians(h + 10.3 / (h + 5.11))) + 0.0019279
   dR = (P / 1010) * (283 / (273 + T))
   RefractionTrueAltitude = R * dR
End Function

Public Function kwd(ByVal TZ As Double, ByVal Bujur As Double) As Double
' Fungsi kwd dimaksudkan untuk menghitug koreksi wilayah daerah sebagai pengaruh dari Zona waktu yang
' diberlakukan di daerah tersebut tidak eksak secara astronomis.
'
' INPUT:
'     TZ    : Zona Waktu
'     Bujur : Koordinat Bujur tempat yang akandihitung kwdnya.
' OUTPUT:
'     Koreksi DIP dinyatakan dalam menit busur
'
' CONTOH:
' Menghitung koreksi kwd di Semarang dengan koorinat Bujur 110.5 dan Zona Waktu 7., sbb:
'      x = kwd(7., 106.5#)
'      x =
'          -0.1
'----------------------------
'*Cibinong, 06 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    kwd = (TZ * 15# - Bujur) / 15#
End Function

Public Function KoreksiDIP(Optional ByVal h As Double = 0#) As Double
' Fungsi KoreksiDIP dimaksudkan untuk koreksi garis ufuk akibat ketinggian tempat
' DIP terjadi karena ketinggian tempat pengamatan mempengaruhi ufuk (horizon).
' Horizon  yang  teramati  pada ketinggian mata sama dengan ketinggian permukaan laut disebut
' horizon sebenarnya (true horizon) atau ufuk hissi. Ufuk ini  sejajar  dengan  ufuk  hakiki  yang melalui
' bumi. Horizon yang teramati oleh mata pada ketinggian tertentu di atas permukaan laut, disebut
' horizon semu atau ufuk mar’i.
' Rumus pendekatan untuk menghitung sudut Dip (D)
'        D = arc Cos ( Rb/( Rb + h) )
' Rb adalah jari-jari bumi, dan h tinggi tempat
' Jika h dinyatakan dalam meter, D karena kecil dihitung
' dalam menit busur, maka: D = 1.76 * sqrt(h)
'
' INPUT:
'     h : Tinggi Tempat (dalam meter) terhadap muka laut
'
' OUTPUT:
'     Koreksi DIP dinyatakan dalam menit busur
'
' CONTOH:
'      x = KoreksiDIP(100#)
'      x =
'          19.2611389886892
'----------------------------
'*Cibinong, 06 Juli 2012
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Informasi Geospasial
' Pusat Pemetaan Batas Wilayah
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
  Dim x, D As Double
  D = Acos(6371008# / (6371008# + h))
  KoreksiDIP = Degrees(D) * 60#
                                
End Function

Public Function KoreksiNewMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function KoreksiNewMoon dipakai untuk menghitung fase-fase Bulan
'  Koreksi NewMoon adalah Koreksi yang diperlukan untuk mendapatkan hasil perhitungan yang akurat
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'     maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  OUTPUT:
'     Koreksi NewMoon dalam Hari. Koreksi ini akan dipakai untuk menghitung saat NewMoon
'     yang lebih teliti.
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'     place in February 1977.
'      maka :
'        Y = 1977
'        M = 2
'        D = 15
'        x =  (Y, M, D)
'        x =
'            -0.289156341420674
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 321 dan 322
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E, M, M1, F, O, dF As Double
   
   E = nilai_E(Year, Month, Day)
   M = Radians(nilai_M(Year, Month, Day))
   M1 = Radians(nilai_M1(Year, Month, Day))
   F = Radians(nilai_F(Year, Month, Day))
   O = Radians(nilai_O(Year, Month, Day))
   
   dF = 0#
   dF = dF - 0.4072 * Sin(M1)
   dF = dF + 0.17241 * E * Sin(M)
   dF = dF + 0.01608 * Sin(2 * M1)
   dF = dF + 0.01039 * Sin(2 * F)
   dF = dF + 0.00739 * E * Sin(M1 - M)
   dF = dF - 0.00514 * E * Sin(M1 + M)
   dF = dF + 0.00208 * E * E * Sin(2 * M)
   dF = dF - 0.00111 * Sin(M1 - 2 * F)
   dF = dF - 0.00057 * Sin(M1 + 2 * F)
   dF = dF + 0.00056 * E * Sin(2 * M1 + M)
   dF = dF - 0.00042 * Sin(3 * M1)
   dF = dF + 0.00042 * E * Sin(M + 2 * F)
   dF = dF + 0.00038 * E * Sin(M - 2 * F)
   dF = dF - 0.00024 * E * Sin(2 * M1 - M)
   dF = dF - 0.00017 * Sin(O)
   dF = dF - 0.00007 * Sin(M1 + 2 * M)
   dF = dF + 0.00004 * Sin(2 * (M1 - F))
   dF = dF + 0.00004 * Sin(3 * M)
   dF = dF + 0.00003 * Sin(M1 + M - 2 * F)
   dF = dF + 0.00003 * Sin(2 * (M1 + F))
   dF = dF - 0.00003 * Sin(M1 + M + 2 * F)
   dF = dF + 0.00003 * Sin(M1 - M + 2 * F)
   dF = dF - 0.00002 * Sin(M1 - M - 2 * F)
   dF = dF - 0.00002 * Sin(3 * M1 + M)
   dF = dF + 0.00002 * Sin(4 * M1)
   KoreksiNewMoon = dF
End Function

Public Function KoreksiFirstQuarter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function KoreksiFirstQuarter dipakai untuk menghitung fase-fase Bulan saat seperempat pertama (First Quarter)
'  Koreksi KoreksiFirstQuarter adalah Koreksi yang diperlukan untuk mendapatkan hasil perhitungan yang akurat
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'     maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  OUTPUT:
'     Koreksi FirstQuarter dalam Hari. Koreksi ini akan dipakai untuk menghitung saat First Quarter
'     yang lebih teliti.
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Analogously Example 4-7.a — Calculate the instant of the First Quarter of the
'                                                        Moon which took place in February 1977.
'      maka :
'        Y = 1977
'        M = 2
'        D = 15
'        x = KoreksiFirstQuarter(Y, M, D)
'        x =
'            0.293146044813627
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 321 dan 322
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E, M, M1, F, O, dF As Double
   
   E = nilai_E(Year, Month, Day, 0.25)
   M = Radians(nilai_M(Year, Month, Day, 0.25))
   M1 = Radians(nilai_M1(Year, Month, Day, 0.25))
   F = Radians(nilai_F(Year, Month, Day, 0.25))
   O = Radians(nilai_O(Year, Month, Day, 0.25))
   
   dF = 0#
   dF = dF - 0.62801 * Sin(M1)
   dF = dF + 0.17172 * E * Sin(M)
   dF = dF - 0.01183 * E * Sin(M1 + M)
   dF = dF + 0.00862 * Sin(2 * M1)
   dF = dF + 0.00804 * Sin(2 * F)
   dF = dF + 0.00454 * E * Sin(M1 - M)
   dF = dF + 0.00204 * E * E * Sin(2 * M)
   dF = dF - 0.0018 * Sin(M1 - 2 * F)
   dF = dF - 0.0007 * Sin(M1 + 2 * F)
   dF = dF - 0.0004 * Sin(3 * M1)
   dF = dF - 0.00034 * E * Sin(2 * M1 - M)
   dF = dF + 0.00032 * E * Sin(M + 2 * F)
   dF = dF + 0.00032 * E * Sin(M - 2 * F)
   dF = dF - 0.00028 * E * E * Sin(M1 + 2 * M)
   dF = dF + 0.00027 * E * Sin(2 * M1 * M)
   dF = dF - 0.00017 * Sin(O)
   dF = dF - 0.00005 * Sin(M1 - M - 2 * F)
   dF = dF + 0.00004 * Sin(2 * M1 + 2 * F)
   dF = dF - 0.00004 * Sin(M1 + M + 2 * F)
   dF = dF + 0.00004 * Sin(M1 - 2 * M)
   dF = dF + 0.00003 * Sin(M1 + M - 2 * F)
   dF = dF + 0.00003 * Sin(3 * M)
   dF = dF + 0.00002 * Sin(2 * M1 - 2 * F)
   dF = dF + 0.00002 * Sin(M1 - M + 2 * F)
   dF = dF - 0.00002 * Sin(3 * M1 + M)
   KoreksiFirstQuarter = dF
End Function

Public Function KoreksiFullMoon(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function KoreksiFullMoon dipakai untuk menghitung fase Bulan saat purnama (Full Moon)
'  Koreksi KoreksiFullMoon adalah Koreksi yang diperlukan untuk mendapatkan hasil perhitungan yang akurat
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'     maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  OUTPUT:
'     Koreksi FullMoon dalam Hari. Koreksi ini akan dipakai untuk menghitung saat Bulan Purnama (Full Moon)
'     yang lebih teliti.
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Analogously Example 4-7.a — Calculate the instant of the Full Moon which took
'     place in February 1977.
'      maka :
'        Y = 1977
'        M = 2
'        D = 15
'        x = KoreksiFullMoon(Y, M, D)
'        x =
'            0.512488444123921
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 321 dan 322
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E, M, M1, F, O, dF As Double
   
   E = nilai_E(Year, Month, Day, 0.5)
   M = Radians(nilai_M(Year, Month, Day, 0.5))
   M1 = Radians(nilai_M1(Year, Month, Day, 0.5))
   F = Radians(nilai_F(Year, Month, Day, 0.5))
   O = Radians(nilai_O(Year, Month, Day, 0.5))
   
   dF = 0#
   dF = dF - 0.40614 * Sin(M1)
   dF = dF + 0.17302 * E * Sin(M)
   dF = dF + 0.01614 * Sin(2 * M1)
   dF = dF + 0.01043 * Sin(2 * F)
   dF = dF + 0.00734 * E * Sin(M1 - M)
   dF = dF - 0.00515 * E * Sin(M1 + M)
   dF = dF + 0.00209 * E * E * Sin(2 * M)
   dF = dF - 0.00111 * Sin(M1 - 2 * F)
   dF = dF - 0.00057 * Sin(M1 + 2 * F)
   dF = dF + 0.00056 * E * Sin(2 * M1 + M)
   dF = dF - 0.00042 * Sin(3 * M1)
   dF = dF + 0.00042 * E * Sin(M + 2 * F)
   dF = dF + 0.00038 * E * Sin(M - 2 * F)
   dF = dF - 0.00024 * E * Sin(2 * M1 - M)
   dF = dF - 0.00017 * Sin(O)
   dF = dF - 0.00007 * Sin(M1 + 2 * M)
   dF = dF + 0.00004 * Sin(2 * (M1 - F))
   dF = dF + 0.00004 * Sin(3 * M)
   dF = dF + 0.00003 * Sin(M1 + M - 2 * F)
   dF = dF + 0.00003 * Sin(2 * (M1 + F))
   dF = dF - 0.00003 * Sin(M1 + M + 2 * F)
   dF = dF + 0.00003 * Sin(M1 - M + 2 * F)
   dF = dF - 0.00002 * Sin(M1 - M - 2 * F)
   dF = dF - 0.00002 * Sin(3 * M1 + M)
   dF = dF + 0.00002 * Sin(4 * M1)
   KoreksiFullMoon = dF
End Function

Public Function KoreksiLastQuarter(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function KoreksiLastQuarter dipakai untuk menghitung fase-fase Bulan saat seperempat terakhir (Last Quarter)
'  Koreksi KoreksiLastQuarter adalah Koreksi yang diperlukan untuk mendapatkan hasil perhitungan yang akurat
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'     maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  OUTPUT:
'     Koreksi LastQuarter dalam Hari. Koreksi ini akan dipakai untuk menghitung saat Last Quarter
'     yang lebih teliti.
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Analogously Example 4-7.a — Calculate the instant of the Last Quarter of the
'                                                        Moon which took place in February 1977.
'      maka :
'        Y = 1977
'        M = 2
'        D = 15
'        x = KoreksiLastQuarter(Y, M, D)
'        x =
'            -0.102383364921171
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 321 dan 322
' -------------------------------------------------------
'*  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E, M, M1, F, O, dF As Double
   
   E = nilai_E(Year, Month, Day, 0.75)
   M = Radians(nilai_M(Year, Month, Day, 0.75))
   M1 = Radians(nilai_M1(Year, Month, Day, 0.75))
   F = Radians(nilai_F(Year, Month, Day, 0.75))
   O = Radians(nilai_O(Year, Month, Day, 0.75))
   
   dF = 0#
   dF = dF - 0.62801 * Sin(M1)
   dF = dF + 0.17172 * E * Sin(M)
   dF = dF - 0.01183 * E * Sin(M1 + M)
   dF = dF + 0.00862 * Sin(2 * M1)
   dF = dF + 0.00804 * Sin(2 * F)
   dF = dF + 0.00454 * E * Sin(M1 - M)
   dF = dF + 0.00204 * E * E * Sin(2 * M)
   dF = dF - 0.0018 * Sin(M1 - 2 * F)
   dF = dF - 0.0007 * Sin(M1 + 2 * F)
   dF = dF - 0.0004 * Sin(3 * M1)
   dF = dF - 0.00034 * E * Sin(2 * M1 - M)
   dF = dF + 0.00032 * E * Sin(M + 2 * F)
   dF = dF + 0.00032 * E * Sin(M - 2 * F)
   dF = dF - 0.00028 * E * E * Sin(M1 + 2 * M)
   dF = dF + 0.00027 * E * Sin(2 * M1 * M)
   dF = dF - 0.00017 * Sin(O)
   dF = dF - 0.00005 * Sin(M1 - M - 2 * F)
   dF = dF + 0.00004 * Sin(2 * M1 + 2 * F)
   dF = dF - 0.00004 * Sin(M1 + M + 2 * F)
   dF = dF + 0.00004 * Sin(M1 - 2 * M)
   dF = dF + 0.00003 * Sin(M1 + M - 2 * F)
   dF = dF + 0.00003 * Sin(3 * M)
   dF = dF + 0.00002 * Sin(2 * M1 - 2 * F)
   dF = dF + 0.00002 * Sin(M1 - M + 2 * F)
   dF = dF - 0.00002 * Sin(3 * M1 + M)
   KoreksiLastQuarter = dF
End Function

Public Function Koreksi_W(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function Koreksi_W (Additional Correction) dipakai untuk menghitung fase seperempat pertama (First Quarter) dan fase seperempat terakhir
'  (Last Quarter).  Koreksi W adalah Koreksi yang diperlukan untuk mendapatkan hasil perhitungan yang akurat
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'     maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  OUTPUT:
'     Koreksi NewMoon dalam Hari. Koreksi ini akan dipakai untuk menghitung saat First Quarter dan Last Quarter
'     yang lebih teliti.
'
'  CATATAN:
'     - Untuk Perhitungan Bulan Baru (New Moon) dan Bulan Purnama (Full Moon), tidak diperlukan koreksi W ini
'     - Untuk first Quarter, koreksi W ini ditambahkan. Sedangkan untuk Last Quarter, koreksi ini dikurangkan.
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'     place in February 1977.
'      maka :
'        Y = 1977
'        M = 2
'        D = 15
'        F = 0.25
'        x = Koreksi_W(Y, M, D, F)
'        x =
'            0.00353463363970559
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 321 dan 322
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim E, M, M1, F, w As Double
   
   If Fase = 0# Or Fase = 0.5 Then
      w = 0#
      Exit Function
   End If
   
   E = nilai_E(Year, Month, Day, Fase)
   M = nilai_M(Year, Month, Day, Fase)
   M1 = nilai_M1(Year, Month, Day, Fase)
   F = nilai_F(Year, Month, Day, Fase)

' Additional correction for  First Quarter (+W) and for last Quarter (-W) page 322
'------------------------------------------------------------------------------
   w = 0.00306
   w = w - 0.00038 * E * Cos(M)
   w = w + 0.00026 * Cos(M1)
   w = w - 0.00002 * Cos(M1 - M)
   w = w + 0.00002 * Cos(M1 + M)
   w = w + 0.00002 * Cos(2 * F)

   Koreksi_W = w
End Function

Public Function Koreksi_PlanetaryArguments(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function Koreksi PlanetaryArguments (Pengaruh Planet-planet) dipakai untuk menghitung fase-fase bulan
'  Koreksi Planetary Arguments adalah Koreksi yang diperlukan untuk mendapatkan hasil perhitungan yang akurat
'
'  INPUT:
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Fase  : Optional. Jika nilainya tidak diberikan maka default-nya adalah 0.
'                * New Moon      :  Fase = 0.00
'                * First Quarter :  Fase = 0.25
'                * Full Moon     :  Fase = 0.50
'                * Last Quarter  :  Fase = 0.75
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'     maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  OUTPUT:
'     Koreksi NewMoon dalam Hari. Koreksi ini akan dipakai untuk menghitung saat First Quarter dan Last Quarter
'     yang lebih teliti.
'
'  CATATAN:
'     - Koreksi ini berlaku untuk semua perhitungan fase-fase bulan.
'
' Contoh pemakaian fungsi:
'     Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'     place in February 1977.
'      maka :
'        Y = 1977
'        M = 2
'        D = 15
'        F = 0.75
'        x = Koreksi_PlanetaryArguments(Y, M, D, F)
'        x =
'            -0.000709617784822135
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 321 dan 322
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim k, T, CP As Double
   Dim A01, A02, A03, A04, A05, A06, A07 As Double
   Dim A08, A09, A10, A11, A12, A13, A14 As Double
   
   k = nilai_K(Year, Month, Day, Fase)
   T = nilai_T(Year, Month, Day, Fase)

' Planetary arguments  page 321:
' ------------------------------
   A01 = Radians(299.77 + 0.107408 * k - 0.009173 * T ^ 2)
   A02 = Radians(251.88 + 0.016321 * k)
   A03 = Radians(251.83 + 26.651886 * k)
   A04 = Radians(349.42 + 36.412478 * k)
   A05 = Radians(84.66 + 18.206239 * k)
   A06 = Radians(141.74 + 53.303771 * k)
   A07 = Radians(207.14 + 2.453732 * k)
   A08 = Radians(154.84 + 7.30686 * k)
   A09 = Radians(34.52 + 27.261239 * k)
   A10 = Radians(207.19 + 0.121824 * k)
   A11 = Radians(291.34 + 1.844379 * k)
   A12 = Radians(161.72 + 24.198154 * k)
   A13 = Radians(239.56 + 25.513099 * k)
   A14 = Radians(331.55 + 3.592518 * k)
    
' Additional correction for all phases
' ------------------------------------
   CP = 0#
   CP = CP + 0.000325 * Sin(A01)
   CP = CP + 0.000165 * Sin(A02)
   CP = CP + 0.000164 * Sin(A03)
   CP = CP + 0.000126 * Sin(A04)
   CP = CP + 0.00011 * Sin(A05)
   CP = CP + 0.000062 * Sin(A06)
   CP = CP + 0.00006 * Sin(A07)
   CP = CP + 0.000056 * Sin(A08)
   CP = CP + 0.000047 * Sin(A09)
   CP = CP + 0.000042 * Sin(A10)
   CP = CP + 0.00004 * Sin(A11)
   CP = CP + 0.000037 * Sin(A12)
   CP = CP + 0.000035 * Sin(A13)
   CP = CP + 0.000023 * Sin(A14)
   Koreksi_PlanetaryArguments = CP
End Function


Public Function MeanObliquityOfEcliptic(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  The mean obliquity of the ecliptic is given by the following
'  formula, adopted by the International Astronomical Union [1] :
'  epsilon0 = 23°26'21".448 - 46".8150T - 0".00059 T^2 + 0".001 813 T^3
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'   MeanObliquityOfEcliptic (Epsilon0) dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Mean Obliquity of the Ecliptic (Delta Epsilon) Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ObliquityOfEcliptic.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = MeanObliquityOfEcliptic(Y, M, D)
'    Hasil:
'           x = 23.4379756875736
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 21. hal. 135
'----------------------------
'*Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   MeanObliquityOfEcliptic = 23 + 26 / 60 + (21.448 - 46.815 * T - 0.00059 * T ^ 2 + 0.001813 * T ^ 3) / 3600#
End Function

Public Function ObliquityOfEcliptic(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' function ObliquityOfEcliptic dimaksudkan untuk menghitung Epsilon (Obliquity of the Ecliptic)
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'   ObliquityOfEcliptic (Epsilon) dalam derajat
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung Obliquity of the Ecliptic (Epsilon) Pada tanggal 12 Februari 2010, jam 10:30:15 WIB
' Langkah-langkah perhitungan yang harus dipersiapkan
'    1. Mengkonversikan waktu lokal ke waktu Universal dengan cara
'       Waktu Universal = Waktu Lokal - Zona Waktu
'       Dalam hal ini WIB berarti Zona Waktunya 7
'    2. Jalankan fungsi ObliquityOfEcliptic.
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = ObliquityOfEcliptic(Y, M, D)
'    Hasil:
'           x = 23.4388802831097
'
' Referensi:
'     Jean Meeus, Astronomical Algorithm bab 21. hal. 131
'     Menurut Jean Meeus : The accuracy of this expression is estimated
'     at 0".01 after 1000 years (that is between A.D. 1000 and 3000)
'     and a few second after 10000 years
'----------------------------
'* Cibinong, 17 Juli 2010
'
' Ditulis oleh : Dr.-Ing. H. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' Pusat Pemetaan Dasar Kelautan dan Kedirgantaraan
' Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   ObliquityOfEcliptic = MeanObliquityOfEcliptic(Year, Month, Day) + NutationInObliquity(Year, Month, Day)
End Function


Public Function nilai_K(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function nilai_K dipakai untuk menghitung fase-fase Bulan
'  nilai k mengindikasikan new moon yang ke k dihitung sejak awal tahun 2000. Sehingga untuk tahun sebelum
'  tahun 2000 nilai k adalah negatif.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Fase  : Optional. Jika nilainya tidak diberikan maka default-nya adalah 0.
'                * New Moon      :  Fase = 0.00
'                * First Quarter :  Fase = 0.25
'                * Full Moon     :  Fase = 0.50
'                * Last Quarter  :  Fase = 0.75
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      F = 0.00
'      k = nilai_K(Y, M, D, F)
'      k =
'          -283
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 320
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Tahun As Double
   
   Tahun = Year + (JulianDatum(Year, Month, Day) - JulianDatum(Year, 1, 1)) / 365.24
   nilai_K = Round((Tahun - 2000#) * 12.3685) + Fase
End Function

Public Function value_K(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function value_K dipakai untuk menghitung maximum deklinasi, opogee, perigee bulan
'  value k mengindikasikan moment yang ke k dihitung sejak awal tahun 2000. Sehingga untuk tahun sebelum
'  tahun 2000 value k adalah negatif.
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 50.a — Greatest northern declination of the Moon in December 1988#
'      maka :
'      Y = 1988
'      M = 12
'      D = 15
'      k = value_K(Y, M, D)
'      k =
'          -148
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 337
' ------------------------------------------
'* Cibinong, 23 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim Tahun As Double
   
   Tahun = Year + (JulianDatum(Year, Month, Day) - JulianDatum(Year, 1, 1)) / 365.24
   value_K = Round((Tahun - 2000#) * 13.3686)
End Function

Public Function nilai_T(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function nilai_T dipakai untuk menghitung fase-fase Bulan
'  nilai T adalah waktu julian century yang dihitung dengan rumus sederhana berdasar nilai k
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Fase  : Optional. Jika nilainya tidak diberikan maka default-nya adalah 0.
'                * New Moon      :  Fase = 0.00
'                * First Quarter :  Fase = 0.25
'                * Full Moon     :  Fase = 0.50
'                * Last Quarter  :  Fase = 0.75
'
' Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 50.a — Greatest northern declination of the Moon in December 1988#
'      maka :
'      Y = 1988
'      M = 12
'      D = 1
'      T = nilai_T(Y, M, D)
'      T =
'          -0.228807050167765
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 320
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    nilai_T = nilai_K(Year, Month, Day, Fase) / 1236.85
End Function

Public Function value_T(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function value_T dipakai untuk menghitung maimum deklinasi, apogee, perigee Bulan
'  value T adalah waktu julian century yang dihitung dengan rumus sederhana berdasar value k
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      F = 0.00
'      T = value_T(Y, M, D, F)
'      T =
'          -0.110707179510196
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 320
' ------------------------------------------
'* Cibinong, 23 Juli 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    value_T = value_K(Year, Month, Day) / 1336.86
End Function


Public Function nilai_E(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function nilai_E dipakai untuk menghitung koordinat ekliptik dan fase-fase Bulan
'  nilai_E  adalah faktor pengali yang merupakan fungsi dari nilai T (lihat fungsi nilai_T).
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Fase  : Optional. Jika nilainya tidak diberikan maka default-nya adalah 0.
'                * New Moon      :  Fase = 0.00
'                * First Quarter :  Fase = 0.25
'                * Full Moon     :  Fase = 0.50
'                * Last Quarter  :  Fase = 0.75
'
' Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      F = 0.00
'      E = nilai_E(Y, M, D, F)
'      E =
'          1.00057529112849
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 308
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
    Dim T As Double
    
    T = nilai_T(Year, Month, Day, Fase)
    nilai_E = 1 - 0.002516 * T - 0.0000074 * T ^ 2
End Function

Public Function nilai_M(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function nilai_M dipakai untuk menghitung fase-fase Bulan
'  nilai M adalah Sun 's mean anomaly at time JDE (Rata-rata anomali Matahari pada saat JDE fase-fase bulan).
'  Lihat fungsi nilai_JDE, nilai_k, nilai_T dll
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Fase  : Optional. Jika nilainya tidak diberikan maka default-nya adalah 0.
'                * New Moon      :  Fase = 0.00
'                * First Quarter :  Fase = 0.25
'                * Full Moon     :  Fase = 0.50
'                * Last Quarter  :  Fase = 0.75
'
' Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      F = 0.00
'      x = nilai_M(Y, M, D, F)
'      x =
'          -8234.26254440997
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 320
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   
   Dim k, T As Double
   
   k = nilai_K(Year, Month, Day, Fase)
   T = nilai_T(Year, Month, Day, Fase)
   nilai_M = 2.5534 + 29.10535669 * k - 0.0000218 * T ^ 2 - 0.00000011 * T ^ 3
End Function

Public Function nilai_M1(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function nilai_M1 dipakai untuk menghitung fase-fase Bulan
'  nilai M adalah Moon 's mean anomaly at time JDE (Rata-rata anomali Bulan pada saat JDE fase-fase bulan).
'  Lihat fungsi nilai_JDE, nilai_k, nilai_T dll
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Fase  : Optional. Jika nilainya tidak diberikan maka default-nya adalah 0.
'                * New Moon      :  Fase = 0.00
'                * First Quarter :  Fase = 0.25
'                * Full Moon     :  Fase = 0.50
'                * Last Quarter  :  Fase = 0.75
'
' Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      F = 0.00
'      x = nilai_M1(Y, M, D, F)
'      x =
'          -108984.627821922
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 320
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
      Dim k, T As Double
   
   k = nilai_K(Year, Month, Day, Fase)
   T = nilai_T(Year, Month, Day, Fase)
   nilai_M1 = 201.5643 + 385.81693528 * k + 0.0107438 * T ^ 2 + 0.00001239 * T ^ 3 - 0.000000058 * T ^ 4
End Function

Public Function nilai_F(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function nilai_F dipakai untuk menghitung fase-fase Bulan
'  nilai F adalah Moon's argument of latitude at time JDE (Argumen Lintang Ekliptik Bulann pada saat JDE fase-fase bulan).
'  Lihat fungsi nilai_JDE, nilai_k, nilai_T dll
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Fase  : Optional. Jika nilainya tidak diberikan maka default-nya adalah 0.
'                * New Moon      :  Fase = 0.00
'                * First Quarter :  Fase = 0.25
'                * Full Moon     :  Fase = 0.50
'                * Last Quarter  :  Fase = 0.75
'
' Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      F = 0.00
'      x = nilai_F(Y, M, D, F)
'      x =
'          -110399.041560942
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 320
' ------------------------------------------
'*  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim k, T As Double
   
   k = nilai_K(Year, Month, Day, Fase)
   T = nilai_T(Year, Month, Day, Fase)
   nilai_F = 160.7108 + 390.67050274 * k - 0.0016341 * T ^ 2 - 0.00000227 * T ^ 3 + 0.000000011 * T ^ 4
End Function

Public Function nilai_O(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function nilai_O dipakai untuk menghitung fase-fase Bulan
'  nilai O(mega) adalah Longitude of the ascending node of the lunar orbit (Bujur Titik pendakian orbit bulan)
'  sekitar JDE fase-fase bulan yang sedang dihitung.
'  Lihat fungsi nilai_JDE, nilai_k, nilai_T dll
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Fase  : Optional. Jika nilainya tidak diberikan maka default-nya adalah 0.
'                * New Moon      :  Fase = 0.00
'                * First Quarter :  Fase = 0.25
'                * Full Moon     :  Fase = 0.50
'                * Last Quarter  :  Fase = 0.75
'
' Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
' Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      F = 0.00
'      x = nilai_O(Y, M, D, F)
'      x =
'          567.317599697147
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 320
' ------------------------------------------
'*  Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim k, T As Double
   
   k = nilai_K(Year, Month, Day, Fase)
   T = nilai_T(Year, Month, Day, Fase)
   nilai_O = 124.7746 - 1.5637558 * k + 0.0020691 * T ^ 2 + 0.00000215 * T ^ 3
End Function

Public Function nilai_Omega(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
' Fungsi nilai_Omega berfungsi untuk menghitung Longitude of the ascending node of the Moon's mean orbit on the
' ecliptic, measured from the mean equinox of the date :
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
' Misal tanggal 12 Februari 2010, jam 10:30:15 WIB
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (10+30/60+15/3600)/24 - 7/24
'           Day   = 12.14600694444440
'
' OUTPUT:
'    nilai_Omega bujur Bulan dalam derajat
'
' CATATAN:
'    Dalam Buku Astronomical Algorithm disimbolkan dengan Omega
'
' CONTOH CARA PEMAKAIAN FUNGSI INI:
' ---------------------------------
' Menghitung  Longitude of the ascending node of the Moon's mean orbit pada 12 Februari 2010, jam 10:30:15 WIB sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = nilai_Omega(Y, M, D)
'    Hasil:
'         x = -70.6008746499872
'           = -70° 36' 3.15"
'           = 289 ° 23 ' 56.70"
'
' Referensi:
'    Jean Meeus, Astronomical Algorithm, hal 308
' -------------------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   nilai_Omega = 125.04452 - 1934.136261 * T + 0.0020708 * T ^ 2 + T ^ 3 / 450000
End Function

Public Function nilai_JDE(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double, Optional ByVal Fase As Double = 0#) As Double
'  Function nilai_JDE dipakai untuk menghitung fase-fase Bulan
'  nilai_JDE  adalah Julian Ephemeris Day saat terjadinya fase-fase bulan. Nilai_JDE dipakai untuk
'  memprediksikan saat fase-fase bulan yang dicari sesuai input yang dimasukkan.
'
'  INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'      - Fase  : Optional. Jika nilainya tidak diberikan maka default-nya adalah 0.
'                * New Moon      :  Fase = 0.00
'                * First Quarter :  Fase = 0.25
'                * Full Moon     :  Fase = 0.50
'                * Last Quarter  :  Fase = 0.75
'
'  Misal tanggal 12 Februari 2010, jam 13:30:15
'    maka   Year  = 2010
'           Month = 2
'           Day   = 12 + (13+30/60+15/3600)/24
'           Day   = 12.56267361111111
'
'  Contoh pemakaian fungsi:
'   Astronomical Algorithm Example 4-7.a — Calculate the instant of the New Moon which took
'   place in February 1977.
'      maka :
'      Y = 1977
'      M = 2
'      D = 15
'      F = 0.00
'      JDE = nilai_JDE(Y, M, D, F)
'      JDE =
'          2443192.9410116
'
'  Referensi:
'    Astronomical Algorithm, Jean Meeus, hal. 319
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim k, T As Double
   
   k = nilai_K(Year, Month, Day, Fase)
   T = nilai_T(Year, Month, Day, Fase)
   nilai_JDE = 2451550.09765 + 29.530588853 * k + 0.0001337 * T ^ 2 - 0.00000015 * T ^ 3 + 0.00000000073 * T ^ 4
End Function

Public Function Argument_A1(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function Argument_A1 dipakai untuk menghitung koordinat Ekliptik Bulan
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  CONTOH:
'  Jalankan fungsi Argument_A1 sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = Argument_A1(Y, M, D)
'    Hasil:
'         x = 133.087042662151
'
' Referensi
'    Astronomical Algorithm Bab 45.
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   Argument_A1 = 119.75 + 131.849 * T
End Function

Public Function Argument_A2(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function Argument_A2 dipakai untuk menghitung koordinat Ekliptik Bulan
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  CONTOH:
'  Jalankan fungsi Argument_A2 sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = Argument_A2(Y, M, D)
'    Hasil:
'         x = 48532.5497014439
'
' Referensi
'    Astronomical Algorithm Bab 45.
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   Argument_A2 = 53.09 + 479264.29 * T
End Function

Public Function Argument_A3(ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Double) As Double
'  Function Argument_A3 dipakai untuk menghitung koordinat Ekliptik Bulan
'
' INPUT:
'      Sistem waktu yang digunakan adalah Waktu Universal atau Universal Time
'      - Year  : tahun Gregorian/Julian (tahun Masehi)
'      - Month : bulan Masehi (1-12)
'      - Day   : hari dalam decimal
'
'  CONTOH:
'  Jalankan fungsi Argument_A3 sbb:
'           Y = 2010
'           M = 2
'           D = 12.14600694444440
'           x = Argument_A3(Y, M, D)
'    Hasil:
'         x = 48995.439465006
'
' Referensi
'    Astronomical Algorithm Bab 45.
' ------------------------------------------
'* Demak, 29 Juni 2012
'
'  Ditulis oleh : Dr.-Ing. H. Khafid
'  Badan Informasi Geospasial
'  Pusat Pemetaan Batas Wilayah
'  Jl. Raya Jakarta-Bogor Km. 46 Cibinong
'-------------------------------------------------
   Dim T As Double
   
   T = JulianEphemerisCentury(Year, Month, Day)
   Argument_A3 = 313.45 + 481266.484 * T
End Function


