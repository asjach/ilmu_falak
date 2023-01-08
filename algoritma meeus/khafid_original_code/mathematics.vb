Public Function Modulus(ByVal A As Double, ByVal B As Double) As Double
    Modulus = A - (B * (A \ B))
    If Modulus < 0# Then Modulus = Modulus + B
End Function

Public Function Radians(ByVal Sudut As Double) As Double
' Fungsi Radians dimaksudkan untuk mengkonversikan sudut dalam
' satuan derajat ke satuan radian.
' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' 15 Juli 2010

   Radians = Sudut * Atn(1#) / 45#
End Function

Public Function Degrees(ByVal Sudut As Double) As Double
' Fungsi Radians dimaksudkan untuk mengkonversikan sudut dalam
' satuan radaian ke satuan derajat.
' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' 15 Juli 2010

   Degrees = Sudut * 45# / Atn(1#)
End Function

Public Function deg2dms(ByVal Sudut As Double, ByVal Code As Byte) As String
   Dim D, M As Integer
   Dim S As Double
' Fungsi deg2dms dimaksudkan untuk mengkonversikan sudut dalam
' satuan derajat (desimal) ke satuan derajat, menit dan detik.
' Hasil konversinya ditulisakan dalam bentuk text,
' Contoh:
'    Juka masukan fungsi ini 125.2323
'    maka luarannya adalah : 125° 13' 56"
' ---------------------------------------------
' Programmer : Dr.-Ing. Khafid
' Badan Koordinasi Survei dan Pemetaan Nasional
' 15 Juli 2010

   D = Int(Abs(Sudut))
   M = Int((Abs(Sudut) - D) * 60#)
   S = (Abs(Sudut) - D - M / 60#) * 3600#
  
 ' Check/koreksi kalau ada hasil pembulatan yang angkanya 60
 ' sehingga hasil konversi seperti 125° 13' 60" dapat dihindarkan
 
   If (S = 60) Then
      S = 0
      M = M + 1
   End If
   If (M = 60) Then
      M = 0
      D = D + 1
   End If
   If (Sudut < 0) Then
     If (Code = 0) Then
        If Format(S, "00") = "60" Then
           S = 0#
           M = M + 1
        End If
        If Format(M, "00") = "60" Then
           M = 0#
           D = D + 1
        End If
        deg2dms = D & "° " & Format(M, "00") & "' " & Format(S, "00") & """ W"
     ElseIf (Code = 1) Then
        If Format(S, "00") = "60" Then
           S = 0#
           M = M + 1
        End If
        If Format(M, "00") = "60" Then
           M = 0#
           D = D + 1
        End If
        deg2dms = D & "° " & Format(M, "00") & "' " & Format(S, "00") & """ S"
     ElseIf (Code = 2) Then
        If Format(S, "00") = "60" Then
           S = 0#
           M = M + 1
        End If
        If Format(M, "00") = "60" Then
           M = 0#
           D = D + 1
        End If
        deg2dms = "-" & D & "° " & Format(M, "00") & "' " & Format(S, "00") & """ "
     Else
        If Format(S, "0.00") = "60.00" Then
           S = 0#
           M = M + 1
        End If
        If Format(M, "00") = "60" Then
           M = 0#
           D = D + 1
        End If
        deg2dms = "-" & D & "° " & Format(M, "00") & "' " & Format(S, "0.00") & """ "
     End If
   Else
     If (Code = 0) Then
        If Format(S, "00") = "60" Then
           S = 0#
           M = M + 1
        End If
        If Format(M, "00") = "60" Then
           M = 0#
           D = D + 1
        End If
        deg2dms = D & "° " & Format(M, "00") & "' " & Format(S, "00") & """ E"
     ElseIf (Code = 1) Then
        If Format(S, "00") = "60" Then
           S = 0#
           M = M + 1
        End If
        If Format(M, "00") = "60" Then
           M = 0#
           D = D + 1
        End If
       deg2dms = D & "° " & Format(M, "00") & "' " & Format(S, "00") & """ N"
     ElseIf (Code = 2) Then
        If Format(S, "00") = "60" Then
           S = 0#
           M = M + 1
        End If
        If Format(M, "00") = "60" Then
           M = 0#
           D = D + 1
        End If
        deg2dms = D & "° " & Format(M, "00") & "' " & Format(S, "00") & """ "
     Else
        If Format(S, "0.00") = "60.00" Then
           S = 0#
           M = M + 1
        End If
        If Format(M, "00") = "60" Then
           M = 0#
           D = D + 1
        End If
        deg2dms = D & "° " & Format(M, "00") & "' " & Format(S, "0.00") & """ "
     End If
   End If
End Function


Public Function deg2hms(ByVal Sudut As Double) As String
   Dim h, M As Integer
   Dim S As Double

   Hours = Sudut / 15#
   h = Int(Abs(Hours))
   M = Int((Abs(Hours) - h) * 60#)
   S = (Abs(Hours) - h - M / 60#) * 3600#
  
   If (S = 60) Then
      S = 0
      M = M + 1
   End If
   If (M = 60) Then
      M = 0
      h = h + 1
   End If
   If Format(S, "0.00") = "60.00" Then
      S = 0#
      M = M + 1
   End If
   If Format(M, "00") = "60" Then
      M = 0#
      h = h + 1
   End If
   deg2hms = h & "h " & Format(M, "00") & "m " & Format(S, "0.000") & "s"
End Function

Public Function Asin(ByVal x As Double) As Double
  Asin = Atn(x / Sqr(1 - x ^ 2))
End Function

Public Function Acos(ByVal x As Double) As Double
   If x = 1 Then
      Acos = 0#
   Else
      Acos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
   End If
End Function

Public Function Atn2(ByVal x As Double, ByVal Y As Double) As Double
   Atn2 = Atn(Y / x)
   If (Y >= 0 And x >= 0) Then
       Atn2 = Abs(Atn2)
   ElseIf (Y >= 0 And x < 0) Then
       Atn2 = 4# * Atn(1#) - Abs(Atn2)
   ElseIf (Y < 0 And x < 0) Then
       Atn2 = 4# * Atn(1#) + Abs(Atn2)
   Else
       Atn2 = 8# * Atn(1#) - Abs(Atn2)
   End If
End Function

Public Function Min(ByVal x As Double, ByVal Y As Double) As Double
   If (x < Y) Then
      Min = x
   Else
      Min = Y
   End If
End Function

Public Function Max(ByVal x As Double, ByVal Y As Double) As Double
   If (x > Y) Then
      Max = x
   Else
      Max = Y
   End If
End Function
