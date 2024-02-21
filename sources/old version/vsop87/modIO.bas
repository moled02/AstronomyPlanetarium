Attribute VB_Name = "modIO"
Private delta(7) As Double

Function StrDate(Datum As tDatum) As String

Dim st As String, hstr As String
Dim i           As Integer

    st = ""
    hstr = Format(Int(Datum.DD), "00")
    st = st + hstr

    hstr = Format(Datum.mm, "00")
    st = st + "-" + hstr + "-"

    hstr = Format(Datum.jj, "0000")
    st = st + hstr
    StrDate = st
End Function



Sub InitIO()
    delta(0) = (0.5): delta(1) = 0.05: delta(2) = 0.00833333333: delta(3) = 0.0008333333333
    delta(4) = 0.00013888889: delta(5) = 0.00001388889: delta(6) = 0.00000138889: delta(7) = 0.00000013889
End Sub

Function StrHMS(time As Double, n As Integer) As String

Dim st, hstr As String
Dim sFormat As String
Dim i As Integer
Call InitIO
    st = ""
    If (time < 0) Then
        st = st + "-"
        time = -time
    End If
    time = time * RToH + delta(n)
    If n = 1 Then
        hstr = Format(Round(time * 10) / 10, "00.0")
        st = st + hstr + "h"
    Else
        hstr = Format(Int(time), "00")
        st = st + hstr + "h"
        time = (time - Int(time)) * 60
        If n = 3 Then
            hstr = Format(time - 60 * delta(3), "00.0")
            st = st + hstr + "m"
        Else
            If (n > 1) Then
                hstr = Format(Int(time), "00")
                st = st + hstr + "m"
                If (n > 3) Then
                    time = (time - Int(time)) * 60
                    If (n = 4) Then
                        hstr = Format(Int(time), "00")
                        st = st + hstr + "s"
                    Else
                        sFormat = String(n - 4, "0")
                        If sFormat <> "" Then
                            sFormat = "." + sFormat
                        End If
                        sFormat = String(n, "0") + sFormat
                        sFormat = Right(sFormat, n - 1)
                        hstr = Format(time - 3600 * delta(n), sFormat)
                        st = st + hstr + "s"
                    End If
                End If
            End If
        End If
    End If
    For i = 1 To Len(st)
      If Mid(st, i, 1) = " " Then st = Left(st, i - 1) + "0" + Mid(st, i + 1)
    Next
    StrHMS = st
End Function


Function StrHMS_DMS(ByVal Angle As Double, ByVal n1 As Long, ByVal n2 As Long, ByVal k90 As Boolean, ByVal k180 As Boolean, ByVal HMSDMS As String, ByVal aantalpos_eerste As Long) As String

'  n1 = optie °, ', °', ", '", °'"; of
'              h , m, hm, s, ms, hms
'  n2 = aantal decimalen in de uitkomst
'  k90 = kleiner 90 gr.

Dim verm1(7) As Double
Dim verm2(7) As Double

'const verm1:array[1..7] of double=(1,  60,  60,  3600,   0, 3600, 3600);
'      verm2:array[1..7] of double=(1,   1,  60,     1,   0,   60, 3600);


Dim st As String, hstr     As String
Dim gmsstr                 As String
Dim sign As Long, i As Long, Pos As Long, decsverm As Long
Dim maxdiep As Long, Max   As Long, n3 As Long
Dim hAngle As Double, tmp  As Double, hFac As Double
Dim modulo                 As Double
Dim gms As Double, decpos  As Long  '{1,2 of 4}
Dim doorgaan               As Boolean

verm1(1) = 1: verm1(2) = 60: verm1(3) = 60: verm1(4) = 3600: verm1(5) = 0: verm1(6) = 3600: verm1(7) = 3600
verm2(1) = 1: verm2(2) = 1: verm2(3) = 60: verm2(4) = 1: verm2(5) = 0: verm2(6) = 60: verm2(7) = 3600

If HMSDMS = "h" Then Angle = Angle / 15
i = 1
decsverm = 1
Do While i <= n2
  decsverm = decsverm * 10
  i = i + 1
Loop

If k180 Then
  Angle = Fmod(Angle, 360#)
  If Angle > 180 Then
    Angle = Angle - 360
  End If
ElseIf k90 Then
  Angle = Fmod(Angle, 180#)
  If Angle > 90 Then
     Angle = Angle - 180
  End If
Else
  Angle = Fmod(Angle, 360#)
End If

If Angle < 0# Then
  sign = -1
  Angle = -Angle
Else
  sign = 1
End If

Select Case n1
Case 1, 3, 7:
    If HMSDMS = "h" Then
        gmsstr = "hms"
    Else
        gmsstr = "°" + Chr(39) + Chr(34)
    End If
Case 2, 6:
    If HMSDMS = "h" Then
        gmsstr = "ms"
    Else
        gmsstr = Chr(39) + Chr(34)
    End If
Case 4:
    If HMSDMS = "h" Then
        gmsstr = "s"
    Else
        gmsstr = Chr(34)
    End If
End Select

Select Case n1
Case 1, 2, 4:
        maxdiep = 1
Case 3, 6:
        maxdiep = 2
Case 7: maxdiep = 3
End Select

'{ n1
'  1     ø
'  3,2   ø' of '
'  7,6,4 ø'" of '" of " }


hAngle = Int(Angle * decsverm * verm1(n1) + 0.5)

hFac = verm2(n1) * decsverm
st = ""
For i = 1 To maxdiep
    If i < maxdiep Then
        tmp = Int(hAngle / hFac)
        hAngle = hAngle - tmp * hFac
        If (i = 1) Then
            hstr = Format(tmp, ZetFormat(aantalpos_eerste, 0))
        Else
            hstr = Format(tmp, "00")
        End If
        st = st + hstr + Mid(gmsstr, i, 1)
        hFac = Int(hFac / 60)
    Else
        If n2 > 0 Then
            n3 = n2 + 1
        Else
            n3 = 0
        End If
        If (i = 1) Then
            hstr = Format(hAngle / decsverm, ZetFormat(aantalpos_eerste, n2))
        Else
            hstr = Format(hAngle / decsverm, ZetFormat(2 + n3, n2))
        End If
        st = st + hstr + Mid(gmsstr, i, 1)
    End If
Next
    i = 2
    Do While (i < Len(st))
        If Mid(st, i, 1) = " " Then
            st = Left(st, i - 1) + "0" + Mid(st, i + 1)
        End If
        i = i + 1
    Loop
    If Left(st, 1) = " " Then
        i = 2
    Else
        i = 1
    End If
    doorgaan = True
    Do While doorgaan
         If i > Len(st) - 1 Then
            doorgaan = False
         Else
            If (Mid(st, i, 1) = "0") And (Mid(st, i + 1, 1) >= "0") And (Mid(st, i + 1, 1) <= "9") Then
               st = Left(st, i - 1) + " " + Mid(st, i + 1)
               i = i + 1
            Else
               doorgaan = False
            End If
        End If
    Loop
    If sign = -1 Then
        If Left(st, 1) = " " Then
            st = "-" + Mid(st, 2)
        Else
            st = "-" + st
        End If
    End If

StrHMS_DMS = st
End Function

Function Fmod(x As Double, Y As Double) As Double
    If Y = 0 Then
      Fmod = x
    Else
      Fmod = x - Y * (Int(x / Y))
    End If
End Function

                        
Function ZetFormat(n1 As Long, n2 As Long) As String
Dim sFormat As String
sFormat = String(n2, "0")
If sFormat <> "" Then
    sFormat = "." + sFormat
End If
ZetFormat = String(n1 - Len(sFormat), "0") + sFormat
If Left(ZetFormat, 1) = "." Then
    ZetFormat = "0" + ZetFormat
End If
End Function
