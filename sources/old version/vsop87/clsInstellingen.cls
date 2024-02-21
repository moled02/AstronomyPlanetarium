VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInstellingen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'unit Const2

'interface
'uses aaio, aatime
Private Type tInst
               aanwezig    As Boolean '{indien gevuld, dan waarde 1'}
               opmaak_hgms As Byte
               aantalpos_eerste_hgms As Byte
               urengraad   As String * 1 '{h of d '}
               k90         As Boolean
               k180        As Boolean
               h_RTS_hgms As Double
               dec_hgms    As Byte

               totaalpos_getal As Byte
                sign_getal As String * 1
               dec_getal As Byte

               Opmaak_datum As Byte

               Opmaak_sterrenbeeld As Byte
               Lengte_sterrenbeeld As Byte

               cTypeGegeven As String * 1 '{n:numeriek of s:string'}
               niettonen As Boolean '{indien gevuld, waarde 1}
End Type

Private cDateOpmaak As String
Private cHgmsOpmaak As Variant '(1 To 11) As String
Private cHgmsRTSOpmaak As String
Private cGetalOpmaak As Variant '(1 To 5) As String
Private cStringOpmaak  As String
Private xInstelling As tInst  ' overdracht gegevens

Const cdDate = 1
Const cTime = 2
Const cJulDat = 3
Const cDeltaT = 4

Const cl_helio = 5
Const cb_helio = 6
Const cr_helio = 7
Const cl_geoc = 8
Const cb_geoc = 9
Const cr_geoc = 10
Const cl_geom = 11
Const cb_geom = 12
Const cr_geom = 13

Const ca_equ2000 = 89
Const cd_equ2000 = 90
Const ca_equm = 14
Const cd_equm = 15
Const ca_equt = 16
Const cd_equt = 17
Const ca_top = 18
Const cd_top = 19

Const cL_eclt = 20
Const cB_eclt = 21
Const cL_olib = 22
Const cB_olib = 23
Const cL_plib = 24
Const cB_plib = 25
Const cL_tlib = 26
Const cB_tlib = 27

Const cAz = 28
Const cHoogte = 29
Const cHoogteAtm = 30
Const cMaxhoogte = 31

Const cRTS_R1 = 32
Const cRTS_T1 = 33
Const cRTS_S1 = 34
Const cRTS_R2 = 35
Const cRTS_T2 = 36
Const cRTS_S2 = 37
Const cRTS_R3 = 38
Const cRTS_T3 = 39
Const cRTS_S3 = 40

Const cL0 = 41
Const cB0 = 42
Const cP = 43

Const cFase = 44
Const cFase_corr = 45
Const cFasehoek = 46
Const cMagnitude = 47
Const cElongatie = 48
Const cParhoek = 49

Const cSemidiam = 50
Const cSemiequ = 51
Const cSemipol = 52
Const cTop_diam = 53
Const cGeoc_diam = 54
Const cPar = 55
Const cAfstand = 56
Const cTerm = 57

Const cDe = 58
Const cDs = 59
Const cQ = 60
Const cQQ = 61
Const cDeff_ill = 62
Const cPos_h = 63
Const cPosh_p = 64
Const cPosh_limb = 65
Const cPosh_Q = 66
Const cPsh_p = 67

Const cc = 68
Const cOm = 69
Const cOm1 = 70
Const cOm2 = 71

Const cRing_P = 72
Const cRing_P1 = 73
Const cRing_U = 74
Const cRing_U1 = 75
Const cRing_B = 76
Const cRing_B1 = 77
Const cRing_SemiAsBinnen = 78
Const cRing_SemiAsBuiten = 79
'{ onderstaande voor tekenen ban diverse schijven in de ring: bv. Cassini-scheiding '}
Const cRing_ioASBinnen = 80
Const cRing_oiASBinnen = 81
Const cRing_iiASBinnen = 82
Const cRing_idASBinnen = 83
Const cRing_ioASBuiten = 84
Const cRing_oiASBuiten = 85
Const cRing_iiASBuiten = 86
Const cRing_idASBuiten = 87

Const cSterrenbeeld = 88
      

'const
'      cDateOpmaak as string       = '|date|'
'      cHgmsOpmaak as array (1 to 11) of string       =
                                   '('|time|l_helio|b_helio|l_geoc|b_geoc|l_geom|b_geom|',
                                   '|a_equm|d_equm|a_equt|d_equt|a_top|d_top|'                                ,
                                   '|l_eclt|b_eclt|l_olib|b_olib|l_plib|b_plib|l_tlib|b_tlib|'                ,
                                   '|az|hoogte|hoogteatm|maxhoogte|'                                          ,
                                   '|l0|b0|p|'                                                                ,
                                   '|fasehoek|elongatie|semidiam|semiequ|semipol|'                            ,
                                   '|top_diam|geoc_diam|par|term|parhoek|'                                    ,
                                   '|de|ds|deff_ill|pos_h|posh_p|posh_limb|posh_q|psp_p|om|om1|om2|'          ,
                                   '|ring_p|ring_p1|ring_u|ring_u1|ring_b|ring_b1|ring_semiasbinnen|ring_semiasbuiten|' ,
                                   '|ring_ioasbinnen|ring_oiasbinnen|ring_iiasbinnen|ring_idasbinnen|' ,
                                   '|ring_ioasbuiten|ring_oiasbuiten|ring_iiasbuiten|ring_idasbuiten|')
'      cHgmsRTSOpmaak as string    = '|rts_r1|rts_t1|rts_s1|rts_r2|rts_t2|rts_s2|rts_r3|rts_t3|rts_s3|'
'      cGetalOpmaak   as array (1 to 5) of string    =
'                                   ('|juldat|deltat|' ,
                                   '|r_helio|r_geoc|r_geom|',
                                   '|fase|fase_corr|' ,
                                   '|magnitude|'      ,
                                   '|afstand|')
'      cStringOpmaak  as string    = '|sterrenbeeld|'

Private Enum ePlaneten
 Zon
 Mercurius
 Venus
 Aarde
 Mars
 Jupiter
 Saturnus
 Uranus
 Neptunus
 Pluto
 maan
'      tPlaneten=(Zon, Mercurius, Venus, Aarde, Mars, Jupiter,
'               Saturnus, Uranus, Neptunus, Pluto, Maan)
End Enum


'(* opties instellingen:
'          graden as 1(g), 2(m), 3(gm), 4(s), 5(ms), 6(gms), 7(h), 8(hm), 9(hms)
'          2e waarde = aantal decimalen
'*)



Private Type tInstHgms
    hgms As Byte
    dec As Byte
End Type

Private Type tInstGetal
    sign As String * 1 '{'J' = ja'}
    totaalPos As Byte
    dec As Byte
End Type
Private Type tRTS
    Rise As tInstHgms
    Transit As tInstHgms
    Setting As tInstHgms
End Type

'{    tDate =
'           record
'               opmaak as string(30)
'           End
'
'    tInstZon = record
'                   Date: tDate
'Time:                       tInstHgms
'JulDat:                     tInstGetal
'DeltaT:                     tInstGetal
'                   l_helio, b_helio, r_helio,
'                   l_geoc,  b_geoc,  r_geoc,
'                   l_geom,  b_geom,  r_geom,
'                   A_equm, D_equm,
'                   A_equt, D_equt,
'                   A_top, D_top,
'                   Az, Hoogte,
'                   HoogteAtm , MaxHoogte: tInstHgms
'                   RTS1 , RTS2, RTS3: tRTS
'                  End
'
'    tInstAarde = record
'                   Date: tDate
'Time:                       tInstHgms
'JulDat:                     tInstGetal
'DeltaT:                     tInstGetal
'                   L_helio , b_helio, r_helio: tInstHgms
'                  End
'
'    tInstOverigePlaneten = record
'                   Date: tDate
'Time:                       tInstHgms
'JulDat:                     tInstGetal
'DeltaT:                     tInstGetal
'                   l_helio, b_helio, r_helio,
'                   l_geoc,  b_geoc,  r_geoc,
'                   l_geom,  b_geom,  r_geom,
'                   A_equm, D_equm,
'                   A_equt, D_equt,
'                   A_top, D_top,
'                   Az, Hoogte,
'                   HoogteAtm , MaxHoogte: tInstHgms
'                   RTS1 , RTS2, RTS3: tRTS
'fase:                     tInstGetal
'fasehoek:                     tInstHgms
'magnitude:                     tInstGetal
'                   elongatie , Semidiam: tInstHgms
'               End
'    tInstMars = record
'                   Date: tDate
'Time:                       tInstHgms
'JulDat:                     tInstGetal
'DeltaT:                     tInstGetal
'                   l_helio, b_helio, r_helio,
'                   l_geoc,  b_geoc,  r_geoc,
'                   l_geom,  b_geom,  r_geom,
'                   A_equm, D_equm,
'                   A_equt, D_equt,
'                   A_top, D_top,
'                   Az, Hoogte,
'                   HoogteAtm , MaxHoogte: tInstHgms
'                   RTS1 , RTS2, RTS3: tRTS
'                   DE, DS,
'                   Posh_p,
'                   def_ill , posh_Q: tInstHgms
'fase:                     tInstGetal
'fasehoek:                     tInstHgms
'magnitude:                     tInstGetal
'                   elongatie , Semidiam: tInstHgms
'               End
'    tInstJupiter =
'               record
'                   Date: tDate
'Time:                       tInstHgms
'JulDat:                     tInstGetal
'DeltaT:                     tInstGetal
'                   l_helio, b_helio, r_helio,
'                   l_geoc,  b_geoc,  r_geoc,
'                   l_geom,  b_geom,  r_geom,
'                   A_equm, D_equm,
'                   A_equt, D_equt,
'                   A_top, D_top,
'                   Az, Hoogte,
'                   HoogteAtm , MaxHoogte: tInstHgms
'                   RTS1 , RTS2, RTS3: tRTS
'                   de, ds, psh_p,
'                   Om1 , Om2: tInstHgms
'fase_corr:                     tInstGetal
'fase:                     tInstGetal
'fasehoek:                     tInstHgms
'magnitude:                     tInstGetal
'                   elongatie , SemiEqu, SemiPol: tInstHgms
'               End
'    tInstSaturnus =
'               record
'                   Date: tDate
'Time:                       tInstHgms
'JulDat:                     tInstGetal
'DeltaT:                     tInstGetal
'                   l_helio, b_helio, r_helio,
'                   l_geoc,  b_geoc,  r_geoc,
'                   l_geom,  b_geom,  r_geom,
'                   A_equm, D_equm,
'                   A_equt, D_equt,
'                   A_top, D_top,
'                   Az, Hoogte,
'                   HoogteAtm , MaxHoogte: tInstHgms
'                   RTS1 , RTS2, RTS3: tRTS
'                   P, P1, U,
'                   U1, B, B1,
'                   psh_p,
'                   SemiAsRingBinnen , SemiAsRingBuiten: tInstHgms
'fase:                     tInstGetal
'fasehoek:                    tInstHgms
'magnitude:                     tInstGetal
'                   elongatie , SemiEqu, SemiPol: tInstHgms
'               End
'    tInstMaan = record
'                   Date   As Byte
'Time:                       tInstHgms
'JulDat:                     tInstGetal
'DeltaT:                     tInstGetal
'                   l_eclt,  b_eclt,
'                   Par , geoc_diam: tInstHgms
'afstand:                     tInstGetal
'                   A_equm, D_equm,
'                   A_equt, D_equt,
'                   A_top, D_top,
'                   top_diam,
'                   Az , Hoogte, HoogteAtm: tInstHgms
'                   pos_h, fase,
'                   posh_Limb,
'                   l_olib, b_olib,
'                   l_plib, b_plib,
'                   l_tlib, b_tlib,
'                   term , posh_Q: tInstHgms
'RTS:                     tRTS
'                End
'}

'(* Alle gegevens worden opgeslagen in een array *)
'(* Met onderstaand array wordt aangegeven welke optie moet worden meegenomen
'   in de uitvoer *)

Private Volgorde(0 To 10, 0 To 90) As Long
Private Type tGegevens
    gegevensDouble(0 To 90) As Double
    GegevensString(0 To 90) As String
End Type
Private Gegevens As tGegevens
Private Instelling(0 To 10, 1 To 90) As tInst
Private StringArray(0 To 90) As String
    '(*

'Date, Time ,Juldat, DeltaT, L_helio, b_helio, r_helio,
'                             l_geom,  b_geom , r_geom,
'                             l_geoc,  b_geoc , r_geoc,
' A_equm , D_equm, A_equt, D_equt, A_top, D_top, Az, Hoogte, HoogteAtm, MaxHoogte'
''
' ZON (l0, b0, p)
' OVERIGE PLANETEN(de, ds, psh_p, Om, def_ill, posh_Q) of
' JUPITER (de, ds, psh_p, om1, om2, fase_corr) of
' SATURNUS (lat_aarde=P, lat_zon=P1, posh_p=U, U1, B, B1,
'           SemiAsRingBinnen,SemiAsRingBuiten) of
' PLUTO(),
' MAAN (l_eclt,  b_eclt, par, geoc_diam, afstand,
'       top_diam, pos_h, posh_Limb, l_olib, b_olib,
'       l_plib, b_plib, l_tlib, b_tlib, term, posh_Q)'
'
'fase , fasehoek, magnitude, elongatie
''
'SemiEqu, SemiPol of
'Semidiam,
'
'Rise1, Trans1, Set1,
'Rise2, Trans2, Set2,
'' Rise3 , Trans3, Set3
'Sterrenbeeld
'*)


Private sIndexInstString(0 To 90) As String
'Private Instelling As tInstellingen
'{    Instellingen as record
'Zon:                         tInstZon
'                       Mercurius,
'                       Venus,
'Aarde:                         tInstOverigePlaneten'
'
'Mars:                         tInstMars
'Jupiter:                         tInstJupiter
'Saturnus:                         tInstSaturnus
'
'                       Uranus,
'                       Neptunus,
'Pluto:                         tInstOverigePlaneten
'
'Maan:                         tInstMaan
'                    End '}
 
Function IndexPlaneet(cPlaneet As String) As Long
Dim nPlaneet As Long
     If cPlaneet = "ZON" Then nPlaneet = 0
     If cPlaneet = "MERCURIUS" Then nPlaneet = 1
     If cPlaneet = "VENUS" Then nPlaneet = 2
     If cPlaneet = "AARDE" Then nPlaneet = 3
     If cPlaneet = "MARS" Then nPlaneet = 4
     If cPlaneet = "JUPITER" Then nPlaneet = 5
     If cPlaneet = "SATURNUS" Then nPlaneet = 6
     If cPlaneet = "URANUS" Then nPlaneet = 7
     If cPlaneet = "NEPTUNUS" Then nPlaneet = 8
     If cPlaneet = "PLUTO" Then nPlaneet = 9
     If cPlaneet = "MAAN" Then nPlaneet = 10
     IndexPlaneet = nPlaneet
End Function

Sub LeesIndexInstellingen(ByVal nFile As Long)

Dim sIndexRegel As String
Dim C As String * 1
Dim doorgaan As Boolean
Dim nIndexInstelling As Long
Dim sIndexInstelling As String
     nIndexInstelling = 0
     doorgaan = Not EOF(nFile)
     Do While doorgaan
          Line Input #nFile, sIndexRegel
          If sIndexRegel = "[EINDE INDEX-INSTELLINGEN]" Then
             doorgaan = False
          Else
              If sIndexRegel <> "" Then
                  nIndexInstelling = nIndexInstelling + 1
                  sIndexInstString(nIndexInstelling) = sIndexRegel
                  doorgaan = Not EOF(nFile)
              End If
          End If
     Loop
End Sub

Function IndexInstelling(cInstelling As String) As Long
Dim nPos As Long
Dim i As Long, code As Long
Dim doorgaan As Boolean
    doorgaan = True
    i = 0
    Do While doorgaan
         i = i + 1
         If sIndexInstString(i) = cInstelling Then
            doorgaan = False
        End If
    Loop
'{    nPos = instr(sIndexInstString, 0,'|'+cInstelling+'|')
'    val(mid(sIndexInstString,nPos-2,2),i,code)'}
    IndexInstelling = i
End Function

Function NaamInstelling(nIndexInstelling As Long) As String
Dim nPos1 As Long, nPos2 As Long
Dim cIndexInstelling As String
'{    str(nIndexInstelling:2,cIndexInstelling)
'    if cIndexInstelling(1) = ' ' then
'       cIndexInstelling(1) = '0'
'    nPos1 = instr(sIndexInstString, 0,'|'+cIndexInstelling+'|')
'    nPos2 = instr(sIndexInstString,nPos1+4,'|')
'    NaamInstelling = mid(sIndexInstString, nPos1 + 4, nPos2 - nPos1 - 4) '}
    NaamInstelling = sIndexInstString(nIndexInstelling)
End Function

Function InstellingInString(cZoekstring As String, cInstelling As String) As Boolean
Dim lResultaat As Boolean
Dim doorgaan      As Boolean
Dim i             As Long
     lResultaat = False
     If cZoekstring = "DateOpmaak" Then lResultaat = InStr(cDateOpmaak, "|" + cInstelling + "|") > 0
     If cZoekstring = "HgmsOpmaak" Then
         doorgaan = True
         i = 0
         Do While doorgaan
            If i = 11 Then
               doorgaan = False
            Else
               i = i + 1
               lResultaat = InStr(cHgmsOpmaak(i), "|" + cInstelling + "|") > 0
               doorgaan = Not lResultaat
            End If
        Loop
     End If
     If cZoekstring = "HgmsRTSOpmaak" Then lResultaat = InStr(cHgmsRTSOpmaak, "|" + cInstelling + "|") > 0
     If cZoekstring = "GetalOpmaak" Then
         doorgaan = True
         i = 0
         Do While doorgaan
            If i = 5 Then
               doorgaan = False
            Else
               i = i + 1
               lResultaat = InStr(cGetalOpmaak(i), "|" + cInstelling + "|") > 0
               doorgaan = Not lResultaat
            End If
        Loop
     End If
     If cZoekstring = "StringOpmaak" Then lResultaat = InStr(cStringOpmaak, "|" + cInstelling + "|") > 0
     InstellingInString = lResultaat
End Function

Sub VerwerkInstelling(nPlaneet As Long, cInstelling As String, cGegevens As String)

Dim nIndexInstelling As Long
     nIndexInstelling = IndexInstelling(cInstelling)
     cGegevens = Trim(cGegevens)
     With Instelling(nPlaneet, nIndexInstelling)
         .niettonen = False
         .aanwezig = True
                                                    '{ date '}
         If InstellingInString("DateOpmaak", cInstelling) Then
            .Opmaak_datum = Int(Val(cGegevens))
            .cTypeGegeven = "n"
            .niettonen = InStr(cGegevens, "nt") > 0
         End If
                                                    '{ hgms '}
         If InstellingInString("HgmsOpmaak", cInstelling) Then
            .opmaak_hgms = Int(Val(Mid(cGegevens, 1, 1)))
            .dec_hgms = Int(Val(Mid(cGegevens, 3, 1)))
            .aantalpos_eerste_hgms = Int(Val(Mid(cGegevens, 5, 1)))
            .urengraad = Mid(cGegevens, 7, 1)
            .h_RTS_hgms = 0
            .cTypeGegeven = "n"
            .k90 = InStr(cGegevens, "k90") > 0
            .k180 = InStr(cGegevens, "k180") > 0
            .niettonen = InStr(cGegevens, "nt") > 0
         End If
                                                    '{ hgms '}
         If InstellingInString("HgmsRTSOpmaak", cInstelling) Then
            .opmaak_hgms = Int(Val(Mid(cGegevens, 1, 1)))
            .dec_hgms = Int(Val(Mid(cGegevens, 3, 1)))
            .aantalpos_eerste_hgms = Int(Val(Mid(cGegevens, 5, 1)))
            .urengraad = Mid(cGegevens, 7, 1)
            .h_RTS_hgms = Val(Mid(cGegevens, 9, 3))
            .cTypeGegeven = "n"
            .k90 = False
            .k180 = False
            .niettonen = InStr(cGegevens, "nt") > 0
         End If
                                                    '{ getal '}
         If InstellingInString("GetalOpmaak", cInstelling) Then
              .sign_getal = Left(cGegevens, 1)
              .totaalpos_getal = Int(Val(Mid(cGegevens, 3, 2)))
              .dec_getal = Int(Val(Mid(cGegevens, 6, 2)))
              .cTypeGegeven = "n"
              .niettonen = InStr(cGegevens, "nt") > 0
         End If
         If InstellingInString("StringOpmaak", cInstelling) Then
              .Opmaak_sterrenbeeld = Int(Val(Mid(cGegevens, 1, 1)))
              .Lengte_sterrenbeeld = Int(Val(Mid(cGegevens, 3, 2)))
              .cTypeGegeven = "s"
              .niettonen = InStr(cGegevens, "nt") > 0
         End If
     End With
End Sub

Sub verwerkregel(nPlaneet As Long, cRegel As String)

Dim cPlaneet As String
Dim nPos        As Long
Dim cInstelling As String
Static nVolgorde As Long
    If Left(cRegel, 1) = "[" Then
       cPlaneet = Mid(cRegel, 2, InStr(cRegel, "]") - 2)
       nPlaneet = IndexPlaneet(cPlaneet)
       nVolgorde = 0
       Volgorde(nPlaneet, nVolgorde + 1) = 0
    Else
       If cRegel <> "" Then
           nPos = InStr(cRegel, " ")
           cInstelling = Mid(cRegel, 1, nPos - 1)
           Call VerwerkInstelling(nPlaneet, cInstelling, Mid(cRegel, nPos + 1))
           nVolgorde = nVolgorde + 1
           Volgorde(nPlaneet, nVolgorde) = IndexInstelling(cInstelling)
           Volgorde(nPlaneet, nVolgorde + 1) = 0
       End If
    End If
End Sub

Sub leesopmaakbestand(sFile)
Dim cRegel As String
Dim nPlaneet As Long
Dim nFile As Long
    nFile = FreeFile
     Open sFile For Input As #nFile
     LeesIndexInstellingen (nFile)
     Do While Not EOF(nFile)
         Line Input #nFile, cRegel
         Call verwerkregel(nPlaneet, cRegel)
     Loop
     Close (nFile)
End Sub

Sub OpslaanGegeven(nPlaneet As Long, cGegeven As String, nWaarde As Double, cWaarde As String)
Dim nIndexGegeven As Long
     nIndexGegeven = IndexInstelling(cGegeven) '{ plaats opslag, gelijk aan instelling '}
     If Instelling(nPlaneet, nIndexGegeven).cTypeGegeven = "n" Then
        Gegevens.gegevensDouble(nIndexGegeven) = nWaarde
        If cWaarde <> "" Then Gegevens.GegevensString(nIndexGegeven) = cWaarde
     Else
        Gegevens.GegevensString(nIndexGegeven) = cWaarde
    End If
End Sub

Sub GeefInstelling(nPlaneet As Long, cInstelling As String, gevonden As Boolean)
'Sub GeefInstelling(nPlaneet As Integer, cInstelling As String, xInstelling As tInst, gevonden As Boolean)

Dim nIndexInstelling As Long
Dim nVolgordeIndex      As Long
Dim doorgaan As Boolean
    nIndexInstelling = IndexInstelling(cInstelling)
    nVolgordeIndex = 0
    gevonden = nIndexInstelling > 0
    If gevonden Then
        doorgaan = True
        Do While doorgaan
            nVolgordeIndex = nVolgordeIndex + 1
            If Volgorde(nPlaneet, nVolgordeIndex) = 0 Then doorgaan = False
            If Volgorde(nPlaneet, nVolgordeIndex) = nIndexInstelling Then doorgaan = False
        Loop
        gevonden = Volgorde(nPlaneet, nVolgordeIndex) > 0
        If gevonden Then
            xInstelling = Instelling(nPlaneet, nIndexInstelling)
            gevonden = Instelling(nPlaneet, nVolgordeIndex).aanwezig
        Else
           xInstelling = Instelling(0, 1)
        End If
    End If
End Sub

Sub InstellingAanUit(nPlaneet As Long, cInstelling As String, InstellingAan As Boolean)

Dim nIndexInstelling As Long
Dim nVolgordeIndex  As Long
Dim doorgaan As Boolean
    nIndexInstelling = IndexInstelling(cInstelling)
    If nIndexInstelling > 0 Then
        nVolgordeIndex = 0
        doorgaan = True
        Do While doorgaan
            nVolgordeIndex = nVolgordeIndex + 1
            If Volgorde(nPlaneet, nVolgordeIndex) = 0 Then doorgaan = False
            If Volgorde(nPlaneet, nVolgordeIndex) = nIndexInstelling Then doorgaan = False
        Loop
        If Volgorde(nPlaneet, nVolgordeIndex) > 0 Then Instelling(nPlaneet, nIndexInstelling).aanwezig = InstellingAan
    End If
End Sub

Public Function GeefGegevenMetOpmaak(ByVal nPlaneet As Long, ByVal cGegeven As String, ByVal nGegeven As Long) As String
Dim nIndexInstelling As Long, nPos As Long
Dim nWaarde As Double
Dim cWaarde As String
Dim cNaamInstelling As String, hStr As String
Dim dDate As tDatum
Dim cRes As String

    If nGegeven > 0 Then
       nIndexInstelling = nGegeven
       cNaamInstelling = NaamInstelling(nIndexInstelling)
    Else
       nIndexInstelling = IndexInstelling(cGegeven)
       cNaamInstelling = cGegeven
    End If
    If Instelling(nPlaneet, nIndexInstelling).aanwezig Then
        nWaarde = Gegevens.gegevensDouble(nIndexInstelling)
        cWaarde = Gegevens.GegevensString(nIndexInstelling)
    Else
        nWaarde = 0
        cWaarde = ""
    End If

    With Instelling(nPlaneet, nIndexInstelling)
                                                    '{ date '}
         If InstellingInString("DateOpmaak", cNaamInstelling) Then
              If .Opmaak_datum = 1 Then
                  dDate = JDNaarKalender(nWaarde)
                  cRes = StrDate(dDate)
              End If
         End If
                                                    '{ hgms '}
         If InstellingInString("HgmsOpmaak", cNaamInstelling) Then
              cRes = StrHMS_DMS(nWaarde, .opmaak_hgms, .dec_hgms, .k90, .k180, Left(.urengraad, 1), .aantalpos_eerste_hgms)
         End If
                                                    '{ hgms '}
         If InstellingInString("HgmsRTSOpmaak", cNaamInstelling) Then
              If (nWaarde = 0) And (cWaarde <> "") Then
                  hStr = "----------"
                  nPos = .aantalpos_eerste_hgms + 1
                  If (.opmaak_hgms = 3) Or (.opmaak_hgms = 6) Then nPos = nPos + 3
                  If .opmaak_hgms = 7 Then nPos = nPos + 5
                  If .dec_hgms > 0 Then nPos = nPos + 1
                  nPos = nPos + .dec_hgms
                  hStr = Left(hStr, nPos)
                  cRes = hStr
              Else
                  cRes = StrHMS_DMS(nWaarde, .opmaak_hgms, .dec_hgms, .k90, .k180, Left(.urengraad, 1), .aantalpos_eerste_hgms)
              End If
         End If
                                                    '{ getal '}
         If InstellingInString("GetalOpmaak", cNaamInstelling) Then
             'call str (nWaarde:totaalpos_getal: dec_getal, cWaarde)
             If .dec_getal = 0 Then
                cWaarde = Format(nWaarde, "0")
             Else
                cWaarde = Format(nWaarde, "0." & String(.dec_getal, "0"))
'                cWaarde = Format(nWaarde, String(.totaalpos_getal - 1 - .dec_getal, "0") & "." & String(.dec_getal, "0"))
             End If
             cWaarde = String(.totaalpos_getal - Len(cWaarde), " ") + cWaarde
             '{ nog wel even sign controleren '}
             cRes = cWaarde
         End If

         If InstellingInString("StringOpmaak", cNaamInstelling) Then
             cRes = Mid(Trim(cWaarde) + "                              ", 1, .Lengte_sterrenbeeld)
         End If
         If .aanwezig Then
            GeefGegevenMetOpmaak = cRes
         Else
            If Len(cRes) > 0 Then
               GeefGegevenMetOpmaak = Mid("                                            ", 1, Len(cRes))
            Else
               GeefGegevenMetOpmaak = ""
            End If
     End If
     If .niettonen Then
        GeefGegevenMetOpmaak = String(Len(GeefGegevenMetOpmaak), " ")
     End If
    End With
End Function

Public Function leesvolgorde(ByVal nPlaneet As Long, ByVal nVolgordeIndex As Long) As Long
     leesvolgorde = Volgorde(nPlaneet, nVolgordeIndex)
End Function



Private Sub Class_Initialize()
cDateOpmaak = "|date|"
cHgmsOpmaak = Array("|geen_instelling|", _
                    "|time|l_helio|b_helio|l_geoc|b_geoc|l_geom|b_geom|", _
                    "|a_equ2000|d_equ2000|a_equm|d_equm|a_equt|d_equt|a_top|d_top|", _
                    "|l_eclt|b_eclt|l_olib|b_olib|l_plib|b_plib|l_tlib|b_tlib|", _
                    "|az|hoogte|hoogteatm|maxhoogte|", _
                    "|l0|b0|p|", _
                    "|fasehoek|elongatie|semidiam|semiequ|semipol|", _
                    "|top_diam|geoc_diam|par|term|parhoek|", _
                    "|de|ds|deff_ill|pos_h|posh_p|posh_limb|posh_q|psp_p|om|om1|om2|", _
                    "|ring_p|ring_p1|ring_u|ring_u1|ring_b|ring_b1|ring_semiasbinnen|ring_semiasbuiten|", _
                    "|ring_ioasbinnen|ring_oiasbinnen|ring_iiasbinnen|ring_idasbinnen|", _
                    "|ring_ioasbuiten|ring_oiasbuiten|ring_iiasbuiten|ring_idasbuiten|")
cHgmsRTSOpmaak = "|rts_r1|rts_t1|rts_s1|rts_r2|rts_t2|rts_s2|rts_r3|rts_t3|rts_s3|"
cGetalOpmaak = Array("|geen_instelling|", _
                     "|juldat|deltat|", _
                     "|r_helio|r_geoc|r_geom|", _
                     "|fase|fase_corr|", _
                     "|magnitude|", _
                     "|afstand|")
cStringOpmaak = "|sterrenbeeld|"

'Call leesopmaakbestand
'    for i=0 to 10 do
'       volgorde^(i, 1) = 0
'    new(sIndexInstString)
'{   dispose(Instelling)
'    dispose (Gegevens)
'    dispose (Volgorde)
'    dispose (sIndexInstString) '}
End Sub
