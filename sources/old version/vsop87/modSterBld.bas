Attribute VB_Name = "modSterBld"
Sub SterBld(ByVal nA As Double, ByVal nD As Double, ByVal T As Double, ByRef snaam As String)

Dim nLower As Double, nUpper As Double, nLowerDec As Double
Dim nFile As Long
Dim doorgaan                  As Boolean
Dim T1 As Double
Dim s                      As String '{ afk. en latijnse naam weglaten}
                                         '{vanaf 26, nederlandse naam}
                                         
Dim asSterBld(357) As String

asSterBld(1) = "00.0000 24.0000  88.0000 UMi Ursa Minor          Kleine Beer"
asSterBld(2) = "08.0000 14.5000  86.5000 UMi Ursa Minor          Kleine Beer"
asSterBld(3) = "21.0000 23.0000  86.1667 UMi Ursa Minor          Kleine Beer"
asSterBld(4) = "18.0000 21.0000  86.0000 UMi Ursa Minor          Kleine Beer"
asSterBld(5) = "00.0000 08.0000  85.0000 Cep Cepheus             Cepheus"
asSterBld(6) = "09.1667 10.6667  82.0000 Cam Camelopardalis      Giraffe"
asSterBld(7) = "00.0000 05.0000  80.0000 Cep Cepheus             Cepheus"
asSterBld(8) = "10.6667 14.5000  80.0000 Cam Camelopardalis      Giraffe"
asSterBld(9) = "17.5000 18.0000  80.0000 UMi Ursa Minor          Kleine Beer"
asSterBld(10) = "20.1667 21.0000  80.0000 Dra Draco               Draak"
asSterBld(11) = "00.0000 03.5083  77.0000 Cep Cepheus             Cepheus"
asSterBld(12) = "11.5000 13.5833  77.0000 Cam Camelopardalis      Giraffe"
asSterBld(13) = "16.5333 17.5000  75.0000 UMi Ursa Minor          Kleine Beer"
asSterBld(14) = "20.1667 20.6667  75.0000 Cep Cepheus             Cepheus"
asSterBld(15) = "07.9667 09.1667  73.5000 Cam Camelopardalis      Giraffe"
asSterBld(16) = "09.1667 11.3333  73.5000 Dra Draco               Draak"
asSterBld(17) = "13.0000 16.5333  70.0000 UMi Ursa Minor          Kleine Beer"
asSterBld(18) = "03.1000 03.4167  68.0000 Cas Cassiopeia          Cassiopeia"
asSterBld(19) = "20.4167 20.6667  67.0000 Dra Draco               Draak"
asSterBld(20) = "11.3333 12.0000  66.5000 Dra Draco               Draak"
asSterBld(21) = "00.0000 00.3333  66.0000 Cep Cepheus             Cepheus"
asSterBld(22) = "14.0000 15.6667  66.0000 UMi Ursa Minor          Kleine Beer"
asSterBld(23) = "23.5833 24.0000  66.0000 Cep Cepheus             Cepheus"
asSterBld(24) = "12.0000 13.5000  64.0000 Dra Draco               Draak"
asSterBld(25) = "13.5000 14.4167  63.0000 Dra Draco               Draak"
asSterBld(26) = "23.1667 23.5833  63.0000 Cep Cepheus             Cepheus"
asSterBld(27) = "06.1000 07.0000  62.0000 Cam Camelopardalis      Giraffe"
asSterBld(28) = "20.0000 20.4167  61.5000 Dra Draco               Draak"
asSterBld(29) = "20.5367 20.6000  60.9167 Cep Cepheus             Cepheus"
asSterBld(30) = "07.0000 07.9667  60.0000 Cam Camelopardalis      Giraffe"
asSterBld(31) = "07.9667 08.4167  60.0000 UMa Ursa Major          Grote Beer"
asSterBld(32) = "19.7667 20.0000  59.5000 Dra Draco               Draak"
asSterBld(33) = "20.0000 20.5367  59.5000 Cep Cepheus             Cepheus"
asSterBld(34) = "22.8667 23.1667  59.0833 Cep Cepheus             Cepheus"
asSterBld(35) = "00.0000 02.4333  58.5000 Cas Cassiopeia          Cassiopeia"
asSterBld(36) = "19.4167 19.7667  58.0000 Dra Draco               Draak"
asSterBld(37) = "01.7000 01.9083  57.5000 Cas Cassiopeia          Cassiopeia"
asSterBld(38) = "02.4333 03.1000  57.0000 Cas Cassiopeia          Cassiopeia"
asSterBld(39) = "03.1000 03.1667  57.0000 Cam Camelopardalis      Giraffe"
asSterBld(40) = "22.3167 22.8667  56.2500 Cep Cepheus             Cepheus"
asSterBld(41) = "05.0000 06.1000  56.0000 Cam Camelopardalis      Giraffe"
asSterBld(42) = "14.0333 14.4167  55.5000 UMa Ursa Major          Grote Beer"
asSterBld(43) = "14.4167 19.4167  55.5000 Dra Draco               Draak"
asSterBld(44) = "03.1667 03.3333  55.0000 Cam Camelopardalis      Giraffe"
asSterBld(45) = "22.1333 22.3167  55.0000 Cep Cepheus             Cepheus"
asSterBld(46) = "20.6000 21.9667  54.8333 Cep Cepheus             Cepheus"
asSterBld(47) = "00.0000 01.7000  54.0000 Cas Cassiopeia          Cassiopeia"
asSterBld(48) = "06.1000 06.5000  54.0000 Lyn Lynx                Lynx"
asSterBld(49) = "12.0833 13.5000  53.0000 UMa Ursa Major          Grote Beer"
asSterBld(50) = "15.2500 15.7500  53.0000 Dra Draco               Draak"
asSterBld(51) = "21.9667 22.1333  52.7500 Cep Cepheus             Cepheus"
asSterBld(52) = "03.3333 05.0000  52.5000 Cam Camelopardalis      Giraffe"
asSterBld(53) = "22.8667 23.3333  52.5000 Cas Cassiopeia          Cassiopeia"
asSterBld(54) = "15.7500 17.0000  51.5000 Dra Draco               Draak"
asSterBld(55) = "02.0417 02.5167  50.5000 Per Perseus             Perseus"
asSterBld(56) = "17.0000 18.2333  50.5000 Dra Draco               Draak"
asSterBld(57) = "00.0000 01.3667  50.0000 Cas Cassiopeia          Cassiopeia"
asSterBld(58) = "01.3667 01.6667  50.0000 Per Perseus             Perseus"
asSterBld(59) = "06.5000 06.8000  50.0000 Lyn Lynx                Lynx"
asSterBld(60) = "23.3333 24.0000  50.0000 Cas Cassiopeia          Cassiopeia"
asSterBld(61) = "13.5000 14.0333  48.5000 UMa Ursa Major          Grote Beer"
asSterBld(62) = "00.0000 01.1167  48.0000 Cas Cassiopeia          Cassiopeia"
asSterBld(63) = "23.5833 24.0000  48.0000 Cas Cassiopeia          Cassiopeia"
asSterBld(64) = "18.1750 18.2333  47.5000 Her Herculus            Herculus"
asSterBld(65) = "18.2333 19.0833  47.5000 Dra Draco               Draak"
asSterBld(66) = "19.0833 19.1667  47.5000 Cyg Cygnus              Zwaan"
asSterBld(67) = "01.6667 02.0417  47.0000 Per Perseus             Perseus"
asSterBld(68) = "08.4167 09.1667  47.0000 UMa Ursa Major          Grote Beer"
asSterBld(69) = "00.1667 00.8667  46.0000 Cas Cassiopeia          Cassiopeia"
asSterBld(70) = "12.0000 12.0833  45.0000 UMa Ursa Major          Grote Beer"
asSterBld(71) = "06.8000 07.3667  44.5000 Lyn Lynx                Lynx"
asSterBld(72) = "21.9083 21.9667  44.0000 Cyg Cygnus              Zwaan"
asSterBld(73) = "21.8750 21.9083  43.7500 Cyg Cygnus              Zwaan"
asSterBld(74) = "19.1667 19.4000  43.5000 Cyg Cygnus              Zwaan"
asSterBld(75) = "09.1667 10.1667  42.0000 UMa Ursa Major          Grote Beer"
asSterBld(76) = "10.1667 10.7833  40.0000 UMa Ursa Major          Grote Beer"
asSterBld(77) = "15.4333 15.7500  40.0000 Boo Bootes              Ossenhoeder"
asSterBld(78) = "15.7500 16.3333  40.0000 Her Herculus            Herculus"
asSterBld(79) = "09.2500 09.5833  39.7500 Lyn Lynx                Lynx"
asSterBld(80) = "00.0000 02.5167  36.7500 And Andromeda           Andromeda"
asSterBld(81) = "02.5167 02.5667  36.7500 Per Perseus             Perseus"
asSterBld(82) = "19.3583 19.4000  36.5000 Lyr Lyra                Lier"
asSterBld(83) = "04.5000 04.6917  36.0000 Per Perseus             Perseus"
asSterBld(84) = "21.7333 21.8750  36.0000 Cyg Cygnus              Zwaan"
asSterBld(85) = "21.8750 22.0000  36.0000 Lac Lacerta             Hagedis"
asSterBld(86) = "06.5333 07.3667  35.5000 Aur Auriga              Voerman"
asSterBld(87) = "07.3667 07.7500  35.5000 Lyn Lynx                Lynx"
asSterBld(88) = "00.0000 02.0000  35.0000 And Andromeda           Andromeda"
asSterBld(89) = "22.0000 22.8167  35.0000 Lac Lacerta             Hagedis"
asSterBld(90) = "22.8167 22.8667  34.5000 Lac Lacerta             Hagedis"
asSterBld(91) = "22.8667 23.5000  34.5000 And Andromeda           Andromeda"
asSterBld(92) = "02.5667 02.7167  34.0000 Per Perseus             Perseus"
asSterBld(93) = "10.7833 11.0000  34.0000 UMa Ursa Major          Grote Beer"
asSterBld(94) = "12.0000 12.3333  34.0000 CVn Canes Venatici      Jachthonden"
asSterBld(95) = "07.7500 09.2500  33.5000 Lyn Lynx                Lynx"
asSterBld(96) = "09.2500 09.8833  33.5000 Lmi Leo Minor           Kleine Leeuw"
asSterBld(97) = "00.7167 01.4083  33.0000 And Andromeda           Andromeda"
asSterBld(98) = "15.1833 15.4333  33.0000 Boo Bootes              Ossenhoeder"
asSterBld(99) = "23.5000 23.7500  32.0833 And Andromeda           Andromeda"
asSterBld(100) = "12.3333 13.2500  32.0000 CVn Canes Venatici      Jachthonden"
asSterBld(101) = "23.7500 24.0000  31.3333 And Andromeda           Andromeda"
asSterBld(102) = "13.9583 14.0333  30.7500 CVn Canes Venatici      Jachthonden"
asSterBld(103) = "02.4167 02.7167  30.6667 Tri Triangulum          Driehoek"
asSterBld(104) = "02.7167 04.5000  30.6667 Per Perseus             Perseus"
asSterBld(105) = "04.5000 04.7500  30.0000 Aur Auriga              Voerman"
asSterBld(106) = "18.1750 19.3583  30.0000 Lyr Lyra                Lier"
asSterBld(107) = "11.0000 12.0000  29.0000 UMa Ursa Major          Grote Beer"
asSterBld(108) = "19.6667 20.9167  29.0000 Cyg Cygnus              Zwaan"
asSterBld(109) = "04.7500 05.8833  28.5000 Aur Auriga              Voerman"
asSterBld(110) = "09.8833 10.5000  28.5000 Lmi Leo Minor           Kleine Leeuw"
asSterBld(111) = "13.2500 13.9583  28.5000 CVn Canes Venatici      Jachthonden"
asSterBld(112) = "00.0000 00.0667  28.0000 And Andromeda           Andromeda"
asSterBld(113) = "01.4083 01.6667  28.0000 Tri Triangulum          Driehoek"
asSterBld(114) = "05.8833 06.5333  28.0000 Aur Auriga              Voerman"
asSterBld(115) = "07.8833 08.0000  28.0000 Gem Gemini              Tweelingen"
asSterBld(116) = "20.9167 21.7333  28.0000 Cyg Cygnus              Zwaan"
asSterBld(117) = "19.2583 19.6667  27.5000 Cyg Cygnus              Zwaan"
asSterBld(118) = "01.9167 02.4167  27.2500 Tri Triangulum          Driehoek"
asSterBld(119) = "16.1667 16.3333  27.0000 CrB Corona borealis     Noorderkroon"
asSterBld(120) = "15.0833 15.1833  26.0000 Boo Bootes              Ossenhoeder"
asSterBld(121) = "15.1833 16.1667  26.0000 CrB Corona borealis     Noorderkroon"
asSterBld(122) = "18.3667 18.8667  26.0000 Lyr Lyra                Lier"
asSterBld(123) = "10.7500 11.0000  25.5000 Lmi Leo Minor           Kleine Leeuw"
asSterBld(124) = "18.8667 19.2583  25.5000 Lyr Lyra                Lier"
asSterBld(125) = "01.6667 01.9167  25.0000 Tri Triangulum          Driehoek"
asSterBld(126) = "00.7167 00.8500  23.7500 Psc Pisces              Vissen"
asSterBld(127) = "10.5000 10.7500  23.5000 Lmi Leo Minor           Kleine Leeuw"
asSterBld(128) = "21.2500 21.4167  23.5000 Vul Vulpecula           Vos"
asSterBld(129) = "05.7000 05.8833  22.8333 Tau Taurus              Stier"
asSterBld(130) = "00.0667 00.1417  22.0000 And Andromeda           Andromeda"
asSterBld(131) = "15.9167 16.0333  22.0000 Ser Serpens             Slang"
asSterBld(132) = "05.8833 06.2167  21.5000 Gem Gemini              Tweelingen"
asSterBld(133) = "19.8333 20.2500  21.2500 Vul Vulpecula           Vos"
asSterBld(134) = "18.8667 19.2500  21.0833 Vul Vulpecula           Vos"
asSterBld(135) = "00.1417 00.8500  21.0000 And Andromeda           Andromeda"
asSterBld(136) = "20.2500 20.5667  20.5000 Vul Vulpecula           Vos"
asSterBld(137) = "07.8083 07.8833  20.0000 Gem Gemini              Tweelingen"
asSterBld(138) = "20.5667 21.2500  19.5000 Vul Vulpecula           Vos"
asSterBld(139) = "19.2500 19.8333  19.1667 Vul Vulpecula           Vos"
asSterBld(140) = "03.2833 03.3667  19.0000 Ari Aries               Ram"
asSterBld(141) = "18.8667 19.0000  18.5000 Sge Sagitta             Pijl"
asSterBld(142) = "05.7000 05.7667  18.0000 Ori Orion               Orion"
asSterBld(143) = "06.2167 06.3083  17.5000 Gem Gemini              Tweelingen"
asSterBld(144) = "19.0000 19.8333  16.1667 Sge Sagitta             Pijl"
asSterBld(145) = "04.9667 05.3333  16.0000 Tau Taurus              Stier"
asSterBld(146) = "15.9167 16.0833  16.0000 Her Herculus            Herculus"
asSterBld(147) = "19.8333 20.2500  15.7500 Sge Sagitta             Pijl"
asSterBld(148) = "04.6167 04.9667  15.5000 Tau Taurus              Stier"
asSterBld(149) = "05.3333 05.6000  15.5000 Tau Taurus              Stier"
asSterBld(150) = "12.8333 13.5000  15.0000 Com Coma Berenices      Haar van Berenice"
asSterBld(151) = "17.2500 18.2500  14.3333 Her Herculus            Herculus"
asSterBld(152) = "11.8667 12.8333  14.0000 Com Coma Berenices      Haar van Berenice"
asSterBld(153) = "07.5000 07.8083  13.5000 Gem Gemini              Tweelingen"
asSterBld(154) = "16.7500 17.2500  12.8333 Her Herculus            Herculus"
asSterBld(155) = "00.0000 00.1417  12.5000 Peg Pegasus             Pegasus"
asSterBld(156) = "05.6000 05.7667  12.5000 Tau Taurus              Stier"
asSterBld(157) = "07.0000 07.5000  12.5000 Gem Gemini              Tweelingen"
asSterBld(158) = "21.1167 21.3333  12.5000 Peg Pegasus             Pegasus"
asSterBld(159) = "06.3083 06.9333  12.0000 Gem Gemini              Tweelingen"
asSterBld(160) = "18.2500 18.8667  12.0000 Her Herculus            Herculus"
asSterBld(161) = "20.8750 21.0500  11.8333 Del Delphinus           Dolfijn"
asSterBld(162) = "21.0500 21.1167  11.8333 Peg Pegasus             Pegasus"
asSterBld(163) = "11.5167 11.8667  11.0000 Leo Leo                 Leeuw"
asSterBld(164) = "06.2417 06.3083  10.0000 Ori Orion               Orion"
asSterBld(165) = "06.9333 07.0000  10.0000 Gem Gemini              Tweelingen"
asSterBld(166) = "07.8083 07.9250  10.0000 Cnc Cancer              Kreeft"
asSterBld(167) = "23.8333 24.0000  10.0000 Peg Pegasus             Pegasus"
asSterBld(168) = "01.6667 03.2833  09.9167 Ari Aries               Ram"
asSterBld(169) = "20.1417 20.3000  08.5000 Del Delphinus           Dolfijn"
asSterBld(170) = "13.5000 15.0833  08.0000 Boo Bootes              Ossenhoeder"
asSterBld(171) = "22.7500 23.8333  07.5000 Peg Pegasus             Pegasus"
asSterBld(172) = "07.9250 09.2500  07.0000 Cnc Cancer              Kreeft"
asSterBld(173) = "09.2500 10.7500  07.0000 Leo Leo                 Leeuw"
asSterBld(174) = "18.2500 18.6622  06.2500 Oph Ophuichus           Slangendrager"
asSterBld(175) = "18.6622 18.8667  06.2500 Aql Aquila              Arend"
asSterBld(176) = "20.8333 20.8750  06.0000 Del Delphinus           Dolfijn"
asSterBld(177) = "07.0000 07.0167  05.5000 CMi Canis Minor         Kleine Hond"
asSterBld(178) = "18.2500 18.4250  04.5000 Ser Serpens             Slang"
asSterBld(179) = "16.0833 16.7500  04.0000 Her Herculus            Herculus"
asSterBld(180) = "18.2500 18.4250  03.0000 Oph Ophuichus           Slangendrager"
asSterBld(181) = "21.4667 21.6667  02.7500 Peg Pegasus             Pegasus"
asSterBld(182) = "00.0000 02.0000  02.0000 Psc Pisces              Vissen"
asSterBld(183) = "18.5833 18.8667  02.0000 Ser Serpens             Slang"
asSterBld(184) = "20.3000 20.8333  02.0000 Del Delphinus           Dolfijn"
asSterBld(185) = "20.8333 21.3333  02.0000 Equ Equuleus            Klein Paard"
asSterBld(186) = "21.3333 21.4667  02.0000 Peg Pegasus             Pegasus"
asSterBld(187) = "22.0000 22.7500  02.0000 Peg Pegasus             Pegasus"
asSterBld(188) = "21.6667 22.0000  01.7500 Peg Pegasus             Pegasus"
asSterBld(189) = "07.0167 07.2000  01.5000 CMi Canis Minor         Kleine Hond"
asSterBld(190) = "03.5833 04.6167  00.0000 Tau Taurus              Stier"
asSterBld(191) = "04.6167 04.6667  00.0000 Ori Orion               Orion"
asSterBld(192) = "07.2000 08.0833  00.0000 CMi Canis Minor         Kleine Hond"
asSterBld(193) = "14.6667 15.0833  00.0000 Vir Virgo               Maagd"
asSterBld(194) = "17.8333 18.2500  00.0000 Oph Ophuichus           Slangendrager"
asSterBld(195) = "02.6500 03.2833 -01.7500 Cet Cetus               Walvis"
asSterBld(196) = "03.2833 03.5833 -01.7500 Tau Taurus              Stier"
asSterBld(197) = "15.0833 16.2667 -03.2500 Ser Serpens             Slang"
asSterBld(198) = "04.6667 05.0833 -04.0000 Ori Orion               Orion"
asSterBld(199) = "05.8333 06.2417 -04.0000 Ori Orion               Orion"
asSterBld(200) = "17.8333 17.9667 -04.0000 Ser Serpens             Slang"
asSterBld(201) = "18.2500 18.5833 -04.0000 Ser Serpens             Slang"
asSterBld(202) = "18.5833 18.8667 -04.0000 Aql Aquila              Arend"
asSterBld(203) = "22.7500 23.8333 -04.0000 Psc Pisces              Vissen"
asSterBld(204) = "10.7500 11.5167 -06.0000 Leo Leo                 Leeuw"
asSterBld(205) = "11.5167 11.8333 -06.0000 Vir Virgo               Maagd"
asSterBld(206) = "00.0000 00.3333 -07.0000 Psc Pisces              Vissen"
asSterBld(207) = "23.8333 24.0000 -07.0000 Psc Pisces              Vissen"
asSterBld(208) = "14.2500 14.6667 -08.0000 Vir Virgo               Maagd"
asSterBld(209) = "15.9167 16.2667 -08.0000 Oph Ophuichus           Slangendrager"
asSterBld(210) = "20.0000 20.5333 -09.0000 Aql Aquila              Arend"
asSterBld(211) = "21.3333 21.8667 -09.0000 Aqr Aquarius            Waterman"
asSterBld(212) = "17.1667 17.9667 -10.0000 Oph Ophuichus           Slangendrager"
asSterBld(213) = "05.8333 08.0833 -11.0000 Mon Monoceros           Eenhoorn"
asSterBld(214) = "04.9167 05.0833 -11.0000 Eri Eridanus            Eridanus"
asSterBld(215) = "05.0833 05.8333 -11.0000 Ori Orion               Orion"
asSterBld(216) = "08.0833 08.3667 -11.0000 Hya Hydra               Waterman"
asSterBld(217) = "09.5833 10.7500 -11.0000 Sex Sextans             Sextant"
asSterBld(218) = "11.8333 12.8333 -11.0000 Vir Virgo               Maagd"
asSterBld(219) = "17.5833 17.6667 -11.6667 Oph Ophuichus           Slangendrager"
asSterBld(220) = "18.8667 20.0000 -12.0333 Aql Aquila              Arend"
asSterBld(221) = "04.8333 04.9167 -14.5000 Eri Eridanus            Eridanus"
asSterBld(222) = "20.5333 21.3333 -15.0000 Aqr Aquarius            Waterman"
asSterBld(223) = "17.1667 18.2500 -16.0000 Ser Serpens             Slang"
asSterBld(224) = "18.2500 18.8667 -16.0000 Sct Scutum              Schild"
asSterBld(225) = "08.3667 08.5833 -17.0000 Hya Hydra               Waterman"
asSterBld(226) = "16.2667 16.3750 -18.2500 Oph Ophuichus           Slangendrager"
asSterBld(227) = "08.5833 09.0833 -19.0000 Hya Hydra               Waterman"
asSterBld(228) = "10.7500 10.8333 -19.0000 Crt Crater              Beker"
asSterBld(229) = "16.2667 16.3750 -19.2500 Oph Ophuichus           Slangendrager"
asSterBld(230) = "15.6667 15.9167 -20.0000 Lib Libra               Weegschaal"
asSterBld(231) = "12.5833 12.8333 -22.0000 Crv Corvus              Raaf"
asSterBld(232) = "12.8333 14.2500 -22.0000 Vir Virgo               Maagd"
asSterBld(233) = "09.0833 09.7500 -24.0000 Hya Hydra               Waterman"
asSterBld(234) = "01.6667 02.6500 -24.3833 Cet Cetus               Walvis"
asSterBld(235) = "02.6500 03.7500 -24.3833 Eri Eridanus            Eridanus"
asSterBld(236) = "10.8333 11.8333 -24.5000 Crt Crater              Beker"
asSterBld(237) = "11.8333 12.5833 -24.5000 Crv Corvus              Raaf"
asSterBld(238) = "14.2500 14.9167 -24.5000 Lib Libra               Weegschaal"
asSterBld(239) = "16.2667 16.7500 -24.5833 Oph Ophuichus           Slangendrager"
asSterBld(240) = "00.0000 01.6667 -25.5000 Cet Cetus               Walvis"
asSterBld(241) = "21.3333 21.8667 -25.5000 Cap Capricornus         Steenbok"
asSterBld(242) = "21.8667 23.8333 -25.5000 Aqr Aquarius            Waterman"
asSterBld(243) = "23.8333 24.0000 -25.5000 Cet Cetus               Walvis"
asSterBld(244) = "09.7500 10.2500 -26.5000 Hya Hydra               Waterman"
asSterBld(245) = "04.7000 04.8333 -27.2500 Eri Eridanus            Eridanus"
asSterBld(246) = "04.8333 06.1167 -27.2500 Lep Lepus               Haas"
asSterBld(247) = "20.0000 21.3333 -28.0000 Cap Capricornus         Steenbok"
asSterBld(248) = "10.2500 10.5833 -29.1667 Hya Hydra               Waterman"
asSterBld(249) = "12.5833 14.9167 -29.5000 Hya Hydra               Waterman"
asSterBld(250) = "14.9167 15.6667 -29.5000 Lib Libra               Weegschaal"
asSterBld(251) = "15.6667 16.0000 -29.5000 Sco Scorpius            Schorpioen"
asSterBld(252) = "04.5833 04.7000 -30.0000 Eri Eridanus            Eridanus"
asSterBld(253) = "16.7500 17.6000 -30.0000 Oph Ophuichus           Slangendrager"
asSterBld(254) = "17.6000 17.8333 -30.0000 Sgr Sagittarius         Schutter"
asSterBld(255) = "10.5833 10.8333 -31.1667 Hya Hydra               Waterman"
asSterBld(256) = "06.1167 07.3667 -33.0000 CMa Canis Major         Grote Hond"
asSterBld(257) = "12.2500 12.5833 -33.0000 Hya Hydra               Waterman"
asSterBld(258) = "10.8333 12.2500 -35.0000 Hya Hydra               Waterman"
asSterBld(259) = "03.5000 03.7500 -36.0000 For Fornax              Oven"
asSterBld(260) = "08.3667 09.3667 -36.7500 Pyx Pyxis               Kompas"
asSterBld(261) = "04.2667 04.5833 -37.0000 Eri Eridanus            Eridanus"
asSterBld(262) = "17.8333 19.1667 -37.0000 Sgr Sagittarius         Schutter"
asSterBld(263) = "21.3333 23.0000 -37.0000 PsA Pisces Austrinus    Zuidervis"
asSterBld(264) = "23.0000 23.3333 -37.0000 Scl Sculptor            Beeldhouwer"
asSterBld(265) = "03.0000 03.5000 -39.5833 For Fornax              Oven"
asSterBld(266) = "09.3667 11.0000 -39.7500 Ant Antlia              Luchtpomp"
asSterBld(267) = "00.0000 01.6667 -40.0000 Scl Sculptor            Beeldhouwer"
asSterBld(268) = "01.6667 03.0000 -40.0000 For Fornax              Oven"
asSterBld(269) = "03.8667 04.2667 -40.0000 Eri Eridanus            Eridanus"
asSterBld(270) = "23.3333 24.0000 -40.0000 Scl Sculptor            Beeldhouwer"
asSterBld(271) = "14.1667 14.9167 -42.0000 Cen Centaurus           Centaur"
asSterBld(272) = "15.6667 16.0000 -42.0000 Lup Lupus               Wolf"
asSterBld(273) = "16.0000 16.4208 -42.0000 Sco Scorpius            Schorpioen"
asSterBld(274) = "04.8333 05.0000 -43.0000 Cae Caelum              Graveerschrift"
asSterBld(275) = "05.0000 06.5833 -43.0000 Col Columba             Duif"
asSterBld(276) = "08.0000 08.3667 -43.0000 Pup Puppis              Achtersteven"
asSterBld(277) = "03.4167 03.8667 -44.0000 Eri Eridanus            Eridanus"
asSterBld(278) = "16.4208 17.8333 -45.5000 Sco Scorpius            Schorpioen"
asSterBld(279) = "17.8333 19.1667 -45.5000 CrA Corona Australis    Zuiderkroon"
asSterBld(280) = "19.1667 20.3333 -45.5000 Sgr Sagittarius         Schutter"
asSterBld(281) = "20.3333 21.3333 -45.5000 Mic Microscopium        Microscoop"
asSterBld(282) = "03.0000 03.4167 -46.0000 Eri Eridanus            Eridanus"
asSterBld(283) = "04.5000 04.8333 -46.5000 Cae Caelum              Graveerschrift"
asSterBld(284) = "15.3333 15.6667 -48.0000 Lup Lupus               Wolf"
asSterBld(285) = "00.0000 02.3333 -48.1667 Phe Phoenix             Phoenix"
asSterBld(286) = "02.6667 03.0000 -49.0000 Eri Eridanus            Eridanus"
asSterBld(287) = "04.0833 04.2667 -49.0000 Hor Horologium          Slingeruurwerk"
asSterBld(288) = "04.2667 04.5000 -49.0000 Cae Caelum              Graveerschrift"
asSterBld(289) = "21.3333 22.0000 -50.0000 Gru Grus                Kraanvogel"
asSterBld(290) = "06.0000 08.0000 -50.7500 Pup Puppis              Achtersteven"
asSterBld(291) = "08.0000 08.1667 -50.7500 Vel Vela                Zeilen"
asSterBld(292) = "02.4167 02.6667 -51.0000 Eri Eridanus            Eridanus"
asSterBld(293) = "03.8333 04.0833 -51.0000 Hor Horologium          Slingeruurwerk"
asSterBld(294) = "00.0000 01.8333 -51.5000 Phe Phoenix             Phoenix"
asSterBld(295) = "06.0000 06.1667 -52.5000 Car Carina              Kiel"
asSterBld(296) = "08.1667 08.4500 -53.0000 Vel Vela                Zeilen"
asSterBld(297) = "03.5000 03.8333 -53.1667 Hor Horologium          Slingeruurwerk"
asSterBld(298) = "03.8333 04.0000 -53.1667 Dor Dorado              Zwaardvis"
asSterBld(299) = "00.0000 01.5833 -53.5000 Phe Phoenix             Phoenix"
asSterBld(300) = "02.1667 02.4167 -54.0000 Eri Eridanus            Eridanus"
asSterBld(301) = "04.5000 05.0000 -54.0000 Pic Pictor              Schilder"
asSterBld(302) = "15.0500 15.3333 -54.0000 Lup Lupus               Wolf"
asSterBld(303) = "08.4500 08.8333 -54.5000 Vel Vela                Zeilen"
asSterBld(304) = "06.1667 06.5000 -55.0000 Car Carina              Kiel"
asSterBld(305) = "11.8333 12.8333 -55.0000 Cen Centaurus           Centaur"
asSterBld(306) = "14.1667 15.0500 -55.0000 Lup Lupus               Wolf"
asSterBld(307) = "15.0500 15.3333 -55.0000 Nor Norma               Winkelhaak"
asSterBld(308) = "04.0000 04.3333 -56.5000 Dor Dorado              Zwaardvis"
asSterBld(309) = "08.8333 11.0000 -56.5000 Vel Vela                Zeilen"
asSterBld(310) = "11.0000 11.2500 -56.5000 Cen Centaurus           Centaur"
asSterBld(311) = "17.5000 18.0000 -57.0000 Ara Ara                 Altaar"
asSterBld(312) = "18.0000 20.3333 -57.0000 Tel Telescopium         Telescoop"
asSterBld(313) = "22.0000 23.3333 -57.0000 Gru Grus                Kraanvogel"
asSterBld(314) = "03.2000 03.5000 -57.5000 Hor Horologium          Slingeruurwerk"
asSterBld(315) = "05.0000 05.5000 -57.5000 Pic Pictor              Schilder"
asSterBld(316) = "06.5000 06.8333 -58.0000 Car Carina              Kiel"
asSterBld(317) = "00.0000 01.3333 -58.5000 Phe Phoenix             Phoenix"
asSterBld(318) = "01.3333 02.1667 -58.5000 Eri Eridanus            Eridanus"
asSterBld(319) = "23.3333 24.0000 -58.5000 Phe Phoenix             Phoenix"
asSterBld(320) = "04.3333 04.5833 -59.0000 Dor Dorado              Zwaardvis"
asSterBld(321) = "15.3333 16.4208 -60.0000 Nor Norma               Winkelhaak"
asSterBld(322) = "20.3333 21.3333 -60.0000 Ind Indus               Indiaan"
asSterBld(323) = "05.5000 06.0000 -61.0000 Pic Pictor              Schilder"
asSterBld(324) = "15.1667 15.3333 -61.0000 Cir Circinus            Passer"
asSterBld(325) = "16.4208 16.5833 -61.0000 Ara Ara                 Altaar"
asSterBld(326) = "14.9167 15.1667 -63.5833 Cir Circinus            Passer"
asSterBld(327) = "16.5833 16.7500 -63.5833 Ara Ara                 Altaar"
asSterBld(328) = "06.0000 06.8333 -64.0000 Pic Pictor              Schilder"
asSterBld(329) = "06.8333 09.0333 -64.0000 Car Carina              Kiel"
asSterBld(330) = "11.2500 11.8333 -64.0000 Cen Centaurus           Centaur"
asSterBld(331) = "11.8333 12.8333 -64.0000 Cru Crux                Zuiderkruis"
asSterBld(332) = "12.8333 14.5333 -64.0000 Cen Centaurus           Centaur"
asSterBld(333) = "13.5000 13.6667 -65.0000 Cir Circinus            Passer"
asSterBld(334) = "16.7500 16.8333 -65.0000 Ara Ara                 Altaar"
asSterBld(335) = "02.1667 03.2000 -67.5000 Hor Horologium          Slingeruurwerk"
asSterBld(336) = "03.2000 04.5833 -67.5000 Ret Reticulum           Net"
asSterBld(337) = "14.7500 14.9167 -67.5000 Cir Circinus            Passer"
asSterBld(338) = "16.8333 17.5000 -67.5000 Ara Ara                 Altaar"
asSterBld(339) = "17.5000 18.0000 -67.5000 Pav Pavo                Pauw"
asSterBld(340) = "22.0000 23.3333 -67.5000 Tuc Tucana              Toekan"
asSterBld(341) = "04.5833 06.5833 -70.0000 Dor Dorado              Zwaardvis"
asSterBld(342) = "13.6667 14.7500 -70.0000 Cir Circinus            Passer"
asSterBld(343) = "14.7500 17.0000 -70.0000 TrA Triangulum Australe Zuiderdriehoek"
asSterBld(344) = "00.0000 01.3333 -75.0000 Tuc Tucana              Toekan"
asSterBld(345) = "03.5000 04.5833 -75.0000 Hyi Hudrus              Kleine Waterslang"
asSterBld(346) = "06.5833 09.0333 -75.0000 Vol Voltans             Vliegende Vis"
asSterBld(347) = "09.0333 11.2500 -75.0000 Car Carina              Kiel"
asSterBld(348) = "11.2500 13.6667 -75.0000 Mus Musca               Vlieg"
asSterBld(349) = "18.0000 21.3333 -75.0000 Pav Pavo                Pauw"
asSterBld(350) = "21.3333 23.3333 -75.0000 Ind Indus               Indiaan"
asSterBld(351) = "23.3333 24.0000 -75.0000 Tuc Tucana              Toekan"
asSterBld(352) = "00.7500 01.3333 -76.0000 Tuc Tucana              Toekan"
asSterBld(353) = "00.0000 03.5000 -82.5000 Hyi Hudrus              Kleine Waterslang"
asSterBld(354) = "07.6667 13.6667 -82.5000 Cha Chameleon           Kameleon"
asSterBld(355) = "13.6667 18.0000 -82.5000 Aps Apus                Paradijsvogel"
asSterBld(356) = "03.5000 07.6667 -85.0000 Men Mensa               Tafelberg"
asSterBld(357) = "00.0000 24.0000 -90.0000 Oct Octans              Octant"

nFile = FreeFile
  T1 = JDToT(2405889.5)
  Call PrecessFK5(T, T1, nA, nD)
  nA = nA * 12 / Pi: nD = nD * 180 / Pi
  doorgaan = True
  i = 1
  While (doorgaan)
      nLower = Val(Left(asSterBld(i), 7))
      nUpper = Val(Mid(asSterBld(i), 9, 7))
      nLowerDec = Val(Mid(asSterBld(i), 17, 8))
      #If FRANS Then
        s = Mid(asSterBld(i), 30, 20)
      #Else
        s = Mid(asSterBld(i), 50)
      #End If
      snaam = Left(s + "                         ", 25)
      snaam = Left(snaam, 25)
      If (nA >= nLower) And (nA <= nUpper) And (nD >= nLowerDec) Then doorgaan = False
      i = i + 1
  Wend
  Close (nFile)
End Sub

Sub T()
Dim snaam As String
Call SterBld(5.5 * Pi / 12, 0, 0, snaam)
End Sub