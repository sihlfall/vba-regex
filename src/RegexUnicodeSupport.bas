Attribute VB_Name = "RegexUnicodeSupport"
Option Explicit

Private UnicodeCanonLookupTable(0 To 65535) As Integer
Public UnicodeCanonRunsTable(0 To 302) As Long

Public UnicodeInitialized As Boolean

Public Sub UnicodeInitialize()
    InitializeUnicodeCanonLookupTable UnicodeCanonLookupTable
    InitializeUnicodeCanonRunsTable UnicodeCanonRunsTable
    UnicodeInitialized = True
End Sub

Public Function ReCanonicalizeChar(ByVal codepoint As Long) As Long
    ' ! This function must leave negative values of codepoint unchanged, as ENDOFINPUT is represented by -1.
    If codepoint And &HFFFF0000 Then ' Codepoint not in [0;&HFFFF&]
        ReCanonicalizeChar = codepoint
    Else
        ReCanonicalizeChar = (codepoint + UnicodeCanonLookupTable(codepoint)) And &HFFFF&
    End If
End Function

Public Function UnicodeIsLineTerminator(ByVal codepoint As Long) As Boolean
    ' ! This function must return False for negative values of codepoint, as ENDOFINPUT is represented by -1.
    If codepoint = &HA& Then
        UnicodeIsLineTerminator = True
    ElseIf codepoint = &HD& Then
        UnicodeIsLineTerminator = True
    ElseIf codepoint < &H2028& Then
        UnicodeIsLineTerminator = False
    ElseIf codepoint > &H2029& Then
        UnicodeIsLineTerminator = False
    Else
        UnicodeIsLineTerminator = True
    End If
End Function

Private Sub InitializeUnicodeCanonLookupTable(ByRef t() As Integer)
    Dim i As Long
    
    For i = 97 To 122: t(i) = -32: Next i
    t(181) = 743
    For i = 224 To 246: t(i) = -32: Next i
    For i = 248 To 254: t(i) = -32: Next i
    t(255) = 121
    For i = 257 To 303 Step 2: t(i) = -1: Next i
    t(307) = -1
    t(309) = -1
    t(311) = -1
    For i = 314 To 328 Step 2: t(i) = -1: Next i
    For i = 331 To 375 Step 2: t(i) = -1: Next i
    t(378) = -1
    t(380) = -1
    t(382) = -1
    t(384) = 195
    t(387) = -1
    t(389) = -1
    t(392) = -1
    t(396) = -1
    t(402) = -1
    t(405) = 97
    t(409) = -1
    t(410) = 163
    t(414) = 130
    t(417) = -1
    t(419) = -1
    t(421) = -1
    t(424) = -1
    t(429) = -1
    t(432) = -1
    t(436) = -1
    t(438) = -1
    t(441) = -1
    t(445) = -1
    t(447) = 56
    t(453) = -1
    t(454) = -2
    t(456) = -1
    t(457) = -2
    t(459) = -1
    t(460) = -2
    For i = 462 To 476 Step 2: t(i) = -1: Next i
    t(477) = -79
    For i = 479 To 495 Step 2: t(i) = -1: Next i
    t(498) = -1
    t(499) = -2
    t(501) = -1
    For i = 505 To 543 Step 2: t(i) = -1: Next i
    For i = 547 To 563 Step 2: t(i) = -1: Next i
    t(572) = -1
    t(575) = 10815
    t(576) = 10815
    t(578) = -1
    For i = 583 To 591 Step 2: t(i) = -1: Next i
    t(592) = 10783
    t(593) = 10780
    t(594) = 10782
    t(595) = -210
    t(596) = -206
    t(598) = -205
    t(599) = -205
    t(601) = -202
    t(603) = -203
    t(604) = -23217
    t(608) = -205
    t(609) = -23221
    t(611) = -207
    t(613) = -23256
    t(614) = -23228
    t(616) = -209
    t(617) = -211
    t(618) = -23228
    t(619) = 10743
    t(620) = -23231
    t(623) = -211
    t(625) = 10749
    t(626) = -213
    t(629) = -214
    t(637) = 10727
    t(640) = -218
    t(642) = -23229
    t(643) = -218
    t(647) = -23254
    t(648) = -218
    t(649) = -69
    t(650) = -217
    t(651) = -217
    t(652) = -71
    t(658) = -219
    t(669) = -23275
    t(670) = -23278
    t(837) = 84
    t(881) = -1
    t(883) = -1
    t(887) = -1
    t(891) = 130
    t(892) = 130
    t(893) = 130
    t(940) = -38
    t(941) = -37
    t(942) = -37
    t(943) = -37
    For i = 945 To 961: t(i) = -32: Next i
    t(962) = -31
    For i = 963 To 971: t(i) = -32: Next i
    t(972) = -64
    t(973) = -63
    t(974) = -63
    t(976) = -62
    t(977) = -57
    t(981) = -47
    t(982) = -54
    t(983) = -8
    For i = 985 To 1007 Step 2: t(i) = -1: Next i
    t(1008) = -86
    t(1009) = -80
    t(1010) = 7
    t(1011) = -116
    t(1013) = -96
    t(1016) = -1
    t(1019) = -1
    For i = 1072 To 1103: t(i) = -32: Next i
    For i = 1104 To 1119: t(i) = -80: Next i
    For i = 1121 To 1153 Step 2: t(i) = -1: Next i
    For i = 1163 To 1215 Step 2: t(i) = -1: Next i
    For i = 1218 To 1230 Step 2: t(i) = -1: Next i
    t(1231) = -15
    For i = 1233 To 1327 Step 2: t(i) = -1: Next i
    For i = 1377 To 1414: t(i) = -48: Next i
    For i = 4304 To 4346: t(i) = 3008: Next i
    t(4349) = 3008
    t(4350) = 3008
    t(4351) = 3008
    For i = 5112 To 5117: t(i) = -8: Next i
    t(7296) = -6254
    t(7297) = -6253
    t(7298) = -6244
    t(7299) = -6242
    t(7300) = -6242
    t(7301) = -6243
    t(7302) = -6236
    t(7303) = -6181
    t(7304) = -30270
    t(7545) = -30204
    t(7549) = 3814
    t(7566) = -30152
    For i = 7681 To 7829 Step 2: t(i) = -1: Next i
    t(7835) = -59
    For i = 7841 To 7935 Step 2: t(i) = -1: Next i
    For i = 7936 To 7943: t(i) = 8: Next i
    For i = 7952 To 7957: t(i) = 8: Next i
    For i = 7968 To 7975: t(i) = 8: Next i
    For i = 7984 To 7991: t(i) = 8: Next i
    For i = 8000 To 8005: t(i) = 8: Next i
    t(8017) = 8
    t(8019) = 8
    t(8021) = 8
    t(8023) = 8
    For i = 8032 To 8039: t(i) = 8: Next i
    t(8048) = 74
    t(8049) = 74
    t(8050) = 86
    t(8051) = 86
    t(8052) = 86
    t(8053) = 86
    t(8054) = 100
    t(8055) = 100
    t(8056) = 128
    t(8057) = 128
    t(8058) = 112
    t(8059) = 112
    t(8060) = 126
    t(8061) = 126
    t(8112) = 8
    t(8113) = 8
    t(8126) = -7205
    t(8144) = 8
    t(8145) = 8
    t(8160) = 8
    t(8161) = 8
    t(8165) = 7
    t(8526) = -28
    For i = 8560 To 8575: t(i) = -16: Next i
    t(8580) = -1
    For i = 9424 To 9449: t(i) = -26: Next i
    For i = 11312 To 11358: t(i) = -48: Next i
    t(11361) = -1
    t(11365) = -10795
    t(11366) = -10792
    t(11368) = -1
    t(11370) = -1
    t(11372) = -1
    t(11379) = -1
    t(11382) = -1
    For i = 11393 To 11491 Step 2: t(i) = -1: Next i
    t(11500) = -1
    t(11502) = -1
    t(11507) = -1
    For i = 11520 To 11557: t(i) = -7264: Next i
    t(11559) = -7264
    t(11565) = -7264
    For i = 42561 To 42605 Step 2: t(i) = -1: Next i
    For i = 42625 To 42651 Step 2: t(i) = -1: Next i
    For i = 42787 To 42799 Step 2: t(i) = -1: Next i
    For i = 42803 To 42863 Step 2: t(i) = -1: Next i
    t(42874) = -1
    t(42876) = -1
    For i = 42879 To 42887 Step 2: t(i) = -1: Next i
    t(42892) = -1
    t(42897) = -1
    t(42899) = -1
    t(42900) = 48
    For i = 42903 To 42921 Step 2: t(i) = -1: Next i
    For i = 42933 To 42943 Step 2: t(i) = -1: Next i
    t(42947) = -1
    t(43859) = -928
    For i = 43888 To 43967: t(i) = 26672: Next i
    For i = 65345 To 65370: t(i) = -32: Next i
End Sub

Private Sub InitializeUnicodeCanonRunsTable(ByRef t() As Long)
    t(0) = 97
    t(1) = 123
    t(2) = 181
    t(3) = 182
    t(4) = 224
    t(5) = 247
    t(6) = 248
    t(7) = 255
    t(8) = 256
    t(9) = 257
    t(10) = 304
    t(11) = 307
    t(12) = 312
    t(13) = 314
    t(14) = 329
    t(15) = 331
    t(16) = 376
    t(17) = 378
    t(18) = 383
    t(19) = 384
    t(20) = 385
    t(21) = 387
    t(22) = 390
    t(23) = 392
    t(24) = 393
    t(25) = 396
    t(26) = 397
    t(27) = 402
    t(28) = 403
    t(29) = 405
    t(30) = 406
    t(31) = 409
    t(32) = 410
    t(33) = 411
    t(34) = 414
    t(35) = 415
    t(36) = 417
    t(37) = 422
    t(38) = 424
    t(39) = 425
    t(40) = 429
    t(41) = 430
    t(42) = 432
    t(43) = 433
    t(44) = 436
    t(45) = 439
    t(46) = 441
    t(47) = 442
    t(48) = 445
    t(49) = 446
    t(50) = 447
    t(51) = 448
    t(52) = 453
    t(53) = 454
    t(54) = 455
    t(55) = 456
    t(56) = 457
    t(57) = 458
    t(58) = 459
    t(59) = 460
    t(60) = 461
    t(61) = 462
    t(62) = 477
    t(63) = 478
    t(64) = 479
    t(65) = 496
    t(66) = 498
    t(67) = 499
    t(68) = 500
    t(69) = 501
    t(70) = 502
    t(71) = 505
    t(72) = 544
    t(73) = 547
    t(74) = 564
    t(75) = 572
    t(76) = 573
    t(77) = 575
    t(78) = 577
    t(79) = 578
    t(80) = 579
    t(81) = 583
    t(82) = 592
    t(83) = 593
    t(84) = 594
    t(85) = 595
    t(86) = 596
    t(87) = 597
    t(88) = 598
    t(89) = 600
    t(90) = 601
    t(91) = 602
    t(92) = 603
    t(93) = 604
    t(94) = 605
    t(95) = 608
    t(96) = 609
    t(97) = 610
    t(98) = 611
    t(99) = 612
    t(100) = 613
    t(101) = 614
    t(102) = 615
    t(103) = 616
    t(104) = 617
    t(105) = 618
    t(106) = 619
    t(107) = 620
    t(108) = 621
    t(109) = 623
    t(110) = 624
    t(111) = 625
    t(112) = 626
    t(113) = 627
    t(114) = 629
    t(115) = 630
    t(116) = 637
    t(117) = 638
    t(118) = 640
    t(119) = 641
    t(120) = 642
    t(121) = 643
    t(122) = 644
    t(123) = 647
    t(124) = 648
    t(125) = 649
    t(126) = 650
    t(127) = 652
    t(128) = 653
    t(129) = 658
    t(130) = 659
    t(131) = 669
    t(132) = 670
    t(133) = 671
    t(134) = 837
    t(135) = 838
    t(136) = 881
    t(137) = 884
    t(138) = 887
    t(139) = 888
    t(140) = 891
    t(141) = 894
    t(142) = 940
    t(143) = 941
    t(144) = 944
    t(145) = 945
    t(146) = 962
    t(147) = 963
    t(148) = 972
    t(149) = 973
    t(150) = 975
    t(151) = 976
    t(152) = 977
    t(153) = 978
    t(154) = 981
    t(155) = 982
    t(156) = 983
    t(157) = 984
    t(158) = 985
    t(159) = 1008
    t(160) = 1009
    t(161) = 1010
    t(162) = 1011
    t(163) = 1012
    t(164) = 1013
    t(165) = 1014
    t(166) = 1016
    t(167) = 1017
    t(168) = 1019
    t(169) = 1020
    t(170) = 1072
    t(171) = 1104
    t(172) = 1120
    t(173) = 1121
    t(174) = 1154
    t(175) = 1163
    t(176) = 1216
    t(177) = 1218
    t(178) = 1231
    t(179) = 1232
    t(180) = 1233
    t(181) = 1328
    t(182) = 1377
    t(183) = 1415
    t(184) = 4304
    t(185) = 4347
    t(186) = 4349
    t(187) = 4352
    t(188) = 5112
    t(189) = 5118
    t(190) = 7296
    t(191) = 7297
    t(192) = 7298
    t(193) = 7299
    t(194) = 7301
    t(195) = 7302
    t(196) = 7303
    t(197) = 7304
    t(198) = 7305
    t(199) = 7545
    t(200) = 7546
    t(201) = 7549
    t(202) = 7550
    t(203) = 7566
    t(204) = 7567
    t(205) = 7681
    t(206) = 7830
    t(207) = 7835
    t(208) = 7836
    t(209) = 7841
    t(210) = 7936
    t(211) = 7944
    t(212) = 7952
    t(213) = 7958
    t(214) = 7968
    t(215) = 7976
    t(216) = 7984
    t(217) = 7992
    t(218) = 8000
    t(219) = 8006
    t(220) = 8017
    t(221) = 8024
    t(222) = 8032
    t(223) = 8040
    t(224) = 8048
    t(225) = 8050
    t(226) = 8054
    t(227) = 8056
    t(228) = 8058
    t(229) = 8060
    t(230) = 8062
    t(231) = 8112
    t(232) = 8114
    t(233) = 8126
    t(234) = 8127
    t(235) = 8144
    t(236) = 8146
    t(237) = 8160
    t(238) = 8162
    t(239) = 8165
    t(240) = 8166
    t(241) = 8526
    t(242) = 8527
    t(243) = 8560
    t(244) = 8576
    t(245) = 8580
    t(246) = 8581
    t(247) = 9424
    t(248) = 9450
    t(249) = 11312
    t(250) = 11359
    t(251) = 11361
    t(252) = 11362
    t(253) = 11365
    t(254) = 11366
    t(255) = 11367
    t(256) = 11368
    t(257) = 11373
    t(258) = 11379
    t(259) = 11380
    t(260) = 11382
    t(261) = 11383
    t(262) = 11393
    t(263) = 11492
    t(264) = 11500
    t(265) = 11503
    t(266) = 11507
    t(267) = 11508
    t(268) = 11520
    t(269) = 11558
    t(270) = 11559
    t(271) = 11560
    t(272) = 11565
    t(273) = 11566
    t(274) = 42561
    t(275) = 42606
    t(276) = 42625
    t(277) = 42652
    t(278) = 42787
    t(279) = 42800
    t(280) = 42803
    t(281) = 42864
    t(282) = 42874
    t(283) = 42877
    t(284) = 42879
    t(285) = 42888
    t(286) = 42892
    t(287) = 42893
    t(288) = 42897
    t(289) = 42900
    t(290) = 42901
    t(291) = 42903
    t(292) = 42922
    t(293) = 42933
    t(294) = 42944
    t(295) = 42947
    t(296) = 42948
    t(297) = 43859
    t(298) = 43860
    t(299) = 43888
    t(300) = 43968
    t(301) = 65345
    t(302) = 65371
End Sub

