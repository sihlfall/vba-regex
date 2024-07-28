Attribute VB_Name = "RegexUnicodeSupport"
Option Explicit

Public UnicodeInitialized As Boolean ' auto-initialized to False
Public RangeTablesInitialized As Boolean ' auto-initialized to False

Public Const AST_TABLE_START As Long = 0

Public Const RANGE_TABLE_DIGIT_START As Long = AST_TABLE_START + RegexAst.AST_TABLE_LENGTH
Public Const RANGE_TABLE_DIGIT_LENGTH As Long = 2

Public Const RANGE_TABLE_WHITE_START As Long = RANGE_TABLE_DIGIT_START + RANGE_TABLE_DIGIT_LENGTH
Public Const RANGE_TABLE_WHITE_LENGTH As Long = 22

Public Const RANGE_TABLE_WORDCHAR_START As Long = RANGE_TABLE_WHITE_START + RANGE_TABLE_WHITE_LENGTH
Public Const RANGE_TABLE_WORDCHAR_LENGTH As Long = 8

Public Const RANGE_TABLE_NOTDIGIT_START As Long = RANGE_TABLE_WORDCHAR_START + RANGE_TABLE_WORDCHAR_LENGTH
Public Const RANGE_TABLE_NOTDIGIT_LENGTH As Long = 4

Public Const RANGE_TABLE_NOTWHITE_START As Long = RANGE_TABLE_NOTDIGIT_START + RANGE_TABLE_NOTDIGIT_LENGTH
Public Const RANGE_TABLE_NOTWHITE_LENGTH As Long = 24

Public Const RANGE_TABLE_NOTWORDCHAR_START As Long = RANGE_TABLE_NOTWHITE_START + RANGE_TABLE_NOTWHITE_LENGTH
Public Const RANGE_TABLE_NOTWORDCHAR_LENGTH As Long = 10

Public Const UNICODE_CANON_LOOKUP_TABLE_START As Long = RANGE_TABLE_NOTWORDCHAR_START + RANGE_TABLE_NOTWORDCHAR_LENGTH
Public Const UNICODE_CANON_LOOKUP_TABLE_LENGTH As Long = 65536

Public Const UNICODE_CANON_RUNS_TABLE_START As Long = UNICODE_CANON_LOOKUP_TABLE_START + UNICODE_CANON_LOOKUP_TABLE_LENGTH
Public Const UNICODE_CANON_RUNS_TABLE_LENGTH As Long = 303

Public Const STATIC_DATA_LENGTH As Long = UNICODE_CANON_RUNS_TABLE_START + UNICODE_CANON_RUNS_TABLE_LENGTH

Public StaticData(0 To STATIC_DATA_LENGTH - 1) As Long

Private Const MIN_LONG As Long = &H80000000
Private Const MAX_LONG As Long = &H7FFFFFFF

Public Sub UnicodeInitialize()
    InitializeUnicodeCanonLookupTable StaticData
    InitializeUnicodeCanonRunsTable StaticData
    UnicodeInitialized = True
End Sub

Public Sub RangeTablesInitialize()
    InitializeRangeTables StaticData
    RangeTablesInitialized = True
End Sub

Private Sub InitializeRangeTables(ByRef t() As Long)
    t(RANGE_TABLE_DIGIT_START + 0) = &H30&
    t(RANGE_TABLE_DIGIT_START + 1) = &H39&
    
    '---------------------------------
    
    t(RANGE_TABLE_WHITE_START + 0) = &H9&
    t(RANGE_TABLE_WHITE_START + 1) = &HD&
    t(RANGE_TABLE_WHITE_START + 2) = &H20&
    t(RANGE_TABLE_WHITE_START + 3) = &H20&
    t(RANGE_TABLE_WHITE_START + 4) = &HA0&
    t(RANGE_TABLE_WHITE_START + 5) = &HA0&
    t(RANGE_TABLE_WHITE_START + 6) = &H1680&
    t(RANGE_TABLE_WHITE_START + 7) = &H1680&
    t(RANGE_TABLE_WHITE_START + 8) = &H180E&
    t(RANGE_TABLE_WHITE_START + 9) = &H180E&
    t(RANGE_TABLE_WHITE_START + 10) = &H2000&
    t(RANGE_TABLE_WHITE_START + 11) = &H200A
    t(RANGE_TABLE_WHITE_START + 12) = &H2028&
    t(RANGE_TABLE_WHITE_START + 13) = &H2029&
    t(RANGE_TABLE_WHITE_START + 14) = &H202F
    t(RANGE_TABLE_WHITE_START + 15) = &H202F
    t(RANGE_TABLE_WHITE_START + 16) = &H205F&
    t(RANGE_TABLE_WHITE_START + 17) = &H205F&
    t(RANGE_TABLE_WHITE_START + 18) = &H3000&
    t(RANGE_TABLE_WHITE_START + 19) = &H3000&
    t(RANGE_TABLE_WHITE_START + 20) = &HFEFF&
    t(RANGE_TABLE_WHITE_START + 21) = &HFEFF&
    
    '---------------------------------
    
    t(RANGE_TABLE_WORDCHAR_START + 0) = &H30&
    t(RANGE_TABLE_WORDCHAR_START + 1) = &H39&
    t(RANGE_TABLE_WORDCHAR_START + 2) = &H41&
    t(RANGE_TABLE_WORDCHAR_START + 3) = &H5A&
    t(RANGE_TABLE_WORDCHAR_START + 4) = &H5F&
    t(RANGE_TABLE_WORDCHAR_START + 5) = &H5F&
    t(RANGE_TABLE_WORDCHAR_START + 6) = &H61&
    t(RANGE_TABLE_WORDCHAR_START + 7) = &H7A&
    
    '---------------------------------
    
    t(RANGE_TABLE_NOTDIGIT_START + 0) = MIN_LONG
    t(RANGE_TABLE_NOTDIGIT_START + 1) = &H2F&
    t(RANGE_TABLE_NOTDIGIT_START + 2) = &H3A&
    t(RANGE_TABLE_NOTDIGIT_START + 3) = MAX_LONG
    
    '---------------------------------
    
    t(RANGE_TABLE_NOTWHITE_START + 0) = MIN_LONG
    t(RANGE_TABLE_NOTWHITE_START + 1) = &H8&
    t(RANGE_TABLE_NOTWHITE_START + 2) = &HE&
    t(RANGE_TABLE_NOTWHITE_START + 3) = &H1F&
    t(RANGE_TABLE_NOTWHITE_START + 4) = &H21&
    t(RANGE_TABLE_NOTWHITE_START + 5) = &H9F&
    t(RANGE_TABLE_NOTWHITE_START + 6) = &HA1&
    t(RANGE_TABLE_NOTWHITE_START + 7) = &H167F&
    t(RANGE_TABLE_NOTWHITE_START + 8) = &H1681&
    t(RANGE_TABLE_NOTWHITE_START + 9) = &H180D&
    t(RANGE_TABLE_NOTWHITE_START + 10) = &H180F&
    t(RANGE_TABLE_NOTWHITE_START + 11) = &H1FFF&
    t(RANGE_TABLE_NOTWHITE_START + 12) = &H200B&
    t(RANGE_TABLE_NOTWHITE_START + 13) = &H2027&
    t(RANGE_TABLE_NOTWHITE_START + 14) = &H202A&
    t(RANGE_TABLE_NOTWHITE_START + 15) = &H202E&
    t(RANGE_TABLE_NOTWHITE_START + 16) = &H2030&
    t(RANGE_TABLE_NOTWHITE_START + 17) = &H205E&
    t(RANGE_TABLE_NOTWHITE_START + 18) = &H2060&
    t(RANGE_TABLE_NOTWHITE_START + 19) = &H2FFF&
    t(RANGE_TABLE_NOTWHITE_START + 20) = &H3001&
    t(RANGE_TABLE_NOTWHITE_START + 21) = &HFEFE&
    t(RANGE_TABLE_NOTWHITE_START + 22) = &HFF00&
    t(RANGE_TABLE_NOTWHITE_START + 23) = MAX_LONG
    
    '---------------------------------
    
    t(RANGE_TABLE_NOTWORDCHAR_START + 0) = MIN_LONG
    t(RANGE_TABLE_NOTWORDCHAR_START + 1) = &H2F&
    t(RANGE_TABLE_NOTWORDCHAR_START + 2) = &H3A&
    t(RANGE_TABLE_NOTWORDCHAR_START + 3) = &H40&
    t(RANGE_TABLE_NOTWORDCHAR_START + 4) = &H5B&
    t(RANGE_TABLE_NOTWORDCHAR_START + 5) = &H5E&
    t(RANGE_TABLE_NOTWORDCHAR_START + 6) = &H60&
    t(RANGE_TABLE_NOTWORDCHAR_START + 7) = &H60&
    t(RANGE_TABLE_NOTWORDCHAR_START + 8) = &H7B&
    t(RANGE_TABLE_NOTWORDCHAR_START + 9) = MAX_LONG
    
    '---------------------------------
End Sub

Public Function ReCanonicalizeChar(ByVal codepoint As Long) As Long
    ' ! This function must not alter codepoint if codepoint is negative, as ENDOFINPUT is represented by -1.
    If codepoint And &HFFFF0000 Then ' Codepoint not in [0;&HFFFF&]
        ReCanonicalizeChar = codepoint
    Else
        ReCanonicalizeChar = (codepoint + StaticData(UNICODE_CANON_LOOKUP_TABLE_START + codepoint)) And &HFFFF&
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

Private Sub InitializeUnicodeCanonLookupTable(ByRef t() As Long)
    ' Array of integers would be sufficient
    Const b As Long = UNICODE_CANON_LOOKUP_TABLE_START
    
    Dim i As Long
    
    For i = b + 97 To b + 122: t(i) = -32: Next i
    t(b + 181) = 743
    For i = b + 224 To b + 246: t(i) = -32: Next i
    For i = b + 248 To b + 254: t(i) = -32: Next i
    t(b + 255) = 121
    For i = b + 257 To b + 303 Step 2: t(i) = -1: Next i
    t(b + 307) = -1
    t(b + 309) = -1
    t(b + 311) = -1
    For i = b + 314 To b + 328 Step 2: t(i) = -1: Next i
    For i = b + 331 To b + 375 Step 2: t(i) = -1: Next i
    t(b + 378) = -1
    t(b + 380) = -1
    t(b + 382) = -1
    t(b + 384) = 195
    t(b + 387) = -1
    t(b + 389) = -1
    t(b + 392) = -1
    t(b + 396) = -1
    t(b + 402) = -1
    t(b + 405) = 97
    t(b + 409) = -1
    t(b + 410) = 163
    t(b + 414) = 130
    t(b + 417) = -1
    t(b + 419) = -1
    t(b + 421) = -1
    t(b + 424) = -1
    t(b + 429) = -1
    t(b + 432) = -1
    t(b + 436) = -1
    t(b + 438) = -1
    t(b + 441) = -1
    t(b + 445) = -1
    t(b + 447) = 56
    t(b + 453) = -1
    t(b + 454) = -2
    t(b + 456) = -1
    t(b + 457) = -2
    t(b + 459) = -1
    t(b + 460) = -2
    For i = b + 462 To b + 476 Step 2: t(i) = -1: Next i
    t(b + 477) = -79
    For i = b + 479 To b + 495 Step 2: t(i) = -1: Next i
    t(b + 498) = -1
    t(b + 499) = -2
    t(b + 501) = -1
    For i = b + 505 To b + 543 Step 2: t(i) = -1: Next i
    For i = b + 547 To b + 563 Step 2: t(i) = -1: Next i
    t(b + 572) = -1
    t(b + 575) = 10815
    t(b + 576) = 10815
    t(b + 578) = -1
    For i = b + 583 To b + 591 Step 2: t(i) = -1: Next i
    t(b + 592) = 10783
    t(b + 593) = 10780
    t(b + 594) = 10782
    t(b + 595) = -210
    t(b + 596) = -206
    t(b + 598) = -205
    t(b + 599) = -205
    t(b + 601) = -202
    t(b + 603) = -203
    t(b + 604) = -23217
    t(b + 608) = -205
    t(b + 609) = -23221
    t(b + 611) = -207
    t(b + 613) = -23256
    t(b + 614) = -23228
    t(b + 616) = -209
    t(b + 617) = -211
    t(b + 618) = -23228
    t(b + 619) = 10743
    t(b + 620) = -23231
    t(b + 623) = -211
    t(b + 625) = 10749
    t(b + 626) = -213
    t(b + 629) = -214
    t(b + 637) = 10727
    t(b + 640) = -218
    t(b + 642) = -23229
    t(b + 643) = -218
    t(b + 647) = -23254
    t(b + 648) = -218
    t(b + 649) = -69
    t(b + 650) = -217
    t(b + 651) = -217
    t(b + 652) = -71
    t(b + 658) = -219
    t(b + 669) = -23275
    t(b + 670) = -23278
    t(b + 837) = 84
    t(b + 881) = -1
    t(b + 883) = -1
    t(b + 887) = -1
    t(b + 891) = 130
    t(b + 892) = 130
    t(b + 893) = 130
    t(b + 940) = -38
    t(b + 941) = -37
    t(b + 942) = -37
    t(b + 943) = -37
    For i = b + 945 To b + 961: t(i) = -32: Next i
    t(b + 962) = -31
    For i = b + 963 To b + 971: t(i) = -32: Next i
    t(b + 972) = -64
    t(b + 973) = -63
    t(b + 974) = -63
    t(b + 976) = -62
    t(b + 977) = -57
    t(b + 981) = -47
    t(b + 982) = -54
    t(b + 983) = -8
    For i = b + 985 To b + 1007 Step 2: t(i) = -1: Next i
    t(b + 1008) = -86
    t(b + 1009) = -80
    t(b + 1010) = 7
    t(b + 1011) = -116
    t(b + 1013) = -96
    t(b + 1016) = -1
    t(b + 1019) = -1
    For i = b + 1072 To b + 1103: t(i) = -32: Next i
    For i = b + 1104 To b + 1119: t(i) = -80: Next i
    For i = b + 1121 To b + 1153 Step 2: t(i) = -1: Next i
    For i = b + 1163 To b + 1215 Step 2: t(i) = -1: Next i
    For i = b + 1218 To b + 1230 Step 2: t(i) = -1: Next i
    t(b + 1231) = -15
    For i = b + 1233 To b + 1327 Step 2: t(i) = -1: Next i
    For i = b + 1377 To b + 1414: t(i) = -48: Next i
    For i = b + 4304 To b + 4346: t(i) = 3008: Next i
    t(b + 4349) = 3008
    t(b + 4350) = 3008
    t(b + 4351) = 3008
    For i = b + 5112 To b + 5117: t(i) = -8: Next i
    t(b + 7296) = -6254
    t(b + 7297) = -6253
    t(b + 7298) = -6244
    t(b + 7299) = -6242
    t(b + 7300) = -6242
    t(b + 7301) = -6243
    t(b + 7302) = -6236
    t(b + 7303) = -6181
    t(b + 7304) = -30270
    t(b + 7545) = -30204
    t(b + 7549) = 3814
    t(b + 7566) = -30152
    For i = b + 7681 To b + 7829 Step 2: t(i) = -1: Next i
    t(b + 7835) = -59
    For i = b + 7841 To b + 7935 Step 2: t(i) = -1: Next i
    For i = b + 7936 To b + 7943: t(i) = 8: Next i
    For i = b + 7952 To b + 7957: t(i) = 8: Next i
    For i = b + 7968 To b + 7975: t(i) = 8: Next i
    For i = b + 7984 To b + 7991: t(i) = 8: Next i
    For i = b + 8000 To b + 8005: t(i) = 8: Next i
    t(b + 8017) = 8
    t(b + 8019) = 8
    t(b + 8021) = 8
    t(b + 8023) = 8
    For i = b + 8032 To b + 8039: t(i) = 8: Next i
    t(b + 8048) = 74
    t(b + 8049) = 74
    t(b + 8050) = 86
    t(b + 8051) = 86
    t(b + 8052) = 86
    t(b + 8053) = 86
    t(b + 8054) = 100
    t(b + 8055) = 100
    t(b + 8056) = 128
    t(b + 8057) = 128
    t(b + 8058) = 112
    t(b + 8059) = 112
    t(b + 8060) = 126
    t(b + 8061) = 126
    t(b + 8112) = 8
    t(b + 8113) = 8
    t(b + 8126) = -7205
    t(b + 8144) = 8
    t(b + 8145) = 8
    t(b + 8160) = 8
    t(b + 8161) = 8
    t(b + 8165) = 7
    t(b + 8526) = -28
    For i = b + 8560 To b + 8575: t(i) = -16: Next i
    t(b + 8580) = -1
    For i = b + 9424 To b + 9449: t(i) = -26: Next i
    For i = b + 11312 To b + 11358: t(i) = -48: Next i
    t(b + 11361) = -1
    t(b + 11365) = -10795
    t(b + 11366) = -10792
    t(b + 11368) = -1
    t(b + 11370) = -1
    t(b + 11372) = -1
    t(b + 11379) = -1
    t(b + 11382) = -1
    For i = b + 11393 To b + 11491 Step 2: t(i) = -1: Next i
    t(b + 11500) = -1
    t(b + 11502) = -1
    t(b + 11507) = -1
    For i = b + 11520 To b + 11557: t(i) = -7264: Next i
    t(b + 11559) = -7264
    t(b + 11565) = -7264
    For i = b + 42561 To b + 42605 Step 2: t(i) = -1: Next i
    For i = b + 42625 To b + 42651 Step 2: t(i) = -1: Next i
    For i = b + 42787 To b + 42799 Step 2: t(i) = -1: Next i
    For i = b + 42803 To b + 42863 Step 2: t(i) = -1: Next i
    t(b + 42874) = -1
    t(b + 42876) = -1
    For i = b + 42879 To b + 42887 Step 2: t(i) = -1: Next i
    t(b + 42892) = -1
    t(b + 42897) = -1
    t(b + 42899) = -1
    t(b + 42900) = 48
    For i = b + 42903 To b + 42921 Step 2: t(i) = -1: Next i
    For i = b + 42933 To b + 42943 Step 2: t(i) = -1: Next i
    t(b + 42947) = -1
    t(b + 43859) = -928
    For i = b + 43888 To b + 43967: t(i) = 26672: Next i
    For i = b + 65345 To b + 65370: t(i) = -32: Next i
End Sub

Private Sub InitializeUnicodeCanonRunsTable(ByRef t() As Long)
    Const b As Long = UNICODE_CANON_RUNS_TABLE_START
    t(b + 0) = 97
    t(b + 1) = 123
    t(b + 2) = 181
    t(b + 3) = 182
    t(b + 4) = 224
    t(b + 5) = 247
    t(b + 6) = 248
    t(b + 7) = 255
    t(b + 8) = 256
    t(b + 9) = 257
    t(b + 10) = 304
    t(b + 11) = 307
    t(b + 12) = 312
    t(b + 13) = 314
    t(b + 14) = 329
    t(b + 15) = 331
    t(b + 16) = 376
    t(b + 17) = 378
    t(b + 18) = 383
    t(b + 19) = 384
    t(b + 20) = 385
    t(b + 21) = 387
    t(b + 22) = 390
    t(b + 23) = 392
    t(b + 24) = 393
    t(b + 25) = 396
    t(b + 26) = 397
    t(b + 27) = 402
    t(b + 28) = 403
    t(b + 29) = 405
    t(b + 30) = 406
    t(b + 31) = 409
    t(b + 32) = 410
    t(b + 33) = 411
    t(b + 34) = 414
    t(b + 35) = 415
    t(b + 36) = 417
    t(b + 37) = 422
    t(b + 38) = 424
    t(b + 39) = 425
    t(b + 40) = 429
    t(b + 41) = 430
    t(b + 42) = 432
    t(b + 43) = 433
    t(b + 44) = 436
    t(b + 45) = 439
    t(b + 46) = 441
    t(b + 47) = 442
    t(b + 48) = 445
    t(b + 49) = 446
    t(b + 50) = 447
    t(b + 51) = 448
    t(b + 52) = 453
    t(b + 53) = 454
    t(b + 54) = 455
    t(b + 55) = 456
    t(b + 56) = 457
    t(b + 57) = 458
    t(b + 58) = 459
    t(b + 59) = 460
    t(b + 60) = 461
    t(b + 61) = 462
    t(b + 62) = 477
    t(b + 63) = 478
    t(b + 64) = 479
    t(b + 65) = 496
    t(b + 66) = 498
    t(b + 67) = 499
    t(b + 68) = 500
    t(b + 69) = 501
    t(b + 70) = 502
    t(b + 71) = 505
    t(b + 72) = 544
    t(b + 73) = 547
    t(b + 74) = 564
    t(b + 75) = 572
    t(b + 76) = 573
    t(b + 77) = 575
    t(b + 78) = 577
    t(b + 79) = 578
    t(b + 80) = 579
    t(b + 81) = 583
    t(b + 82) = 592
    t(b + 83) = 593
    t(b + 84) = 594
    t(b + 85) = 595
    t(b + 86) = 596
    t(b + 87) = 597
    t(b + 88) = 598
    t(b + 89) = 600
    t(b + 90) = 601
    t(b + 91) = 602
    t(b + 92) = 603
    t(b + 93) = 604
    t(b + 94) = 605
    t(b + 95) = 608
    t(b + 96) = 609
    t(b + 97) = 610
    t(b + 98) = 611
    t(b + 99) = 612
    t(b + 100) = 613
    t(b + 101) = 614
    t(b + 102) = 615
    t(b + 103) = 616
    t(b + 104) = 617
    t(b + 105) = 618
    t(b + 106) = 619
    t(b + 107) = 620
    t(b + 108) = 621
    t(b + 109) = 623
    t(b + 110) = 624
    t(b + 111) = 625
    t(b + 112) = 626
    t(b + 113) = 627
    t(b + 114) = 629
    t(b + 115) = 630
    t(b + 116) = 637
    t(b + 117) = 638
    t(b + 118) = 640
    t(b + 119) = 641
    t(b + 120) = 642
    t(b + 121) = 643
    t(b + 122) = 644
    t(b + 123) = 647
    t(b + 124) = 648
    t(b + 125) = 649
    t(b + 126) = 650
    t(b + 127) = 652
    t(b + 128) = 653
    t(b + 129) = 658
    t(b + 130) = 659
    t(b + 131) = 669
    t(b + 132) = 670
    t(b + 133) = 671
    t(b + 134) = 837
    t(b + 135) = 838
    t(b + 136) = 881
    t(b + 137) = 884
    t(b + 138) = 887
    t(b + 139) = 888
    t(b + 140) = 891
    t(b + 141) = 894
    t(b + 142) = 940
    t(b + 143) = 941
    t(b + 144) = 944
    t(b + 145) = 945
    t(b + 146) = 962
    t(b + 147) = 963
    t(b + 148) = 972
    t(b + 149) = 973
    t(b + 150) = 975
    t(b + 151) = 976
    t(b + 152) = 977
    t(b + 153) = 978
    t(b + 154) = 981
    t(b + 155) = 982
    t(b + 156) = 983
    t(b + 157) = 984
    t(b + 158) = 985
    t(b + 159) = 1008
    t(b + 160) = 1009
    t(b + 161) = 1010
    t(b + 162) = 1011
    t(b + 163) = 1012
    t(b + 164) = 1013
    t(b + 165) = 1014
    t(b + 166) = 1016
    t(b + 167) = 1017
    t(b + 168) = 1019
    t(b + 169) = 1020
    t(b + 170) = 1072
    t(b + 171) = 1104
    t(b + 172) = 1120
    t(b + 173) = 1121
    t(b + 174) = 1154
    t(b + 175) = 1163
    t(b + 176) = 1216
    t(b + 177) = 1218
    t(b + 178) = 1231
    t(b + 179) = 1232
    t(b + 180) = 1233
    t(b + 181) = 1328
    t(b + 182) = 1377
    t(b + 183) = 1415
    t(b + 184) = 4304
    t(b + 185) = 4347
    t(b + 186) = 4349
    t(b + 187) = 4352
    t(b + 188) = 5112
    t(b + 189) = 5118
    t(b + 190) = 7296
    t(b + 191) = 7297
    t(b + 192) = 7298
    t(b + 193) = 7299
    t(b + 194) = 7301
    t(b + 195) = 7302
    t(b + 196) = 7303
    t(b + 197) = 7304
    t(b + 198) = 7305
    t(b + 199) = 7545
    t(b + 200) = 7546
    t(b + 201) = 7549
    t(b + 202) = 7550
    t(b + 203) = 7566
    t(b + 204) = 7567
    t(b + 205) = 7681
    t(b + 206) = 7830
    t(b + 207) = 7835
    t(b + 208) = 7836
    t(b + 209) = 7841
    t(b + 210) = 7936
    t(b + 211) = 7944
    t(b + 212) = 7952
    t(b + 213) = 7958
    t(b + 214) = 7968
    t(b + 215) = 7976
    t(b + 216) = 7984
    t(b + 217) = 7992
    t(b + 218) = 8000
    t(b + 219) = 8006
    t(b + 220) = 8017
    t(b + 221) = 8024
    t(b + 222) = 8032
    t(b + 223) = 8040
    t(b + 224) = 8048
    t(b + 225) = 8050
    t(b + 226) = 8054
    t(b + 227) = 8056
    t(b + 228) = 8058
    t(b + 229) = 8060
    t(b + 230) = 8062
    t(b + 231) = 8112
    t(b + 232) = 8114
    t(b + 233) = 8126
    t(b + 234) = 8127
    t(b + 235) = 8144
    t(b + 236) = 8146
    t(b + 237) = 8160
    t(b + 238) = 8162
    t(b + 239) = 8165
    t(b + 240) = 8166
    t(b + 241) = 8526
    t(b + 242) = 8527
    t(b + 243) = 8560
    t(b + 244) = 8576
    t(b + 245) = 8580
    t(b + 246) = 8581
    t(b + 247) = 9424
    t(b + 248) = 9450
    t(b + 249) = 11312
    t(b + 250) = 11359
    t(b + 251) = 11361
    t(b + 252) = 11362
    t(b + 253) = 11365
    t(b + 254) = 11366
    t(b + 255) = 11367
    t(b + 256) = 11368
    t(b + 257) = 11373
    t(b + 258) = 11379
    t(b + 259) = 11380
    t(b + 260) = 11382
    t(b + 261) = 11383
    t(b + 262) = 11393
    t(b + 263) = 11492
    t(b + 264) = 11500
    t(b + 265) = 11503
    t(b + 266) = 11507
    t(b + 267) = 11508
    t(b + 268) = 11520
    t(b + 269) = 11558
    t(b + 270) = 11559
    t(b + 271) = 11560
    t(b + 272) = 11565
    t(b + 273) = 11566
    t(b + 274) = 42561
    t(b + 275) = 42606
    t(b + 276) = 42625
    t(b + 277) = 42652
    t(b + 278) = 42787
    t(b + 279) = 42800
    t(b + 280) = 42803
    t(b + 281) = 42864
    t(b + 282) = 42874
    t(b + 283) = 42877
    t(b + 284) = 42879
    t(b + 285) = 42888
    t(b + 286) = 42892
    t(b + 287) = 42893
    t(b + 288) = 42897
    t(b + 289) = 42900
    t(b + 290) = 42901
    t(b + 291) = 42903
    t(b + 292) = 42922
    t(b + 293) = 42933
    t(b + 294) = 42944
    t(b + 295) = 42947
    t(b + 296) = 42948
    t(b + 297) = 43859
    t(b + 298) = 43860
    t(b + 299) = 43888
    t(b + 300) = 43968
    t(b + 301) = 65345
    t(b + 302) = 65371
End Sub

