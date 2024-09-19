'シート名
    Global Const inputSheet = "INPUT_LIST"
    Global Const inputR117NoSheet = "TS_R117_NO"
    Global Const inputAppovalNoSheet = "TS_EC_APPROVAL_NO"
    Global Const inputR117MainSheet = "TS_R117_MAIN"
    Global Const inputR117SiezDataSheet = "TS_R117_SIZEDATA"
    Global Const outputR117ExtDataSheet = "OUTPUT_TS_R117_EXTDATA"
    Global Const outputR117MainSheet = "OUTPUT_TS_R117_MAIN"
    Global Const outputR117SizeDataSheet = "OUTPUT_TS_R117_SIZEDATA"
    
'INPUT_登録リスト
    '項目の開始位置
    Global Const inputSheetItemCol = 2
    Global Const inputSheetItmeRow = 3
    '認可グループ
    Global Const inputSheetNinkaGrpCol = 3
    '認可No
    Global Const inputSheetNinkaNoCol = 4
    'motoExt
    Global Const inputSheetMotoExtCol = 5
    'Ext
    Global Const inputSheetExtCol = 6
    'SECTION
    Global Const inputSheetSectionCol = 8
    '速度構造
    Global Const inputSheetSpeedCol = 9
    '??径
    Global Const inputSheetRimDiaCol = 10
    '偏平率
    Global Const inputSheetHenpeiCol = 11
    'LI
    Global Const inputSheetLICol = 12
    'SS
    Global Const inputSheetSSCol = 13
    '補強構造
    Global Const inputSheetCMCol = 14
    '生産立上り日
    Global Const inputSheetStartYMDCol = 15
    '国
    Global Const inputSheetNationCol = 16
    '極区分
    Global Const inputSheetPetBurKbnCol = 17
    'クラス
    Global Const inputSheetClassCol = 18
    'グループ1
    Global Const inputSheetGrp1Col = 19
    'パターンカテゴリー
    Global Const inputSheetPatterntypeCol = 20
    'トレードネーム
    Global Const inputSheetBrdCol = 21
    '正式商品名1
    Global Const inputSheetPatternLCol = 22
    'パターン略称1
    Global Const inputSheetPatternS1Col = 23
    '発信元/設計部署
    Global Const inputSheetSendFrCol = 24
    '発信担当
    Global Const inputSheetSendTantouCol = 25
    ' R117タイプ
    Global Const inputSheetR117TypeCol = 26

'INPUT_TS_R117_NO
    '項目の開始位置
    Global Const inputR117NoSheetItemCol = 2
    Global Const inputR117NoSheetItemRow = 3

    Global Const inputR117NoSheetColR117_RECID = 2
    Global Const inputR117NoSheetColR117_MAX_NO = 3

'INPUT_TS_EC_APPROVAL_NO
    '項目の開始位置
    Global Const inputAppovalNoSheetItemCol = 2
    Global Const inputAppovalNoSheetItemRow = 3
    
    Global Const inputAppovalNoSheetColEC_REG_KBN = 2
    Global Const inputAppovalNoSheetColEC_FROM = 3
    Global Const inputAppovalNoSheetColEC_TO = 4
    Global Const inputAppovalNoSheetColEC_USEFLG = 5
    Global Const inputAppovalNoSheetColEC_LENGTH = 6

    Global Const inputAppovalNoSheetColNINKA_NO = 8
    Global Const inputAppovalNoSheetColREG_KBN = 9
    Global Const inputAppovalNoSheetColUSEFLG = 10


'INPUT_TS_R117_MAIN
    '項目の開始位置
    Global Const inputR117MainSheetItemCol = 2
    Global Const inputR117MainSheetItemRow = 3

    Global Const inputR117MainSheetColR117_RECNO = 2
    Global Const inputR117MainSheetColR117_NINKA_NO_ALL = 3
    Global Const inputR117MainSheetColR117_NINKA_NO = 4
    Global Const inputR117MainSheetColR117_NATION = 5
    Global Const inputR117MainSheetColR117_EXTENTION = 6
    Global Const inputR117MainSheetColR117_PET_BUR_KBN = 7
    Global Const inputR117MainSheetColR117_CLASS = 8
    Global Const inputR117MainSheetColR117_GRP1 = 9
    Global Const inputR117MainSheetColR117_GRP2 = 10
    Global Const inputR117MainSheetColR117_GRP3 = 11
    Global Const inputR117MainSheetColR117_PURPOSE_KBN = 12
    Global Const inputR117MainSheetColR117_BRD = 13
    Global Const inputR117MainSheetColR117_PATTERN_L1 = 14
    Global Const inputR117MainSheetColR117_PATTERN_L2 = 15
    Global Const inputR117MainSheetColR117_PATTERN_L3 = 16
    Global Const inputR117MainSheetColR117_PATTERN_L4 = 17
    Global Const inputR117MainSheetColR117_PATTERN_L5 = 18
    Global Const inputR117MainSheetColR117_PATTERN_L6 = 19
    Global Const inputR117MainSheetColR117_PATTERN_S1 = 20
    Global Const inputR117MainSheetColR117_PATTERN_S2 = 21
    Global Const inputR117MainSheetColR117_PATTERN_S3 = 22
    Global Const inputR117MainSheetColR117_PATTERN_S4 = 23
    Global Const inputR117MainSheetColR117_PATTERN_S5 = 24
    Global Const inputR117MainSheetColR117_PATTERN_S6 = 25
    Global Const inputR117MainSheetColR117_PATTERN_S7 = 26
    Global Const inputR117MainSheetColR117_PATTERN_S8 = 27
    Global Const inputR117MainSheetColR117_PATTERN_S9 = 28
    Global Const inputR117MainSheetColR117_PATTERN_S10 = 29
    Global Const inputR117MainSheetColR117_PATTERN_S11 = 30
    Global Const inputR117MainSheetColR117_PATTERN_S12 = 31
    Global Const inputR117MainSheetColR117_PATTERN_S13 = 32
    Global Const inputR117MainSheetColR117_PATTERN_S14 = 33
    Global Const inputR117MainSheetColR117_PATTERN_S15 = 34
    Global Const inputR117MainSheetColR117_PATTERN_S16 = 35
    Global Const inputR117MainSheetColR117_PATTERN_S17 = 36
    Global Const inputR117MainSheetColR117_PATTERN_S18 = 37
    Global Const inputR117MainSheetColR117_STS = 38
    Global Const inputR117MainSheetColR117_DELKBN = 39
    Global Const inputR117MainSheetColR117_MAKE_YMD = 40
    Global Const inputR117MainSheetColR117_ALTER_YMD = 41
    Global Const inputR117MainSheetColR117_USERID = 42
    Global Const inputR117MainSheetColR117_CODE_NAME = 43
    Global Const inputR117MainSheetColR117_NINKA_SETKBN = 44
    Global Const inputR117MainSheetColR117_CHECK_DATA = 45
    Global Const inputR117MainSheetColR117_RECNO_ORG = 46
    Global Const inputR117MainSheetColR117_HNO = 47
    Global Const inputR117MainSheetColR117_IRI_SEND_FR = 48
    Global Const inputR117MainSheetColR117_IRI_SEND_TANTOU = 49
    Global Const inputR117MainSheetColR117_IRI_SEND_TO = 50
    Global Const inputR117MainSheetColR117_IRI_COPY_SECT1 = 51
    Global Const inputR117MainSheetColR117_IRI_COPY_SECT2 = 52
    Global Const inputR117MainSheetColR117_IRI_SEND_MEMO = 53
    Global Const inputR117MainSheetColR117_IRI_HREQ_YMD = 54
    Global Const inputR117MainSheetColR117_JIM_SEND_FR = 55
    Global Const inputR117MainSheetColR117_JIM_SEND_TANTOU = 56
    Global Const inputR117MainSheetColR117_JIM_SEND_TO = 57
    Global Const inputR117MainSheetColR117_JIM_COPY_SECT1 = 58
    Global Const inputR117MainSheetColR117_JIM_COPY_SECT2 = 59
    Global Const inputR117MainSheetColR117_JIM_SEND_MEMO = 60
    Global Const inputR117MainSheetColR117_JIM_HREQ_YMD = 61
    Global Const inputR117MainSheetColR117_TYPE = 62
    Global Const inputR117MainSheetColR117_PATTERN_S19 = 63
    Global Const inputR117MainSheetColR117_PATTERN_S20 = 64
    Global Const inputR117MainSheetColR117_PATTERN_S21 = 65
    Global Const inputR117MainSheetColR117_PATTERN_S22 = 66
    Global Const inputR117MainSheetColR117_PATTERN_S23 = 67
    Global Const inputR117MainSheetColR117_PATTERN_S24 = 68
    Global Const inputR117MainSheetColR117_PATTERN_S25 = 69
    Global Const inputR117MainSheetColR117_PATTERN_S26 = 70
    Global Const inputR117MainSheetColR117_PATTERN_S27 = 71
    Global Const inputR117MainSheetColR117_PATTERN_S28 = 72
    Global Const inputR117MainSheetColR117_PATTERN_S29 = 73
    Global Const inputR117MainSheetColR117_PATTERN_S30 = 74
    Global Const inputR117MainSheetColR117_PATTERN_S31 = 75
    Global Const inputR117MainSheetColR117_PATTERN_S32 = 76
    Global Const inputR117MainSheetColR117_PATTERN_S33 = 77
    Global Const inputR117MainSheetColR117_PATTERN_S34 = 78
    Global Const inputR117MainSheetColR117_PATTERN_S35 = 79
    Global Const inputR117MainSheetColR117_PATTERN_S36 = 80
    Global Const inputR117MainSheetColR117_PATTERN_S37 = 81
    Global Const inputR117MainSheetColR117_PATTERN_S38 = 82
    Global Const inputR117MainSheetColR117_PATTERN_S39 = 83
    Global Const inputR117MainSheetColR117_PATTERN_S40 = 84
    Global Const inputR117MainSheetColR117_PATTERN_S41 = 85
    Global Const inputR117MainSheetColR117_PATTERN_S42 = 86
    Global Const inputR117MainSheetColR117_PATTERN_S43 = 87
    Global Const inputR117MainSheetColR117_PATTERN_S44 = 88
    Global Const inputR117MainSheetColR117_PATTERN_S45 = 89
    Global Const inputR117MainSheetColR117_PATTERN_S46 = 90
    Global Const inputR117MainSheetColR117_PATTERN_S47 = 91
    Global Const inputR117MainSheetColR117_PATTERN_S48 = 92
    Global Const inputR117MainSheetColR117_PATTERN_S49 = 93
    Global Const inputR117MainSheetColR117_PATTERN_S50 = 94
    Global Const inputR117MainSheetColR117_PATTERN_S51 = 95
    Global Const inputR117MainSheetColR117_PATTERN_S52 = 96
    Global Const inputR117MainSheetColR117_PATTERN_S53 = 97
    Global Const inputR117MainSheetColR117_PATTERN_S54 = 98
    Global Const inputR117MainSheetColR117_PATTERN_S55 = 99
    Global Const inputR117MainSheetColR117_PATTERN_S56 = 100
    Global Const inputR117MainSheetColR117_PATTERN_S57 = 101
    Global Const inputR117MainSheetColR117_PATTERN_S58 = 102
    Global Const inputR117MainSheetColR117_PATTERN_S59 = 103
    Global Const inputR117MainSheetColR117_PATTERN_S60 = 104
    Global Const inputR117MainSheetColR117_PATTERN_S61 = 105
    Global Const inputR117MainSheetColR117_PATTERN_S62 = 106
    Global Const inputR117MainSheetColR117_PATTERN_S63 = 107
    Global Const inputR117MainSheetColR117_PATTERN_S64 = 108
    Global Const inputR117MainSheetColR117_PATTERN_S65 = 109
    Global Const inputR117MainSheetColR117_PATTERN_S66 = 110
    Global Const inputR117MainSheetColR117_PATTERN_S67 = 111
    Global Const inputR117MainSheetColR117_PATTERN_S68 = 112
    Global Const inputR117MainSheetColR117_PATTERN_S69 = 113
    Global Const inputR117MainSheetColR117_PATTERN_S70 = 114
    Global Const inputR117MainSheetColR117_PATTERN_S71 = 115
    Global Const inputR117MainSheetColR117_PATTERN_S72 = 116

'INPUT_TS_R117_SIZEDATA
    '項目の開始位置
    Global Const inputR117SizeDataSheetItemCol = 2
    Global Const inputR117SizeDataSheetItemRow = 3

    Global Const inputR117SizeDataSheetColR117S_NINKA_NO = 2
    Global Const inputR117SizeDataSheetColR117S_EXTENTION = 3
    Global Const inputR117SizeDataSheetColR117S_SECTION = 4
    Global Const inputR117SizeDataSheetColR117S_SPEED = 5
    Global Const inputR117SizeDataSheetColR117S_RIM_DIA = 6
    Global Const inputR117SizeDataSheetColR117S_HENPEI = 7
    Global Const inputR117SizeDataSheetColR117S_CM = 8
    Global Const inputR117SizeDataSheetColR117S_START_YMD = 9
    Global Const inputR117SizeDataSheetColR117S_STS = 10
    Global Const inputR117SizeDataSheetColR117S_DELKBN = 11
    Global Const inputR117SizeDataSheetColR117S_MAKE_YMD = 12
    Global Const inputR117SizeDataSheetColR117S_ALTER_YMD = 13
    Global Const inputR117SizeDataSheetColR117S_USERID = 14
    Global Const inputR117SizeDataSheetColR117S_CODE_NAME = 15
    Global Const inputR117SizeDataSheetColR117S_SIZE = 16
    Global Const inputR117SizeDataSheetColR117S_LI = 17
    Global Const inputR117SizeDataSheetColR117S_SS = 18
    Global Const inputR117SizeDataSheetColR117S_DESIGN_SECT = 19

'OUTPUT_TS_R117_SIZEDATA
    '項目の開始位置
    Const outputR117ExtDataSheetItemCol = 2
    Const outputR117ExtDataSheetItemRow = 3

    Const outputR117ExtDataSheetColR117E_NINKA_NO = 2
    Const outputR117ExtDataSheetColR117E_EXTENTION = 3
    Const outputR117ExtDataSheetColR117E_EXT_OI = 4
    Const outputR117ExtDataSheetColR117E_PLAN_SECT = 5
    Const outputR117ExtDataSheetColR117E_PLAN_TANTOU = 6
    Const outputR117ExtDataSheetColR117E_IRI_YMD = 7
    Const outputR117ExtDataSheetColR117E_IRI_NO = 8
    Const outputR117ExtDataSheetColR117E_APPLI_YMD = 9
    Const outputR117ExtDataSheetColR117E_APPRD_YMD = 10
    Const outputR117ExtDataSheetColR117E_RES_DB = 11
    Const outputR117ExtDataSheetColR117E_SIZE = 12
    Const outputR117ExtDataSheetColR117E_LISS = 13
    Const outputR117ExtDataSheetColR117E_PATTERN_S = 14
    Const outputR117ExtDataSheetColR117E_CM = 15
    Const outputR117ExtDataSheetColR117E_TEST_YMD = 16
    Const outputR117ExtDataSheetColR117E_TEST_NO = 17
    Const outputR117ExtDataSheetColR117E_STS = 18
    Const outputR117ExtDataSheetColR117E_DELKBN = 19
    Const outputR117ExtDataSheetColR117E_MAKE_YMD = 20
    Const outputR117ExtDataSheetColR117E_ALTER_YMD = 21
    Const outputR117ExtDataSheetColR117E_USERID = 22
    Const outputR117ExtDataSheetColR117E_CODE_NAME = 23
    Const outputR117ExtDataSheetColR117E_TESTFLG = 24
    Const outputR117ExtDataSheetColR117E_HON_SECT = 25
    Const outputR117ExtDataSheetColR117E_HON_TANTOU = 26
    Const outputR117ExtDataSheetColR117E_IRI_MEMO = 27
    Const outputR117ExtDataSheetColR117E_JIM_MEMO = 28
    Const outputR117ExtDataSheetColR117E_DB_TYPE = 29

'OUTPUT_TS_R117_MAIN
    '項目の開始位置
    Const outputR117MainSheetItemCol = 2
    Const outputR117MainSheetItemRow = 3

    Const outputR117MainSheetColR117_RECNO = 2
    Const outputR117MainSheetColR117_NINKA_NO_ALL = 3
    Const outputR117MainSheetColR117_NINKA_NO = 4
    Const outputR117MainSheetColR117_NATION = 5
    Const outputR117MainSheetColR117_EXTENTION = 6
    Const outputR117MainSheetColR117_PET_BUR_KBN = 7
    Const outputR117MainSheetColR117_CLASS = 8
    Const outputR117MainSheetColR117_GRP1 = 9
    Const outputR117MainSheetColR117_GRP2 = 10
    Const outputR117MainSheetColR117_GRP3 = 11
    Const outputR117MainSheetColR117_PURPOSE_KBN = 12
    Const outputR117MainSheetColR117_BRD = 13
    Const outputR117MainSheetColR117_PATTERN_L1 = 14
    Const outputR117MainSheetColR117_PATTERN_L2 = 15
    Const outputR117MainSheetColR117_PATTERN_L3 = 16
    Const outputR117MainSheetColR117_PATTERN_L4 = 17
    Const outputR117MainSheetColR117_PATTERN_L5 = 18
    Const outputR117MainSheetColR117_PATTERN_L6 = 19
    Const outputR117MainSheetColR117_PATTERN_S1 = 20
    Const outputR117MainSheetColR117_PATTERN_S2 = 21
    Const outputR117MainSheetColR117_PATTERN_S3 = 22
    Const outputR117MainSheetColR117_PATTERN_S4 = 23
    Const outputR117MainSheetColR117_PATTERN_S5 = 24
    Const outputR117MainSheetColR117_PATTERN_S6 = 25
    Const outputR117MainSheetColR117_PATTERN_S7 = 26
    Const outputR117MainSheetColR117_PATTERN_S8 = 27
    Const outputR117MainSheetColR117_PATTERN_S9 = 28
    Const outputR117MainSheetColR117_PATTERN_S10 = 29
    Const outputR117MainSheetColR117_PATTERN_S11 = 30
    Const outputR117MainSheetColR117_PATTERN_S12 = 31
    Const outputR117MainSheetColR117_PATTERN_S13 = 32
    Const outputR117MainSheetColR117_PATTERN_S14 = 33
    Const outputR117MainSheetColR117_PATTERN_S15 = 34
    Const outputR117MainSheetColR117_PATTERN_S16 = 35
    Const outputR117MainSheetColR117_PATTERN_S17 = 36
    Const outputR117MainSheetColR117_PATTERN_S18 = 37
    Const outputR117MainSheetColR117_STS = 38
    Const outputR117MainSheetColR117_DELKBN = 39
    Const outputR117MainSheetColR117_MAKE_YMD = 40
    Const outputR117MainSheetColR117_ALTER_YMD = 41
    Const outputR117MainSheetColR117_USERID = 42
    Const outputR117MainSheetColR117_CODE_NAME = 43
    Const outputR117MainSheetColR117_NINKA_SETKBN = 44
    Const outputR117MainSheetColR117_CHECK_DATA = 45
    Const outputR117MainSheetColR117_RECNO_ORG = 46
    Const outputR117MainSheetColR117_HNO = 47
    Const outputR117MainSheetColR117_IRI_SEND_FR = 48
    Const outputR117MainSheetColR117_IRI_SEND_TANTOU = 49
    Const outputR117MainSheetColR117_IRI_SEND_TO = 50
    Const outputR117MainSheetColR117_IRI_COPY_SECT1 = 51
    Const outputR117MainSheetColR117_IRI_COPY_SECT2 = 52
    Const outputR117MainSheetColR117_IRI_SEND_MEMO = 53
    Const outputR117MainSheetColR117_IRI_HREQ_YMD = 54
    Const outputR117MainSheetColR117_JIM_SEND_FR = 55
    Const outputR117MainSheetColR117_JIM_SEND_TANTOU = 56
    Const outputR117MainSheetColR117_JIM_SEND_TO = 57
    Const outputR117MainSheetColR117_JIM_COPY_SECT1 = 58
    Const outputR117MainSheetColR117_JIM_COPY_SECT2 = 59
    Const outputR117MainSheetColR117_JIM_SEND_MEMO = 60
    Const outputR117MainSheetColR117_JIM_HREQ_YMD = 61
    Const outputR117MainSheetColR117_TYPE = 62
    Const outputR117MainSheetColR117_PATTERN_S19 = 63
    Const outputR117MainSheetColR117_PATTERN_S20 = 64
    Const outputR117MainSheetColR117_PATTERN_S21 = 65
    Const outputR117MainSheetColR117_PATTERN_S22 = 66
    Const outputR117MainSheetColR117_PATTERN_S23 = 67
    Const outputR117MainSheetColR117_PATTERN_S24 = 68
    Const outputR117MainSheetColR117_PATTERN_S25 = 69
    Const outputR117MainSheetColR117_PATTERN_S26 = 70
    Const outputR117MainSheetColR117_PATTERN_S27 = 71
    Const outputR117MainSheetColR117_PATTERN_S28 = 72
    Const outputR117MainSheetColR117_PATTERN_S29 = 73
    Const outputR117MainSheetColR117_PATTERN_S30 = 74
    Const outputR117MainSheetColR117_PATTERN_S31 = 75
    Const outputR117MainSheetColR117_PATTERN_S32 = 76
    Const outputR117MainSheetColR117_PATTERN_S33 = 77
    Const outputR117MainSheetColR117_PATTERN_S34 = 78
    Const outputR117MainSheetColR117_PATTERN_S35 = 79
    Const outputR117MainSheetColR117_PATTERN_S36 = 80
    Const outputR117MainSheetColR117_PATTERN_S37 = 81
    Const outputR117MainSheetColR117_PATTERN_S38 = 82
    Const outputR117MainSheetColR117_PATTERN_S39 = 83
    Const outputR117MainSheetColR117_PATTERN_S40 = 84
    Const outputR117MainSheetColR117_PATTERN_S41 = 85
    Const outputR117MainSheetColR117_PATTERN_S42 = 86
    Const outputR117MainSheetColR117_PATTERN_S43 = 87
    Const outputR117MainSheetColR117_PATTERN_S44 = 88
    Const outputR117MainSheetColR117_PATTERN_S45 = 89
    Const outputR117MainSheetColR117_PATTERN_S46 = 90
    Const outputR117MainSheetColR117_PATTERN_S47 = 91
    Const outputR117MainSheetColR117_PATTERN_S48 = 92
    Const outputR117MainSheetColR117_PATTERN_S49 = 93
    Const outputR117MainSheetColR117_PATTERN_S50 = 94
    Const outputR117MainSheetColR117_PATTERN_S51 = 95
    Const outputR117MainSheetColR117_PATTERN_S52 = 96
    Const outputR117MainSheetColR117_PATTERN_S53 = 97
    Const outputR117MainSheetColR117_PATTERN_S54 = 98
    Const outputR117MainSheetColR117_PATTERN_S55 = 99
    Const outputR117MainSheetColR117_PATTERN_S56 = 100
    Const outputR117MainSheetColR117_PATTERN_S57 = 101
    Const outputR117MainSheetColR117_PATTERN_S58 = 102
    Const outputR117MainSheetColR117_PATTERN_S59 = 103
    Const outputR117MainSheetColR117_PATTERN_S60 = 104
    Const outputR117MainSheetColR117_PATTERN_S61 = 105
    Const outputR117MainSheetColR117_PATTERN_S62 = 106
    Const outputR117MainSheetColR117_PATTERN_S63 = 107
    Const outputR117MainSheetColR117_PATTERN_S64 = 108
    Const outputR117MainSheetColR117_PATTERN_S65 = 109
    Const outputR117MainSheetColR117_PATTERN_S66 = 110
    Const outputR117MainSheetColR117_PATTERN_S67 = 111
    Const outputR117MainSheetColR117_PATTERN_S68 = 112
    Const outputR117MainSheetColR117_PATTERN_S69 = 113
    Const outputR117MainSheetColR117_PATTERN_S70 = 114
    Const outputR117MainSheetColR117_PATTERN_S71 = 115
    Const outputR117MainSheetColR117_PATTERN_S72 = 116

'OUTPUT_TS_R117_SIZEDATA
    '項目の開始位置
    Const outputR117SizeDataSheetItemCol = 2
    Const outputR117SizeDataSheetItemRow = 3

    Const outputR117SizeDataSheetColR117S_NINKA_NO = 2
    Const outputR117SizeDataSheetColR117S_EXTENTION = 3
    Const outputR117SizeDataSheetColR117S_SECTION = 4
    Const outputR117SizeDataSheetColR117S_SPEED = 5
    Const outputR117SizeDataSheetColR117S_RIM_DIA = 6
    Const outputR117SizeDataSheetColR117S_HENPEI = 7
    Const outputR117SizeDataSheetColR117S_CM = 8
    Const outputR117SizeDataSheetColR117S_START_YMD = 9
    Const outputR117SizeDataSheetColR117S_STS = 10
    Const outputR117SizeDataSheetColR117S_DELKBN = 11
    Const outputR117SizeDataSheetColR117S_MAKE_YMD = 12
    Const outputR117SizeDataSheetColR117S_ALTER_YMD = 13
    Const outputR117SizeDataSheetColR117S_USERID = 14
    Const outputR117SizeDataSheetColR117S_CODE_NAME = 15
    Const outputR117SizeDataSheetColR117S_SIZE = 16
    Const outputR117SizeDataSheetColR117S_LI = 17
    Const outputR117SizeDataSheetColR117S_SS = 18
    Const outputR117SizeDataSheetColR117S_DESIGN_SECT = 19
    
'グローバル変数
Global GlbNinkaNoAllList As Object

Sub input_Click()

On Error GoTo lavelErr
    
    Dim maxRow As Long 'データ最終行
    Dim ninkaNoPreValue As String '前の認可No
    Dim ninkaNoAll As String '認可No
    Dim recNo As String 'レコードNo
    Dim ninkaGrpValue As String '認可グループ
    Dim ninkaNoValue As String '認可No
    Dim extValue As String
    Dim sectionValue As String
    Dim speedValue As String
    Dim rimDiaValue As String
    Dim henpeiValue As String
    Dim liValue As String
    Dim ssValue As String
    Dim cmValue As String
    Dim startYMDalue As String
    Dim nationoValue As String
    Dim petBurKbnValue As String
    Dim classValue As String
    Dim grpValue As String
    Dim patterntypeValue As String
    Dim brdValue As String
    Dim patternLValue As String
    Dim patternS1Value As String
    Dim sendFrValue As String
    Dim sendTantouValue As String
    Dim r117TypeValue As String


    Set GlbNinkaNoAllList = CreateObject("scripting.dictionary")

    'Call clearOutput

    Set wk = ThisWorkbook.Worksheets(inputSheet)
    With wk
        maxRow = .Cells(Rows.Count, inputSheetNinkaNoCol).End(xlUp).Row
        For i = inputSheetItmeRow To maxRow
            ninkaGrpValue = .Cells(i, inputSheetNinkaGrpCol).Value
            ninkaNoValue = .Cells(i, inputSheetNinkaNoCol).Value
            extValue = .Cells(i, inputSheetExtCol).Value
            sectionValue = .Cells(i, inputSheetSectionCol).Value
            speedValue = .Cells(i, inputSheetSpeedCol).Value
            rimDiaValue = .Cells(i, inputSheetRimDiaCol).Value
            henpeiValue = .Cells(i, inputSheetHenpeiCol).Value
            liValue = .Cells(i, inputSheetLICol).Value
            ssValue = .Cells(i, inputSheetSSCol).Value
            cmValue = .Cells(i, inputSheetCMCol).Value
            startYMDalue = .Cells(i, inputSheetStartYMDCol).Value
            nationoValue = .Cells(i, inputSheetNationCol).Value
            petBurKbnValue = .Cells(i, inputSheetPetBurKbnCol).Value
            classValue = .Cells(i, inputSheetClassCol).Value
            grpValue = .Cells(i, inputSheetGrp1Col).Value
            patterntypeValue = .Cells(i, inputSheetPatterntypeCol).Value
            brdValue = .Cells(i, inputSheetBrdCol).Value
            patternLValue = .Cells(i, inputSheetPatternLCol).Value
            patternS1Value = .Cells(i, inputSheetPatternS1Col).Value
            sendFrValue = .Cells(i, inputSheetSendFrCol).Value
            sendTantouValue = .Cells(i, inputSheetSendTantouCol).Value
            r117TypeValue = .Cells(i, inputSheetR117TypeCol).Value

            If ninkaNoValue Like "NEW*" Or ninkaNoValue Like "New*" Then
                If ninkaNoPreValue = Null Or ninkaNoPreValue <> ninkaNoValue Then
                    Dim ninkaNo As String
                    ninkaNo = getNinkaNo(nationoValue, classValue, r117TypeValue)
                    GlbNinkaNoAllList.Add ninkaNo, ""
                    ninkaNoAll = nationoValue + "-" + ninkaNo + " " + r117TypeValue
                    recNo = getNextRecNo("01")
                End If
            Else
                ninkaNoAll = ninkaNoValue
            End If
            'InputRowNo,認可No,レコードNo,国,極,クラス,グループ1,パターンカテゴリー,トレードネーム,正式商品名1,パターン略称1,発信担当,R117タイプ
            Call insertOrUpdate117Main(ninkaNoAll, ninkaNo, recNo, nationoValue, petBurKbnValue, classValue, grpValue, patterntypeValue, brdValue, patternLValue, patternS1Value, sendTantouValue, r117TypeValue)
            
            ninkaNoPreValue = ninkaNoValue
        Next i
    End With

Exit Sub
    
lavelErr:
    OnErrorMsg ("input_Click")
End Sub

Function getNinkaNo(strNATION, strCLASS, strTYPE)
    
On Error GoTo lavelErr
    Dim maxRow As Long 'データ最終行
    Dim regKbn As String
    Dim useFlg As Integer
    Dim strNinkaNo As String
    Dim strRegKbn As String
    Dim strR117Series As String
    Dim retNinkaNo As String

    If ExistsWorksheet(inputAppovalNoSheet) = False Then
        MsgBox "シート「" + inputAppovalNoSheet + "」が存在しません。", vbCritical, "エラーメッセージ"
        End
    End If

    If strTYPE = "S" Or strTYPE = "SW" Then
        strR117Series = "01"
    ElseIf strTYPE = "S2WR2B" Then
        strR117Series = "03"
    ElseIf strTYPE = "S2W2R3B" Or strTYPE = "S2W1R2B" Then
        strR117Series = "04"
    Else
        strR117Series = "02"
    End If

    If strCLASS = "C1" Then
        strRegKbn = strR117Series & "-C1"
    Else
        strRegKbn = strR117Series & "-C2"
    End If

    Set wk1 = ThisWorkbook.Worksheets(inputAppovalNoSheet)
    With wk1
        maxRow = .Cells(Rows.Count, inputAppovalNoSheetColNINKA_NO).End(xlUp).Row
        For i = inputAppovalNoSheetItemRow To maxRow
            regKbn = .Cells(i, inputAppovalNoSheetColREG_KBN).Value
            useFlg = .Cells(i, inputAppovalNoSheetColUSEFLG).Value
            strNinkaNo = .Cells(i, inputAppovalNoSheetColNINKA_NO).Value
            If regKbn = strRegKbn And useFlg <> 1 Then
                If GlbNinkaNoAllList.exists(strNinkaNo) = False Then
                    retNinkaNo = strNinkaNo
                    Exit For
                End If
            End If
        Next i
    End With
    If retNinkaNo = Null Then
        MsgBox "認可Noの採番にエラーが発生しました", vbCritical, "エラーメッセージ"
    End If

    getNinkaNo = retNinkaNo

Exit Function

lavelErr:

    OnErrorMsg ("レコードNoの採番にエラーが発生しました")

End Function


Function getNextRecNo(sRECID)
    
On Error GoTo lavelErr
    Dim maxRow As Long 'データ最終行
    Dim recId As String
    Dim maxNo As Long
    Dim nextNo As Long

    If ExistsWorksheet(inputR117NoSheet) = False Then
        MsgBox "シート「" + inputR117NoSheet + "」が存在しません。", vbCritical, "エラーメッセージ"
    End If

    Set wk = ThisWorkbook.Worksheets(inputR117NoSheet)
    With wk
        maxRow = .Cells(Rows.Count, inputR117NoSheetColR117_RECID).End(xlUp).Row
        For i = inputR117NoSheetItemRow To maxRow

            recId = .Cells(i, inputR117NoSheetColR117_RECID).Value
            maxNo = .Cells(i, inputR117NoSheetColR117_MAX_NO).Value

            If .Cells(i, inputR117NoSheetColR117_MAX_NO + 1).Value <> "" Then
                maxNo = .Cells(i, inputR117NoSheetColR117_MAX_NO + 1).Value
            End If

            If recId = sRECID Then
                nextNo = maxNo + 1
                .Cells(i, inputR117NoSheetColR117_MAX_NO + 1).Value = nextNo
            End If

        Next i
    End With

    If sRECID = Null Then
        MsgBox "レコードNoの採番にエラーが発生しました", vbCritical, "エラーメッセージ"
    End If

    getNextRecNo = nextNo

Exit Function

lavelErr:

    OnErrorMsg ("レコードNoの採番にエラーが発生しました")
    
End Function

Sub insertOrUpdate117Main(ninkaNoAll, ninkaNo, recNo, nationoValue, petBurKbnValue, classValue, grpValue, patterntypeValue, brdValue, patternLValue, patternS1Value, sendTantouValue, r117TypeValue)
    
On Error GoTo lavelErr
    Dim maxRow As Long 'データ最終行
    Dim motoNinkaNoAll As String
    Dim motoNinkaRow As Long
    Dim outputNinkaNoAll As String
    
    
    If ExistsWorksheet(inputR117MainSheet) = False Then
        MsgBox "シート「" + inputR117MainSheet + "」が存在しません。", vbCritical, "エラーメッセージ"
        End
    End If

     If ExistsWorksheet(outputR117MainSheet) = False Then
        MsgBox "シート「" + outputR117MainSheet + "」が存在しません。", vbCritical, "エラーメッセージ"
        End
    End If

    Set wk1 = ThisWorkbook.Worksheets(inputR117MainSheet)
    Set wk2 = ThisWorkbook.Worksheets(outputR117MainSheet)
    
    
    With wk1
        maxRow = .Cells(Rows.Count, inputR117MainSheetColR117_NINKA_NO_ALL).End(xlUp).Row
        For i = inputR117MainSheetItemRow To maxRow

            motoNinkaNoAll = .Cells(i, inputR117MainSheetColR117_NINKA_NO_ALL).Value

            If ninkaNoAll = motoNinkaNoAll Then
              motoNinkaRow = i
              Exit For
            Else
                motoNinkaNoAll = ""
            End If
        Next i
    End With

    With wk2
        maxRow = .Cells(Rows.Count, inputR117MainSheetColR117_NINKA_NO_ALL).End(xlUp).Row
        If maxRow = outputR117MainSheetItemRow - 1 Then
            If motoNinkaRow <> 0 Then
                wk2.Rows(maxRow + 1).Insert
                wk1.Range(wk1.Cells(motoNinkaRow, inputR117MainSheetColR117_RECNO), wk1.Cells(motoNinkaRow, inputR117MainSheetColR117_PATTERN_S72)).Copy wk2.Range(wk2.Cells(maxRow + 1, outputR117MainSheetColR117_RECNO), wk2.Cells(maxRow + 1, outputR117MainSheetColR117_PATTERN_S72))
                outputNinkaNoAll = motoNinkaNoAll
                'updata to moto
                .Cells(maxRow + 1, 1).Value = "CopyFromMoto"
                .Cells(maxRow + 1, inputR117MainSheetColR117_EXTENTION).Value = .Cells(i, inputR117MainSheetColR117_EXTENTION).Value + "+ * "
            Else
                .Rows(maxRow + 1).Insert
                .Cells(maxRow, 1).Value = "NEW"
                .Cells(maxRow, inputR117MainSheetColR117_RECNO).Value = recNo
                .Cells(maxRow, inputR117MainSheetColR117_NINKA_NO_ALL).Value = ninkaNoAll
                .Cells(maxRow, inputR117MainSheetColR117_NINKA_NO).Value = ninkaNo
                .Cells(maxRow, inputR117MainSheetColR117_NATION).Value = nationoValue
                .Cells(maxRow, inputR117MainSheetColR117_EXTENTION).Value = ""
                .Cells(maxRow, inputR117MainSheetColR117_PET_BUR_KBN).Value = petBurKbnValue
                .Cells(maxRow, inputR117MainSheetColR117_CLASS).Value = classValue
                .Cells(maxRow, inputR117MainSheetColR117_GRP1).Value = grpValue
                .Cells(maxRow, inputR117MainSheetColR117_CLASS).Value = ""
                .Cells(maxRow, inputR117MainSheetColR117_GRP2).Value = ""
                .Cells(maxRow, inputR117MainSheetColR117_GRP3).Value = ""
                .Cells(maxRow, inputR117MainSheetColR117_PURPOSE_KBN).Value = ""
                .Cells(maxRow, inputR117MainSheetColR117_BRD).Value = brdValue
   
            End If
            
        End If
        For i = outputR117MainSheetItemRow To maxRow
            '
            outputNinkaNoAll = .Cells(i, inputR117MainSheetColR117_NINKA_NO_ALL).Value
            
            If motoNinkaRow <> Null And outputNinkaNoAll <> motoNinkaNoAll And i = maxRow Then
                wk2.Rows(maxRow + 1).Insert
                wk1.Range(wk1.Cells(motoNinkaRow, inputR117MainSheetColR117_RECNO), wk1.Cells(motoNinkaRow, inputR117MainSheetColR117_PATTERN_S72)).Copy wk2.Range(wk2.Cells(i + 1, outputR117MainSheetColR117_RECNO), wk2.Cells(i + 1, outputR117MainSheetColR117_PATTERN_S72))
                'updata to moto
                .Cells(i + 1, 1).Value = "CopyFromMoto"
                .Cells(i + 1, inputR117MainSheetColR117_EXTENTION).Value = .Cells(i, inputR117MainSheetColR117_EXTENTION).Value + "+ * "
            End If
            
            If ninkaNoAll = outputNinkaNoAll Then
                .Cells(i, inputR117MainSheetColR117_EXTENTION).Value = .Cells(i, inputR117MainSheetColR117_EXTENTION).Value + "+ * "
                Exit For
            ElseIf i = maxRow Then
                .Rows(i + 1).Insert
                .Cells(i + 1, 1).Value = "NEW"
                .Cells(maxRow, inputR117MainSheetColR117_RECNO).Value = recNo
                .Cells(i + 1, inputR117MainSheetColR117_NINKA_NO_ALL).Value = ninkaNoAll
                .Cells(i + 1, inputR117MainSheetColR117_NINKA_NO).Value = ninkaNo
                .Cells(i + 1, inputR117MainSheetColR117_NATION).Value = nationoValue
                .Cells(i + 1, inputR117MainSheetColR117_EXTENTION).Value = ""
                .Cells(i + 1, inputR117MainSheetColR117_PET_BUR_KBN).Value = petBurKbnValue
                .Cells(i + 1, inputR117MainSheetColR117_CLASS).Value = classValue
                .Cells(i + 1, inputR117MainSheetColR117_GRP1).Value = grpValue
                .Cells(i + 1, inputR117MainSheetColR117_GRP2).Value = ""
                .Cells(i + 1, inputR117MainSheetColR117_GRP3).Value = ""
                .Cells(i + 1, inputR117MainSheetColR117_PURPOSE_KBN).Value = ""
                .Cells(i + 1, inputR117MainSheetColR117_BRD).Value = brdValue
            End If
            
        Next i
    End With
    
On Error GoTo lavelErr

Exit Sub

lavelErr:

    OnErrorMsg ("outputR117Main")
    
End Sub

Sub clearOutput()
    
On Error GoTo lavelErr
   
    If ExistsWorksheet(outputR117MainSheet) Then
        ThisWorkbook.Worksheets(outputR117MainSheet).Rows((outputR117MainSheetItemRow + 1) & ":" & Rows.Count).Delete
        ThisWorkbook.Worksheets(outputR117MainSheet).Rows(outputR117MainSheetItemRow & ":" & outputR117MainSheetItemRow).ClearContents
    End If

    If ExistsWorksheet(outputR117ExtDataSheet) Then
        ThisWorkbook.Worksheets(outputR117ExtDataSheet).Rows((outputR117ExtDataSheetItemRow + 1) & ":" & Rows.Count).Delete
        ThisWorkbook.Worksheets(outputR117ExtDataSheet).Rows(outputR117ExtDataSheetItemRow & ":" & outputR117ExtDataSheetItemRow).ClearContents
    End If

    If ExistsWorksheet(outputR117SizeDataSheet) Then
        ThisWorkbook.Worksheets(outputR117SizeDataSheet).Rows((outputR117SizeDataSheetItemRow + 1) & ":" & Rows.Count).Delete
        ThisWorkbook.Worksheets(outputR117SizeDataSheet).Rows(outputR117SizeDataSheetItemRow & ":" & outputR117SizeDataSheetItemRow).ClearContents
    End If
    'レコードNoの最終採番No
    If ExistsWorksheet(inputR117NoSheet) Then
        ThisWorkbook.Worksheets(inputR117NoSheet).Columns(inputR117NoSheetColR117_MAX_NO + 1).ClearContents
    End If

Exit Sub

lavelErr:

    OnErrorMsg ("Sub clearOutput")
    
End Sub


'シート存在チェック
Public Function ExistsWorksheet(ByVal name As String)

    Dim ws As Worksheet

On Error GoTo lavelErr

    For Each ws In Sheets
    
        If ws.name = name Then
        
            ' 存在する
            ExistsWorksheet = True
            
            Exit Function
            
        End If
        
    Next
    
    ' 存在しない
    ExistsWorksheet = False
    
Exit Function

lavelErr:
    OnErrorMsg ("ExistsWorksheet")
    
End Function

'on errorのメッセージ
Public Function OnErrorMsg(ByVal functionName As String)

    MsgBox "エラー処理：" & functionName & vbLf & _
            "エラー番号：" & Err.Number & vbLf & _
            "エラーメッセージ：" & Err.Description _
            , vbCritical, "システムエラー"
    
    ThisWorkbook.Activate
    
    'OUTPUTシートをクリックする
    'DeleteCopyWorkSheets
    
    '処理を終了する
    End
    
End Function









