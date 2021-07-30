*  WdMailSystem
#DEFINE wdNoMailSystem         0
#DEFINE wdMAPI                 1
#DEFINE wdPowerTalk            2
#DEFINE wdMAPIandPowerTalk     3

*  WdTemplateType
#DEFINE wdNormalTemplate       0
#DEFINE wdGlobalTemplate       1
#DEFINE wdAttachedTemplate     2

*  WdContinue
#DEFINE wdContinueDisabled     0
#DEFINE wdResetList            1
#DEFINE wdContinueList         2

*  WdIMEMode
#DEFINE wdIMEModeNoControl     0
#DEFINE wdIMEModeOn            1
#DEFINE wdIMEModeOff           2
#DEFINE wdIMEModeHiragana      4
#DEFINE wdIMEModeKatakana      5
#DEFINE wdIMEModeKatakanaHalf  6
#DEFINE wdIMEModeAlphaFull     7
#DEFINE wdIMEModeAlpha         8
#DEFINE wdIMEModeHangulFull    9
#DEFINE wdIMEModeHangul        10

*  WdBaselineAlignment
#DEFINE wdBaselineAlignTop        0
#DEFINE wdBaselineAlignCenter     1
#DEFINE wdBaselineAlignBaseline   2
#DEFINE wdBaselineAlignFarEast50  3
#DEFINE wdBaselineAlignAuto       4

*  WdIndexFilter
#DEFINE wdIndexFilterNone         0
#DEFINE wdIndexFilterAiueo        1
#DEFINE wdIndexFilterAkasatana    2
#DEFINE wdIndexFilterChosung      3
#DEFINE wdIndexFilterLow          4
#DEFINE wdIndexFilterMedium       5
#DEFINE wdIndexFilterFull         6

*  WdIndexSortBy
#DEFINE wdIndexSortByStroke       0
#DEFINE wdIndexSortBySyllable     1

*  WdJustificationMode
#DEFINE wdJustificationModeExpand        0
#DEFINE wdJustificationModeCompress      1
#DEFINE wdJustificationModeCompressKana  2

*  WdFarEastLineBreakLevel
#DEFINE wdFarEastLineBreakLevelNormal    0
#DEFINE wdFarEastLineBreakLevelStrict    1
#DEFINE wdFarEastLineBreakLevelCustom    2

*  WdMultipleWordConversionsMode
#DEFINE wdHangulToHanja            0
#DEFINE wdHanjaToHangul            1

*  WdColorIndex
#DEFINE wdAuto            0
#DEFINE wdBlack           1
#DEFINE wdBlue            2
#DEFINE wdTurquoise       3
#DEFINE wdBrightGreen     4
#DEFINE wdPink            5
#DEFINE wdRed             6
#DEFINE wdYellow          7
#DEFINE wdWhite           8
#DEFINE wdDarkBlue        9
#DEFINE wdTeal            10
#DEFINE wdGreen           11
#DEFINE wdViolet          12
#DEFINE wdDarkRed         13
#DEFINE wdDarkYellow      14
#DEFINE wdGray50          15
#DEFINE wdGray25          16
#DEFINE wdByAuthor        -1
#DEFINE wdNoHighlight     0

*  WdTextureIndex
#DEFINE wdTextureNone                 0
#DEFINE wdTexture2Pt5Percent          25
#DEFINE wdTexture5Percent             50
#DEFINE wdTexture7Pt5Percent          75
#DEFINE wdTexture10Percent            100
#DEFINE wdTexture12Pt5Percent         125
#DEFINE wdTexture15Percent            150
#DEFINE wdTexture17Pt5Percent         175
#DEFINE wdTexture20Percent            200
#DEFINE wdTexture22Pt5Percent         225
#DEFINE wdTexture25Percent            250
#DEFINE wdTexture27Pt5Percent         275
#DEFINE wdTexture30Percent            300
#DEFINE wdTexture32Pt5Percent         325
#DEFINE wdTexture35Percent            350
#DEFINE wdTexture37Pt5Percent         375
#DEFINE wdTexture40Percent            400
#DEFINE wdTexture42Pt5Percent         425
#DEFINE wdTexture45Percent            450
#DEFINE wdTexture47Pt5Percent         475
#DEFINE wdTexture50Percent            500
#DEFINE wdTexture52Pt5Percent         525
#DEFINE wdTexture55Percent            550
#DEFINE wdTexture57Pt5Percent         575
#DEFINE wdTexture60Percent            600
#DEFINE wdTexture62Pt5Percent         625
#DEFINE wdTexture65Percent            650
#DEFINE wdTexture67Pt5Percent         675
#DEFINE wdTexture70Percent            700
#DEFINE wdTexture72Pt5Percent         725
#DEFINE wdTexture75Percent            750
#DEFINE wdTexture77Pt5Percent         775
#DEFINE wdTexture80Percent            800
#DEFINE wdTexture82Pt5Percent         825
#DEFINE wdTexture85Percent            850
#DEFINE wdTexture87Pt5Percent         875
#DEFINE wdTexture90Percent            900
#DEFINE wdTexture92Pt5Percent         925
#DEFINE wdTexture95Percent            950
#DEFINE wdTexture97Pt5Percent         975
#DEFINE wdTextureSolid                1000
#DEFINE wdTextureDarkHorizontal       -1
#DEFINE wdTextureDarkVertical         -2
#DEFINE wdTextureDarkDiagonalDown     -3
#DEFINE wdTextureDarkDiagonalUp       -4
#DEFINE wdTextureDarkCross            -5
#DEFINE wdTextureDarkDiagonalCross    -6
#DEFINE wdTextureHorizontal           -7
#DEFINE wdTextureVertical             -8
#DEFINE wdTextureDiagonalDown         -9
#DEFINE wdTextureDiagonalUp           -10
#DEFINE wdTextureCross                -11
#DEFINE wdTextureDiagonalCross        -12

*  WdUnderline
#DEFINE wdUnderlineNone          0
#DEFINE wdUnderlineSingle        1
#DEFINE wdUnderlineWords         2
#DEFINE wdUnderlineDouble        3
#DEFINE wdUnderlineDotted        4
#DEFINE wdUnderlineThick         6
#DEFINE wdUnderlineDash          7
#DEFINE wdUnderlineDotDash       9
#DEFINE wdUnderlineDotDotDash    10
#DEFINE wdUnderlineWavy          11

*  WdEmphasisMark
#DEFINE wdEmphasisMarkNone               0
#DEFINE wdEmphasisMarkOverSolidCircle    1
#DEFINE wdEmphasisMarkOverComma          2
#DEFINE wdEmphasisMarkOverWhiteCircle    3
#DEFINE wdEmphasisMarkUnderSolidCircle   4

*  WdInternationalIndex
#DEFINE wdListSeparator            17
#DEFINE wdDecimalSeparator         18
#DEFINE wdThousandsSeparator       19
#DEFINE wdCurrencyCode             20
#DEFINE wd24HourClock              21
#DEFINE wdInternationalAM          22
#DEFINE wdInternationalPM          23
#DEFINE wdTimeSeparator            24
#DEFINE wdDateSeparator            25
#DEFINE wdProductLanguageID        26

*  WdAutoMacros
#DEFINE wdAutoExec        0
#DEFINE wdAutoNew         1
#DEFINE wdAutoOpen        2
#DEFINE wdAutoClose       3
#DEFINE wdAutoExit        4

*  WdCaptionPosition
#DEFINE wdCaptionPositionAbove     0
#DEFINE wdCaptionPositionBelow     1

*  WdCountry
#DEFINE wdUS                1
#DEFINE wdCanada            2
#DEFINE wdLatinAmerica      3
#DEFINE wdNetherlands       31
#DEFINE wdFrance            33
#DEFINE wdSpain             34
#DEFINE wdItaly             39
#DEFINE wdUK                44
#DEFINE wdDenmark           45
#DEFINE wdSweden            46
#DEFINE wdNorway            47
#DEFINE wdGermany           49
#DEFINE wdPeru              51
#DEFINE wdMexico            52
#DEFINE wdArgentina         54
#DEFINE wdBrazil            55
#DEFINE wdChile             56
#DEFINE wdVenezuela         58
#DEFINE wdJapan             81
#DEFINE wdTaiwan            886
#DEFINE wdChina             86
#DEFINE wdKorea             82
#DEFINE wdFinland           358
#DEFINE wdIceland           354

*  WdHeadingSeparator
#DEFINE wdHeadingSeparatorNone            0
#DEFINE wdHeadingSeparatorBlankLine       1
#DEFINE wdHeadingSeparatorLetter          2
#DEFINE wdHeadingSeparatorLetterLow       3
#DEFINE wdHeadingSeparatorLetterFull      4

*  WdSeparatorType
#DEFINE wdSeparatorHyphen    0
#DEFINE wdSeparatorPeriod    1
#DEFINE wdSeparatorColon     2
#DEFINE wdSeparatorEmDash    3
#DEFINE wdSeparatorEnDash    4

*  WdPageNumberAlignment
#DEFINE wdAlignPageNumberLeft            0
#DEFINE wdAlignPageNumberCenter          1
#DEFINE wdAlignPageNumberRight           2
#DEFINE wdAlignPageNumberInside          3
#DEFINE wdAlignPageNumberOutside         4

*  WdBorderType
#DEFINE wdBorderTop          -1
#DEFINE wdBorderLeft         -2
#DEFINE wdBorderBottom       -3
#DEFINE wdBorderRight        -4
#DEFINE wdBorderHorizontal   -5
#DEFINE wdBorderVertical     -6

*  WdBorderTypeHID
#DEFINE wdBorderDiagonalDown    -7
#DEFINE wdBorderDiagonalUp      -8

*  WdFramePosition
#DEFINE wdFrameTop           -999999
#DEFINE wdFrameLeft          -999998
#DEFINE wdFrameBottom        -999997
#DEFINE wdFrameRight         -999996
#DEFINE wdFrameCenter        -999995
#DEFINE wdFrameInside        -999994
#DEFINE wdFrameOutside       -999993

*  WdAnimation
#DEFINE wdAnimationNone                   0
#DEFINE wdAnimationLasVegasLights         1
#DEFINE wdAnimationBlinkingBackground     2
#DEFINE wdAnimationSparkleText            3
#DEFINE wdAnimationMarchingBlackAnts      4
#DEFINE wdAnimationMarchingRedAnts        5
#DEFINE wdAnimationShimmer                6

*  WdCharacterCase
#DEFINE wdNextCase            -1
#DEFINE wdLowerCase            0
#DEFINE wdUpperCase            1
#DEFINE wdTitleWord            2
#DEFINE wdTitleSentence        4
#DEFINE wdToggleCase           5

*  WdCharacterCaseHID
#DEFINE wdHalfWidth            6
#DEFINE wdFullWidth            7
#DEFINE wdKatakana             8
#DEFINE wdHiragana             9

*  WdSummaryMode
#DEFINE wdSummaryModeHighlight             0
#DEFINE wdSummaryModeHideAllButSummary     1
#DEFINE wdSummaryModeInsert                2
#DEFINE wdSummaryModeCreateNew             3

*  WdSummaryLength
#DEFINE wd10Sentences          -2
#DEFINE wd20Sentences          -3
#DEFINE wd100Words             -4
#DEFINE wd500Words             -5
#DEFINE wd10Percent            -6
#DEFINE wd25Percent            -7
#DEFINE wd50Percent            -8
#DEFINE wd75Percent            -9

*  WdStyleType
#DEFINE wdStyleTypeParagraph   1
#DEFINE wdStyleTypeCharacter   2

*  WdUnits
#DEFINE wdCharacter            1
#DEFINE wdWord                 2
#DEFINE wdSentence             3
#DEFINE wdParagraph            4
#DEFINE wdLine                 5
#DEFINE wdStory                6
#DEFINE wdScreen               7
#DEFINE wdSection              8
#DEFINE wdColumn               9
#DEFINE wdRow                  10
#DEFINE wdWindow               11
#DEFINE wdCell                 12
#DEFINE wdCharacterFormatting  13
#DEFINE wdParagraphFormatting  14
#DEFINE wdTable                15
#DEFINE wdItem                 16

*  WdGoToItem
#DEFINE wdGoToBookmark             -1
#DEFINE wdGoToSection              0
#DEFINE wdGoToPage                 1
#DEFINE wdGoToTable                2
#DEFINE wdGoToLine                 3
#DEFINE wdGoToFootnote             4
#DEFINE wdGoToEndnote              5
#DEFINE wdGoToComment              6
#DEFINE wdGoToField                7
#DEFINE wdGoToGraphic              8
#DEFINE wdGoToObject               9
#DEFINE wdGoToEquation             10
#DEFINE wdGoToHeading              11
#DEFINE wdGoToPercent              12
#DEFINE wdGoToSpellingError        13
#DEFINE wdGoToGrammaticalError     14
#DEFINE wdGoToProofreadingError    15

*  WdGoToDirection
#DEFINE wdGoToFirst           1
#DEFINE wdGoToLast            -1
#DEFINE wdGoToNext            2
#DEFINE wdGoToRelative        2
#DEFINE wdGoToPrevious        3
#DEFINE wdGoToAbsolute        1

*  WdCollapseDirection
#DEFINE wdCollapseStart       1
#DEFINE wdCollapseEnd         0

*  WdRowHeightRule
#DEFINE wdRowHeightAuto       0
#DEFINE wdRowHeightAtLeast    1
#DEFINE wdRowHeightExactly    2

*  WdFrameSizeRule
#DEFINE wdFrameAuto           0
#DEFINE wdFrameAtLeast        1
#DEFINE wdFrameExact          2

*  WdInsertCells
#DEFINE wdInsertCellsShiftRight        0
#DEFINE wdInsertCellsShiftDown         1
#DEFINE wdInsertCellsEntireRow         2
#DEFINE wdInsertCellsEntireColumn      3

*  WdDeleteCells
#DEFINE wdDeleteCellsShiftLeft         0
#DEFINE wdDeleteCellsShiftUp           1
#DEFINE wdDeleteCellsEntireRow         2
#DEFINE wdDeleteCellsEntireColumn      3

*  WdListApplyTo
#DEFINE wdListApplyToWholeList         0
#DEFINE wdListApplyToThisPointForward  1
#DEFINE wdListApplyToSelection         2

*  WdAlertLevel
#DEFINE wdAlertsNone              0
#DEFINE wdAlertsMessageBox        -2
#DEFINE wdAlertsAll               -1

*  WdCursorType
#DEFINE wdCursorWait              0
#DEFINE wdCursorIBeam             1
#DEFINE wdCursorNormal            2
#DEFINE wdCursorNorthwestArrow    3

*  WdEnableCancelKey
#DEFINE wdCancelDisabled          0
#DEFINE wdCancelInterrupt         1

*  WdRulerStyle
#DEFINE wdAdjustNone              0
#DEFINE wdAdjustProportional      1
#DEFINE wdAdjustFirstColumn       2
#DEFINE wdAdjustSameWidth         3

*  WdParagraphAlignment
#DEFINE wdAlignParagraphLeft      0
#DEFINE wdAlignParagraphCenter    1
#DEFINE wdAlignParagraphRight     2
#DEFINE wdAlignParagraphJustify   3

*  WdParagraphAlignmentHID
#DEFINE wdAlignParagraphDistribute  4

*  WdListLevelAlignment
#DEFINE wdListLevelAlignLeft      0
#DEFINE wdListLevelAlignCenter    1
#DEFINE wdListLevelAlignRight     2

*  WdRowAlignment
#DEFINE wdAlignRowLeft            0
#DEFINE wdAlignRowCenter          1
#DEFINE wdAlignRowRight           2

*  WdTabAlignment
#DEFINE wdAlignTabLeft            0
#DEFINE wdAlignTabCenter          1
#DEFINE wdAlignTabRight           2
#DEFINE wdAlignTabDecimal         3
#DEFINE wdAlignTabBar             4
#DEFINE wdAlignTabList            6

*  WdVerticalAlignment
#DEFINE wdAlignVerticalTop        0
#DEFINE wdAlignVerticalCenter     1
#DEFINE wdAlignVerticalJustify    2
#DEFINE wdAlignVerticalBottom     3

*  WdCellVerticalAlignment
#DEFINE wdCellAlignVerticalTop    0
#DEFINE wdCellAlignVerticalCenter 1
#DEFINE wdCellAlignVerticalBottom 3

*  WdTrailingCharacter
#DEFINE wdTrailingTab             0
#DEFINE wdTrailingSpace           1
#DEFINE wdTrailingNone            2

*  WdListGalleryType
#DEFINE wdBulletGallery           1
#DEFINE wdNumberGallery           2
#DEFINE wdOutlineNumberGallery    3

*  WdListNumberStyle
#DEFINE wdListNumberStyleArabic                 0
#DEFINE wdListNumberStyleUppercaseRoman         1
#DEFINE wdListNumberStyleLowercaseRoman         2
#DEFINE wdListNumberStyleUppercaseLetter        3
#DEFINE wdListNumberStyleLowercaseLetter        4
#DEFINE wdListNumberStyleOrdinal                5
#DEFINE wdListNumberStyleCardinalText           6
#DEFINE wdListNumberStyleOrdinalText            7
#DEFINE wdListNumberStyleArabicLZ               22
#DEFINE wdListNumberStyleBullet                 23
#DEFINE wdListNumberStyleLegal                  253
#DEFINE wdListNumberStyleLegalLZ                254
#DEFINE wdListNumberStyleNone                   255

*  WdListNumberStyleHID
#DEFINE wdListNumberStyleKanji                  10
#DEFINE wdListNumberStyleKanjiDigit             11
#DEFINE wdListNumberStyleAiueoHalfWidth         12
#DEFINE wdListNumberStyleIrohaHalfWidth         13
#DEFINE wdListNumberStyleArabicFullWidth        14
#DEFINE wdListNumberStyleKanjiTraditional       16
#DEFINE wdListNumberStyleKanjiTraditional2      17
#DEFINE wdListNumberStyleNumberInCircle         18
#DEFINE wdListNumberStyleAiueo                  20
#DEFINE wdListNumberStyleIroha                  21
#DEFINE wdListNumberStyleGanada                 24
#DEFINE wdListNumberStyleChosung                25
#DEFINE wdListNumberStyleGBNum1                 26
#DEFINE wdListNumberStyleGBNum2                 27
#DEFINE wdListNumberStyleGBNum3                 28
#DEFINE wdListNumberStyleGBNum4                 29
#DEFINE wdListNumberStyleZodiac1                30
#DEFINE wdListNumberStyleZodiac2                31
#DEFINE wdListNumberStyleZodiac3                32
#DEFINE wdListNumberStyleTradChinNum1           33
#DEFINE wdListNumberStyleTradChinNum2           34
#DEFINE wdListNumberStyleTradChinNum3           35
#DEFINE wdListNumberStyleTradChinNum4           36
#DEFINE wdListNumberStyleSimpChinNum1           37
#DEFINE wdListNumberStyleSimpChinNum2           38
#DEFINE wdListNumberStyleSimpChinNum3           39
#DEFINE wdListNumberStyleSimpChinNum4           40
#DEFINE wdListNumberStyleHanjaRead              41
#DEFINE wdListNumberStyleHanjaReadDigit         42
#DEFINE wdListNumberStyleHangul                 43
#DEFINE wdListNumberStyleHanja                  44


*  WdNoteNumberStyle
#DEFINE wdNoteNumberStyleArabic                 0
#DEFINE wdNoteNumberStyleUppercaseRoman         1
#DEFINE wdNoteNumberStyleLowercaseRoman         2
#DEFINE wdNoteNumberStyleUppercaseLetter        3
#DEFINE wdNoteNumberStyleLowercaseLetter        4
#DEFINE wdNoteNumberStyleSymbol                 9

*  WdNoteNumberStyleHID
#DEFINE wdNoteNumberStyleArabicFullWidth        14
#DEFINE wdNoteNumberStyleKanji                  10
#DEFINE wdNoteNumberStyleKanjiDigit             11
#DEFINE wdNoteNumberStyleKanjiTraditional       16
#DEFINE wdNoteNumberStyleNumberInCircle         18
#DEFINE wdNoteNumberStyleHanjaRead              41
#DEFINE wdNoteNumberStyleHanjaReadDigit         42
#DEFINE wdNoteNumberStyleTradChinNum1           33
#DEFINE wdNoteNumberStyleTradChinNum2           34
#DEFINE wdNoteNumberStyleSimpChinNum1           37
#DEFINE wdNoteNumberStyleSimpChinNum2           38

*  WdCaptionNumberStyle
#DEFINE wdCaptionNumberStyleArabic              0
#DEFINE wdCaptionNumberStyleUppercaseRoman      1
#DEFINE wdCaptionNumberStyleLowercaseRoman      2
#DEFINE wdCaptionNumberStyleUppercaseLetter     3
#DEFINE wdCaptionNumberStyleLowercaseLetter     4

*  WdCaptionNumberStyleHID
#DEFINE wdCaptionNumberStyleArabicFullWidth     14
#DEFINE wdCaptionNumberStyleKanji               10
#DEFINE wdCaptionNumberStyleKanjiDigit          11
#DEFINE wdCaptionNumberStyleKanjiTraditional    16
#DEFINE wdCaptionNumberStyleNumberInCircle      18
#DEFINE wdCaptionNumberStyleGanada              24
#DEFINE wdCaptionNumberStyleChosung             25
#DEFINE wdCaptionNumberStyleZodiac1             30
#DEFINE wdCaptionNumberStyleZodiac2             31
#DEFINE wdCaptionNumberStyleHanjaRead           41
#DEFINE wdCaptionNumberStyleHanjaReadDigit      42
#DEFINE wdCaptionNumberStyleTradChinNum2        34
#DEFINE wdCaptionNumberStyleTradChinNum3        35
#DEFINE wdCaptionNumberStyleSimpChinNum2        38
#DEFINE wdCaptionNumberStyleSimpChinNum3        39

*  WdPageNumberStyle
#DEFINE wdPageNumberStyleArabic                 0
#DEFINE wdPageNumberStyleUppercaseRoman         1
#DEFINE wdPageNumberStyleLowercaseRoman         2
#DEFINE wdPageNumberStyleUppercaseLetter        3
#DEFINE wdPageNumberStyleLowercaseLetter        4

*  WdPageNumberStyleHID
#DEFINE wdPageNumberStyleArabicFullWidth        14
#DEFINE wdPageNumberStyleKanji                  10
#DEFINE wdPageNumberStyleKanjiDigit             11
#DEFINE wdPageNumberStyleKanjiTraditional       16
#DEFINE wdPageNumberStyleNumberInCircle         18
#DEFINE wdPageNumberStyleHanjaRead              41
#DEFINE wdPageNumberStyleHanjaReadDigit         42
#DEFINE wdPageNumberStyleTradChinNum1           33
#DEFINE wdPageNumberStyleTradChinNum2           34
#DEFINE wdPageNumberStyleSimpChinNum1           37
#DEFINE wdPageNumberStyleSimpChinNum2           38

*  WdStatistic
#DEFINE wdStatisticWords                        0
#DEFINE wdStatisticLines                        1
#DEFINE wdStatisticPages                        2
#DEFINE wdStatisticCharacters                   3
#DEFINE wdStatisticParagraphs                   4
#DEFINE wdStatisticCharactersWithSpaces         5

*  WdStatisticHID
#DEFINE wdStatisticFarEastCharacters            6

*  WdBuiltInProperty
#DEFINE wdPropertyTitle              1
#DEFINE wdPropertySubject            2
#DEFINE wdPropertyAuthor             3
#DEFINE wdPropertyKeywords           4
#DEFINE wdPropertyComments           5
#DEFINE wdPropertyTemplate           6
#DEFINE wdPropertyLastAuthor         7
#DEFINE wdPropertyRevision           8
#DEFINE wdPropertyAppName            9
#DEFINE wdPropertyTimeLastPrinted    10
#DEFINE wdPropertyTimeCreated        11
#DEFINE wdPropertyTimeLastSaved      12
#DEFINE wdPropertyVBATotalEdit       13
#DEFINE wdPropertyPages              14
#DEFINE wdPropertyWords              15
#DEFINE wdPropertyCharacters         16
#DEFINE wdPropertySecurity           17
#DEFINE wdPropertyCategory           18
#DEFINE wdPropertyFormat             19
#DEFINE wdPropertyManager            20
#DEFINE wdPropertyCompany            21
#DEFINE wdPropertyBytes              22
#DEFINE wdPropertyLines              23
#DEFINE wdPropertyParas              24
#DEFINE wdPropertySlides             25
#DEFINE wdPropertyNotes              26
#DEFINE wdPropertyHiddenSlides       27
#DEFINE wdPropertyMMClips            28
#DEFINE wdPropertyHyperlinkBase      29
#DEFINE wdPropertyCharsWSpaces       30

*  WdLineSpacing
#DEFINE wdLineSpaceSingle            0
#DEFINE wdLineSpace1pt5              1
#DEFINE wdLineSpaceDouble            2
#DEFINE wdLineSpaceAtLeast           3
#DEFINE wdLineSpaceExactly           4
#DEFINE wdLineSpaceMultiple          5

*  WdNumberType
#DEFINE wdNumberParagraph            1
#DEFINE wdNumberListNum              2
#DEFINE wdNumberAllNumbers           3

*  WdListType
#DEFINE wdListNoNumbering            0
#DEFINE wdListListNumOnly            1
#DEFINE wdListBullet                 2
#DEFINE wdListSimpleNumbering        3
#DEFINE wdListOutlineNumbering       4
#DEFINE wdListMixedNumbering         5

*  WdStoryType
#DEFINE wdMainTextStory              1
#DEFINE wdFootnotesStory             2
#DEFINE wdEndnotesStory              3
#DEFINE wdCommentsStory              4
#DEFINE wdTextFrameStory             5
#DEFINE wdEvenPagesHeaderStory       6
#DEFINE wdPrimaryHeaderStory         7
#DEFINE wdEvenPagesFooterStory       8
#DEFINE wdPrimaryFooterStory         9
#DEFINE wdFirstPageHeaderStory       10
#DEFINE wdFirstPageFooterStory       11

*  WdSaveFormat
#DEFINE wdFormatDocument             0
#DEFINE wdFormatTemplate             1
#DEFINE wdFormatText                 2
#DEFINE wdFormatTextLineBreaks       3
#DEFINE wdFormatDOSText              4
#DEFINE wdFormatDOSTextLineBreaks    5
#DEFINE wdFormatRTF                  6
#DEFINE wdFormatUnicodeText          7

*  WdOpenFormat
#DEFINE wdOpenFormatAuto             0
#DEFINE wdOpenFormatDocument         1
#DEFINE wdOpenFormatTemplate         2
#DEFINE wdOpenFormatRTF              3
#DEFINE wdOpenFormatText             4
#DEFINE wdOpenFormatUnicodeText      5

*  WdHeaderFooterIndex
#DEFINE wdHeaderFooterPrimary        1
#DEFINE wdHeaderFooterFirstPage      2
#DEFINE wdHeaderFooterEvenPages      3

*  WdTocFormat
#DEFINE wdTOCTemplate        0
#DEFINE wdTOCClassic         1
#DEFINE wdTOCDistinctive     2
#DEFINE wdTOCFancy           3
#DEFINE wdTOCModern          4
#DEFINE wdTOCFormal          5
#DEFINE wdTOCSimple          6

*  WdTofFormat
#DEFINE wdTOFTemplate        0
#DEFINE wdTOFClassic         1
#DEFINE wdTOFDistinctive     2
#DEFINE wdTOFCentered        3
#DEFINE wdTOFFormal          4
#DEFINE wdTOFSimple          5

*  WdToaFormat
#DEFINE wdTOATemplate        0
#DEFINE wdTOAClassic         1
#DEFINE wdTOADistinctive     2
#DEFINE wdTOAFormal          3
#DEFINE wdTOASimple          4

*  WdLineStyle
#DEFINE wdLineStyleNone                       0
#DEFINE wdLineStyleSingle                     1
#DEFINE wdLineStyleDot                        2
#DEFINE wdLineStyleDashSmallGap               3
#DEFINE wdLineStyleDashLargeGap               4
#DEFINE wdLineStyleDashDot                    5
#DEFINE wdLineStyleDashDotDot                 6
#DEFINE wdLineStyleDouble                     7
#DEFINE wdLineStyleTriple                     8
#DEFINE wdLineStyleThinThickSmallGap          9
#DEFINE wdLineStyleThickThinSmallGap          10
#DEFINE wdLineStyleThinThickThinSmallGap      11
#DEFINE wdLineStyleThinThickMedGap            12
#DEFINE wdLineStyleThickThinMedGap            13
#DEFINE wdLineStyleThinThickThinMedGap        14
#DEFINE wdLineStyleThinThickLargeGap          15
#DEFINE wdLineStyleThickThinLargeGap          16
#DEFINE wdLineStyleThinThickThinLargeGap      17
#DEFINE wdLineStyleSingleWavy                 18
#DEFINE wdLineStyleDoubleWavy                 19
#DEFINE wdLineStyleDashDotStroked             20
#DEFINE wdLineStyleEmboss3D                   21
#DEFINE wdLineStyleEngrave3D                  22

*  WdLineWidth
#DEFINE wdLineWidth025pt            2
#DEFINE wdLineWidth050pt            4
#DEFINE wdLineWidth075pt            6
#DEFINE wdLineWidth100pt            8
#DEFINE wdLineWidth150pt            12
#DEFINE wdLineWidth225pt            18
#DEFINE wdLineWidth300pt            24
#DEFINE wdLineWidth450pt            36
#DEFINE wdLineWidth600pt            48

*  WdBreakType
#DEFINE wdSectionBreakNextPage      2
#DEFINE wdSectionBreakContinuous    3
#DEFINE wdSectionBreakEvenPage      4
#DEFINE wdSectionBreakOddPage       5
#DEFINE wdLineBreak                 6
#DEFINE wdPageBreak                 7
#DEFINE wdColumnBreak               8

*  WdTabLeader
#DEFINE wdTabLeaderSpaces           0
#DEFINE wdTabLeaderDots             1
#DEFINE wdTabLeaderDashes           2
#DEFINE wdTabLeaderLines            3

*  WdTabLeaderHID
#DEFINE wdTabLeaderHeavy            4
#DEFINE wdTabLeaderMiddleDot        5

*  WdMeasurementUnits
#DEFINE wdInches            0
#DEFINE wdCentimeters       1
#DEFINE wdPoints            3
#DEFINE wdPicas             4

*  WdMeasurementUnitsHID
#DEFINE wdMillimeters       2

*  WdDropPosition
#DEFINE wdDropNone          0
#DEFINE wdDropNormal        1
#DEFINE wdDropMargin        2

*  WdNumberingRule
#DEFINE wdRestartContinuous 0
#DEFINE wdRestartSection    1
#DEFINE wdRestartPage       2

*  WdFootnoteLocation
#DEFINE wdBottomOfPage      0
#DEFINE wdBeneathText       1

*  WdEndnoteLocation
#DEFINE wdEndOfSection      0
#DEFINE wdEndOfDocument     1

*  WdSortSeparator
#DEFINE wdSortSeparateByTabs                    0
#DEFINE wdSortSeparateByCommas                  1
#DEFINE wdSortSeparateByDefaultTableSeparator   2

*  WdTableFieldSeparator
#DEFINE wdSeparateByParagraphs                  0
#DEFINE wdSeparateByTabs                        1
#DEFINE wdSeparateByCommas                      2
#DEFINE wdSeparateByDefaultListSeparator        3

*  WdSortFieldType
#DEFINE wdSortFieldAlphanumeric    0
#DEFINE wdSortFieldNumeric         1
#DEFINE wdSortFieldDate            2

*  WdSortFieldTypeHID
#DEFINE wdSortFieldSyllable        3
#DEFINE wdSortFieldJapanJIS        4
#DEFINE wdSortFieldStroke          5
#DEFINE wdSortFieldKoreaKS         6

*  WdSortOrder
#DEFINE wdSortOrderAscending       0
#DEFINE wdSortOrderDescending      1

*  WdTableFormat
#DEFINE wdTableFormatNone               0
#DEFINE wdTableFormatSimple1            1
#DEFINE wdTableFormatSimple2            2
#DEFINE wdTableFormatSimple3            3
#DEFINE wdTableFormatClassic1           4
#DEFINE wdTableFormatClassic2           5
#DEFINE wdTableFormatClassic3           6
#DEFINE wdTableFormatClassic4           7
#DEFINE wdTableFormatColorful1          8
#DEFINE wdTableFormatColorful2          9
#DEFINE wdTableFormatColorful3          10
#DEFINE wdTableFormatColumns1           11
#DEFINE wdTableFormatColumns2           12
#DEFINE wdTableFormatColumns3           13
#DEFINE wdTableFormatColumns4           14
#DEFINE wdTableFormatColumns5           15
#DEFINE wdTableFormatGrid1              16
#DEFINE wdTableFormatGrid2              17
#DEFINE wdTableFormatGrid3              18
#DEFINE wdTableFormatGrid4              19
#DEFINE wdTableFormatGrid5              20
#DEFINE wdTableFormatGrid6              21
#DEFINE wdTableFormatGrid7              22
#DEFINE wdTableFormatGrid8              23
#DEFINE wdTableFormatList1              24
#DEFINE wdTableFormatList2              25
#DEFINE wdTableFormatList3              26
#DEFINE wdTableFormatList4              27
#DEFINE wdTableFormatList5              28
#DEFINE wdTableFormatList6              29
#DEFINE wdTableFormatList7              30
#DEFINE wdTableFormatList8              31
#DEFINE wdTableFormat3DEffects1         32
#DEFINE wdTableFormat3DEffects2         33
#DEFINE wdTableFormat3DEffects3         34
#DEFINE wdTableFormatContemporary       35
#DEFINE wdTableFormatElegant            36
#DEFINE wdTableFormatProfessional       37
#DEFINE wdTableFormatSubtle1            38
#DEFINE wdTableFormatSubtle2            39

*  WdTableFormatApply
#DEFINE wdTableFormatApplyBorders       1
#DEFINE wdTableFormatApplyShading       2
#DEFINE wdTableFormatApplyFont          4
#DEFINE wdTableFormatApplyColor         8
#DEFINE wdTableFormatApplyAutoFit       16
#DEFINE wdTableFormatApplyHeadingRows   32
#DEFINE wdTableFormatApplyLastRow       64
#DEFINE wdTableFormatApplyFirstColumn   128
#DEFINE wdTableFormatApplyLastColumn    256

*  WdLanguageID
#DEFINE wdLanguageNone          0
#DEFINE wdNoProofing            1024
#DEFINE wdDanish                1030
#DEFINE wdGerman                1031
#DEFINE wdSwissGerman           2055
#DEFINE wdEnglishAUS            3081
#DEFINE wdEnglishUK             2057
#DEFINE wdEnglishUS             1033
#DEFINE wdEnglishCanadian       4105
#DEFINE wdEnglishNewZealand     5129
#DEFINE wdEnglishSouthAfrica    7177
#DEFINE wdSpanish               1034
#DEFINE wdFrench                1036
#DEFINE wdFrenchCanadian        3084
#DEFINE wdItalian               1040
#DEFINE wdDutch                 1043
#DEFINE wdNorwegianBokmol       1044
#DEFINE wdNorwegianNynorsk      2068
#DEFINE wdBrazilianPortuguese   1046
#DEFINE wdPortuguese            2070
#DEFINE wdFinnish               1035
#DEFINE wdSwedish               1053
#DEFINE wdCatalan               1027
#DEFINE wdGreek                 1032
#DEFINE wdTurkish               1055
#DEFINE wdRussian               1049
#DEFINE wdCzech                 1029
#DEFINE wdHungarian             1038
#DEFINE wdPolish                1045
#DEFINE wdSlovenian             1060
#DEFINE wdBasque                1069
#DEFINE wdMalaysian             1086
#DEFINE wdJapanese              1041
#DEFINE wdKorean                1042
#DEFINE wdSimplifiedChinese     2052
#DEFINE wdTraditionalChinese    1028
#DEFINE wdSwissFrench           4108
#DEFINE wdSesotho               1072
#DEFINE wdTsonga                1073
#DEFINE wdTswana                1074
#DEFINE wdVenda                 1075
#DEFINE wdXhosa                 1076
#DEFINE wdZulu                  1077
#DEFINE wdAfrikaans             1078
#DEFINE wdArabic                1025
#DEFINE wdHebrew                1037
#DEFINE wdSlovak                1051
#DEFINE wdFarsi                 1065
#DEFINE wdRomanian              1048
#DEFINE wdCroatian              1050
#DEFINE wdUkrainian             1058
#DEFINE wdByelorussian          1059
#DEFINE wdEstonian              1061
#DEFINE wdLatvian               1062
#DEFINE wdMacedonian            1071
#DEFINE wdSerbianLatin          2074
#DEFINE wdSerbianCyrillic       3098
#DEFINE wdIcelandic             1039
#DEFINE wdBelgianFrench         2060
#DEFINE wdBelgianDutch          2067
#DEFINE wdBulgarian             1026
#DEFINE wdMexicanSpanish        2058
#DEFINE wdSpanishModernSort     3082
#DEFINE wdSwissItalian          2064

*  WdFieldType
#DEFINE wdFieldEmpty            -1
#DEFINE wdFieldRef              3
#DEFINE wdFieldIndexEntry       4
#DEFINE wdFieldFootnoteRef      5
#DEFINE wdFieldSet              6
#DEFINE wdFieldIf               7
#DEFINE wdFieldIndex            8
#DEFINE wdFieldTOCEntry         9
#DEFINE wdFieldStyleRef         10
#DEFINE wdFieldRefDoc           11
#DEFINE wdFieldSequence         12
#DEFINE wdFieldTOC              13
#DEFINE wdFieldInfo             14
#DEFINE wdFieldTitle            15
#DEFINE wdFieldSubject          16
#DEFINE wdFieldAuthor           17
#DEFINE wdFieldKeyWord          18
#DEFINE wdFieldComments         19
#DEFINE wdFieldLastSavedBy      20
#DEFINE wdFieldCreateDate       21
#DEFINE wdFieldSaveDate         22
#DEFINE wdFieldPrintDate        23
#DEFINE wdFieldRevisionNum      24
#DEFINE wdFieldEditTime         25
#DEFINE wdFieldNumPages         26
#DEFINE wdFieldNumWords         27
#DEFINE wdFieldNumChars         28
#DEFINE wdFieldFileName         29
#DEFINE wdFieldTemplate         30
#DEFINE wdFieldDate             31
#DEFINE wdFieldTime             32
#DEFINE wdFieldPage             33
#DEFINE wdFieldExpression       34
#DEFINE wdFieldQuote            35
#DEFINE wdFieldInclude          36
#DEFINE wdFieldPageRef          37
#DEFINE wdFieldAsk              38
#DEFINE wdFieldFillIn           39
#DEFINE wdFieldData             40
#DEFINE wdFieldNext             41
#DEFINE wdFieldNextIf           42
#DEFINE wdFieldSkipIf           43
#DEFINE wdFieldMergeRec         44
#DEFINE wdFieldDDE              45
#DEFINE wdFieldDDEAuto          46
#DEFINE wdFieldGlossary         47
#DEFINE wdFieldPrint            48
#DEFINE wdFieldFormula          49
#DEFINE wdFieldGoToButton       50
#DEFINE wdFieldMacroButton      51
#DEFINE wdFieldAutoNumOutline   52
#DEFINE wdFieldAutoNumLegal     53
#DEFINE wdFieldAutoNum          54
#DEFINE wdFieldImport           55
#DEFINE wdFieldLink             56
#DEFINE wdFieldSymbol           57
#DEFINE wdFieldEmbed            58
#DEFINE wdFieldMergeField       59
#DEFINE wdFieldUserName         60
#DEFINE wdFieldUserInitials     61
#DEFINE wdFieldUserAddress      62
#DEFINE wdFieldBarCode          63
#DEFINE wdFieldDocVariable      64
#DEFINE wdFieldSection          65
#DEFINE wdFieldSectionPages     66
#DEFINE wdFieldIncludePicture   67
#DEFINE wdFieldIncludeText      68
#DEFINE wdFieldFileSize         69
#DEFINE wdFieldFormTextInput    70
#DEFINE wdFieldFormCheckBox     71
#DEFINE wdFieldNoteRef          72
#DEFINE wdFieldTOA              73
#DEFINE wdFieldTOAEntry         74
#DEFINE wdFieldMergeSeq         75
#DEFINE wdFieldPrivate          77
#DEFINE wdFieldDatabase         78
#DEFINE wdFieldAutoText         79
#DEFINE wdFieldCompare          80
#DEFINE wdFieldAddin            81
#DEFINE wdFieldSubscriber       82
#DEFINE wdFieldFormDropDown     83
#DEFINE wdFieldAdvance          84
#DEFINE wdFieldDocProperty      85
#DEFINE wdFieldOCX              87
#DEFINE wdFieldHyperlink        88
#DEFINE wdFieldAutoTextList     89
#DEFINE wdFieldListNum          90
#DEFINE wdFieldHTMLActiveX      91

*  WdBuiltinStyle
#DEFINE wdStyleNormal                  -1
#DEFINE wdStyleEnvelopeAddress         -37
#DEFINE wdStyleEnvelopeReturn          -38
#DEFINE wdStyleBodyText                -67
#DEFINE wdStyleHeading1                -2
#DEFINE wdStyleHeading2                -3
#DEFINE wdStyleHeading3                -4
#DEFINE wdStyleHeading4                -5
#DEFINE wdStyleHeading5                -6
#DEFINE wdStyleHeading6                -7
#DEFINE wdStyleHeading7                -8
#DEFINE wdStyleHeading8                -9
#DEFINE wdStyleHeading9                -10
#DEFINE wdStyleIndex1                  -11
#DEFINE wdStyleIndex2                  -12
#DEFINE wdStyleIndex3                  -13
#DEFINE wdStyleIndex4                  -14
#DEFINE wdStyleIndex5                  -15
#DEFINE wdStyleIndex6                  -16
#DEFINE wdStyleIndex7                  -17
#DEFINE wdStyleIndex8                  -18
#DEFINE wdStyleIndex9                  -19
#DEFINE wdStyleTOC1                    -20
#DEFINE wdStyleTOC2                    -21
#DEFINE wdStyleTOC3                    -22
#DEFINE wdStyleTOC4                    -23
#DEFINE wdStyleTOC5                    -24
#DEFINE wdStyleTOC6                    -25
#DEFINE wdStyleTOC7                    -26
#DEFINE wdStyleTOC8                    -27
#DEFINE wdStyleTOC9                    -28
#DEFINE wdStyleNormalIndent            -29
#DEFINE wdStyleFootnoteText            -30
#DEFINE wdStyleCommentText             -31
#DEFINE wdStyleHeader                  -32
#DEFINE wdStyleFooter                  -33
#DEFINE wdStyleIndexHeading            -34
#DEFINE wdStyleCaption                 -35
#DEFINE wdStyleTableOfFigures          -36
#DEFINE wdStyleFootnoteReference       -39
#DEFINE wdStyleCommentReference        -40
#DEFINE wdStyleLineNumber              -41
#DEFINE wdStylePageNumber              -42
#DEFINE wdStyleEndnoteReference        -43
#DEFINE wdStyleEndnoteText             -44
#DEFINE wdStyleTableOfAuthorities      -45
#DEFINE wdStyleMacroText               -46
#DEFINE wdStyleTOAHeading              -47
#DEFINE wdStyleList                    -48
#DEFINE wdStyleListBullet              -49
#DEFINE wdStyleListNumber              -50
#DEFINE wdStyleList2                   -51
#DEFINE wdStyleList3                   -52
#DEFINE wdStyleList4                   -53
#DEFINE wdStyleList5                   -54
#DEFINE wdStyleListBullet2             -55
#DEFINE wdStyleListBullet3             -56
#DEFINE wdStyleListBullet4             -57
#DEFINE wdStyleListBullet5             -58
#DEFINE wdStyleListNumber2             -59
#DEFINE wdStyleListNumber3             -60
#DEFINE wdStyleListNumber4             -61
#DEFINE wdStyleListNumber5             -62
#DEFINE wdStyleTitle                   -63
#DEFINE wdStyleClosing                 -64
#DEFINE wdStyleSignature               -65
#DEFINE wdStyleDefaultParagraphFont    -66
#DEFINE wdStyleBodyTextIndent          -68
#DEFINE wdStyleListContinue            -69
#DEFINE wdStyleListContinue2           -70
#DEFINE wdStyleListContinue3           -71
#DEFINE wdStyleListContinue4           -72
#DEFINE wdStyleListContinue5           -73
#DEFINE wdStyleMessageHeader           -74
#DEFINE wdStyleSubtitle                -75
#DEFINE wdStyleSalutation              -76
#DEFINE wdStyleDate                    -77
#DEFINE wdStyleBodyTextFirstIndent     -78
#DEFINE wdStyleBodyTextFirstIndent2    -79
#DEFINE wdStyleNoteHeading             -80
#DEFINE wdStyleBodyText2               -81
#DEFINE wdStyleBodyText3               -82
#DEFINE wdStyleBodyTextIndent2         -83
#DEFINE wdStyleBodyTextIndent3         -84
#DEFINE wdStyleBlockQuotation          -85
#DEFINE wdStyleHyperlink               -86
#DEFINE wdStyleHyperlinkFollowed       -87
#DEFINE wdStyleStrong                  -88
#DEFINE wdStyleEmphasis                -89
#DEFINE wdStyleNavPane                 -90
#DEFINE wdStylePlainText               -91

*  WdWordDialogTab
#DEFINE wdDialogToolsOptionsTabView                             204
#DEFINE wdDialogToolsOptionsTabGeneral                          203
#DEFINE wdDialogToolsOptionsTabEdit                             224
#DEFINE wdDialogToolsOptionsTabPrint                            208
#DEFINE wdDialogToolsOptionsTabSave                             209
#DEFINE wdDialogToolsOptionsTabProofread                        211
#DEFINE wdDialogToolsOptionsTabTrackChanges                     386
#DEFINE wdDialogToolsOptionsTabUserInfo                         213
#DEFINE wdDialogToolsOptionsTabCompatibility                    525
#DEFINE wdDialogToolsOptionsTabFileLocations                    225
#DEFINE wdDialogFilePageSetupTabMargins                         150000
#DEFINE wdDialogFilePageSetupTabPaperSize                       150001
#DEFINE wdDialogFilePageSetupTabPaperSource                     150002
#DEFINE wdDialogFilePageSetupTabLayout                          150003
#DEFINE wdDialogInsertSymbolTabSymbols                          200000
#DEFINE wdDialogInsertSymbolTabSpecialCharacters                200001
#DEFINE wdDialogNoteOptionsTabAllFootnotes                      300000
#DEFINE wdDialogNoteOptionsTabAllEndnotes                       300001
#DEFINE wdDialogInsertIndexAndTablesTabIndex                    400000
#DEFINE wdDialogInsertIndexAndTablesTabTableOfContents          400001
#DEFINE wdDialogInsertIndexAndTablesTabTableOfFigures           400002
#DEFINE wdDialogInsertIndexAndTablesTabTableOfAuthorities       400003
#DEFINE wdDialogOrganizerTabStyles                              500000
#DEFINE wdDialogOrganizerTabAutoText                            500001
#DEFINE wdDialogOrganizerTabCommandBars                         500002
#DEFINE wdDialogOrganizerTabMacros                              500003
#DEFINE wdDialogFormatFontTabFont                               600000
#DEFINE wdDialogFormatFontTabCharacterSpacing                   600001
#DEFINE wdDialogFormatFontTabAnimation                          600002
#DEFINE wdDialogFormatBordersAndShadingTabBorders               700000
#DEFINE wdDialogFormatBordersAndShadingTabPageBorder            700001
#DEFINE wdDialogFormatBordersAndShadingTabShading               700002
#DEFINE wdDialogToolsEnvelopesAndLabelsTabEnvelopes             800000
#DEFINE wdDialogToolsEnvelopesAndLabelsTabLabels                800001
#DEFINE wdDialogFormatParagraphTabIndentsAndSpacing             1000000
#DEFINE wdDialogFormatParagraphTabTextFlow                      1000001
#DEFINE wdDialogFormatDrawingObjectTabColorsAndLines            1200000
#DEFINE wdDialogFormatDrawingObjectTabSize                      1200001
#DEFINE wdDialogFormatDrawingObjectTabPosition                  1200002
#DEFINE wdDialogFormatDrawingObjectTabWrapping                  1200003
#DEFINE wdDialogFormatDrawingObjectTabPicture                   1200004
#DEFINE wdDialogFormatDrawingObjectTabTextbox                   1200005
#DEFINE wdDialogToolsAutoCorrectExceptionsTabFirstLetter        1400000
#DEFINE wdDialogToolsAutoCorrectExceptionsTabInitialCaps        1400001
#DEFINE wdDialogFormatBulletsAndNumberingTabBulleted            1500000
#DEFINE wdDialogFormatBulletsAndNumberingTabNumbered            1500001
#DEFINE wdDialogFormatBulletsAndNumberingTabOutlineNumbered     1500002
#DEFINE wdDialogLetterWizardTabLetterFormat                     1600000
#DEFINE wdDialogLetterWizardTabRecipientInfo                    1600001
#DEFINE wdDialogLetterWizardTabOtherElements                    1600002
#DEFINE wdDialogLetterWizardTabSenderInfo                       1600003
#DEFINE wdDialogToolsAutoManagerTabAutoCorrect                  1700000
#DEFINE wdDialogToolsAutoManagerTabAutoFormatAsYouType          1700001
#DEFINE wdDialogToolsAutoManagerTabAutoText                     1700002
#DEFINE wdDialogToolsAutoManagerTabAutoFormat                   1700003

*  WdWordDialogTabHID
#DEFINE wdDialogToolsOptionsTabTypography                       739
#DEFINE wdDialogToolsOptionsTabFuzzy                            790
#DEFINE wdDialogToolsOptionsTabHangulHanjaConversion            786
#DEFINE wdDialogFilePageSetupTabCharsLines                      150004
#DEFINE wdDialogFormatParagraphTabTeisai                        1000002
#DEFINE wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet  1400002

*  WdWordDialog
#DEFINE wdDialogHelpAbout                      9
#DEFINE wdDialogHelpWordPerfectHelp            10
#DEFINE wdDialogHelpWordPerfectHelpOptions     511
#DEFINE wdDialogFormatChangeCase               322
#DEFINE wdDialogToolsWordCount                 228
#DEFINE wdDialogDocumentStatistics             78
#DEFINE wdDialogFileNew                        79
#DEFINE wdDialogFileOpen                       80
#DEFINE wdDialogMailMergeOpenDataSource        81
#DEFINE wdDialogMailMergeOpenHeaderSource      82
#DEFINE wdDialogMailMergeUseAddressBook        779
#DEFINE wdDialogFileSaveAs                     84
#DEFINE wdDialogFileSummaryInfo                86
#DEFINE wdDialogToolsTemplates                 87
#DEFINE wdDialogOrganizer                      222
#DEFINE wdDialogFilePrint                      88
#DEFINE wdDialogMailMerge                      676
#DEFINE wdDialogMailMergeCheck                 677
#DEFINE wdDialogMailMergeQueryOptions          681
#DEFINE wdDialogMailMergeFindRecord            569
#DEFINE wdDialogMailMergeInsertIf              4049
#DEFINE wdDialogMailMergeInsertNextIf          4053
#DEFINE wdDialogMailMergeInsertSkipIf          4055
#DEFINE wdDialogMailMergeInsertFillIn          4048
#DEFINE wdDialogMailMergeInsertAsk             4047
#DEFINE wdDialogMailMergeInsertSet             4054
#DEFINE wdDialogMailMergeHelper                680
#DEFINE wdDialogLetterWizard                   821
#DEFINE wdDialogFilePrintSetup                 97
#DEFINE wdDialogFileFind                       99
#DEFINE wdDialogMailMergeCreateDataSource      642
#DEFINE wdDialogMailMergeCreateHeaderSource    643
#DEFINE wdDialogEditPasteSpecial               111
#DEFINE wdDialogEditFind                       112
#DEFINE wdDialogEditReplace                    117
#DEFINE wdDialogEditGoToOld                    811
#DEFINE wdDialogEditGoTo                       896
#DEFINE wdDialogCreateAutoText                 872
#DEFINE wdDialogEditAutoText                   985
#DEFINE wdDialogEditLinks                      124
#DEFINE wdDialogEditObject                     125
#DEFINE wdDialogConvertObject                  392
#DEFINE wdDialogTableToText                    128
#DEFINE wdDialogTextToTable                    127
#DEFINE wdDialogTableInsertTable               129
#DEFINE wdDialogTableInsertCells               130
#DEFINE wdDialogTableInsertRow                 131
#DEFINE wdDialogTableDeleteCells               133
#DEFINE wdDialogTableSplitCells                137
#DEFINE wdDialogTableFormula                   348
#DEFINE wdDialogTableAutoFormat                563
#DEFINE wdDialogTableFormatCell                612
#DEFINE wdDialogViewZoom                       577
#DEFINE wdDialogNewToolbar                     586
#DEFINE wdDialogInsertBreak                    159
#DEFINE wdDialogInsertFootnote                 370
#DEFINE wdDialogInsertSymbol                   162
#DEFINE wdDialogInsertPicture                  163
#DEFINE wdDialogInsertFile                     164
#DEFINE wdDialogInsertDateTime                 165
#DEFINE wdDialogInsertField                    166
#DEFINE wdDialogInsertDatabase                 341
#DEFINE wdDialogInsertMergeField               167
#DEFINE wdDialogInsertBookmark                 168
#DEFINE wdDialogMarkIndexEntry                 169
#DEFINE wdDialogMarkCitation                   463
#DEFINE wdDialogEditTOACategory                625
#DEFINE wdDialogInsertIndexAndTables           473
#DEFINE wdDialogInsertIndex                    170
#DEFINE wdDialogInsertTableOfContents          171
#DEFINE wdDialogMarkTableOfContentsEntry       442
#DEFINE wdDialogInsertTableOfFigures           472
#DEFINE wdDialogInsertTableOfAuthorities       471
#DEFINE wdDialogInsertObject                   172
#DEFINE wdDialogFormatCallout                  610
#DEFINE wdDialogDrawSnapToGrid                 633
#DEFINE wdDialogDrawAlign                      634
#DEFINE wdDialogToolsEnvelopesAndLabels        607
#DEFINE wdDialogToolsCreateEnvelope            173
#DEFINE wdDialogToolsCreateLabels              489
#DEFINE wdDialogToolsProtectDocument           503
#DEFINE wdDialogToolsProtectSection            578
#DEFINE wdDialogToolsUnprotectDocument         521
#DEFINE wdDialogFormatFont                     174
#DEFINE wdDialogFormatParagraph                175
#DEFINE wdDialogFormatSectionLayout            176
#DEFINE wdDialogFormatColumns                  177
#DEFINE wdDialogFileDocumentLayout             178
#DEFINE wdDialogFileMacPageSetup               685
#DEFINE wdDialogFilePrintOneCopy               445
#DEFINE wdDialogFileMacPageSetupGX             444
#DEFINE wdDialogFileMacCustomPageSetupGX       737
#DEFINE wdDialogFilePageSetup                  178
#DEFINE wdDialogFormatTabs                     179
#DEFINE wdDialogFormatStyle                    180
#DEFINE wdDialogFormatStyleGallery             505
#DEFINE wdDialogFormat#DEFINEStyleFont         181
* #DEFINE wdDialogFormat#DEFINEStylePara         182
* #DEFINE wdDialogFormat#DEFINEStyleTabs         183
* #DEFINE wdDialogFormat#DEFINEStyleFrame        184
* #DEFINE wdDialogFormat#DEFINEStyleBorders      185
* #DEFINE wdDialogFormat#DEFINEStyleLang         186
#DEFINE wdDialogFormatPicture                  187
#DEFINE wdDialogToolsLanguage                  188
#DEFINE wdDialogFormatBordersAndShading        189
#DEFINE wdDialogFormatDrawingObject            960
#DEFINE wdDialogFormatFrame                    190
#DEFINE wdDialogFormatDropCap                  488
#DEFINE wdDialogFormatBulletsAndNumbering      824
#DEFINE wdDialogToolsHyphenation               195
#DEFINE wdDialogToolsBulletsNumbers            196
#DEFINE wdDialogToolsHighlightChanges          197
#DEFINE wdDialogToolsAcceptRejectChanges       506
#DEFINE wdDialogToolsMergeDocuments            435
#DEFINE wdDialogToolsCompareDocuments          198
#DEFINE wdDialogTableSort                      199
#DEFINE wdDialogToolsCustomizeMenuBar          615
#DEFINE wdDialogToolsCustomize                 152
#DEFINE wdDialogToolsCustomizeKeyboard         432
#DEFINE wdDialogToolsCustomizeMenus            433
#DEFINE wdDialogListCommands                   723
#DEFINE wdDialogToolsOptions                   974
#DEFINE wdDialogToolsOptionsGeneral            203
#DEFINE wdDialogToolsAdvancedSettings          206
#DEFINE wdDialogToolsOptionsCompatibility      525
#DEFINE wdDialogToolsOptionsPrint              208
#DEFINE wdDialogToolsOptionsSave               209
#DEFINE wdDialogToolsOptionsSpellingAndGrammar 211
#DEFINE wdDialogToolsSpellingAndGrammar        828
#DEFINE wdDialogToolsThesaurus                 194
#DEFINE wdDialogToolsOptionsUserInfo           213
#DEFINE wdDialogToolsOptionsAutoFormat         959
#DEFINE wdDialogToolsOptionsTrackChanges       386
#DEFINE wdDialogToolsOptionsEdit               224
#DEFINE wdDialogToolsMacro                     215
#DEFINE wdDialogInsertPageNumbers              294
#DEFINE wdDialogFormatPageNumber               298
#DEFINE wdDialogNoteOptions                    373
#DEFINE wdDialogCopyFile                       300
#DEFINE wdDialogFormatAddrFonts                103
#DEFINE wdDialogFormatRetAddrFonts             221
#DEFINE wdDialogToolsOptionsFileLocations      225
#DEFINE wdDialogToolsCreateDirectory           833
#DEFINE wdDialogUpdateTOC                      331
#DEFINE wdDialogInsertFormField                483
#DEFINE wdDialogFormFieldOptions               353
#DEFINE wdDialogInsertCaption                  357
#DEFINE wdDialogInsertAutoCaption              359
#DEFINE wdDialogInsertAddCaption               402
#DEFINE wdDialogInsertCaptionNumbering         358
#DEFINE wdDialogInsertCrossReference           367
#DEFINE wdDialogToolsManageFields              631
#DEFINE wdDialogToolsAutoManager               915
#DEFINE wdDialogToolsAutoCorrect               378
#DEFINE wdDialogToolsAutoCorrectExceptions     762
#DEFINE wdDialogConnect                        420
#DEFINE wdDialogToolsOptionsView               204
#DEFINE wdDialogInsertSubdocument              583
#DEFINE wdDialogFileRoutingSlip                624
#DEFINE wdDialogFontSubstitution               581
#DEFINE wdDialogEditCreatePublisher            732
#DEFINE wdDialogEditSubscribeTo                733
#DEFINE wdDialogEditPublishOptions             735
#DEFINE wdDialogEditSubscribeOptions           736
#DEFINE wdDialogControlRun                     235
#DEFINE wdDialogFileVersions                   945
#DEFINE wdDialogToolsAutoSummarize             874
#DEFINE wdDialogFileSaveVersion                1007
#DEFINE wdDialogWindowActivate                 220
#DEFINE wdDialogToolsMacroRecord               214
#DEFINE wdDialogToolsRevisions                 197

*  WdWordDialogHID
#DEFINE wdDialogToolsOptionsFuzzy               790
#DEFINE wdDialogToolsOptionsTypography          739
#DEFINE wdDialogToolsOptionsAutoFormatAsYouType 778

*  WdFieldKind
#DEFINE wdFieldKindNone            0
#DEFINE wdFieldKindHot             1
#DEFINE wdFieldKindWarm            2
#DEFINE wdFieldKindCold            3

*  WdTextFormFieldType
#DEFINE wdRegularText              0
#DEFINE wdNumberText               1
#DEFINE wdDateText                 2
#DEFINE wdCurrentDateText          3
#DEFINE wdCurrentTimeText          4
#DEFINE wdCalculationText          5

*  WdChevronConvertRule
#DEFINE wdNeverConvert             0
#DEFINE wdAlwaysConvert            1
#DEFINE wdAskToNotConvert          2
#DEFINE wdAskToConvert             3

*  WdMailMergeMainDocType
#DEFINE wdNotAMergeDocument        -1
#DEFINE wdFormLetters              0
#DEFINE wdMailingLabels            1
#DEFINE wdEnvelopes                2
#DEFINE wdCatalog                  3

*  WdMailMergeState
#DEFINE wdNormalDocument           0
#DEFINE wdMainDocumentOnly         1
#DEFINE wdMainAndDataSource        2
#DEFINE wdMainAndHeader            3
#DEFINE wdMainAndSourceAndHeader   4
#DEFINE wdDataSource               5

*  WdMailMergeDestination
#DEFINE wdSendToNewDocument        0
#DEFINE wdSendToPrinter            1
#DEFINE wdSendToEmail              2
#DEFINE wdSendToFax                3

*  WdMailMergeActiveRecord
#DEFINE wdNoActiveRecord           -1
#DEFINE wdNextRecord               -2
#DEFINE wdPreviousRecord           -3
#DEFINE wdFirstRecord              -4
#DEFINE wdLastRecord               -5

*  WdMailMergeDefaultRecord
#DEFINE wdDefaultFirstRecord       1
#DEFINE wdDefaultLastRecord        -16

*  WdMailMergeDataSource
#DEFINE wdNoMergeInfo              -1
#DEFINE wdMergeInfoFromWord        0
#DEFINE wdMergeInfoFromAccessDDE   1
#DEFINE wdMergeInfoFromExcelDDE    2
#DEFINE wdMergeInfoFromMSQueryDDE  3
#DEFINE wdMergeInfoFromODBC        4

*  WdMailMergeComparison
#DEFINE wdMergeIfEqual                0
#DEFINE wdMergeIfNotEqual             1
#DEFINE wdMergeIfLessThan             2
#DEFINE wdMergeIfGreaterThan          3
#DEFINE wdMergeIfLessThanOrEqual      4
#DEFINE wdMergeIfGreaterThanOrEqual   5
#DEFINE wdMergeIfIsBlank              6
#DEFINE wdMergeIfIsNotBlank           7

*  WdBookmarkSortBy
#DEFINE wdSortByName            0
#DEFINE wdSortByLocation        1

*  WdWindowState
#DEFINE wdWindowStateNormal     0
#DEFINE wdWindowStateMaximize   1
#DEFINE wdWindowStateMinimize   2

*  WdPictureLinkType
#DEFINE wdLinkNone              0
#DEFINE wdLinkDataInDoc         1
#DEFINE wdLinkDataOnDisk        2

*  WdLinkType
#DEFINE wdLinkTypeOLE           0
#DEFINE wdLinkTypePicture       1
#DEFINE wdLinkTypeText          2
#DEFINE wdLinkTypeReference     3
#DEFINE wdLinkTypeInclude       4
#DEFINE wdLinkTypeImport        5
#DEFINE wdLinkTypeDDE           6
#DEFINE wdLinkTypeDDEAuto       7

*  WdWindowType
#DEFINE wdWindowDocument        0
#DEFINE wdWindowTemplate        1

*  WdViewType
#DEFINE wdNormalView            1
#DEFINE wdOutlineView           2
#DEFINE wdPageView              3
#DEFINE wdPrintPreview          4
#DEFINE wdMasterView            5
#DEFINE wdOnlineView            6

*  WdSeekView
#DEFINE wdSeekMainDocument      0
#DEFINE wdSeekPrimaryHeader     1
#DEFINE wdSeekFirstPageHeader   2
#DEFINE wdSeekEvenPagesHeader   3
#DEFINE wdSeekPrimaryFooter     4
#DEFINE wdSeekFirstPageFooter   5
#DEFINE wdSeekEvenPagesFooter   6
#DEFINE wdSeekFootnotes         7
#DEFINE wdSeekEndnotes          8
#DEFINE wdSeekCurrentPageHeader 9
#DEFINE wdSeekCurrentPageFooter 10

*  WdSpecialPane
#DEFINE wdPaneNone                              0
#DEFINE wdPanePrimaryHeader                     1
#DEFINE wdPaneFirstPageHeader                   2
#DEFINE wdPaneEvenPagesHeader                   3
#DEFINE wdPanePrimaryFooter                     4
#DEFINE wdPaneFirstPageFooter                   5
#DEFINE wdPaneEvenPagesFooter                   6
#DEFINE wdPaneFootnotes                         7
#DEFINE wdPaneEndnotes                          8
#DEFINE wdPaneFootnoteContinuationNotice        9
#DEFINE wdPaneFootnoteContinuationSeparator     10
#DEFINE wdPaneFootnoteSeparator                 11
#DEFINE wdPaneEndnoteContinuationNotice         12
#DEFINE wdPaneEndnoteContinuationSeparator      13
#DEFINE wdPaneEndnoteSeparator                  14
#DEFINE wdPaneComments                          15
#DEFINE wdPaneCurrentPageHeader                 16
#DEFINE wdPaneCurrentPageFooter                 17

*  WdPageFit
#DEFINE wdPageFitNone           0
#DEFINE wdPageFitFullPage       1
#DEFINE wdPageFitBestFit        2

*  WdBrowseTarget
#DEFINE wdBrowsePage            1
#DEFINE wdBrowseSection         2
#DEFINE wdBrowseComment         3
#DEFINE wdBrowseFootnote        4
#DEFINE wdBrowseEndnote         5
#DEFINE wdBrowseField           6
#DEFINE wdBrowseTable           7
#DEFINE wdBrowseGraphic         8
#DEFINE wdBrowseHeading         9
#DEFINE wdBrowseEdit            10
#DEFINE wdBrowseFind            11
#DEFINE wdBrowseGoTo            12

*  WdPaperTray
#DEFINE wdPrinterDefaultBin             0 
#DEFINE wdPrinterUpperBin               1
#DEFINE wdPrinterOnlyBin                1
#DEFINE wdPrinterLowerBin               2
#DEFINE wdPrinterMiddleBin              3
#DEFINE wdPrinterManualFeed             4
#DEFINE wdPrinterEnvelopeFeed           5
#DEFINE wdPrinterManualEnvelopeFeed     6
#DEFINE wdPrinterAutomaticSheetFeed     7
#DEFINE wdPrinterTractorFeed            8
#DEFINE wdPrinterSmallFormatBin         9
#DEFINE wdPrinterLargeFormatBin         10
#DEFINE wdPrinterLargeCapacityBin       11
#DEFINE wdPrinterPaperCassette          14
#DEFINE wdPrinterFormSource             15

*  WdOrientation
#DEFINE wdOrientPortrait       0
#DEFINE wdOrientLandscape      1

*  WdSelectionType
#DEFINE wdNoSelection          0
#DEFINE wdSelectionIP          1
#DEFINE wdSelectionNormal      2
#DEFINE wdSelectionFrame       3
#DEFINE wdSelectionColumn      4
#DEFINE wdSelectionRow         5
#DEFINE wdSelectionBlock       6
#DEFINE wdSelectionInlineShape 7
#DEFINE wdSelectionShape       8

*  WdCaptionLabelID
#DEFINE wdCaptionFigure        -1
#DEFINE wdCaptionTable         -2
#DEFINE wdCaptionEquation      -3

*  WdReferenceType
#DEFINE wdRefTypeNumberedItem  0
#DEFINE wdRefTypeHeading       1
#DEFINE wdRefTypeBookmark      2
#DEFINE wdRefTypeFootnote      3
#DEFINE wdRefTypeEndnote       4

*  WdReferenceKind
#DEFINE wdContentText             -1
#DEFINE wdNumberRelativeContext   -2
#DEFINE wdNumberNoContext         -3
#DEFINE wdNumberFullContext       -4
#DEFINE wdEntireCaption           2
#DEFINE wdOnlyLabelAndNumber      3
#DEFINE wdOnlyCaptionText         4
#DEFINE wdFootnoteNumber          5
#DEFINE wdEndnoteNumber           6
#DEFINE wdPageNumber              7
#DEFINE wdPosition                15
#DEFINE wdFootnoteNumberFormatted 16
#DEFINE wdEndnoteNumberFormatted  17

#DEFINE wdIndexTemplate           0
#DEFINE wdIndexClassic            1
#DEFINE wdIndexFancy              2
#DEFINE wdIndexModern             3
#DEFINE wdIndexBulleted           4
#DEFINE wdIndexFormal             5
#DEFINE wdIndexSimple             6
*  WdIndexFormat

*  WdIndexType
#DEFINE wdIndexIndent             0
#DEFINE wdIndexRunin              1

*  WdRevisionsWrap
#DEFINE wdWrapNever               0
#DEFINE wdWrapAlways              1
#DEFINE wdWrapAsk                 2

*  WdRevisionType
#DEFINE wdNoRevision               0
#DEFINE wdRevisionInsert           1
#DEFINE wdRevisionDelete           2
#DEFINE wdRevisionProperty         3
#DEFINE wdRevisionParagraphNumber  4
#DEFINE wdRevisionDisplayField     5
#DEFINE wdRevisionReconcile        6
#DEFINE wdRevisionConflict         7
#DEFINE wdRevisionStyle            8
#DEFINE wdRevisionReplace          9

*  WdRoutingSlipDelivery
#DEFINE wdOneAfterAnother          0
#DEFINE wdAllAtOnce                1

*  WdRoutingSlipStatus
#DEFINE wdNotYetRouted             0
#DEFINE wdRouteInProgress          1
#DEFINE wdRouteComplete            2

*  WdSectionStart
#DEFINE wdSectionContinuous        0
#DEFINE wdSectionNewColumn         1
#DEFINE wdSectionNewPage           2
#DEFINE wdSectionEvenPage          3
#DEFINE wdSectionOddPage           4

*  WdSaveOptions
#DEFINE wdDoNotSaveChanges         0
#DEFINE wdSaveChanges              -1
#DEFINE wdPromptToSaveChanges      -2

*  WdDocumentKind
#DEFINE wdDocumentNotSpecified     0
#DEFINE wdDocumentLetter           1
#DEFINE wdDocumentEmail            2

*  WdDocumentType
#DEFINE wdTypeDocument             0
#DEFINE wdTypeTemplate             1

*  WdOriginalFormat
#DEFINE wdWordDocument             0
#DEFINE wdOriginalDocumentFormat   1
#DEFINE wdPromptUser               2

*  WdRelocate
#DEFINE wdRelocateUp               0
#DEFINE wdRelocateDown             1

*  WdInsertedTextMark
#DEFINE wdInsertedTextMarkNone               0
#DEFINE wdInsertedTextMarkBold               1
#DEFINE wdInsertedTextMarkItalic             2
#DEFINE wdInsertedTextMarkUnderline          3
#DEFINE wdInsertedTextMarkDoubleUnderline    4

*  WdRevisedLinesMark
#DEFINE wdRevisedLinesMarkNone               0
#DEFINE wdRevisedLinesMarkLeftBorder         1
#DEFINE wdRevisedLinesMarkRightBorder        2
#DEFINE wdRevisedLinesMarkOutsideBorder      3

*  WdDeletedTextMark
#DEFINE wdDeletedTextMarkHidden              0
#DEFINE wdDeletedTextMarkStrikeThrough       1
#DEFINE wdDeletedTextMarkCaret               2
#DEFINE wdDeletedTextMarkPound               3

*  WdRevisedPropertiesMark
#DEFINE wdRevisedPropertiesMarkNone              0
#DEFINE wdRevisedPropertiesMarkBold              1
#DEFINE wdRevisedPropertiesMarkItalic            2
#DEFINE wdRevisedPropertiesMarkUnderline         3
#DEFINE wdRevisedPropertiesMarkDoubleUnderline   4

*  WdFieldShading
#DEFINE wdFieldShadingNever            0
#DEFINE wdFieldShadingAlways           1
#DEFINE wdFieldShadingWhenSelected     2

*  WdDefaultFilePath
#DEFINE wdDocumentsPath                0
#DEFINE wdPicturesPath                 1
#DEFINE wdUserTemplatesPath            2
#DEFINE wdWorkgroupTemplatesPath       3
#DEFINE wdUserOptionsPath              4
#DEFINE wdAutoRecoverPath              5
#DEFINE wdToolsPath                    6
#DEFINE wdTutorialPath                 7
#DEFINE wdStartupPath                  8
#DEFINE wdProgramPath                  9
#DEFINE wdGraphicsFiltersPath          10
#DEFINE wdTextConvertersPath           11
#DEFINE wdProofingToolsPath            12
#DEFINE wdTempFilePath                 13
#DEFINE wdCurrentFolderPath            14
#DEFINE wdStyleGalleryPath             15
#DEFINE wdBorderArtPath                19

*  WdCompatibility
#DEFINE wdNoTabHangIndent                        1
#DEFINE wdNoSpaceRaiseLower                      2
#DEFINE wdPrintColBlack                          3
#DEFINE wdWrapTrailSpaces                        4
#DEFINE wdNoColumnBalance                        5
#DEFINE wdConvMailMergeEsc                       6
#DEFINE wdSuppressSpBfAfterPgBrk                 7
#DEFINE wdSuppressTopSpacing                     8
#DEFINE wdOrigWordTableRules                     9
#DEFINE wdTransparentMetafiles                   10
#DEFINE wdShowBreaksInFrames                     11
#DEFINE wdSwapBordersFacingPages                 12
#DEFINE wdLeaveBackslashAlone                    13
#DEFINE wdExpandShiftReturn                      14
#DEFINE wdDontULTrailSpace                       15
#DEFINE wdDontBalanceSingleByteDoubleByteWidth   16
#DEFINE wdSuppressTopSpacingMac5                 17
#DEFINE wdSpacingInWholePoints                   18
#DEFINE wdPrintBodyTextBeforeHeader              19
#DEFINE wdNoLeading                              20
#DEFINE wdNoSpaceForUL                           21
#DEFINE wdMWSmallCaps                            22
#DEFINE wdNoExtraLineSpacing                     23
#DEFINE wdTruncateFontHeight                     24
#DEFINE wdSubFontBySize                          25
#DEFINE wdUsePrinterMetrics                      26
#DEFINE wdWW6BorderRules                         27
#DEFINE wdExactOnTop                             28
#DEFINE wdSuppressBottomSpacing                  29
#DEFINE wdWPSpaceWidth                           30
#DEFINE wdWPJustification                        31
#DEFINE wdLineWrapLikeWord6                      32

*  WdPaperSize
#DEFINE wdPaper10x14                0
#DEFINE wdPaper11x17                1
#DEFINE wdPaperLetter               2
#DEFINE wdPaperLetterSmall          3
#DEFINE wdPaperLegal                4
#DEFINE wdPaperExecutive            5
#DEFINE wdPaperA3                   6
#DEFINE wdPaperA4                   7
#DEFINE wdPaperA4Small              8
#DEFINE wdPaperA5                   9
#DEFINE wdPaperB4                   10
#DEFINE wdPaperB5                   11
#DEFINE wdPaperCSheet               12
#DEFINE wdPaperDSheet               13
#DEFINE wdPaperESheet               14
#DEFINE wdPaperFanfoldLegalGerman   15
#DEFINE wdPaperFanfoldStdGerman     16
#DEFINE wdPaperFanfoldUS            17
#DEFINE wdPaperFolio                18
#DEFINE wdPaperLedger               19
#DEFINE wdPaperNote                 20
#DEFINE wdPaperQuarto               21
#DEFINE wdPaperStatement            22
#DEFINE wdPaperTabloid              23
#DEFINE wdPaperEnvelope9            24
#DEFINE wdPaperEnvelope10           25
#DEFINE wdPaperEnvelope11           26
#DEFINE wdPaperEnvelope12           27
#DEFINE wdPaperEnvelope14           28
#DEFINE wdPaperEnvelopeB4           29
#DEFINE wdPaperEnvelopeB5           30
#DEFINE wdPaperEnvelopeB6           31
#DEFINE wdPaperEnvelopeC3           32
#DEFINE wdPaperEnvelopeC4           33
#DEFINE wdPaperEnvelopeC5           34
#DEFINE wdPaperEnvelopeC6           35
#DEFINE wdPaperEnvelopeC65          36
#DEFINE wdPaperEnvelopeDL           37
#DEFINE wdPaperEnvelopeItaly        38
#DEFINE wdPaperEnvelopeMonarch      39
#DEFINE wdPaperEnvelopePersonal     40
#DEFINE wdPaperCustom               41

*  WdCustomLabelPageSize
#DEFINE wdCustomLabelLetter         0
#DEFINE wdCustomLabelLetterLS       1
#DEFINE wdCustomLabelA4             2
#DEFINE wdCustomLabelA4LS           3
#DEFINE wdCustomLabelA5             4
#DEFINE wdCustomLabelA5LS           5
#DEFINE wdCustomLabelB5             6
#DEFINE wdCustomLabelMini           7
#DEFINE wdCustomLabelFanfold        8

*  WdProtectionType
#DEFINE wdNoProtection              -1
#DEFINE wdAllowOnlyRevisions        0
#DEFINE wdAllowOnlyComments         1
#DEFINE wdAllowOnlyFormFields       2

*  WdPartOfSpeech
#DEFINE wdAdjective       0
#DEFINE wdNoun            1
#DEFINE wdAdverb          2
#DEFINE wdVerb            3

*  WdSubscriberFormats
#DEFINE wdSubscriberBestFormat      0
#DEFINE wdSubscriberRTF             1
#DEFINE wdSubscriberText            2
#DEFINE wdSubscriberPict            4

*  WdEditionType
#DEFINE wdPublisher       0
#DEFINE wdSubscriber      1

*  WdEditionOption
#DEFINE wdCancelPublisher           0
#DEFINE wdSendPublisher             1
#DEFINE wdSelectPublisher           2
#DEFINE wdAutomaticUpdate           3
#DEFINE wdManualUpdate              4
#DEFINE wdChangeAttributes          5
#DEFINE wdUpdateSubscriber          6
#DEFINE wdOpenSource                7

*  WdRelativeHorizontalPosition
#DEFINE wdRelativeHorizontalPositionMargin    0
#DEFINE wdRelativeHorizontalPositionPage      1
#DEFINE wdRelativeHorizontalPositionColumn    2

*  WdRelativeVerticalPosition
#DEFINE wdRelativeVerticalPositionMargin      0
#DEFINE wdRelativeVerticalPositionPage        1
#DEFINE wdRelativeVerticalPositionParagraph   2

*  WdHelpType
#DEFINE wdHelp                      0
#DEFINE wdHelpAbout                 1
#DEFINE wdHelpActiveWindow          2
#DEFINE wdHelpContents              3
#DEFINE wdHelpExamplesAndDemos      4
#DEFINE wdHelpIndex                 5
#DEFINE wdHelpKeyboard              6
#DEFINE wdHelpPSSHelp               7
#DEFINE wdHelpQuickPreview          8
#DEFINE wdHelpSearch                9
#DEFINE wdHelpUsingHelp             10

*  WdHelpTypeHID
#DEFINE wdHelpIchitaro              11
#DEFINE wdHelpPE2                   12

*  WdKeyCategory
#DEFINE wdKeyCategoryNil            -1
#DEFINE wdKeyCategoryDisable        0
#DEFINE wdKeyCategoryCommand        1
#DEFINE wdKeyCategoryMacro          2
#DEFINE wdKeyCategoryFont           3
#DEFINE wdKeyCategoryAutoText       4
#DEFINE wdKeyCategoryStyle          5
#DEFINE wdKeyCategorySymbol         6
#DEFINE wdKeyCategoryPrefix         7

*  WdKey
#DEFINE wdNoKey                  255
#DEFINE wdKeyShift               256
#DEFINE wdKeyControl             512
#DEFINE wdKeyCommand             512
#DEFINE wdKeyAlt                 1024
#DEFINE wdKeyOption              1024
#DEFINE wdKeyA                   65
#DEFINE wdKeyB                   66
#DEFINE wdKeyC                   67
#DEFINE wdKeyD                   68
#DEFINE wdKeyE                   69
#DEFINE wdKeyF                   70
#DEFINE wdKeyG                   71
#DEFINE wdKeyH                   72
#DEFINE wdKeyI                   73
#DEFINE wdKeyJ                   74
#DEFINE wdKeyK                   75
#DEFINE wdKeyL                   76
#DEFINE wdKeyM                   77
#DEFINE wdKeyN                   78
#DEFINE wdKeyO                   79
#DEFINE wdKeyP                   80
#DEFINE wdKeyQ                   81
#DEFINE wdKeyR                   82
#DEFINE wdKeyS                   83
#DEFINE wdKeyT                   84
#DEFINE wdKeyU                   85
#DEFINE wdKeyV                   86
#DEFINE wdKeyW                   87
#DEFINE wdKeyX                   88
#DEFINE wdKeyY                   89
#DEFINE wdKeyZ                   90
#DEFINE wdKey0                   48
#DEFINE wdKey1                   49
#DEFINE wdKey2                   50
#DEFINE wdKey3                   51
#DEFINE wdKey4                   52
#DEFINE wdKey5                   53
#DEFINE wdKey6                   54
#DEFINE wdKey7                   55
#DEFINE wdKey8                   56
#DEFINE wdKey9                   57
#DEFINE wdKeyBackspace           8
#DEFINE wdKeyTab                 9
#DEFINE wdKeyNumeric5Special     12
#DEFINE wdKeyReturn              13
#DEFINE wdKeyPause               19
#DEFINE wdKeyEsc                 27
#DEFINE wdKeySpacebar            32
#DEFINE wdKeyPageUp              33
#DEFINE wdKeyPageDown            34
#DEFINE wdKeyEnd                 35
#DEFINE wdKeyHome                36
#DEFINE wdKeyInsert              45
#DEFINE wdKeyDelete              46
#DEFINE wdKeyNumeric0            96
#DEFINE wdKeyNumeric1            97
#DEFINE wdKeyNumeric2            98
#DEFINE wdKeyNumeric3            99
#DEFINE wdKeyNumeric4            100
#DEFINE wdKeyNumeric5            101
#DEFINE wdKeyNumeric6            102
#DEFINE wdKeyNumeric7            103
#DEFINE wdKeyNumeric8            104
#DEFINE wdKeyNumeric9            105
#DEFINE wdKeyNumericMultiply     106
#DEFINE wdKeyNumericAdd          107
#DEFINE wdKeyNumericSubtract     109
#DEFINE wdKeyNumericDecimal      110
#DEFINE wdKeyNumericDivide       111
#DEFINE wdKeyF1                  112
#DEFINE wdKeyF2                  113
#DEFINE wdKeyF3                  114
#DEFINE wdKeyF4                  115
#DEFINE wdKeyF5                  116
#DEFINE wdKeyF6                  117
#DEFINE wdKeyF7                  118
#DEFINE wdKeyF8                  119
#DEFINE wdKeyF9                  120
#DEFINE wdKeyF10                 121
#DEFINE wdKeyF11                 122
#DEFINE wdKeyF12                 123
#DEFINE wdKeyF13                 124
#DEFINE wdKeyF14                 125
#DEFINE wdKeyF15                 126
#DEFINE wdKeyF16                 127
#DEFINE wdKeyScrollLock          145
#DEFINE wdKeySemiColon           186
#DEFINE wdKeyEquals              187
#DEFINE wdKeyComma               188
#DEFINE wdKeyHyphen              189
#DEFINE wdKeyPeriod              190
#DEFINE wdKeySlash               191
#DEFINE wdKeyBackSingleQuote     192
#DEFINE wdKeyOpenSquareBrace     219
#DEFINE wdKeyBackSlash           220
#DEFINE wdKeyCloseSquareBrace    221
#DEFINE wdKeySingleQuote         222

*  WdOLEType
#DEFINE wdOLELink      0
#DEFINE wdOLEEmbed     1
#DEFINE wdOLEControl   2

*  WdOLEVerb
#DEFINE wdOLEVerbPrimary             0
#DEFINE wdOLEVerbShow                -1
#DEFINE wdOLEVerbOpen                -2
#DEFINE wdOLEVerbHide                -3
#DEFINE wdOLEVerbUIActivate          -4
#DEFINE wdOLEVerbInPlaceActivate     -5
#DEFINE wdOLEVerbDiscardUndoState    -6

*  WdOLEPlacement
#DEFINE wdInLine            0
#DEFINE wdFloatOverText     1

*  WdEnvelopeOrientation
#DEFINE wdLeftPortrait      0
#DEFINE wdCenterPortrait    1
#DEFINE wdRightPortrait     2
#DEFINE wdLeftLandscape     3
#DEFINE wdCenterLandscape   4
#DEFINE wdRightLandscape    5
#DEFINE wdLeftClockwise     6
#DEFINE wdCenterClockwise   7
#DEFINE wdRightClockwise    8

*  WdLetterStyle
#DEFINE wdFullBlock         0
#DEFINE wdModifiedBlock     1
#DEFINE wdSemiBlock         2

*  WdLetterheadLocation
#DEFINE wdLetterTop         0
#DEFINE wdLetterBottom      1
#DEFINE wdLetterLeft        2
#DEFINE wdLetterRight       3

*  WdSalutationType
#DEFINE wdSalutationInformal  0
#DEFINE wdSalutationFormal    1
#DEFINE wdSalutationBusiness  2
#DEFINE wdSalutationOther     3

*  WdSalutationGender
#DEFINE wdGenderFemale     0
#DEFINE wdGenderMale       1
#DEFINE wdGenderNeutral    2
#DEFINE wdGenderUnknown    3

*  WdMovementType
#DEFINE wdMove             0
#DEFINE wdExtend           1

*  WdConstants
#DEFINE wdUn#DEFINEd       9999999
#DEFINE wdToggle           9999998
#DEFINE wdForward          1073741823
#DEFINE wdBackward         -1073741823
#DEFINE wdAutoPosition     0
#DEFINE wdFirst            1
#DEFINE wdCreatorCode      1297307460

*  WdPasteDataType
#DEFINE wdPasteOLEObject                 0
#DEFINE wdPasteRTF                       1
#DEFINE wdPasteText                      2
#DEFINE wdPasteMetafilePicture           3
#DEFINE wdPasteBitmap                    4
#DEFINE wdPasteDeviceIndependentBitmap   5
#DEFINE wdPasteHyperlink                 7
#DEFINE wdPasteShape                     8
#DEFINE wdPasteEnhancedMetafile          9

*  WdPrintOutItem
#DEFINE wdPrintDocumentContent    0
#DEFINE wdPrintProperties         1
#DEFINE wdPrintComments           2
#DEFINE wdPrintStyles             3
#DEFINE wdPrintAutoTextEntries    4
#DEFINE wdPrintKeyAssignments     5
#DEFINE wdPrintEnvelope           6

*  WdPrintOutPages
#DEFINE wdPrintAllPages           0
#DEFINE wdPrintOddPagesOnly       1
#DEFINE wdPrintEvenPagesOnly      2

*  WdPrintOutRange
#DEFINE wdPrintAllDocument        0
#DEFINE wdPrintSelection          1
#DEFINE wdPrintCurrentPage        2
#DEFINE wdPrintFromTo             3
#DEFINE wdPrintRangeOfPages       4

*  WdDictionaryType
#DEFINE wdSpelling                0
#DEFINE wdGrammar                 1
#DEFINE wdThesaurus               2
#DEFINE wdHyphenation             3
#DEFINE wdSpellingComplete        4
#DEFINE wdSpellingCustom          5
#DEFINE wdSpellingLegal           6
#DEFINE wdSpellingMedical         7

*  WdDictionaryTypeHID
#DEFINE wdHangulHanjaConversion         8
#DEFINE wdHangulHanjaConversionCustom   9

*  WdSpellingWordType
#DEFINE wdSpellword           0
#DEFINE wdWildcard            1
#DEFINE wdAnagram             2

*  WdSpellingErrorType
#DEFINE wdSpellingCorrect            0
#DEFINE wdSpellingNotInDictionary    1
#DEFINE wdSpellingCapitalization     2

*  WdProofreadingErrorType
#DEFINE wdSpellingError       0
#DEFINE wdGrammaticalError    1

*  WdInlineShapeType
#DEFINE wdInlineShapeEmbeddedOLEObject  1
#DEFINE wdInlineShapeLinkedOLEObject    2
#DEFINE wdInlineShapePicture            3
#DEFINE wdInlineShapeLinkedPicture      4
#DEFINE wdInlineShapeOLEControlObject   5

*  WdArrangeStyle
#DEFINE wdTiled            0
#DEFINE wdIcons            1

*  WdSelectionFlags
#DEFINE wdSelStartActive   1
#DEFINE wdSelAtEOL         2
#DEFINE wdSelOvertype      4
#DEFINE wdSelActive        8
#DEFINE wdSelReplace       16

*  WdAutoVersions
#DEFINE wdAutoVersionOff       0
#DEFINE wdAutoVersionOnClose   1

*  WdOrganizerObject
#DEFINE wdOrganizerObjectStyles         0
#DEFINE wdOrganizerObjectAutoText       1
#DEFINE wdOrganizerObjectCommandBars    2
#DEFINE wdOrganizerObjectProjectItems   3

*  WdFindMatch
#DEFINE wdMatchParagraphMark       65551
#DEFINE wdMatchTabCharacter        9
#DEFINE wdMatchCommentMark         5
#DEFINE wdMatchAnyCharacter        65599
#DEFINE wdMatchAnyDigit            65567
#DEFINE wdMatchAnyLetter           65583
#DEFINE wdMatchCaretCharacter      11
#DEFINE wdMatchColumnBreak         14
#DEFINE wdMatchEmDash              8212
#DEFINE wdMatchEnDash              8211
#DEFINE wdMatchEndnoteMark         65555
#DEFINE wdMatchField               19
#DEFINE wdMatchFootnoteMark        65554
#DEFINE wdMatchGraphic             1
#DEFINE wdMatchManualLineBreak     65551
#DEFINE wdMatchManualPageBreak     65564
#DEFINE wdMatchNonbreakingHyphen   30
#DEFINE wdMatchNonbreakingSpace    160
#DEFINE wdMatchOptionalHyphen      31
#DEFINE wdMatchSectionBreak        65580
#DEFINE wdMatchWhiteSpace          65655

*  WdFindWrap
#DEFINE wdFindStop       0
#DEFINE wdFindContinue   1
#DEFINE wdFindAsk        2

*  WdInformation
#DEFINE wdActiveEndAdjustedPageNumber                 1
#DEFINE wdActiveEndSectionNumber                      2
#DEFINE wdActiveEndPageNumber                         3
#DEFINE wdNumberOfPagesInDocument                     4
#DEFINE wdHorizontalPositionRelativeToPage            5
#DEFINE wdVerticalPositionRelativeToPage              6
#DEFINE wdHorizontalPositionRelativeToTextBoundary    7
#DEFINE wdVerticalPositionRelativeToTextBoundary      8
#DEFINE wdFirstCharacterColumnNumber                  9
#DEFINE wdFirstCharacterLineNumber                    10
#DEFINE wdFrameIsSelected                             11
#DEFINE wdWithInTable                                 12
#DEFINE wdStartOfRangeRowNumber                       13
#DEFINE wdEndOfRangeRowNumber                         14
#DEFINE wdMaximumNumberOfRows                         15
#DEFINE wdStartOfRangeColumnNumber                    16
#DEFINE wdEndOfRangeColumnNumber                      17
#DEFINE wdMaximumNumberOfColumns                      18
#DEFINE wdZoomPercentage                              19
#DEFINE wdSelectionMode                               20
#DEFINE wdCapsLock                                    21
#DEFINE wdNumLock                                     22
#DEFINE wdOverType                                    23
#DEFINE wdRevisionMarking                             24
#DEFINE wdInFootnoteEndnotePane                       25
#DEFINE wdInCommentPane                               26
#DEFINE wdInHeaderFooter                              28
#DEFINE wdAtEndOfRowMarker                            31
#DEFINE wdReferenceOfType                             32
#DEFINE wdHeaderFooterType                            33
#DEFINE wdInMasterDocument                            34
#DEFINE wdInFootnote                                  35
#DEFINE wdInEndnote                                   36
#DEFINE wdInWordMail                                  37
#DEFINE wdInClipboard                                 38

*  WdWrapType
#DEFINE wdWrapSquare          0
#DEFINE wdWrapTight           1
#DEFINE wdWrapThrough         2
#DEFINE wdWrapNone            3
#DEFINE wdWrapTopBottom       4

*  WdWrapSideType
#DEFINE wdWrapBoth            0
#DEFINE wdWrapLeft            1
#DEFINE wdWrapRight           2
#DEFINE wdWrapLargest         3

*  WdOutlineLevel
#DEFINE wdOutlineLevel1            1
#DEFINE wdOutlineLevel2            2
#DEFINE wdOutlineLevel3            3
#DEFINE wdOutlineLevel4            4
#DEFINE wdOutlineLevel5            5
#DEFINE wdOutlineLevel6            6
#DEFINE wdOutlineLevel7            7
#DEFINE wdOutlineLevel8            8
#DEFINE wdOutlineLevel9            9
#DEFINE wdOutlineLevelBodyText     10

*  WdTextOrientation
#DEFINE wdTextOrientationHorizontal   0
#DEFINE wdTextOrientationUpward       2
#DEFINE wdTextOrientationDownward     3

*  WdTextOrientationHID
#DEFINE wdTextOrientationVerticalFarEast            1
#DEFINE wdTextOrientationHorizontalRotatedFarEast   4

*  WdPageBorderArt
#DEFINE wdArtApples              1
#DEFINE wdArtMapleMuffins        2
#DEFINE wdArtCakeSlice           3
#DEFINE wdArtCandyCorn           4
#DEFINE wdArtIceCreamCones       5
#DEFINE wdArtChampagneBottle     6
#DEFINE wdArtPartyGlass          7
#DEFINE wdArtChristmasTree       8
#DEFINE wdArtTrees               9
#DEFINE wdArtPalmsColor          10
#DEFINE wdArtBalloons3Colors     11
#DEFINE wdArtBalloonsHotAir      12
#DEFINE wdArtPartyFavor          13
#DEFINE wdArtConfettiStreamers   14
#DEFINE wdArtHearts              15
#DEFINE wdArtHeartBalloon        16
#DEFINE wdArtStars3D             17
#DEFINE wdArtStarsShadowed       18
#DEFINE wdArtStars               19
#DEFINE wdArtSun                 20
#DEFINE wdArtEarth2              21
#DEFINE wdArtEarth1              22
#DEFINE wdArtPeopleHats          23
#DEFINE wdArtSombrero            24
#DEFINE wdArtPencils             25
#DEFINE wdArtPackages            26
#DEFINE wdArtClocks              27
#DEFINE wdArtFirecrackers        28
#DEFINE wdArtRings               29
#DEFINE wdArtMapPins             30
#DEFINE wdArtConfetti            31
#DEFINE wdArtCreaturesButterfly  32
#DEFINE wdArtCreaturesLadyBug    33
#DEFINE wdArtCreaturesFish       34
#DEFINE wdArtBirdsFlight         35
#DEFINE wdArtScaredCat           36
#DEFINE wdArtBats                37
#DEFINE wdArtFlowersRoses        38
#DEFINE wdArtFlowersRedRose      39
#DEFINE wdArtPoinsettias         40
#DEFINE wdArtHolly               41
#DEFINE wdArtFlowersTiny         42
#DEFINE wdArtFlowersPansy        43
#DEFINE wdArtFlowersModern2      44
#DEFINE wdArtFlowersModern1      45
#DEFINE wdArtWhiteFlowers        46
#DEFINE wdArtVine                47
#DEFINE wdArtFlowersDaisies      48
#DEFINE wdArtFlowersBlockPrint   49
#DEFINE wdArtDecoArchColor       50
#DEFINE wdArtFans                51
#DEFINE wdArtFilm                52
#DEFINE wdArtLightning1          53
#DEFINE wdArtCompass             54
#DEFINE wdArtDoubleD             55
#DEFINE wdArtClassicalWave       56
#DEFINE wdArtShadowedSquares     57
#DEFINE wdArtTwistedLines1       58
#DEFINE wdArtWaveline            59
#DEFINE wdArtQuadrants           60
#DEFINE wdArtCheckedBarColor     61
#DEFINE wdArtSwirligig           62
#DEFINE wdArtPushPinNote1        63
#DEFINE wdArtPushPinNote2        64
#DEFINE wdArtPumpkin1            65
#DEFINE wdArtEggsBlack           66
#DEFINE wdArtCup                 67
#DEFINE wdArtHeartGray           68
#DEFINE wdArtGingerbreadMan      69
#DEFINE wdArtBabyPacifier        70
#DEFINE wdArtBabyRattle          71
#DEFINE wdArtCabins              72
#DEFINE wdArtHouseFunky          73
#DEFINE wdArtStarsBlack          74
#DEFINE wdArtSnowflakes          75
#DEFINE wdArtSnowflakeFancy      76
#DEFINE wdArtSkyrocket           77
#DEFINE wdArtSeattle             78
#DEFINE wdArtMusicNotes          79
#DEFINE wdArtPalmsBlack          80
#DEFINE wdArtMapleLeaf           81
#DEFINE wdArtPaperClips          82
#DEFINE wdArtShorebirdTracks     83
#DEFINE wdArtPeople              84
#DEFINE wdArtPeopleWaving        85
#DEFINE wdArtEclipsingSquares2   86
#DEFINE wdArtHypnotic            87
#DEFINE wdArtDiamondsGray        88
#DEFINE wdArtDecoArch            89
#DEFINE wdArtDecoBlocks          90
#DEFINE wdArtCirclesLines        91
#DEFINE wdArtPapyrus             92
#DEFINE wdArtWoodwork            93
#DEFINE wdArtWeavingBraid        94
#DEFINE wdArtWeavingRibbon       95
#DEFINE wdArtWeavingAngles       96
#DEFINE wdArtArchedScallops      97
#DEFINE wdArtSafari              98
#DEFINE wdArtCelticKnotwork      99
#DEFINE wdArtCrazyMaze           100
#DEFINE wdArtEclipsingSquares1   101
#DEFINE wdArtBirds               102
#DEFINE wdArtFlowersTeacup       103
#DEFINE wdArtNorthwest           104
#DEFINE wdArtSouthwest           105
#DEFINE wdArtTribal6             106
#DEFINE wdArtTribal4             107
#DEFINE wdArtTribal3             108
#DEFINE wdArtTribal2             109
#DEFINE wdArtTribal5             110
#DEFINE wdArtXIllusions          111
#DEFINE wdArtZanyTriangles       112
#DEFINE wdArtPyramids            113
#DEFINE wdArtPyramidsAbove       114
#DEFINE wdArtConfettiGrays       115
#DEFINE wdArtConfettiOutline     116
#DEFINE wdArtConfettiWhite       117
#DEFINE wdArtMosaic              118
#DEFINE wdArtLightning2          119
#DEFINE wdArtHeebieJeebies       120
#DEFINE wdArtLightBulb           121
#DEFINE wdArtGradient            122
#DEFINE wdArtTriangleParty       123
#DEFINE wdArtTwistedLines2       124
#DEFINE wdArtMoons               125
#DEFINE wdArtOvals               126
#DEFINE wdArtDoubleDiamonds      127
#DEFINE wdArtChainLink           128
#DEFINE wdArtTriangles           129
#DEFINE wdArtTribal1             130
#DEFINE wdArtMarqueeToothed      131
#DEFINE wdArtSharksTeeth         132
#DEFINE wdArtSawtooth            133
#DEFINE wdArtSawtoothGray        134
#DEFINE wdArtPostageStamp        135
#DEFINE wdArtWeavingStrips       136
#DEFINE wdArtZigZag              137
#DEFINE wdArtCrossStitch         138
#DEFINE wdArtGems                139
#DEFINE wdArtCirclesRectangles   140
#DEFINE wdArtCornerTriangles     141
#DEFINE wdArtCreaturesInsects    142
#DEFINE wdArtZigZagStitch        143
#DEFINE wdArtCheckered           144
#DEFINE wdArtCheckedBarBlack     145
#DEFINE wdArtMarquee             146
#DEFINE wdArtBasicWhiteDots      147
#DEFINE wdArtBasicWideMidline    148
#DEFINE wdArtBasicWideOutline    149
#DEFINE wdArtBasicWideInline     150
#DEFINE wdArtBasicThinLines      151
#DEFINE wdArtBasicWhiteDashes    152
#DEFINE wdArtBasicWhiteSquares   153
#DEFINE wdArtBasicBlackSquares   154
#DEFINE wdArtBasicBlackDashes    155
#DEFINE wdArtBasicBlackDots      156
#DEFINE wdArtStarsTop            157
#DEFINE wdArtCertificateBanner   158
#DEFINE wdArtHandmade1           159
#DEFINE wdArtHandmade2           160
#DEFINE wdArtTornPaper           161
#DEFINE wdArtTornPaperBlack      162
#DEFINE wdArtCouponCutoutDashes  163
#DEFINE wdArtCouponCutoutDots    164

*  WdBorderDistanceFrom
#DEFINE wdBorderDistanceFromText       0
#DEFINE wdBorderDistanceFromPageEdge   1

*  WdReplace
#DEFINE wdReplaceNone         0
#DEFINE wdReplaceOne          1
#DEFINE wdReplaceAll          2

*  WdFontBias
#DEFINE wdFontBiasDontCare    255
#DEFINE wdFontBiasDefault     0
#DEFINE wdFontBiasFareast     1

