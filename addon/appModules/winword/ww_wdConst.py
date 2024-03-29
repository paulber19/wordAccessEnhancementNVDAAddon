# appModules\winword\ww_wdConst.py
# A part of WordAccessEnhancement add-on
# Copyright (C) 2019-2022 paulber19
# This file is covered by the GNU General Public License.

import addonHandler
addonHandler.initTranslation()

# wdUnits
wdUndefined = 9999999
# Specifies the unit of measure to use.
wdChar = 1  # a character
wdWord = 2  # A word
wdSentence = 3  # A sentence
wdParagraph = 4  # A paragraph
wdLine = 5  # A line
wdStory = 6  # A story
wdScreen = 7  # The screen dimensions
wdSection = 8  # A section
wdColumn = 9  # A column
wdRow = 10  # A row
wdWindow = 11  # A window
wdCell = 12  # A cell
wdCharacterFormattingCharacter = 13  # formatting
wdParagraphFormatting = 14  # Paragraph formatting
wdTable = 15  # A table
wdItem = 16  # The selected item

# wdGoToItem
# Specifies the type of item to move the insertion point or selection just prior to.

wdGoToBookmark = -1  # A bookmark
wdGoToSection = 0  # A section
wdGoToPage = 1  # A page
wdGoToTable = 2  # A table
wdGoToLine = 3  # A line
wdGoToFootnote = 4  # A footnote
wdGoToEndnote = 5  # An endnote
wdGoToComment = 6  # A comment
wdGoToField = 7  # A field
wdGoToGraphic = 8  # A graphic
wdGoToObject = 9  # An object
wdGoToEquation = 10  # An equation
wdGoToHeading = 11  # A heading
wdGoToPercent = 12  # A percent
wdGoToSpellingError = 13  # A spelling error
wdGoToGrammaticalError = 14  # A grammatical error
wdGoToProofReadingError = 15  # ErrorA proof reading error
# added but not surely
wdGoToRevision = 16  # a revision
wdGoToHyperlink = 17
wdGoToFrame = 18

# wdGoToDirection
# Specifies the position to which a selection or the insertion point is moved in relation to an object
# or to itself.

wdGoToAbsolute = 1  # An absolute position
wdGoToRelative = 2  # The position relative to the current position
wdGoToLast = -1  # The last instance of the specified object
wdGoToFirst = 1  # The first instance of the specified object
wdGoToNext = 2  # The next instance of the specified object
wdGoToPrevious = 3  # The previous instance of the specified object

# WdInformation Enumeration
# Specifies the type of information returned about a specified selection or range.
wdActiveEndAdjustedPageNumber = 1  # returns the number of the page that contains the active end
# of the specified selection or range.
wdActiveEndSectionNumber = 2  # Returns the number of the section that contains the active end
# of the specified selection or range.
wdActiveEndPageNumber = 3  # Returns the number of the page that contains the active end
# of the specified selection or range,
wdNumberOfPagesInDocument = 4  # Returns the number of pages in the document associated with
# the selection or range.
wdHorizontalPositionRelativeToPage = 5  # Returns the horizontal position of the specified selection or range;
wdVerticalPositionRelativeToPage = 6  # Returns the vertical position of the selection or range
wdHorizontalPositionRelativeToTextBoundary = 7  # Returns the horizontal position of the specified
# selection or range relative to the left
wdVerticalPositionRelativeToTextBoundary = 8  # Returns the vertical position of the selection or range
# relative to the top edge of the nearest text boundary enclosing it
wdFirstCharacterColumnNumber = 9  # Returns the character position of the first character
# in the specified selection or range.
wdFirstCharacterLineNumber = 10  # Returns the character position of the first character
# in the specified selection or range. If the selection
wdFrameIsSelected = 11  # Returns True if the selection or range is an entire frame or text box.
wdWithInTable = 12  # Returns True if the selection is in a table.
wdStartOfRangeRowNumber = 13  # Returns the table row number that contains the beginning
# of the selection or range.
wdEndOfRangeRowNumber = 14  # Returns the table row number that contains the end of
# the specified selection or range.
wdMaximumNumberOfRows = 15  # Returns the greatest number of table rows within the table in
# the specified selection or range.
wdStartOfRangeColumnNumber = 16  # Returns the table column number that contains the beginning of
# the selection or range.
wdEndOfRangeColumnNumber = 17  # Returns the table column number that contains the end
# of the specified selection
# or range.
wdMaximumNumberOfColumns = 18  # Returns the greatest number of table columns within any row
# in the selection or range.
wdZoomPercentage = 19  # Returns the current percentage of magnification as set by the Percentageproperty.
wdSelectionMode = 20  # Returns a value that indicates the current selection mode,
# as shown in the following table.
wdCapsLock = 21  # Returns True if Caps Lock is in effect.
wdNumLock = 22  # Returns True if Num Lock is in effect.
wdOverType = 23  # Returns True if Overtype mode is in effect.
# The Overtype property can be used to change the state of the Overtype mode.
wdRevisionMarking = 24  # Returns True if change tracking is in effect.
wdInFootnoteEndnotePane = 25  # Returns True if the specified selection or range is
# in the footnote or endnote pane
wdInCommentPane = 26  # Returns True if the specified selection or range is in a comment pane.
wdInHeaderFooter = 28  # Returns True if the selection or range is in the header or footer pane or
# in a header or footer in print layout view.
wdAtEndOfRowMarker = 31  # Returns True if the specified selection or range is at the end-of-row mark
# in a table.
wdReferenceOfType = 32  # Returns a value that indicates where the selection is in relation to a footnote,
# endnote, or
# comment reference,
wdHeaderFooterType = 33  # Returns a value that indicates the type of header
# or footer that contains the specified selection or range.
wdInMasterDocument = 34  # Returns True if the selection or range is in a master document
wdInFootnote = 35  # Returns True if the specified selection or range is in a footnote area
wdInEndnote = 36  # Returns True if the specified selection or range is in an endnote area
wdInWordMail = 37  # Returns True if the selection or range is in the header or footer pane
# or in a header or footer in print layout view.
wdInClipboard = 38

wdStoryValueToText = {
	1: "Main text story",  # wdMainTextStory
	2: "Footnotes story",  # wdFootnotesStory
	3: "Endnotes story",  # wdEndnotesStory
	4: "Comments story",  # wdCommentsStory
	5: "Text frame story",  # wdTextFrameStory
	6: "Even pages header story",  # wdEvenPagesHeaderStory
	7: "Primary header story",  # wdPrimaryHeaderStory
	8: "Even pages footer story",  # wdEvenPagesFooterStory
	9: "Primary footer story",  # wdPrimaryFooterStory
	10: "First page header story",  # wdFirstPageHeaderStory
	11: "First page footer story",  # wdFirstPageFooterStory
	12: "Footnote separator story",  # wdFootnoteSeparatorStory
	13: "Footnote continuation separator  story",  # wdFootnoteContinuationSeparatorStory
	14: "Footnote continuation notice story",  # wdFootnoteContinuationNoticeStory
	15: "Endnote separator story",  # wdEndnoteSeparatorStory
	16: "Endnote continuation separator story",  # wdEndnoteContinuationSeparatorStory
	17: "Endnote continuation notice story",  # wdEndnoteContinuationNoticeStory")
}

# wdStoryType
wdMainTextStory = 1  # Main text story.
wdFootnotesStory = 2  # Footnotes story.
wdEndnotesStory = 3  # Endnotes story.
wdCommentsStory = 4  # Comments story.
wdTextFrameStory = 5  # Text frame story.
wdEvenPagesHeaderStory = 6  # Even pages header story.
wdPrimaryHeaderStory = 7  # Primary header story.
wdEvenPagesFooterStory = 8  # Even pages footer story.
wdPrimaryFooterStory = 9  # Primary footer story.
wdFirstPageHeaderStory = 10  # First page header story.
wdFirstPageFooterStory = 11  # First page footer story.
wdFootnoteSeparatorStory = 12  # Footnote separator story.
wdFootnoteContinuationSeparatorStory = 13  # Footnote continuation separator story.
wdFootnoteContinuationNoticeStory = 14  # Footnote continuation notice story.
wdEndnoteSeparatorStory = 15  # Endnote separator story.
wdEndnoteContinuationSeparatorStory = 16  # Endnote continuation separator story.
wdEndnoteContinuationNoticeStory = 17  # Endnote continuation notice story

# wdHeaderFooterIndex
wdHeaderFooterPrimary = 1
wdHeaderFooterFirstPage = 2
wdHeaderFooterEvenPages = 3

# WdStatistic
wdStatisticWords = 0  # Count of words.
wdStatisticLines = 1  # Count of lines.
wdStatisticPages = 2  # Count of pages.
wdStatisticCharacters = 3  # Count of characters.
wdStatisticParagraphs = 4  # Count of paragraphs.
wdStatisticCharactersWithSpaces = 5  # Count of characters including spaces.
wdStatisticFarEastCharacters = 6  # Count of characters for Asian languages.

# view types
wdNormalView = 1
wdOutlineView = 2
wdPrintView = 3
wdPrintPreview = 4
wdMasterView = 5
wdWebView = 6
wdReadingView = 7

# protection type
wdNoProtection = -1
wdAllowOnlyRevisions = 0
wdAllowOnlyComments = 1
wdAllowOnlyFormFields = 2
wdAllowOnlyReading = 3

# border types
wdBorderTop = -1
wdBorderLeft = -2
wdBorderBottom = -3
wdBorderRight = -4
wdBorderHorizontal = -5
wdBorderVertical = -6
wdBorderDiagonalDown = -7
wdBorderDiagonalUp = -8

# WDLineStyle constants
wdLineStyleNone = 0
wdLineStyleSingle = 1
wdLineStyleDot = 2
wdLineStyleDashSmallGap = 3
wdLineStyleDashLargeGap = 4
wdLineStyleDashDot = 5
wdLineStyleDashDotDot = 6
wdLineStyleDouble = 7
wdLineStyleTriple = 8
wdLineStyleThinThickSmallGap = 9
wdLineStyleThickThinSmallGap = 10
wdLineStyleThinThickThinSmallGap = 11
wdLineStyleThinThickMedGap = 12
wdLineStyleThickThinMedGap = 13
wdLineStyleThinThickThinMedGap = 14
wdLineStyleThinThickLargeGap = 15
wdLineStyleThickThinLargeGap = 16
wdLineStyleThinThickThinLargeGap = 17
wdLineStyleSingleWavy = 18
wdLineStyleDoubleWavy = 19
wdLineStyleDashDotStroked = 20
wdLineStyleEmboss3D = 21
wdLineStyleEngrave3D = 22
wdLineStyleOutset = 23
wdLineStyleInset = 24


# wdColor constants
wdColorAqua = 0xCCCC33  # 13421619
wdColorAutomatic = 0xFF000000  # -16777216
wdColorBlack = 0
wdColorBlue = 0xFF0000  # 16711680
wdColorBlueGray = 0x996666  # 10053222
wdColorBrightGreen = 0xFF00  # 65280
wdColorBrown = 0x3399  # 13209
wdColorDarkBlue = 0x800000  # 8388608
wdColorDarkGreen = 0x3300  # 13056
wdColorDarkRed = 0x80  # 128
wdColorDarkTeal = 0x663300  # 6697728
wdColorDarkYellow = 0x8080  # 32896
wdColorGold = 0xCCFF   # 52479
wdColorGray05 = 0xF3F3F3  # 15987699
wdColorGray10 = 0xE6E6E6  # 15132390
wdColorGray125 = 0xE0E0E0  # 14737632
wdColorGray15 = 0xD9D9D9  # 14277081
wdColorGray20 = 0xCCCCCC  # 13421772
wdColorGray25 = 0xC0C0C0  # 12632256
wdColorGray30 = 0xB3B3B3  # 11776947
wdColorGray35 = 0xA6A6A6  # 10921638
wdColorGray375 = 0xA0A0A0  # 10526880
wdColorGray40 = 0x999999  # 10066329
wdColorGray45 = 0x8C8C8C  # 9211020
wdColorGray50 = 0x808080  # 8421504
wdColorGray55 = 0x737373  # 7566195
wdColorGray60 = 0x666666  # 6710886
wdColorGray625 = 0x606060  # 6316128
wdColorGray65 = 0x595959  # 5855577
wdColorGray70 = 0x4C4C4C  # 5000268
wdColorGray75 = 0x404040  # 4210752
wdColorGray80 = 0x333333  # 3355443,
wdColorGray85 = 0x262626  # 2500134
wdColorGray875 = 0x202020  # 2105376
wdColorGray90 = 0x191919  # 1644825
wdColorGray95 = 0xC0C0C  # 789516
wdColorGreen = 0x8000  # 32768
wdColorIndigo = 0x993333  # 10040115
wdColorLavender = 0xFF99CC  # 16751052
wdColorLightBlue = 0xFF6633  # 16737843
wdColorLightGreen = 0xCCFFCC  # 13434828
wdColorLightOrange = 0x99FF  # 39423
wdColorLightTurquoise = 0xFFFFCC  # 16777164
wdColorLightYellow = 0x99FFFF  # 10092543
wdColorLime = 0xCC99  # 52377
wdColorOliveGreen = 0x3333  # 13107
wdColorOrange = 0x66FF  # 26367
wdColorPaleBlue = 0xFFCC99  # 16764057
wdColorPink = 0xFF00FF  # 16711935
wdColorPlum = 0x663399  # 6697881
wdColorRed = 0xFF  # 255
wdColorRose = 0xCC99FF  # 13408767
wdColorSeaGreen = 0x669933  # 6723891
wdColorSkyBlue = 0xFFCC00  # 16763904
wdColorTan = 0x99CCFF  # 10079487
wdColorTeal = 0x808000  # 8421376
wdColorTurquoise = 0xFFFF00  # 16776960,
wdColorViolet = 0x800080  # 8388736,
wdColorWhite = 0xFFFFFF  # 16777215,
wdColorYellow = 0xFFFF  # 65535,
wdColorIndexAuto = 0,
wdColorIndexBlack = 1,
wdColorIndexBlue = 2,
wdColorIndexBrightGreen = 4,
wdColorIndexByAuthor = 0xffffffff,
wdColorIndexDarkBlue = 9,
wdColorIndexDarkRed = 13,
wdColorIndexDarkYellow = 14,
wdColorIndexGray25 = 16,
wdColorIndexGray50 = 15,
wdColorIndexGreen = 11,
wdColorIndexNoHighlight = 0,
wdColorIndexPink = 5,
wdColorIndexRed = 6,
wdColorIndexTeal = 10,
wdColorIndexTurquoise = 3,
wdColorIndexViolet = 12,
wdColorIndexWhite = 8,
wdColorIndexYellow = 7,

# wdLineWidth constants
wdLineWidth025pt = 2
wdLineWidth050pt = 4
wdLineWidth075pt = 6
wdLineWidth100pt = 8
wdLineWidth150pt = 12
wdLineWidth225pt = 18
wdLineWidth300pt = 24
wdLineWidth450pt = 36
wdLineWidth600pt = 48

# wdPageArt
wdArtApples = 1
wdArtArchedScallops = 97
wdArtBabyPacifier = 70
wdArtBabyRattle = 71
wdArtBalloons3Colors = 11
wdArtBalloonsHotAir = 12
wdArtBasicBlackDashes = 155
wdArtBasicBlackDots = 156
wdArtBasicBlackSquares = 154
wdArtBasicThinLines = 151
wdArtBasicWhiteDashes = 152
wdArtBasicWhiteDots = 147
wdArtBasicWhiteSquares = 153
wdArtBasicWideInline = 150
wdArtBasicWideMidline = 148
wdArtBasicWideOutline = 149
wdArtBats = 37
wdArtBirds = 102
wdArtBirdsFlight = 35
wdArtCabins = 72
wdArtCakeSlice = 3
wdArtCandyCorn = 4
wdArtCelticKnotwork = 99
wdArtCertificateBanner = 158
wdArtChainLink = 128
wdArtChampagneBottle = 6
wdArtCheckedBarBlack = 145
wdArtCheckedBarColor = 61
wdArtCheckered = 144
wdArtChristmasTree = 8
wdArtCirclesLines = 91
wdArtCirclesRectangles = 140
wdArtClassicalWave = 56
wdArtClocks = 27
wdArtCompass = 54
wdArtConfetti = 31
wdArtConfettiGrays = 115
wdArtConfettiOutline = 116
wdArtConfettiStreamers = 14
wdArtConfettiWhite = 117
wdArtCornerTriangles = 141
wdArtCouponCutoutDashes = 163
wdArtCouponCutoutDots = 164
wdArtCrazyMaze = 100
wdArtCreaturesButterfly = 32
wdArtCreaturesFish = 34
wdArtCreaturesInsects = 142
wdArtCreaturesLadyBug = 33
wdArtCrossStitch = 138
wdArtCup = 67
wdArtDecoArch = 89
wdArtDecoArchColor = 50
wdArtDecoBlocks = 90
wdArtDiamondsGray = 88
wdArtDoubleD = 55
wdArtDoubleDiamonds = 127
wdArtEarth1 = 22
wdArtEarth2 = 21
wdArtEclipsingSquares1 = 101
wdArtEclipsingSquares2 = 86
wdArtEggsBlack = 66
wdArtFans = 51
wdArtFilm = 52
wdArtFirecrackers = 28
wdArtFlowersBlockPrint = 49
wdArtFlowersDaisies = 48
wdArtFlowersModern1 = 45
wdArtFlowersModern2 = 44
wdArtFlowersPansy = 43
wdArtFlowersRedRose = 39
wdArtFlowersRoses = 38
wdArtFlowersTeacup = 103
wdArtFlowersTiny = 42
wdArtGems = 139
wdArtGingerbreadMan = 69
wdArtGradient = 122
wdArtHandmade1 = 159
wdArtHandmade2 = 160
wdArtHeartBalloon = 16
wdArtHeartGray = 68
wdArtHearts = 15
wdArtHeebieJeebies = 120
wdArtHolly = 41
wdArtHouseFunky = 73
wdArtHypnotic = 87
wdArtIceCreamCones = 5
wdArtLightBulb = 121
wdArtLightning1 = 53
wdArtLightning2 = 119
wdArtMapleLeaf = 81
wdArtMapleMuffins = 2
wdArtMapPins = 30
wdArtMarquee = 146
wdArtMarqueeToothed = 131
wdArtMoons = 125
wdArtMosaic = 118
wdArtMusicNotes = 79
wdArtNorthwest = 104
wdArtOvals = 126
wdArtPackages = 26
wdArtPalmsBlack = 80
wdArtPalmsColor = 10
wdArtPaperClips = 82
wdArtPapyrus = 92
wdArtPartyFavor = 13
wdArtPartyGlass = 7
wdArtPencils = 25
wdArtPeople = 84
wdArtPeopleHats = 23
wdArtPeopleWaving = 85
wdArtPoinsettias = 40
wdArtPostageStamp = 135
wdArtPumpkin1 = 65
wdArtPushPinNote1 = 63
wdArtPushPinNote2 = 64
wdArtPyramids = 113
wdArtPyramidsAbove = 114
wdArtQuadrants = 60
wdArtRings = 29
wdArtSafari = 98
wdArtSawtooth = 133
wdArtSawtoothGray = 134
wdArtScaredCat = 36
wdArtSeattle = 78
wdArtShadowedSquares = 57
wdArtSharksTeeth = 132
wdArtShorebirdTracks = 83
wdArtSkyrocket = 77
wdArtSnowflakeFancy = 76
wdArtSnowflakes = 75
wdArtSombrero = 24
wdArtSouthwest = 105
wdArtStars = 19
wdArtStars3D = 17
wdArtStarsBlack = 74
wdArtStarsShadowed = 18
wdArtStarsTop = 157
wdArtSun = 20
wdArtSwirligig = 62
wdArtTornPaper = 161
wdArtTornPaperBlack = 162
wdArtTrees = 9
wdArtTriangleParty = 123
wdArtTriangles = 129
wdArtTribal1 = 130
wdArtTribal2 = 109
wdArtTribal3 = 108
wdArtTribal4 = 107
wdArtTribal5 = 110
wdArtTribal6 = 106
wdArtTwistedLines1 = 58
wdArtTwistedLines2 = 124
wdArtVine = 47
wdArtWaveline = 59
wdArtWeavingAngles = 96
wdArtWeavingBraid = 94
wdArtWeavingRibbon = 95
wdArtWeavingStrips = 136
wdArtWhiteFlowers = 46
wdArtWoodwork = 93
wdArtXIllusions = 111
wdArtZanyTriangles = 112
wdArtZigZag = 137
wdArtZigZagStitch = 143

# wdOrientations
wdOrientLandscape = 1
wdOrientPortrait = 0

# WdVerticalAlignment
wdAlignVerticalTop = 0
wdAlignVerticalCenter = 1
wdAlignVerticalJustify = 2
wdAlignVerticalBottom = 3

# wdPaperSize
wdPaper10x14 = 0  # 10 inches wide, 14 inches long.
wdPaper11x17 = 1  # Legal 11 inches wide, 17 inches long.
wdPaperA3 = 6  # A3 dimensions.
wdPaperA4 = 7  # A4 dimensions.
wdPaperA4Small = 8  # Small A4 dimensions.
wdPaperA5 = 9  # A5 dimensions.
wdPaperB4 = 10  # B4 dimensions.
wdPaperB5 = 11  # B5 dimensions.
wdPaperCSheet = 12  # C sheet dimensions.
wdPaperCustom = 41  # Custom paper size.
wdPaperDSheet = 13  # D sheet dimensions.
wdPaperEnvelope10 = 25  # Legal envelope, size 10.
wdPaperEnvelope11 = 26  # Envelope, size 11.
wdPaperEnvelope12 = 27  # Envelope, size 12.
wdPaperEnvelope14 = 28  # Envelope, size 14.
wdPaperEnvelope9 = 24  # Envelope, size 9.
wdPaperEnvelopeB4 = 29  # B4 envelope.
wdPaperEnvelopeB5 = 30  # B5 envelope.
wdPaperEnvelopeB6 = 31  # B6 envelope.
wdPaperEnvelopeC3 = 32  # C3 envelope.
wdPaperEnvelopeC4 = 33  # C4 envelope.
wdPaperEnvelopeC5 = 34  # C5 envelope.
wdPaperEnvelopeC6 = 35  # C6 envelope.
wdPaperEnvelopeC65 = 36  # C65 envelope.
wdPaperEnvelopeDL = 37  # DL envelope.
wdPaperEnvelopeItaly = 38  # Italian envelope.
wdPaperEnvelopeMonarch = 39  # Monarch envelope.
wdPaperEnvelopePersonal = 40  # Personal envelope.
wdPaperESheet = 14  # E sheet dimensions.
wdPaperExecutive = 5  # Executive dimensions.
wdPaperFanfoldLegalGerman = 15  # German legal fanfold dimensions.
wdPaperFanfoldStdGerman = 16  # German standard fanfold dimensions.
wdPaperFanfoldUS = 17  # United States fanfold dimensions.
wdPaperFolio = 18  # Folio dimensions.
wdPaperLedger = 19  # Ledger dimensions.
wdPaperLegal = 4  # Legal dimensions.
wdPaperLetter = 2  # Letter dimensions.
wdPaperLetterSmall = 3  # Small letter dimensions.
wdPaperNote = 20  # Note dimensions.
wdPaperQuarto = 21  # Quarto dimensions.
wdPaperStatement = 22  # Statement dimensions.
wdPaperTabloid = 23  # Tabloid dimensions.

# wdGutterStyle
wdGutterPosLeft = 0  # On the left side.
wdGutterPosRight = 2  # On the right side.
wdGutterPosTop = 1  # At the top.

# wdGutterStyleOld
wdGutterStyleBidi = 2  # Bidirectional gutter
wdGutterStyleLatin = -10  # Latin gutter

# wdSectionStart
wdSectionContinuous = 0  # Continuous section break.
wdSectionEvenPage = 3  # Even pages section break.
wdSectionNewColumn = 1  # New column section break.
wdSectionNewPage = 2  # New page section break.
wdSectionOddPage = 4  # Odd pages section break.
# wd movement
wdMove = 0
wdExtend = 1
wdCollapseEnd = 0
wdCollapseStar = 1
# WDCONTENTCONTROLTYPE ENUMERATION
# Specifies a rich-text content control.
wdContentControlRichText = 0
# Specifies a text content control
wdContentControlText = 1
# Specifies a picture content control.
wdContentControlPicture = 2
# Specifies a combo box content control.
wdContentControlComboBox = 3
# Specifies a drop-down list content control.
wdContentControlDropdownList = 4
# Specifies a building block gallery content control.
wdContentControlBuildingBlockGallery = 5
# Specifies a date content control.
wdContentControlDate = 6
# Specifies a group content control.
wdContentControlGroup = 7
# Specifies a checkbox content control.
wdContentControlCheckbox = 8
# Specifies a repeating section content control.
wdContentControlRepeatingSection = 9
# labels
paperSizeDescriptions = {
	wdPaper10x14: _(" 10 inches wide, 14 inches long"),
	wdPaper11x17: _("Legal 11 inches wide, 17 inches long"),
	wdPaperA3: _("A3 dimensions"),
	wdPaperA4: _("A4 dimensions"),
	wdPaperA4Small: _("Small A4 dimensions"),
	wdPaperA5: _("A5 dimensions"),
	wdPaperB4: _("B4 dimensions"),
	wdPaperB5: _("B5 dimensions"),
	wdPaperCSheet: _("C sheet dimensions"),
	wdPaperCustom: _("Custom paper size"),
	wdPaperDSheet: _("D sheet dimensions"),
	wdPaperEnvelope10: _("Legal envelope, size 10"),
	wdPaperEnvelope11: _("Envelope, size 11"),
	wdPaperEnvelope12: _("Envelope, size 12"),
	wdPaperEnvelope14: _("Envelope, size 14"),
	wdPaperEnvelope9: _("Envelope, size 9"),
	wdPaperEnvelopeB4: _("B4 envelope"),
	wdPaperEnvelopeB5: _("B5 envelope"),
	wdPaperEnvelopeB6: _("B6 envelope"),
	wdPaperEnvelopeC3: _("C3 envelope"),
	wdPaperEnvelopeC4: _("C4 envelope"),
	wdPaperEnvelopeC5: _("C5 envelope"),
	wdPaperEnvelopeC6: _("C6 envelope"),
	wdPaperEnvelopeC65: _("C65 envelope"),
	wdPaperEnvelopeDL: _("DL envelope"),
	wdPaperEnvelopeItaly: _("Italian envelope"),
	wdPaperEnvelopeMonarch: _("Monarch envelope"),
	wdPaperEnvelopePersonal: _("Personal envelope"),
	wdPaperESheet: _("E sheet dimensions"),
	wdPaperExecutive: _("Executive dimensions"),
	wdPaperFanfoldLegalGerman: _("German legal fanfold dimensions"),
	wdPaperFanfoldStdGerman: _("German standard fanfold dimensions"),
	wdPaperFanfoldUS: _("United States fanfold dimensions"),
	wdPaperFolio: _("Folio dimensions"),
	wdPaperLedger: _("Ledger dimensions"),
	wdPaperLegal: _("Legal dimensions"),
	wdPaperLetter: _("Letter dimensions"),
	wdPaperLetterSmall: _("Small letter dimensions"),
	wdPaperNote: _("Note dimensions"),
	wdPaperQuarto: _("Quarto dimensions"),
	wdPaperStatement: _("Statement dimensions"),
	wdPaperTabloid: _("Tabloid dimensions"),
}

colorNames = {
	wdColorAqua: _("Aqua"),
	wdColorBlack: _("Black"),
	wdColorBlue: _("Blue"),
	wdColorBlueGray: _("BlueGray"),
	wdColorBrightGreen: _("BrightGreen"),
	wdColorBrown: _("Brown"),
	wdColorDarkBlue: _("DarkBlue"),
	wdColorDarkGreen: _("DarkGreen"),
	wdColorDarkRed: _("DarkRed"),
	wdColorDarkTeal: _("DarkTeal"),
	wdColorDarkYellow: _("DarkYellow"),
	wdColorGold: _("Gold"),
	wdColorGray05: _("Gray05"),
	wdColorGray10: _("Gray10"),
	wdColorGray125: _("Gray125"),
	wdColorGray15: _("Gray15"),
	wdColorGray20: _("Gray20"),
	wdColorGray25: _("Gray25"),
	wdColorGray30: _("Gray30"),
	wdColorGray35: _("Gray35"),
	wdColorGray375: _("Gray375"),
	wdColorGray40: _("Gray40"),
	wdColorGray45: _("Gray45"),
	wdColorGray50: _("Gray50"),
	wdColorGray55: _("Gray55"),
	wdColorGray60: _("Gray60"),
	wdColorGray625: _("Gray625"),
	wdColorGray65: _("Gray65"),
	wdColorGray70: _("Gray70"),
	wdColorGray75: _("Gray75"),
	wdColorGray80: _("Gray80"),
	wdColorGray85: _("Gray85"),
	wdColorGray875: _("Gray875"),
	wdColorGray90: _("Gray90"),
	wdColorGray95: _("Gray95"),
	wdColorGreen: _("Green"),
	wdColorIndigo: _("Indigo"),
	wdColorLavender: _("Lavender"),
	wdColorLightBlue: _("LightBlue"),
	wdColorLightGreen: _("LightGreen"),
	wdColorLightOrange: _("LightOrange"),
	wdColorLightTurquoise: _("LightTurquoise"),
	wdColorLightYellow: _("LightYellow"),
	wdColorLime: _("Lime"),
	wdColorOliveGreen: _("OliveGreen"),
	wdColorOrange: _("Orange"),
	wdColorPaleBlue: _("PaleBlue"),
	wdColorPink: _("Pink"),
	wdColorPlum: _("Plum"),
	wdColorRed: _("Red"),
	wdColorRose: _("Rose"),
	wdColorSeaGreen: _("SeaGreen"),
	wdColorSkyBlue: _("SkyBlue"),
	wdColorTan: _("Tan"),
	wdColorTeal: _("Teal"),
	wdColorTurquoise: _("Turquoise"),
	wdColorViolet: _("Violet"),
	wdColorWhite: _("White"),
	wdColorYellow: _("Yellow"),
}
wdColorIndex2wdColor = {
	wdColorIndexBlack: wdColorBlack,
	wdColorIndexBlue: wdColorBlue,
	wdColorIndexBrightGreen: wdColorBrightGreen,
	wdColorIndexDarkBlue: wdColorDarkBlue,
	wdColorIndexDarkRed: wdColorDarkRed,
	wdColorIndexDarkYellow: wdColorDarkYellow,
	wdColorIndexGray25: wdColorGray25,
	wdColorIndexGray50: wdColorGray50,
	wdColorIndexGreen: wdColorGreen,
	wdColorIndexPink: wdColorPink,
	wdColorIndexRed: wdColorRed,
	wdColorIndexTeal: wdColorTeal,
	wdColorIndexTurquoise: wdColorTurquoise,
	wdColorIndexViolet: wdColorViolet,
	wdColorIndexWhite: wdColorWhite,
	wdColorIndexYellow: wdColorYellow,
}
borderNames = {
	wdBorderBottom: _("Bottom"),
	wdBorderDiagonalDown: _("Diagonal down"),
	wdBorderDiagonalUp: _("Diagonal up"),
	wdBorderHorizontal: _("Horizontal"),
	wdBorderLeft: _("Left"),
	wdBorderRight: _("Right"),
	wdBorderTop: _("Top"),
	wdBorderVertical: _("Vertical"),
}

lineStyleDescriptions = {
	wdLineStyleDashDot: _("Dash Dot"),
	wdLineStyleDashDotDot: _("Dash Dot Dot"),
	wdLineStyleDashDotStroked: _("Dash Dot Stroked"),
	wdLineStyleDashLargeGap: _("Dash Large Gap"),
	wdLineStyleDashSmallGap: _("Dash Small Gap"),
	# wdLineStyleDot: _("Dotted"),
	wdLineStyleDouble: _("Double"),
	wdLineStyleDoubleWavy: _("Double Wavy"),
	wdLineStyleEmboss3D: _("Emboss 3D"),
	wdLineStyleEngrave3D: _("Engrave 3D"),
	wdLineStyleInset: _("Inset"),
	wdLineStyleNone: _("None"),
	wdLineStyleOutset: _("Outset"),
	wdLineStyleSingle: _("Single"),
	wdLineStyleSingleWavy: _("Single Wavy"),
	wdLineStyleThickThinLargeGap: _("Thick Thin Large Gap"),
	wdLineStyleThickThinMedGap: _("Thick Thin Medium Gap"),
	wdLineStyleThickThinSmallGap: _("Thick Thin Small Gap"),
	wdLineStyleThinThickLargeGap: _("Thin Thick Large Gap"),
	wdLineStyleThinThickMedGap: _("Thin Thick Medium Gap"),
	wdLineStyleThinThickSmallGap: _("Thin Thick Small Gap"),
	wdLineStyleThinThickThinLargeGap: _("Thin Thick Thin Large Gap"),
	wdLineStyleThinThickThinMedGap: _("Thin Thick Thin Medium Gap"),
	wdLineStyleThinThickThinSmallGap: _("Thin Thick Thin Small Gap"),
	wdLineStyleTriple: _("Triple"),
}
lineWidthDescriptions = {
	wdLineWidth025pt: _("0.25 points"),
	wdLineWidth050pt: _("0.5 points"),
	wdLineWidth075pt: _("0.75 points"),
	wdLineWidth100pt: _("1 point"),
	wdLineWidth150pt: _("1.5 points"),
	wdLineWidth225pt: _("2.25 points"),
	wdLineWidth300pt: _("3 points"),
	wdLineWidth450pt: _("4.5 points"),
	wdLineWidth600pt: _("6 points"),
	-1: _("custom width"),
}
artStyleDescriptions = {
	wdArtApples: _("Apples"),
	wdArtArchedScallops: _("ArchedScallops"),
	wdArtBabyPacifier: _("BabyPacifier"),
	wdArtBabyRattle: _("BabyRattle"),
	wdArtBalloons3Colors: _("Balloons3Colors"),
	wdArtBalloonsHotAir: _("BalloonsHotAir"),
	wdArtBasicBlackDashes: _("BasicBlackDashes"),
	wdArtBasicBlackDots: _("BasicBlackDots"),
	wdArtBasicBlackSquares: _("BasicBlackSquares"),
	wdArtBasicThinLines: _("BasicThinLines"),
	wdArtBasicWhiteDashes: _("BasicWhiteDashes"),
	wdArtBasicWhiteDots: _("BasicWhiteDots"),
	wdArtBasicWhiteSquares: _("BasicWhiteSquares"),
	wdArtBasicWideInline: _("BasicWideInline"),
	wdArtBasicWideMidline: _("BasicWideMidline"),
	wdArtBasicWideOutline: _("BasicWideOutline"),
	wdArtBats: _("Bats"),
	wdArtBirds: _("Birds"),
	wdArtBirdsFlight: _("BirdsFlight"),
	wdArtCabins: _("Cabins"),
	wdArtCakeSlice: _("CakeSlice"),
	wdArtCandyCorn: _("CandyCorn"),
	wdArtCelticKnotwork: _("CelticKnotwork"),
	wdArtCertificateBanner: _("CertificateBanner"),
	wdArtChainLink: _("ChainLink"),
	wdArtChampagneBottle: _("ChampagneBottle"),
	wdArtCheckedBarBlack: _("CheckedBarBlack"),
	wdArtCheckedBarColor: _("CheckedBarColor"),
	wdArtCheckered: _("Checkered"),
	wdArtChristmasTree: _("ChristmasTree"),
	wdArtCirclesLines: _("CirclesLines"),
	wdArtCirclesRectangles: _("CirclesRectangles"),
	wdArtClassicalWave: _("ClassicalWave"),
	wdArtClocks: _("Clocks"),
	wdArtCompass: _("Compass"),
	wdArtConfetti: _("Confetti"),
	wdArtConfettiGrays: _("ConfettiGrays"),
	wdArtConfettiOutline: _("ConfettiOutline"),
	wdArtConfettiStreamers: _("ConfettiStreamers"),
	wdArtConfettiWhite: _("ConfettiWhite"),
	wdArtCornerTriangles: _("CornerTriangles"),
	wdArtCouponCutoutDashes: _("CouponCutoutDashes"),
	wdArtCouponCutoutDots: _("CouponCutoutDots"),
	wdArtCrazyMaze: _("CrazyMaze"),
	wdArtCreaturesButterfly: _("CreaturesButterfly"),
	wdArtCreaturesFish: _("CreaturesFish"),
	wdArtCreaturesInsects: _("CreaturesInsects"),
	wdArtCreaturesLadyBug: _("CreaturesLadyBug"),
	wdArtCrossStitch: _("CrossStitch"),
	wdArtCup: _("Cup"),
	wdArtDecoArch: _("DecoArch"),
	wdArtDecoArchColor: _("DecoArchColor"),
	wdArtDecoBlocks: _("DecoBlocks"),
	wdArtDiamondsGray: _("DiamondsGray"),
	wdArtDoubleD: _("DoubleD"),
	wdArtDoubleDiamonds: _("DoubleDiamonds"),
	wdArtEarth1: _("Earth1"),
	wdArtEarth2: _("Earth2"),
	wdArtEclipsingSquares1: _("EclipsingSquares1"),
	wdArtEclipsingSquares2: _("EclipsingSquares2"),
	wdArtEggsBlack: _("EggsBlack"),
	wdArtFans: _("Fans"),
	wdArtFilm: _("Film"),
	wdArtFirecrackers: _("Firecrackers"),
	wdArtFlowersBlockPrint: _("FlowersBlockPrint"),
	wdArtFlowersDaisies: _("FlowersDaisies"),
	wdArtFlowersModern1: _("FlowersModern1"),
	wdArtFlowersModern2: _("FlowersModern2"),
	wdArtFlowersPansy: _("FlowersPansy"),
	wdArtFlowersRedRose: _("FlowersRedRose"),
	wdArtFlowersRoses: _("FlowersRoses"),
	wdArtFlowersTeacup: _("FlowersTeacup"),
	wdArtFlowersTiny: _("FlowersTiny"),
	wdArtGems: _("Gems"),
	wdArtGingerbreadMan: _("GingerbreadMan"),
	wdArtGradient: _("Gradient"),
	wdArtHandmade1: _("Handmade1"),
	wdArtHandmade2: _("Handmade2"),
	wdArtHeartBalloon: _("HeartBalloon"),
	wdArtHeartGray: _("HeartGray"),
	wdArtHearts: _("Hearts"),
	wdArtHeebieJeebies: _("HeebieJeebies"),
	wdArtHolly: _("Holly"),
	wdArtHouseFunky: _("HouseFunky"),
	wdArtHypnotic: _("Hypnotic"),
	wdArtIceCreamCones: _("IceCreamCones"),
	wdArtLightBulb: _("LightBulb"),
	wdArtLightning1: _("Lightning1"),
	wdArtLightning2: _("Lightning2"),
	wdArtMapleLeaf: _("MapleLeaf"),
	wdArtMapleMuffins: _("MapleMuffins"),
	wdArtMapPins: _("MapPins"),
	wdArtMarquee: _("Marquee"),
	wdArtMarqueeToothed: _("MarqueeToothed"),
	wdArtMoons: _("Moons"),
	wdArtMosaic: _("Mosaic"),
	wdArtMusicNotes: _("MusicNotes"),
	wdArtNorthwest: _("Northwest"),
	wdArtOvals: _("Ovals"),
	wdArtPackages: _("Packages"),
	wdArtPalmsBlack: _("PalmsBlack"),
	wdArtPalmsColor: _("PalmsColor"),
	wdArtPaperClips: _("PaperClips"),
	wdArtPapyrus: _("Papyrus"),
	wdArtPartyFavor: _("PartyFavor"),
	wdArtPartyGlass: _("PartyGlass"),
	wdArtPencils: _("Pencils"),
	wdArtPeople: _("People"),
	wdArtPeopleHats: _("PeopleHats"),
	wdArtPeopleWaving: _("PeopleWaving"),
	wdArtPoinsettias: _("Poinsettias"),
	wdArtPostageStamp: _("PostageStamp"),
	wdArtPumpkin1: _("Pumpkin1"),
	wdArtPushPinNote1: _("PushPinNote1"),
	wdArtPushPinNote2: _("PushPinNote2"),
	wdArtPyramids: _("Pyramids"),
	wdArtPyramidsAbove: _("PyramidsAbove"),
	wdArtQuadrants: _("Quadrants"),
	wdArtRings: _("Rings"),
	wdArtSafari: _("Safari"),
	wdArtSawtooth: _("Sawtooth"),
	wdArtSawtoothGray: _("SawtoothGray"),
	wdArtScaredCat: _("ScaredCat"),
	wdArtSeattle: _("Seattle"),
	wdArtShadowedSquares: _("ShadowedSquares"),
	wdArtSharksTeeth: _("SharksTeeth"),
	wdArtShorebirdTracks: _("ShorebirdTracks"),
	wdArtSkyrocket: _("Skyrocket"),
	wdArtSnowflakeFancy: _("SnowflakeFancy"),
	wdArtSnowflakes: _("Snowflakes"),
	wdArtSombrero: _("Sombrero"),
	wdArtSouthwest: _("Southwest"),
	wdArtStars: _("Stars"),
	wdArtStars3D: _("Stars3D"),
	wdArtStarsBlack: _("StarsBlack"),
	wdArtStarsShadowed: _("StarsShadowed"),
	wdArtStarsTop: _("StarsTop"),
	wdArtSun: _("Sun"),
	wdArtSwirligig: _("Swirligig"),
	wdArtTornPaper: _("TornPaper"),
	wdArtTornPaperBlack: _("TornPaperBlack"),
	wdArtTrees: _("Trees"),
	wdArtTriangleParty: _("TriangleParty"),
	wdArtTriangles: _("Triangles"),
	wdArtTribal1: _("Tribal1"),
	wdArtTribal2: _("Tribal2"),
	wdArtTribal3: _("Tribal3"),
	wdArtTribal4: _("Tribal4"),
	wdArtTribal5: _("Tribal5"),
	wdArtTribal6: _("Tribal6"),
	wdArtTwistedLines1: _("TwistedLines1"),
	wdArtTwistedLines2: _("TwistedLines2"),
	wdArtVine: _("Vine"),
	wdArtWaveline: _("Waveline"),
	wdArtWeavingAngles: _("WeavingAngles"),
	wdArtWeavingBraid: _("WeavingBraid"),
	wdArtWeavingRibbon: _("WeavingRibbon"),
	wdArtWeavingStrips: _("WeavingStrips"),
	wdArtWhiteFlowers: _("WhiteFlowers"),
	wdArtWoodwork: _("Woodwork"),
	wdArtXIllusions: _("XIllusions"),
	wdArtZanyTriangles: _("ZanyTriangles"),
	wdArtZigZag: _("ZigZag"),
	wdArtZigZagStitch: _("ZigZagStitch"),
}
