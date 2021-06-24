# appModules\winword\ww_sections.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import ui
from .ww_wdConst import wdActiveEndPageNumber, wdFirstCharacterLineNumber, wdGoToSection  # noqa:E501
from .ww_collection import Collection, CollectionElement, ReportDialog, convertPixelToUnit  # noqa:E501

_curAddon = addonHandler.getCodeAddon()


addonHandler.initTranslation()


class Section(CollectionElement):
	def __init__(self, parent, item):
		super(Section, self).__init__(parent, item)
		self.borders = item.borders
		self.footers = item.Footers
		self.headers = item.Headers
		self.pageSetup = PageSetup(self, item.PageSetup)
		self.protectedForForms = item.ProtectedForForms
		self.range = item.Range
		self.start = self.range.Start
		r = self.doc.range(self.start, self.start)
		try:
			self.lineStart = r.information(wdFirstCharacterLineNumber)
			self.pageStart = r.Information(wdActiveEndPageNumber)
		except:  # noqa:E722
			self.lineStart = ""
			self.pageStart = ""
		r = self.doc.range(self.range.End-1, self.range.End-1)
		try:
			self.lineEnd = r.information(wdFirstCharacterLineNumber)
			self.pageEnd = r.Information(wdActiveEndPageNumber)
		except:  # noqa:E722
			self.lineEnd = ""
			self.pageEnd = ""

	def formatInfos(self):
		(yes, no) = (_("yes"), _("no"))
		footer = header = no
		if self.footers.Count:
			footer = yes
		if self.headers.Count:
			header = yes
		protected = no
		if self.protectedForForms:
			protected = yes
		pageSetup = self.pageSetup.formatInfos()

		sInfo = _("""Start: Page {pageStart}, line {lineStart}
End: Page {pageEnd}, line {lineEnd}
With header: {header}
With footer: {footer}
Protected for forms: {protected}
{pageSetup}
		""").format(
			pageStart=self.pageStart, lineStart=self.lineStart,
			ageEnd=self.pageEnd, lineEnd=self.lineEnd,
			footer=footer, header=header,
			protected=protected, pageSetup=pageSetup)
		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo

	def reportText(self):
		ui.message(_("page {page} line {line}").format(
			page=self.pageStart, line=self.lineStart))


wdOrientations = {
	0: _("Portrait"),  # wdOrientPortrait,
	1: _("land scape")  # wdOrientLandscape
	}
wdSectionStart = {
		0: _("continue"),  # wdSectionContinuous
		1: _("new column"),  # wdSectionNewColumn
		2: _("new page"),  # wdSectionNewPage
		3: _("even page"),  # wdSectionEvenPage
		4: _("odd page")  # wdSectionOddPage
		}
wdVerticalAlignment = {
	0: _("top"),  # wdAlignVerticalTop
	1: _("center"),  # wdAlignVerticalCenter
	2: _("justify"),  # wdAlignVerticalJustify
	3: _("bottom")  # wdAlignVerticalBottom
}
wdPaperSize = {
	0: "10x14",  # wdPaper10x14
	1: "11x17",  # wdPaper11x17
	2: "letter",  #: wdPaperLetter
	3: "small",  # wdPaperLetterSmall
	4: "legal",  # wdPaperLegal
	5: "executive",  # wdPaperExecutive
	6: "A3",  # wdPaperA3
	7: "A4",  # wdPaperA4
	8: "A4 small",  # wdPaperA4Small
	9: "A5",  # wdPaperA5
	10: "B4",  # wdPaperB4
	11: "B5",  # wdPaperB5
	12: "C sheet",  # wdPaperCSheet
	13: "D sheet",  # wdPaperDSheet
	14: "E sheet",  # wdPaperESheet
	15: "Fa nfold Legal German",  # wdPaperFanfoldLegalGerman
	16: "Fan fold Std German",  # wdPaperFanfoldStdGerman
	17: "Fan fold US",  # wdPaperFanfoldUS
	18: "folio",  # wdPaperFolio
	19: "Ledger",  # wdPaperLedger
	20: "note",  # wdPaperNote
	21: "quarto",  # wdPaperQuarto
	22: "statement",  # wdPaperStatement
	23: "tabloid",  # wdPaperTabloid
	24: "envelope 9",  # wdPaperEnvelope9
	25: "envelope 10",  # wdPaperEnvelope10
	26: "envelope 11",  # wdPaperEnvelope11
	27: "envelope 12",  # wdPaperEnvelope12
	28: "envelope 14",  # wdPaperEnvelope14
	29: "envelope B4",  # wdPaperEnvelopeB4
	30: "envelope B5",  # wdPaperEnvelopeB5
	31: "envelope B6",  # wdPaperEnvelopeB6
	32: "envelope C3",  # wdPaperEnvelopeC3
	33: "envelope C4",  # wdPaperEnvelopeC4
	34: "envelope C5",  # wdPaperEnvelopeC5
	35: "envelope C6",  # wdPaperEnvelopeC6
	36: "envelope 65",  # wdPaperEnvelopeC65
	37: "envelope DL",  # wdPaperEnvelopeDL
	38: "envelope Italy",  # wdPaperEnvelopeItaly
	39: "envelope Monarch",  # wdPaperEnvelopeMonarch
	41: _("custom"),  # wdPaperCustom
	40: "envelope personal",  # wdPaperEnvelopePersonal
	}
wdSectionDirection = {
	0: _("Right alignment, right to left reading"),  # wdSectionDirectionRtl
	1: _("Left alignment, left to right reading")  # wdSectionDirectionLtr
	}


class Sections(Collection):
	_propertyName = (("Sections", Section),)
	_name = (_("Section"), _("Sections"))
	_wdGoToItem = wdGoToSection
	_elementUnit = None

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = SectionsDialog
		self.noElement = _("No section")
		super(Sections, self).__init__(parent, focus)


class PageSetup(object):
	def __init__(self, parent, item):
		self.parent = parent
		self.item = item
		# disposition
		self.sectionStart = item.SectionStart  # WdSectionStart
		# oddAndEvenPagesHeaderFooter : true if has different headers and footers
		self.oddAndEvenPagesHeaderFooter = item.OddAndEvenPagesHeaderFooter
		# for odd-numbered and even-numbered pages.
		# Can be True, False, or wdUndefined.
		self.differentFirstPageHeaderFooter = item.DifferentFirstPageHeaderFooter
		# distance between footer and bottom of the page(in points)
		self.footerDistance = item.FooterDistance
		# distance between header and top of the page (in points)
		self.headerDistance = item.HeaderDistance
		self.verticalAlignment = item.VerticalAlignment
		self.lineNumbering = item.LineNumbering  # object
		# margin (in points)
		self.bottomMargin = item.BottomMargin
		self.topMargin = item.TopMargin
		self.leftMargin = item.LeftMargin
		self.rightMargin = item.RightMargin
		self.orientation = item.Orientation  # WdOrientation
		# paper
		self.paperSize = item.PaperSize  # WdPaperSize
		self.pageHeight = item.PageHeight  # in points
		self.pageWidth = item.PageWidth  # in points
		self.twoPagesOnOne = item.TwoPagesOnOne
		# two pages per sheet.
		self.firstPageTray = item.FirstPageTray  # WdPaperTray
		self.otherPagesTray = item.OtherPagesTray  # WdPaperTray
		# page

		self.charsLine = item.CharsLine  # nb chars by line
		self.linesPage = item.LinesPage  # number of lines per  page
		self.sectionDirection = item.SectionDirection  # WdSectionDirection

	def formatInfos(self):
		(yes, no) = (_("yes"), _("no"))
		start = wdSectionStart[self.sectionStart]
		oddAndEven = firstPage = no
		if self.oddAndEvenPagesHeaderFooter:
			oddAndEven = yes
		if self.differentFirstPageHeaderFooter:
			firstPage = yes
		headerDistance = "%.2f cm" % convertPixelToUnit(self.headerDistance)
		footerDistance = "%.2f cm" % convertPixelToUnit(self.footerDistance)
		verticalAlignment = wdVerticalAlignment[self.verticalAlignment]
		disposition = _("""Disposition:
	Section's start: {start}
	Header / footer's even or odd pages different: {oddAndEven}
	Header/footer's first page different: {firstPage}
	Header distance from top: {headerDistance}
	Footer distance from bottom: {footerDistance}
	Vertical alignment: {verticalAlignment}
		""") .format(
			start=start, oddAndEven=oddAndEven,
			firstPage=firstPage, headerDistance=headerDistance,
			footerDistance=footerDistance, verticalAlignment=verticalAlignment)
		# margins
		left = "%.2f cm" % convertPixelToUnit(self.leftMargin)
		right = "%.2f cm" % convertPixelToUnit(self.rightMargin)
		top = "%.2f cm" % convertPixelToUnit(self.topMargin)
		bottom = "%.2f cm" % convertPixelToUnit(self.bottomMargin)
		orientation = wdOrientations[self.orientation]
		margins = _("""Margins:
	Left,right: {left}, {right}
	Top, bottom: {top}, {bottom}
	Orientation: {orientation}
		""").format(
			left=left, right=right, top=top, bottom=bottom, orientation=orientation)
		# paper
		paperSize = wdPaperSize[self.paperSize]
		width = "%.2f cm" % convertPixelToUnit(self.pageWidth)
		height = "%.2f cm" % convertPixelToUnit(self.pageHeight)
		sectionDirection = wdSectionDirection[self.sectionDirection]
		paper = _("""Paper:
	Format: {paperSize}
	Height's heet: {height}
	Width's sheet: {width}
	Two pages per sheet: {twoPages}
		""") .format(
			paperSize=paperSize, height=height, width=width,
			twoPages=yes if self.twoPagesOnOne else no)
		# page
		chars = int(self.charsLine)
		lines = int(self.linesPage)
		pageFormat = _("""Page:
	Chars number per line: {chars}
	Lines number per page: {lines}
	{sectionDirection}
""").format(chars=chars, lines=lines, sectionDirection=sectionDirection)

		sInfo = ("""{disposition}
{margins}
{paper}
{pageFormat}
		""").format(
			disposition=disposition, margins=margins,
			paper=paper, pageFormat=pageFormat)
		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo


class SectionsDialog(ReportDialog):
	def __init__(self, parent, obj):
		super(SectionsDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		self.lcLabel = _("Sections")
		self.lcColumns = (
			(_("Section"), 100),
			(_("Start"), 150),
			(_("End"), 150),
			)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]

		self.lcSize = (lcWidth, self._defaultLCWidth)

		self.buttons = (
			(100, _("&Go to"), self.goTo),
			)
		self.tc1 = None
		self.tc2 = None

	def get_lcColumnsDatas(self, element):
		start = (_("Page {page}, line {line}")).format(
			page=element.pageStart, line=element.lineStart)
		end = (_("Page {page}, line {line}")).format(
			page=element.pageEnd, line=element.lineEnd)
		index = self.collection.index(element)+1
		datas = (index, start, end)
		return datas
