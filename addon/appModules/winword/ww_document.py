# appModules\winword\ww_document.py
# A part of WordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import speech
import api
from .ww_wdConst import *  # noqa:F403
import sys
import os
_curAddon = addonHandler.getCodeAddon()
path = os.path.join(_curAddon.path, "shared")
sys.path.append(path)
from ww_utils import InformationDialog  # noqa:E402
from ww_NVDAStrings import NVDAString  # noqa:E402
del sys.path[-1]

addonHandler.initTranslation()


def isEvenNumber(number):
	return (number == 0) or (number % 2) == 0


def getPageNumberInfo(obj):
	oPageNumbers = obj.pagenumbers
	if oPageNumbers.Count == 0:
		return ""
	alignmentToMsg = {
		# Translators: text to indicate page number alignment to the left.
		0: _("Page's number on the Left"),
		# Translators: text to indicate page number alignment to the center.
		1: _("Page's number Centered"),
		# Translators: text to indicate page numer alignment to the right.
		2: _("Page's number on the Right"),
		# Translators: text to indicate page number alignment to inside.
		3: _("Page number inside"),
		# Translators: text to indicate page number alignment to outside.
		4: _("Page's number Outside"),
		}
	alignment = oPageNumbers(1).Alignment
	if alignment in alignmentToMsg:
		return alignmentToMsg[alignment]
	else:
		return ""


class Application(object):
	def __init__(self, winwordApplicationObject):
		self.winwordApplicationObject = winwordApplicationObject

	def _getLocalizedMeasurementTextForPointSize(self, offset):
		from NVDAObjects.window.winword import wdInches, wdCentimeters, wdMillimeters, wdPoints, wdPicas  # noqa:E501
		options = self.winwordApplicationObject.Options
		useCharacterUnit = options.useCharacterUnit
		if useCharacterUnit:
			offset = offset/self.winwordApplicationObject.Selection.Font.Size
			# Translators: a measurement in Microsoft Word
			return NVDAString("{offset:.3g} characters").format(offset=offset)
		else:
			unit = options.MeasurementUnit
			if unit == wdInches:
				offset = offset/72.0
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} inches").format(offset=offset)
			elif unit == wdCentimeters:
				offset = offset/28.35
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} centimeters").format(offset=offset)
			elif unit == wdMillimeters:
				offset = offset/2.835
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} millimeters").format(offset=offset)
			elif unit == wdPoints:
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} points").format(offset=offset)
			elif unit == wdPicas:
				offset = offset/12.0
				# Translators: a measurement in Microsoft Word
				# See http://support.microsoft.com/kb/76388 for details.
				return NVDAString("{offset:.3g} picas").format(offset=offset)

	def pointsToDefaultUnits(self, points):
		return self._getLocalizedMeasurementTextForPointSize(points)

	def isCheckSpellingAsYouTypeEnabled(self):
		return self.winwordApplicationObject.Options.CheckSpellingAsYouType

	def isCheckGrammarAsYouTypeEnabled(self):
		return self.winwordApplicationObject.Options.CheckGrammarAsYouType

	def getOverType(self):
		return self.winwordApplicationObject.Options.OverType

	def isCheckGrammarWithSpellingEnabled(self):
		return self.winwordApplicationObject.Options.CheckGrammarWithSpelling

	def isContextualSpellerEnabled(self):
		return self.winwordApplicationObject.Options.ContextualSpeller


class HeaderFooter(object):
	def __init__(self, winwordHeaderFooterObject):
		self.winwordHeaderFooterObject = winwordHeaderFooterObject

	def getRangeText(self):
		if self.winwordHeaderFooterObject.Range.Characters.Count > 1:
			text = self.winwordHeaderFooterObject.range.text
			return text[:-2]
		return ""

	def getPageNumberAlignment(self):
		pageNumbers = self.winwordHeaderFooterObject.PageNumbers
		if pageNumbers.Count == 0:
			return ""
		alignmentToMsg = {
			# Translators: text to indicate page number alignment to the left.
			0: _("Page's number on the Left"),
			# Translators: text to indicate page number alignment to the center.
			1: _("Page's number Centered"),
			# Translators: text to indicate page numer alignment to the right.
			2: _("Page's number on the Right"),
			# Translators: text to indicate page number alignment to inside.
			3: _("Page number inside"),
			# Translators: text to indicate page number alignment to outside.
			4: _("Page's number Outside"),
			}
		alignment = pageNumbers(1).Alignment
		try:
			return alignmentToMsg[alignment]
		except:  # noqa:E722
			return ""


class PageSetup(object):
	def __init__(self, activeDocument, winwordPageSetupObject):
		self.activeDocument = activeDocument
		self.winwordPageSetupObject = winwordPageSetupObject
		self.textColumns = winwordPageSetupObject.Textcolumns
		# differentFirstPageHeaderFooter: different header or footer is used on
		# the first page.
		# Can be True, False, or wdUndefined.
		self.differentFirstPageHeaderFooter = winwordPageSetupObject.DifferentFirstPageHeaderFooter  # noqa:E501
		# oddAndEvenPagesHeaderFooter : True if the specified PageSetup object
		# has different headers and footers for odd-numbered and even-numbered pages.
		# Can be True, False, or wdUndefined.
		self.oddAndEvenPagesHeaderFooter = winwordPageSetupObject.OddAndEvenPagesHeaderFooter  # noqa:E501
		# footerDistance : the distance (in points)
		# between the footer and the bottom of the page.
		self.footerDistance = winwordPageSetupObject.FooterDistance
		# headerDistance : the distance (in points)
		# between the header and the top of the page.
		self.headerDistance = winwordPageSetupObject.HeaderDistance
		# mirrorMargins : True if the inside and outside margins of facing pages
		# are the same width.
		# Can be True, False, or wdUndefined
		self.mirrorMargins = winwordPageSetupObject.MirrorMargins
		# orientation : the orientation of the page.
		self.orientation = winwordPageSetupObject.Orientation
		self.pageHeight = winwordPageSetupObject.PageHeight
		self.pageWidth = winwordPageSetupObject.PageWidth
		self.paperSize = winwordPageSetupObject.PaperSize
		# the reading order and alignment for the specified sections.
		self.sectionDirection = winwordPageSetupObject.SectionDirection
		# sectionStart : type of section break for the specified object.
		self.sectionStart = winwordPageSetupObject.SectionStart
		# twoPagesOnOne: True if specified document is printed two pages per sheet.
		self.twoPagesOnOne = winwordPageSetupObject.TwoPagesOnOne
		# verticalAlignment of text on each page in a document or section.
		self.verticalAlignment = winwordPageSetupObject.VerticalAlignment
		self.gutter = winwordPageSetupObject.Gutter
		self.gutterPos = winwordPageSetupObject.GutterPos
		self.gutterStyle = winwordPageSetupObject.GutterStyle
		self.linesPage = winwordPageSetupObject.LinesPage

	def isMultipleTextColumn(self):
		count = self.textColumns.Count
		return (count > 1)\
			and (count != wdUndefined) and self.activeDocument.isPrintView()

	def getColumnTextStyleInfos(self, textColumns, indent=""):
		pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
		i = 1
		while i <= textColumns.Count:
			width = pointsToDefaultUnits(textColumns(i).Width)
			# Translators: column information.
			text = _("Column {index}: {wide} wide").format(index=str(i), wide=width)
			try:
				spaceAfter = pointsToDefaultUnits(textColumns(i).SpaceAfter)
				# Translators: column information.
				text = text + " " + _("with %s spacing after") % spaceAfter
			except:  # noqa:E722
				pass
			textList.append(indent + text)
			i = i + 1
		return textList

	def getColumnTextInfos(self, indent=""):
		textList = []
		# Translators: title of section's text column paragraph.
		msg = _("Text columns:")
		textList.append(indent + msg)
		curIndent = indent + "\t"
		textColumns = self.winwordPageSetupObject.TextColumns
		count = textColumns.Count
		# Translators: text to indicate text column index.
		msg = _("Number of Text columns: %s") % count
		textList.append(curIndent + msg)
		lineBetween = _("yes") if textColumns.LineBetween else _("no")
		# Translators: text to indicate lines between text columns.
		msg = _("Line between columns: %s") % lineBetween
		textList.append(curIndent + msg)
		if textColumns.EvenlySpaced:
			pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
			# Translators: text to indicate text column style
			msg = _("Style: columns' wide = {wide} (evenly spaced), gap between columns = {spacing}")  # noqa:E501
			msg = msg.format(
				wide=pointsToDefaultUnits(textColumns.Width),
				spacing=pointsToDefaultUnits(textColumns.Spacing))
			textList.append(curIndent + msg)
		else:
			# Translators: text to indicate text column style.
			msg = _("Style: various column wide")
			textList.append(curIndent + msg)
			# Translators: text to indicate text column style.
			textList.extend(self.getColumnTextStyleInfos(textColumns, curIndent))
		return textList

	def getGutterInfos(self, indent=""):
		gutterPosDescriptions = {
			wdGutterPosLeft: _("on the left side"),
			wdGutterPosRight: _("on the right side"),
			wdGutterPosTop: _("on the top side"),
			}

		textList = []
		pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
		pos = gutterPosDescriptions[self.gutterPos]
		size = pointsToDefaultUnits(self.gutter)
		# Translators: gutter information.
		text = _("Gutter: {pos}, {size} large")
		text = text.format(pos=pos, size=size)
		textList.append(indent + text)
		return textList

	def getMarginInfos(self, indent=""):
		pageSetup = self.winwordPageSetupObject
		pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
		leftMargin = pointsToDefaultUnits(pageSetup.LeftMargin)
		rightMargin = pointsToDefaultUnits(pageSetup.RightMargin)
		topMargin = pointsToDefaultUnits(pageSetup.TopMargin)
		bottomMargin = pointsToDefaultUnits(pageSetup.BottomMargin)
		textList = []
		if pageSetup.MirrorMargins:
			# Translators: miror margins informations.
			text = _("Miror margins:")
		else:
			# Translators: margins information.
			text = _("Margins:")
		textList.append(indent + text)
		curIndent = indent + "\t"
		if pageSetup.MirrorMargins:
			# Translators: miror inside margin information:
			text1 = _("Inside: %s") % leftMargin
			# Translators: miror outside margin information.
			text2 = _("Outside: %s") % rightMargin
		else:
			# Translators: left margin information.
			text1 = _("Left: %s") % leftMargin
			# translators: right margin information.
			text2 = _("Right: %s") % rightMargin
		textList.append(curIndent + text1)
		textList.append(curIndent + text2)
		# Translators: top margin information.
		text = _("Top: %s") % topMargin
		textList.append(curIndent + text)
		# Translators: bottom margin information.
		text = _("Bottom: %s") % bottomMargin
		textList.append(curIndent + text)
		# Translators: orientation information.
		orientation = self.getOrientation()
		if orientation != "":
			text = _("Orientation: %s") % orientation
			textList.append(curIndent + text)
		twoPagesOnOneText = _("yes") if self.twoPagesOnOne else _("no")
		# Translators: self.twoPagesOnOne information.
		text = _("Print two pages per sheet: %s") % twoPagesOnOneText
		textList.append(curIndent + text)
		if self.gutter:
			textList.extend(self.getGutterInfos(curIndent))
		return textList

	def getOrientation(self):
		wdOrientationDescriptions = {
			wdOrientLandscape: _("Landscape"),
			wdOrientPortrait: _("Portrait"),
		}
		try:
			return wdOrientationDescriptions[self.orientation]
		except:  # noqa:E722
			return ""

	def getVerticalAlignment(self):
		verticalAlignmentDescriptions = {
			wdAlignVerticalTop: _("top"),
			wdAlignVerticalCenter: _("centered"),
			wdAlignVerticalJustify: _("justified"),
			wdAlignVerticalBottom: _("bottom"),
			}
		try:
			return verticalAlignmentDescriptions[self.verticalAlignment]
		except:  # noqa:E722
			return ""

	def getDispositionInfos(self, indent=""):
		sectionStartDescriptions = {
			wdSectionContinuous: _("Continuous"),
			wdSectionEvenPage: _("Even pages "),
			wdSectionNewColumn: _("New column"),
			wdSectionNewPage: _("New page"),
			wdSectionOddPage: _("Odd pages"),
			}
		pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
		textList = []
		# Translators: title of disposition information .
		text = _("Disposition:")
		textList.append(indent + text)
		curIndent = indent + "\t"
		# Translators: section startinformation.
		text = _("Sectionstart: %s") % sectionStartDescriptions[self.sectionStart]
		textList.append(curIndent + text)
		differentFirstPageText = _("yes")\
			if self.differentFirstPageHeaderFooter else _("no")
		# Translators: different First Page information.
		text = _("Different first page: %s") % differentFirstPageText
		textList.append(curIndent + text)
		# Translators: odd And Even Pages Header Footer information.
		oddAndEvenPagesText = _("yes")\
			if self.oddAndEvenPagesHeaderFooter else _("no")
		# Translators: different odd and even page informations.
		text = _("Different Odd and even pages: %s") % oddAndEvenPagesText
		textList.append(curIndent + text)
		# Translators: distance between footer and bottom information.
		text = _("Distance between footer and bottom: %s")
		text = text % pointsToDefaultUnits(self.footerDistance)
		textList.append(curIndent + text)
		# Translators: distance between header and top information.
		text = _("Distance between header and top: %s") % pointsToDefaultUnits(self.headerDistance)  # noqa:E501
		textList.append(curIndent + text)
		verticalAlignmentText = self.getVerticalAlignment()
		if verticalAlignmentText != "":
			# Translators: vertical alignment information.
			text = _("Vertical alignment: %s") % verticalAlignmentText
			textList.append(curIndent + text)

		return textList

	def getPaperSize(self):
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
		try:
			return paperSizeDescriptions[self.paperSize]
		except:  # noqa:E722
			return ""

	def getPaperInfos(self, indent=""):
		pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
		textList = []
		# Translators: paper information title.
		text = _("Paper:")
		textList.append(indent + text)
		curIndent = indent + "\t"
		# Translators: paper size information.
		text = _("Paper size: %s") % self.getPaperSize()
		textList.append(curIndent + text)
		# Translators: # height and width page information.
		text = _("Page's size: {height}height, {width} width")
		text = text.format(
			height=pointsToDefaultUnits(self.pageHeight),
			width=pointsToDefaultUnits(self.pageWidth))
		textList.append(curIndent + text)
		return textList

	def getInfos(self, indent):
		# pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
		textList = []
		curIndent = indent
		# disposition informations
		textList.extend(self.getDispositionInfos(curIndent))
		# margins informations
		textList.extend(self.getMarginInfos(indent))
		# paper informations
		textList.extend(self.getPaperInfos(curIndent))
		if self.isMultipleTextColumn():
			textList.extend(self.getColumnTextInfos(indent))
		return textList


class Section(object):
	def __init__(self, sectionsCollection, winwordSectionObject):
		self.sectionsCollection = sectionsCollection
		self.activeDocument = self.sectionsCollection.activeDocument
		self.winwordSectionObject = winwordSectionObject
		self.pageSetup = winwordSectionObject.PageSetup
		self.protectedForForms = winwordSectionObject.ProtectedForForms

	def getSectionInfos(self, indent=""):
		textList = []
		protectedText = _("yes") if self.protectedForForms else _("no")
		# Translators: protection forforms information.
		text = _("Text modification only in form fields: %s") % protectedText
		textList.append(indent + text)
		# Translators: section's margin informations.
		# msg_margins = _("{indent}Section' margins:\n{indent}\tLeft: {left}\n{indent}\tRight: {right}\n{indent}\tTop: {top}\n{indent}\tBottom: {bottom}") # noqa:E501
		# Translators: section's miror margin informations
		# msg_mirrorMargins = _("{indent}Section's argins:\n{indent}\tMirror Margins:\n{indent}\tinside: {left}\n{indent}\tOutside: {right}\n{indent}\ttop: {top}\n{indent}\tbottom {bottom}")  # noqa:E501
		pageSetup = self.winwordSectionObject.PageSetup
		headers = self.winwordSectionObject.Headers
		footers = self.winwordSectionObject.Footers
		# Translators: title of section paragraph.
		text = _("Section {index} of {count}:")
		text = text.format(
			index=str(self.winwordSectionObject.index),
			count=str(self.winwordSectionObject.Parent.Sections.Count))
		textList.append(indent + text)
		curIndent = indent + "\t"
		if self.pageSetup.DifferentFirstPageHeaderFooter:
			if headers(wdHeaderFooterFirstPage).exists:
				headerFooter = HeaderFooter(headers(wdHeaderFooterFirstPage))
				text = headerFooter.getRangeText()
				if text != "":
					alignment = headerFooter.getPageNumberAlignment()
					# Translators: first page header information
					msg = curIndent + _("First Page Header: {text}, {alignment}").format(
						text=text, aligment=alignment)
					textList.append(curIndent + msg)
			if footers(wdHeaderFooterFirstPage).exists:
				headerFooter = HeaderFooter(footers(wdHeaderFooterFirstPage))
				text = headerFooter.getRangeText()
				if text != "":
					alignment = headerFooter.getPageNumberAllignment()
					# Translators: first page footer informations.
					msg = _("First Page Footer: {text}, {alignment}").format(
						text=text, alignment=alignment)
					textList.append(curIndent + msg)
		else:
			if pageSetup.OddAndEvenPagesHeaderFooter\
				and isEvenNumber(self.winwordSectionObject.Range.Information(wdActiveEndAdjustedPageNumber)):  # noqa:E501
				if headers(wdHeaderFooterEvenPages).exists:
					headerFooter = HeaderFooter(headers(wdHeaderFooterEvenPages))
					text = headerFooter.getRangeText()
					if text != "":
						alignment = headerFooter.getPageNumberAlignment()
						# Translators: event page header information
						msg = _("Even Page Header: {text}, {alignment}").format(
							text=text, alignment=alignment)
						textList.append(curIndent + msg)
				if footers(wdHeaderFooterEvenPages).exists:
					headerFooter = HeaderFooter(footers(wdHeaderFooterEvenPages))
					text = headerFooter.getRangeText()
					if text != "":
						alignment = headerFooter.getPageNumberAlignment()
						# Translators: Even Page Footer informations
						msg = _("Even Page Footer: {text}, {alignment}").format(
							text=text, alignment=alignment)
						textList.append(curIndent + msg)
			else:
				headerFooter = HeaderFooter(headers(wdHeaderFooterPrimary))
				text = headerFooter.getRangeText()
				if text != "":
					alignment = headerFooter.getPageNumberAlignment()
					# Translators: page header informations
					msg = curIndent + _("Page header: {text}, {alignment}").format(
						text=text, aligment=alignment)
					textList.append(curIndent + msg)
				headerFooter = HeaderFooter(footers(wdHeaderFooterPrimary))
				text = headerFooter.getRangeText()
				if text != "":
					alignment = headerFooter.getPageNumberAlignment()
					# Translators: page footer informations.
					msg = _("Page Footer: {text}, {alignment}").format(
						text=text, alignment=alignment)
					textList.append(curIndent + msg)
		pageSetup = PageSetup(self.activeDocument, self.pageSetup)
		textList.extend(pageSetup.getInfos(curIndent))
		if not self.activeDocument.isProtected()\
			and self.winwordSectionObject.Borders.Enable:
			textList.append(Borders(
				self.activeDocument,
				self.winwordSectionObject.borders).getBordersDescription(curIndent))

		return textList


class Sections(object):
	def __init__(self, activeDocument):
		self.activeDocument = activeDocument
		self.winwordSelectionObject = self.activeDocument.winwordSelectionObject
		self.winwordSectionsObject = self.activeDocument.winwordDocumentObject.Sections  # noqa:E501
		self.pageSetup = self.winwordSectionsObject.PageSetup

	def getAllSectionsInfos(self):
		sections = self.activeDocument.winwordDocumentObject.Sections
		textList = []
		# Translators: title of section paragraph.
		text = _("Document's sections:")
		textList.append(text)
		curIndent = "\t"
		for sct in sections:
			section = Section(self, sct)
			textList.extend(section.getSectionInfos(curIndent))
		return textList


class Selection (object):
	def __init__(self, activeDocument):
		self.activeDocument = activeDocument
		self.winwordSelectionObject = activeDocument.winwordSelectionObject

	def isMultipleTextColumn(self):
		textColumnCount = self.winwordSelectionObject.PageSetup.Textcolumns.Count
		return (textColumnCount > 1)\
			and (textColumnCount != 9999999)\
			and (self.activeDocument.view.Type == wdPrintView)

	def getRangeTextColumnNumber(self):

		if self.inTable()\
			or not self.activeDocument.isPrintView\
			or (self.winwordSelectionObject.storytype != wdMainTextStory):
			return 0
		textColumns = self.winwordSelectionObject.PageSetup.TextColumns
		try:
			textColumnsCount = textColumns.Count
		except:  # noqa:E722
			textColumnsCount = 0
		if textColumnsCount <= 1:
			return 0
		horizontalPositionRelativeToPage = self.winwordSelectionObject.Information(
			wdHorizontalPositionRelativeToPage)
		leftmargin = self.winwordSelectionObject.PageSetup.Leftmargin
		position = horizontalPositionRelativeToPage - leftmargin
		if textColumns.EvenlySpaced == -1:
			width = textColumns.Width + textColumns.Spacing
			return (int(position / width) + 1)
		else:
			width = 0
			for col in range(1, textColumnsCount+1):
				textColumn = textColumns(col)
				width = width + textColumn.Width + textColumn.SpaceAfter
				if position < width:
					return col
		return textColumnsCount

	def inTable(self):
		return self.winwordSelectionObject.Information(wdWithInTable)

	def getCurrentTable(self):
		r = self.winwordSelectionObject.Range
		r.Collapse()
		r.Expand(wdTable)
		return r.Tables(1)

	def getPositionInfos(self):
		selection = self.winwordSelectionObject
		page = selection.Information(wdActiveEndPageNumber)
		textList = []
		# Translators: title of position paragraph.
		text = _("Position in the document:")
		textList.append(text)
		if self.inTable():
			# Translators: position informations.
			text = _("In section {section}, page {page}") .format(
				section=selection.Sections(1).index, page=page)
			textList.append("\t" + text)
			tableIndex = self.activeDocument.winwordDocumentObject.Range(
				0, selection.Tables(1).Range.End).Tables.Count
			table = self.getCurrentTable()
			uniform = table.Uniform
			# Translators: table uniform information
			text = _("uniform") if uniform else _("non-uniform")
			# Translators: In table informations.
			text = _("In a {uniform} table") .format(uniform=text)
			textList.append("\t" + text)
			cell = selection.Cells[0]
			(rowIndex, colIndex) = (cell.rowIndex, cell.columnIndex)
			# Translators: cell positions in table.
			text = _(" Cell row {row} , column {col} of table {tableIndex}") .format(
				row=rowIndex, col=colIndex, tableIndex=tableIndex)
			textList.append("\t" + text)
		else:
			textColumnIndex = self.getRangeTextColumnNumber()
			if textColumnIndex:
				# Translators: in section and column position.
				text = _("In section {section}, text column {textColumn}").format(
					section=selection.Sections(1).index, textColumn=textColumnIndex)
				textList.append("\t" + text)
			else:
				# Translators: just in section.
				text = _("In section {%s") % selection.Sections(1).index
				textList.append("\t" + text)
			line = selection.information(wdFirstCharacterLineNumber)
			column = selection.information(wdFirstCharacterColumnNumber)
			# Translators: indicate line and column position in the page.
			text = _("On line {line} of the page {page}, column {col}") .format(
				line=line, page=page, col=column)
			textList.append("\t" + text)
		return textList


class ActiveDocument(object):
	def getColorName(self, color):
		colorName = self.focus.winwordColorToNVDAColor(color)
		return colorName

	def __init__(self, focus):
		self.focus = focus
		self.application = Application(self.focus.WinwordApplicationObject)
		self.winwordDocumentObject = focus.WinwordDocumentObject
		self.winwordSelectionObject = focus.WinwordSelectionObject
		self._getMainProperties()

	def isPrintView(self):
		view = self.winwordDocumentObject.ActiveWindow.View
		return view.Type == wdPrintView

	def _getMainProperties(self):
		doc = self.winwordDocumentObject
		self.name = doc.Name
		self.readOnly = doc.ReadOnly
		self.protectionType = doc.ProtectionType
		self.inTable = self.winwordSelectionObject.information(wdWithInTable)

	def isProtected(self):
		return self.protectionType >= 0

	def getDocumentProtection(self):
		protectionTypeToText = {
			wdAllowOnlyRevisions: _("Allow only revisions"),
			wdAllowOnlyComments: _("Allow only comments"),
			wdAllowOnlyFormFields: _("Allow only form fields"),
			wdAllowOnlyReading: _("Read-only"),
			}
		try:
			return protectionTypeToText[self.protectionType]
		except:  # noqa:E722
			return _("No protection")

	def isProtectedForm(self):
		return self.protectionType == wdAllowOnlyFormFields

	def _getPositionInfos(self):
		selection = Selection(self)
		return selection.getPositionInfos()

	def getStatistics(self):
		textList = []
		# Translators: title of statistics paragraph.
		text = _("Statistics:")
		textList.append(text)
		doc = self.winwordDocumentObject
		# Translators: text to indicate number of pages
		text = _("Pages: %s") % doc.ComputeStatistics(wdStatisticPages)
		textList.append("\t" + text)
		# Translators: text to indicate number of lines.
		text = _("Lines: %s") % doc.ComputeStatistics(wdStatisticLines)
		textList.append("\t" + text)
		# Translators: text to indicate number of words
		text = _("Words: %s") % doc.ComputeStatistics(wdStatisticWords)
		textList.append("\t" + text)
		# Translators: text to indicate number of characters
		text = _("Characters: %s") % doc.ComputeStatistics(wdStatisticCharacters)
		textList.append("\t" + text)
		# Translators: text to indicate number of paragraphs
		text = _("Paragraphs: %s") % doc.ComputeStatistics(wdStatisticParagraphs)
		textList.append("\t" + text)
		if doc.Sections.Count:
			# Translators: text to indicate number of sections.
			text = _("Sections: %s") % doc.Sections.Count
			textList.append("\t" + text)
		if doc.Comments.Count:
			# Translators: text to indicate number of comments.
			text = _("Comments: %s") % doc.Comments.Count
			textList.append("\t" + text)
		if doc.Tables.Count:
			# Translators:text to indicate number of tables.
			text = _("Tables: %s") % doc.Tables.Count
			textList.append("\t" + text)
		if doc.Shapes.Count:
			# Translators: text to indicate number of shape objects.
			text = _("Shape objects: %s") % doc.Shapes.Count
			textList.append("\t" + text)
		if doc.InLineShapes.Count:
			# Translators: text to indicate number of inLine shape objects.
			text = _("InLineShape objects: %s") % doc.InLineShapes.Count
			textList.append("\t" + text)
		if doc.Endnotes.Count:
			# Translators: text to indicate number of endnotes .
			text = _("Endnotes: %s") % doc.Endnotes.Count
			textList.append("\t" + text)
		if doc.Footnotes.Count:
			# Translators: text to indicate number of footnotes.
			text = _("Footnotes: %s") % doc.Footnotes.Count
			textList.append("\t" + text)
		if doc.Fields.Count:
			# Translators: text to indicate number of fields.
			text = _("Fields: %s") % doc.Fields.Count
			textList.append("\t" + text)
		if doc.FormFields.Count:
			# Translators: text to indicate number of formfields.
			text = _("Formfields: %s") % doc.FormFields.Count
			textList.append("\t" + text)
		if doc.Frames.Count:
			# Translators: text to indicate number of frames.
			text = _("Frames: %s") % doc.Frames.Count
			textList.append("\t" + text)
		if doc.Hyperlinks.Count:
			# Translators: text to indicate number of hyperlinks.
			text = _("Hyperlinks: %s") % doc.Hyperlinks.Count
			textList.append("\t" + text)
		if doc.Lists.Count:
			# Translators: text to indicate number of lists.
			text = _("Lists: %s") % doc.Lists.Count
			textList.append("\t" + text)
		if doc.Bookmarks.Count:
			# Translators: text to indicate number of bookmarks.
			text = _("Bookmarks: %s") % doc.Bookmarks.Count
			textList.append("\t" + text)
		if doc.SpellingErrors.Count:
			# Translators: text to indicate number of spelling errors.
			text = _("Spelling errors: %s") % doc.SpellingErrors.Count
			textList.append("\t" + text)
		if False and doc.GrammaticalErrors.Count:
			# Translators: text to indicate number of grammatical errors.
			text = _("Grammatical errors: %s") % doc.GrammaticalErrors.Count
			textList.append("\t" + text)
		return textList

	def getDocumentProperties(self):
		documentProperties = [
			("Author", _("Author")),
			("Creation Date", _("Creation date")),
			("Revision Number", _("Revision number")),
			("Title", _("Title")),
			("Subject", _("Subject")),
			("Company", _("Compagny"))
			]
		textList = []
		# Translators: document informations
		text = _("Document's properties:")
		textList.append(text)
		text = _("File name: %s") % self.name
		textList.append("\t" + text)
		for propertyName, propertyLabel in documentProperties:
			value = self.winwordDocumentObject.BuiltInDocumentProperties(
				propertyName).Value()
			if value:
				text = "%s: %s" % (propertyLabel, value)
				textList.append("\t" + text)
		text = _("Protection: %s") % self.getDocumentProtection()
		textList.append("\t" + text)
		return textList

	def getOptionsInfos(self):
		textList = []
		# Translators: title of options paragraph.
		text = _("Specific options:")
		textList.append(text)
		# Translators: text to indicate check spelling as you type.
		text = _("Check spelling as you type: %s") % (
			_("Yes") if self.application.isCheckSpellingAsYouTypeEnabled() else _("No"))
		textList.append("\t" + text)
		# Translators: text to indicate check grammar as you type.
		text = _("Check grammar as you type: %s") % (
			_("Yes") if self.application.isCheckGrammarAsYouTypeEnabled() else _("No"))
		textList.append("\t" + text)
		checkGrammarWithSpellingText = _("yes")\
			if self.application.isCheckGrammarWithSpellingEnabled() else _("no")
		# Translators: check grammar with speeling option information.
		text = _("Check grammar with spelling: %s") % checkGrammarWithSpellingText
		textList.append(text)
		# Translators: text to indicate rack revision is activated.
		text = _("Track revision: %s") % (
			_("Yes") if self.winwordDocumentObject.TrackRevisions else _("No"))
		textList.append("\t" + text)
		return textList

	def getMainDocumentInformations(self):
		textList = []
		textList.extend(self._getPositionInfos())
		textList.extend(self.getDocumentProperties())
		textList.extend(self.getStatistics())
		textList.extend(Sections(self).getAllSectionsInfos())
		textList.extend(self.getTablesInformations())
		textList.extend(self.getOptionsInfos())

		return "\n".join(textList)

	def getTablesInformations(self):
		tables = self.winwordDocumentObject.Tables
		textList = []
		if tables.Count == 0:
			return textList
		# Translators: title of tables informations
		text = _("Document's tables:")
		textList.append(text)
		curIndent = "\t"
		count = tables.Count
		for i in range(1, count+1):
			# Translators: table information
			text = curIndent + _("Table {index} of {count}:").format(
				index=str(i), count=str(count))
			textList.append(text)
			table = Table(self, tables(i))
			text = table.getInfos(curIndent + "\t")
			textList.append(text)
		return textList

	def reportDocumentInformations(self):
		# Translators: message to user for waiting.
		speech.speakMessage(_("Please wait"))
		from .ww_tones import RepeatBeep
		rb = RepeatBeep()
		rb.start()
		saved = self.winwordDocumentObject.Saved
		infos = self.getMainDocumentInformations()
		self.winwordDocumentObject.Saved = saved
		rb.stop()
		InformationDialog.run(None, _("Document's informations"), infos)


def getColorDescription(color, colorIndex=None):
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
	obj = api.getFocusObject()
	colorName = obj.winwordColorToNVDAColor(color)
	return colorName
	try:
		return colorNames[color]
	except:  # noqa:E722
		pass
	if colorIndex is not None:
		try:
			return colorNames[wdColorIndex2wdColor[colorIndex]]
		except:  # noqa:E722
			pass
	return ""


class Table(object):
	def __init__(self, activeDocument, winwordTableObject):
		self.activeDocument = activeDocument
		self.winwordTableObject = winwordTableObject
		self.description = winwordTableObject.Descr
		self.title = winwordTableObject.Title
		self.nestingLevel = winwordTableObject.NestingLevel
		self.spacing = winwordTableObject.Spacing
		self.tablesCount = winwordTableObject.TAbles.Count
		self.columnsCount = winwordTableObject.Columns.Count
		self.allowAutoFit = winwordTableObject.AllowAutoFit
		self.topPadding = winwordTableObject.TopPadding
		self.bottomPadding = winwordTableObject.BottomPadding

		self.rowsCount = winwordTableObject.Rows.Count
		self.range = winwordTableObject.Range
		self.title = ""
		self.start = winwordTableObject.range.Start
		doc = activeDocument.winwordDocumentObject
		r = doc.range(self.start, self.start)
		self.line = r.information(wdFirstCharacterLineNumber)
		self.page = r.Information(wdActiveEndPageNumber)

	def isUniform(self):
		return self.winwordTableObject.Uniform

	def isNested(self):
		return self.nestingLevel > 1

	def getInfos(self, indent=""):
		pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
		textList = []
		uniformText = _("uniform") if self.isUniform() else _("non-uniform")
		nestedText = _(" nested") if self.isNested() else ""
		# Translators:  table informations.
		text = _("table {nested} {uniform} of {rows} rows, {columns} columns")
		text = text.format(
			nested=nestedText, uniform=uniformText,
			rows=self.rowsCount, columns=self.columnsCount)
		textList.append(indent + text)
		text = _("Localized at {page} page , {line} line").format(
			page=self.page, line=self.line)
		textList.append(indent + text)
		if self.title and self.title != "":
			# Translators: title of table.
			text = _("Title: %s") % self.title
			textList.append(indent + text)
		if self.description and self.description != "":
			# Translators: table's description.
			text = _("Description: %s") % self.description
			textList.append(indent + text)
		allowAutoFit = _("yes") if self.allowAutoFit else _("no")
		# Translators: allow auto fit information.
		text = _("Automatic resize to fit content: %s") % allowAutoFit
		textList.append(indent + text)
		# Translators: top paddding information.
		text = _("Top padding: %s") % pointsToDefaultUnits(self.topPadding)
		textList.append(indent + text)
		# Translators: bottom padding information.
		text = _("Bottom padding: %s") % pointsToDefaultUnits(self.bottomPadding)
		textList.append(indent + text)
		# Translators: spacing between cells.
		text = _("Spacing between cells: %s") % pointsToDefaultUnits(self.spacing)
		textList.append(indent + text)
		if self.winwordTableObject.Borders.Enable:
			textList.append(Borders(
				self.activeDocument,
				self.winwordTableObject.borders).getBordersDescription(indent))
		if self.tablesCount:
			# Translators: number of tables contained in this table.
			text = _("Contains %s tables") % self.tablesCount
			textList.append(indent + text)
		return "\n".join(textList)


class Borders (object):
	_borderNames = {
		wdBorderBottom: _("Bottom"),
		wdBorderDiagonalDown: _("Diagonal down"),
		wdBorderDiagonalUp: _("Diagonal up"),
		wdBorderHorizontal: _("Horizontal"),
		wdBorderLeft: _("Left"),
		wdBorderRight: _("Right"),
		wdBorderTop: _("Top"),
		wdBorderVertical: _("Vertical"),
		}

	def __init__(self, activeDocument, winwordBordersObject):
		self.activeDocument = activeDocument
		self.winwordBordersObject = winwordBordersObject
		self.enable = winwordBordersObject.Enable

	def getBorderName(self, borderIndex):
		return self._borderNames[borderIndex]

	def getBordersDescription(self, indent=""):
		borderCollection = self.winwordBordersObject
		count = borderCollection.Count
		if count == 0:
			return ""
		leftBorder = borderCollection(wdBorderLeft)
		topBorder = borderCollection(wdBorderTop)
		rightBorder = borderCollection(wdBorderRight)
		bottomBorder = borderCollection(wdBorderBottom)
		# see if oorders are uniform (four surrounding borders are the same)
		lineStyle = leftBorder.LineStyle
		lineWidth = leftBorder.LineWidth
		description = []
		# Translators: title of borders description
		text = _("Borders:")
		description.append(indent + text)
		if (
			leftBorder.Visible
			and topBorder.Visible
			and rightBorder.Visible
			and bottomBorder.Visible
			and leftBorder.LineStyle == lineStyle
			and topBorder.LineStyle == lineStyle
			and rightBorder.LineStyle == lineStyle
			and bottomBorder.LineStyle == lineStyle
			and leftBorder.LineWidth == lineWidth
			and topBorder.LineWidth == lineWidth
			and rightBorder.LineWidth == lineWidth
			and bottomBorder.LineWidth == lineWidth):
			border = Border(topBorder)
			color = self.activeDocument.getColorName(topBorder.Color)
			lineStyleText = border.getLineStyle()
			lineWidthText = border.getLineWidth()
			# check if border has art style
			try:
				artStyle = topBorder.ArtStyle
			except:  # noqa:E722
				artStyle = False

			if artStyle:
				artStyleText = border.getArtStyle()
				msg = _("surrounding {color} {lineStyle} with width of {lineWidth} and art style {artStyle}")  # noqa:E501
				description.append(msg.format(
					color=color, lineStyle=lineStyleText,
					lineWidth=lineWidthText, artStyle=artStyleText))
			else:
				msg = _("surrounding {color} {lineStyle} with width of {lineWidth}")
				description.append(msg.format(
					color=color, lineStyle=lineStyleText, lineWidth=lineWidthText))
			return " ".join(description)
		# the borders are not uniform:
		foundVisibleBorder = False
		curIndent = indent + "\t"
		for index in self._borderNames:
			if borderCollection(index).Visible:
				foundVisibleBorder = True
				border = borderCollection(index)
				b = Border(border)
				name = self.getBorderName(index)
				color = self.activeDocument.getColorName(border.Color)
				lineStyleText = b.getLineStyle()
				lineWidthText = b.getLineWidth()
				try:
					artStyle = border.ArtStyle
				except:  # noqa:E722
					artStyle = False
				if artStyle:
					artStyleText = b.getArtStyle()
					msg = _("{name}= {color} {lineStyle} with width of {lineWidth} and art style {artStyle}")  # noqa:E501
					text = msg.format(
						name=name, color=color,
						lineStyle=lineStyleText, lineWidth=lineWidthText, artStyle=artStyleText)
				else:
					msg = _("{name}= {color} {lineStyle} with width of {lineWidth}")
					text = msg.format(
						name=name, color=color,
						lineStyle=lineStyleText, lineWidth=lineWidthText)
				description.append(curIndent + text)
		if foundVisibleBorder:
			return "\n".join(description)
		# Translators: no border is visible
			text = _("None")
			description.append(text)
			return " ".join(description)


class Border(object):
	def __init__(self, winwordBorderObject):
		self.winwordBorderObject = winwordBorderObject

	def getLineStyle(self):
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
		lineStyle = self.winwordBorderObject.LineStyle
		try:
			desc = lineStyleDescriptions[lineStyle]
		except:  # noqa:E722
			desc = _("mixed")
		return _("%s line") % desc

	def getLineWidth(self):
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
		lineWidth = self.winwordBorderObject.LineWidth
		try:
			desc = lineWidthDescriptions[lineWidth]
		except:  # noqa:E722
			desc = ""
		return desc

	def getArtStyle(self):
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

		artStyle = self.winwordBorderObject.ArtStyle
		if artStyle == 0:
			return ""
		try:
			desc = artStyleDescriptions[artStyle]
		except:  # noqa:E722
			desc = ""
		return desc
