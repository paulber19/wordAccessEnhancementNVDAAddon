# appModules\winword\ww_document.py
# A part of WordAccessEnhancement add-on
# Copyright (C) 2019-2022 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import speech
import api
from .ww_wdConst import wdUndefined
import sys
import os
_curAddon = addonHandler.getCodeAddon()
path = os.path.join(_curAddon.path, "shared")
sys.path.append(path)
from ww_informationDialog import InformationDialog
from ww_NVDAStrings import NVDAString
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
		from NVDAObjects.window.winword import (
			wdInches, wdCentimeters, wdMillimeters, wdPoints, wdPicas)
		options = self.winwordApplicationObject.Options
		useCharacterUnit = options.useCharacterUnit
		if useCharacterUnit:
			offset = offset / self.winwordApplicationObject.Selection.Font.Size
			# Translators: a measurement in Microsoft Word
			return NVDAString("{offset:.3g} characters").format(offset=offset)
		else:
			unit = options.MeasurementUnit
			if unit == wdInches:
				offset = offset / 72.0
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} inches").format(offset=offset)
			elif unit == wdCentimeters:
				offset = offset / 28.35
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} centimeters").format(offset=offset)
			elif unit == wdMillimeters:
				offset = offset / 2.835
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} millimeters").format(offset=offset)
			elif unit == wdPoints:
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} points").format(offset=offset)
			elif unit == wdPicas:
				offset = offset / 12.0
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

	def getRangeText(self, firstSection=False):
		if not firstSection and self.winwordHeaderFooterObject.LinkToPrevious:
			# Translators: text to report a header or footer linked to previous.
			return _("linked to previous")
		if self.winwordHeaderFooterObject.Range.Characters.Count > 1:
			text = self.winwordHeaderFooterObject.range.text
			return text
		return ""

	def getPageNumberAlignment(self):
		alignmentToMsg = {
			# for wdAlignPageNumberLeft
			# Translators: text to indicate page number alignment to the left.
			0: _("Page's number on the Left"),
			# for wdAlignPageNumberCenter
			# Translators: text to indicate page number alignment to the center.
			1: _("Page's number Centered"),
			# for wdAlignPageNumberRight
			# Translators: text to indicate page numer alignment to the right.
			2: _("Page's number on the Right"),
			# for wdAlignPageNumberInside
			# Translators: text to indicate page number alignment to inside.
			3: _("Page number inside"),
			# for wdAlignPageNumberOutside
			# Translators: text to indicate page number alignment to outside.
			4: _("Page's number Outside"),
		}
		pageNumbers = self.winwordHeaderFooterObject.PageNumbers
		if pageNumbers.Count == 0:
			return None
		alignment = pageNumbers(1).Alignment
		alignmentMsg = alignmentToMsg.get(alignment)
		return alignmentMsg


class PageSetup(object):
	def __init__(self, activeDocument, winwordPageSetupObject):
		self.activeDocument = activeDocument
		self.winwordPageSetupObject = winwordPageSetupObject
		self.textColumns = winwordPageSetupObject.Textcolumns
		# differentFirstPageHeaderFooter: different header or footer is used on
		# the first page.
		# Can be True, False, or wdUndefined.
		self.differentFirstPageHeaderFooter = winwordPageSetupObject.DifferentFirstPageHeaderFooter
		# oddAndEvenPagesHeaderFooter : True if the specified PageSetup object
		# has different headers and footers for odd-numbered and even-numbered pages.
		# Can be True, False, or wdUndefined.
		self.oddAndEvenPagesHeaderFooter = winwordPageSetupObject.OddAndEvenPagesHeaderFooter
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
		return (
			count > 1
			and count != wdUndefined
			and self.activeDocument.isPrintView())

	def getColumnTextStyleInfos(self, textColumns, indent=""):
		pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
		i = 1
		textList = []
		while i <= textColumns.Count:
			width = pointsToDefaultUnits(textColumns(i).Width)
			# Translators: column information.
			text = _("Column {index}: {wide} wide").format(index=str(i), wide=width)
			try:
				spaceAfter = pointsToDefaultUnits(textColumns(i).SpaceAfter)
				# Translators: column information.
				text = text + " " + _("with %s spacing after") % spaceAfter
				textList.append(indent + text)
			except Exception:
				pass
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
			msg = _("Style: columns' wide = {wide} (evenly spaced), gap between columns = {spacing}")
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
		from .ww_wdConst import (
			wdGutterPosLeft, wdGutterPosRight, wdGutterPosTop)
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
		from .ww_wdConst import wdOrientLandscape, wdOrientPortrait
		wdOrientationDescriptions = {
			wdOrientLandscape: _("Landscape"),
			wdOrientPortrait: _("Portrait"),
		}
		try:
			return wdOrientationDescriptions[self.orientation]
		except KeyError:
			return ""

	def getVerticalAlignment(self):
		from .ww_wdConst import (
			wdAlignVerticalTop, wdAlignVerticalCenter, wdAlignVerticalJustify, wdAlignVerticalBottom)
		verticalAlignmentDescriptions = {
			wdAlignVerticalTop: _("top"),
			wdAlignVerticalCenter: _("centered"),
			wdAlignVerticalJustify: _("justified"),
			wdAlignVerticalBottom: _("bottom"),
		}
		try:
			return verticalAlignmentDescriptions[self.verticalAlignment]
		except KeyError:
			return ""

	def getDispositionInfos(self, indent=""):
		from .ww_wdConst import (
			wdSectionContinuous, wdSectionEvenPage, wdSectionNewColumn,
			wdSectionNewPage, wdSectionOddPage)
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
		text = _("Distance between header and top: %s") % pointsToDefaultUnits(self.headerDistance)
		textList.append(curIndent + text)
		verticalAlignmentText = self.getVerticalAlignment()
		if verticalAlignmentText != "":
			# Translators: vertical alignment information.
			text = _("Vertical alignment: %s") % verticalAlignmentText
			textList.append(curIndent + text)

		return textList

	def getPaperSize(self):
		from .ww_wdConst import paperSizeDescriptions
		try:
			return paperSizeDescriptions[self.paperSize]
		except KeyError:
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
		self.sectionNumber = str(self.winwordSectionObject.index)
		self.pageSetup = winwordSectionObject.PageSetup
		self.protectedForForms = winwordSectionObject.ProtectedForForms

	def getHeaderFootersInfos(self, pageSetup, indent=""):
		from .ww_wdConst import (
			wdHeaderFooterFirstPage, wdHeaderFooterEvenPages, wdHeaderFooterPrimary)
		curIndent = indent
		textList = []
		headers = self.winwordSectionObject.Headers
		footers = self.winwordSectionObject.Footers
		firstSection = self.sectionNumber == 1
		if self.pageSetup.DifferentFirstPageHeaderFooter:
			if headers(wdHeaderFooterFirstPage).exists:
				headerFooter = HeaderFooter(headers(wdHeaderFooterFirstPage))
				text = headerFooter.getRangeText(firstSection)
				if text != "":
					alignment = headerFooter.getPageNumberAlignment()
					# Translators: first page header information
					msg = _("First Page Header: %s") % text
					if alignment:
						msg = "%s, %s" % (msg, alignment)
					textList.append(curIndent + msg)
			if footers(wdHeaderFooterFirstPage).exists:
				headerFooter = HeaderFooter(footers(wdHeaderFooterFirstPage))
				text = headerFooter.getRangeText(firstSection)
				if text != "":
					alignment = headerFooter.getPageNumberAlignment()
					# Translators: first page footer informations.
					msg = _("First Page Footer: %s") % text
					if alignment:
						msg = "%s, %s" % (msg, alignment)
					textList.append(curIndent + msg)
		if pageSetup.OddAndEvenPagesHeaderFooter:
			if headers(wdHeaderFooterEvenPages).exists:
				headerFooter = HeaderFooter(headers(wdHeaderFooterEvenPages))
				text = headerFooter.getRangeText(firstSection)
				if text != "":
					alignment = headerFooter.getPageNumberAlignment()
					# Translators: event page header information
					msg = _("Even Page Header: %s") % text
					if alignment:
						msg = "%s, %s" % (msg, alignment)
					textList.append(curIndent + msg)
			if footers(wdHeaderFooterEvenPages).exists:
				headerFooter = HeaderFooter(footers(wdHeaderFooterEvenPages))
				text = headerFooter.getRangeText(firstSection)
				if text != "":
					alignment = headerFooter.getPageNumberAlignment()
					# Translators: Even Page Footer informations
					msg = _("Even Page Footer: %s") % text
					if alignment:
						msg = "%s, %s" % (msg, alignment)
					textList.append(curIndent + msg)
		headerFooter = HeaderFooter(headers(wdHeaderFooterPrimary))
		text = headerFooter.getRangeText(firstSection)
		if text != "":
			alignment = headerFooter.getPageNumberAlignment()
			# Translators: page header informations
			msg = _("Page header: %s") % text
			if alignment:
				msg = "%s, %s" % (msg, alignment)
			textList.append(curIndent + msg)
		headerFooter = HeaderFooter(footers(wdHeaderFooterPrimary))
		text = headerFooter.getRangeText(firstSection)
		if text != "":
			alignment = headerFooter.getPageNumberAlignment()
			# Translators: page footer informations.
			msg = _("Page Footer: %s") % text
			if alignment:
				msg = "%s, %s" % (msg, alignment)
			textList.append(curIndent + msg)
		return textList

	def getSectionInfos(self, indent=""):
		textList = []
		# Translators: title of section paragraph.
		text = _("Section {index} of {count}:")
		text = text.format(
			index=self.sectionNumber,
			count=str(self.winwordSectionObject.Parent.Sections.Count))
		textList.append(indent + text)
		curIndent = indent + "\t"
		protectedText = _("yes") if self.protectedForForms else _("no")
		# Translators: protection forforms information.
		text = _("Text modification only in form fields: %s") % protectedText
		textList.append(curIndent + text)
		pageSetup = self.winwordSectionObject.PageSetup
		textList.extend(self.getHeaderFootersInfos(pageSetup, curIndent))
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
		self.winwordSectionsObject = self.activeDocument.winwordDocumentObject.Sections
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
		from .ww_wdConst import wdPrintView
		textColumnCount = self.winwordSelectionObject.PageSetup.Textcolumns.Count
		return (textColumnCount > 1)\
			and (textColumnCount != 9999999)\
			and (self.activeDocument.view.Type == wdPrintView)

	def getRangeTextColumnNumber(self):
		from .ww_wdConst import wdMainTextStory, wdHorizontalPositionRelativeToPage
		if self.inTable()\
			or not self.activeDocument.isPrintView\
			or (self.winwordSelectionObject.storytype != wdMainTextStory):
			return 0
		textColumns = self.winwordSelectionObject.PageSetup.TextColumns
		try:
			textColumnsCount = textColumns.Count
		except Exception:
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
			for col in range(1, textColumnsCount + 1):
				textColumn = textColumns(col)
				width = width + textColumn.Width + textColumn.SpaceAfter
				if position < width:
					return col
		return textColumnsCount

	def inTable(self):
		from .ww_wdConst import wdWithInTable
		return self.winwordSelectionObject.Information(wdWithInTable)

	def getCurrentTable(self):
		r = self.winwordSelectionObject.Range
		r.Collapse()
		from .ww_wdConst import wdTable
		r.Expand(wdTable)
		return r.Tables(1)

	def getPositionInfos(self):
		selection = self.winwordSelectionObject
		from .ww_wdConst import wdActiveEndPageNumber
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
			from .ww_wdConst import wdFirstCharacterLineNumber, wdFirstCharacterColumnNumber
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
		from .ww_wdConst import wdPrintView
		return view.Type == wdPrintView

	def _getMainProperties(self):
		doc = self.winwordDocumentObject
		self.name = doc.Name
		self.readOnly = doc.ReadOnly
		self.protectionType = doc.ProtectionType
		from .ww_wdConst import wdWithInTable
		self.inTable = self.winwordSelectionObject.information(wdWithInTable)

	def isProtected(self):
		return self.protectionType >= 0

	def getDocumentProtection(self):
		from .ww_wdConst import (
			wdAllowOnlyRevisions, wdAllowOnlyComments, wdAllowOnlyFormFields, wdAllowOnlyReading)
		protectionTypeToText = {
			wdAllowOnlyRevisions: _("Allow only revisions"),
			wdAllowOnlyComments: _("Allow only comments"),
			wdAllowOnlyFormFields: _("Allow only form fields"),
			wdAllowOnlyReading: _("Read-only"),
		}
		try:
			return protectionTypeToText[self.protectionType]
		except KeyError:
			return _("No protection")

	def isProtectedForm(self):
		from .ww_wdConst import wdAllowOnlyFormFields
		return self.protectionType == wdAllowOnlyFormFields

	def _getPositionInfos(self):
		selection = Selection(self)
		return selection.getPositionInfos()

	def getStatistics(self):
		from .ww_wdConst import (
			wdStatisticPages, wdStatisticLines, wdStatisticWords,
			wdStatisticCharacters, wdStatisticParagraphs)
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
		textList.append("\t" + text)
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
		for i in range(1, count + 1):
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
		# Translation: title of information dialog.
		InformationDialog.run(None, _("Document's informations"), "", infos, False)


def getColorDescription(color, colorIndex=None):
	from .ww_wdConst import colorNames, wdColorIndex2wdColor
	obj = api.getFocusObject()
	colorName = obj.winwordColorToNVDAColor(color)
	return colorName
	try:
		return colorNames[color]
	except KeyError:
		pass
	if colorIndex is not None:
		try:
			return colorNames[wdColorIndex2wdColor[colorIndex]]
		except Exception:
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
		from .ww_wdConst import wdFirstCharacterLineNumber, wdActiveEndPageNumber
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

		textList.append(
			Borders(
				self.activeDocument,
				self.winwordTableObject.borders
			).getBordersDescription(indent)
		)
		if self.tablesCount:
			# Translators: number of tables contained in this table.
			text = _("Contains %s tables") % self.tablesCount
			textList.append(indent + text)
		return "\n".join(textList)


class Borders (object):
	def __init__(self, activeDocument, winwordBordersObject):
		self.activeDocument = activeDocument
		self.winwordBordersObject = winwordBordersObject
		self.enable = winwordBordersObject.Enable

	def getBorderName(self, borderIndex):
		from .ww_wdConst import borderNames
		return borderNames[borderIndex]

	def getBordersDescription(self, indent=""):
		if not self.enable:
			return _("Without border")
		borderCollection = self.winwordBordersObject
		count = borderCollection.Count
		if count == 0:
			return ""
		description = []
		# Translators: title of borders description
		text = _("Borders:")
		description.append(indent + text)
		from .ww_wdConst import wdBorderLeft, wdBorderTop, wdBorderRight, wdBorderBottom
		leftBorder = borderCollection(wdBorderLeft)
		topBorder = borderCollection(wdBorderTop)
		rightBorder = borderCollection(wdBorderRight)
		bottomBorder = borderCollection(wdBorderBottom)
		# see if orders are uniform (four surrounding borders are the same)
		lineStyle = leftBorder.LineStyle
		lineWidth = leftBorder.LineWidth
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
			except Exception:
				artStyle = False
			if artStyle:
				artStyleText = border.getArtStyle()
				# Translators:  colore and line style of borders.
				msg = _(
					"surrounding {color} {lineStyle}"
					" with width of {lineWidth} and art style {artStyle}"
				)
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
		from .ww_wdConst import borderNames
		for index in borderNames:
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
				except Exception:
					artStyle = False
				if artStyle:
					artStyleText = b.getArtStyle()
					msg = _("{name}= {color} {lineStyle} with width of {lineWidth} and art style {artStyle}")
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
		return _("Without borders")


class Border(object):
	def __init__(self, winwordBorderObject):
		self.winwordBorderObject = winwordBorderObject

	def getLineStyle(self):
		from .ww_wdConst import lineStyleDescriptions
		lineStyle = self.winwordBorderObject.LineStyle
		try:
			desc = lineStyleDescriptions[lineStyle]
		except KeyError:
			desc = _("mixed")
		return _("%s line") % desc

	def getLineWidth(self):
		from .ww_wdConst import lineWidthDescriptions
		lineWidth = self.winwordBorderObject.LineWidth
		try:
			desc = lineWidthDescriptions[lineWidth]
		except KeyError:
			desc = ""
		return desc

	def getArtStyle(self):
		from .ww_wdConst import artStyleDescriptions
		artStyle = self.winwordBorderObject.ArtStyle
		if artStyle == 0:
			return ""
		try:
			desc = artStyleDescriptions[artStyle]
		except KeyError:
			desc = ""
		return desc
