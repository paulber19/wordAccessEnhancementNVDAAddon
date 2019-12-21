# appModules\winword\ww_document.py
# A part of WordAccessEnhancement add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
import speech
from .ww_wdConst import *
import sys
import os
_curAddon = addonHandler.getCodeAddon()
path = os.path.join(_curAddon.path, "shared")
sys.path.append(path)
from ww_utils import printDebug,  InformationDialog
from ww_NVDAStrings import NVDAString
del sys.path[-1]

def isEvenNumber ( number):
	return   (number == 0) or (number% 2) == 0

def getPageNumberInfo ( obj ):
		oPageNumbers = obj.pagenumbers
		if  oPageNumbers.Count == 0 : 
			return ""
		alignmentToMsg = {
			# Translators: text to indicate page number  alignment to the left.
			0 : _("Page's number on the Left"),
			# Translators: text to indicate page number alignment  to the center.
			1 : _("Page's number Centered"),
			# Translators: text to indicate page numer alignment to the right.
			2 : _("Page's number on the Right"),
			# Translators: text to indicate page number alignment to inside.
			3 : _("Page number inside"),
			# Translators: text to indicate page number alignment to outside.
			4 : _("Page's number Outside"),
			}
		alignment = oPageNumbers (1).Alignment
		if alignment in alignmentToMsg:
			return alignmentToMsg[alignment]
		else:
			return ""

	

class Application(object):
	def __init__(self,winwordApplicationObject):
		self.winwordApplicationObject =  winwordApplicationObject
	
	def _getLocalizedMeasurementTextForPointSize(self,offset):
		from NVDAObjects.window.winword import wdInches, wdCentimeters, wdMillimeters, wdPoints, wdPicas
		options=self.winwordApplicationObject.Options
		useCharacterUnit=options.useCharacterUnit
		if useCharacterUnit:
			offset=offset/self.winwordApplicationObject.Selection.Font.Size
			# Translators: a measurement in Microsoft Word
			return NVDAString("{offset:.3g} characters").format(offset=offset)
		else:
			unit=options.MeasurementUnit
			if unit==wdInches:
				offset=offset/72.0
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} inches").format(offset=offset)
			elif unit==wdCentimeters:
				offset=offset/28.35
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} centimeters").format(offset=offset)
			elif unit==wdMillimeters:
				offset=offset/2.835
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} millimeters").format(offset=offset)
			elif unit==wdPoints:
				# Translators: a measurement in Microsoft Word
				return NVDAString("{offset:.3g} points").format(offset=offset)
			elif unit==wdPicas:
				offset=offset/12.0
				# Translators: a measurement in Microsoft Word
				# See http://support.microsoft.com/kb/76388 for details.
				return NVDAString("{offset:.3g} picas").format(offset=offset)

	def pointsToDefaultUnits (self,  points):
		return self._getLocalizedMeasurementTextForPointSize(points)
	
	def isCheckSpellingAsYouTypeEnabled (self):
		return self.winwordApplicationObject.Options.CheckSpellingAsYouType
	
	def isCheckGrammarAsYouTypeEnabled (self):
		return self.winwordApplicationObject.Options.CheckGrammarAsYouType
	
	def getOverType (self):
		return self.winwordApplicationObject.Options.OverType
	def isCheckGrammarWithSpellingEnabled(self):
		return self.winwordApplicationObject.Options.CheckGrammarWithSpelling
	def isContextualSpellerEnabled(self):
		return self.winwordApplicationObject.Options.ContextualSpeller

class HeaderFooter(object):
	def __init__(self,winwordHeaderFooterObject):
		self.winwordHeaderFooterObject = winwordHeaderFooterObject
	
	def getRangeText(self):
		if self.winwordHeaderFooterObject.Range.Characters.Count > 1:
			text =  self.winwordHeaderFooterObject.range.text
			return text[:-2]
		return ""
	def getPageNumberAlignment( self):
		pageNumbers = self.winwordHeaderFooterObject.PageNumbers
		if  pageNumbers.Count == 0 : 
			return ""
		alignmentToMsg = {
			# Translators: text to indicate page number  alignment to the left.
			0 : _("Page's number on the Left"),
			# Translators: text to indicate page number alignment  to the center.
			1 : _("Page's number Centered"),
			# Translators: text to indicate page numer alignment to the right.
			2 : _("Page's number on the Right"),
			# Translators: text to indicate page number alignment to inside.
			3 : _("Page number inside"),
			# Translators: text to indicate page number alignment to outside.
			4 : _("Page's number Outside"),
			}
		alignment = pageNumbers (1).Alignment
		try:
			return alignmentToMsg[alignment]
		except:
			return ""

class PageSetup(object):
	def __init__(self, activeDocument, winwordPageSetupObject):
		self.activeDocument = activeDocument
		self.winwordPageSetupObject = winwordPageSetupObject
		self.textColumns = winwordPageSetupObject.Textcolumns
		self.differentFirstPageHeaderFooter	 = winwordPageSetupObject.DifferentFirstPageHeaderFooter# different header or footer is used on the first page. Can be True, False, or wdUndefined.
		self.oddAndEvenPagesHeaderFooter = winwordPageSetupObject.OddAndEvenPagesHeaderFooter #True if the specified PageSetup object has different headers and footers for odd-numbered and even-numbered pages. Can be True, False, or wdUndefined.
		self.footerDistance = winwordPageSetupObject.FooterDistance ##the distance (in points) between the footer and the bottom of the page.
		self.headerDistance = winwordPageSetupObject.HeaderDistance#the distance (in points) between the header and the top of the page.
		self.mirrorMargins = winwordPageSetupObject.MirrorMargins#True if the inside and outside margins of facing pages are the same width. Can be True, False, or 
		self.orientation = winwordPageSetupObject.Orientation#the orientation of the page.
		self.pageHeight = winwordPageSetupObject.PageHeight
		self.pageWidth = winwordPageSetupObject.PageWidth
		self.paperSize = winwordPageSetupObject.PaperSize
		self.sectionDirection = winwordPageSetupObject.SectionDirection#the reading order and alignment for the specified sections.
		self.sectionStart = winwordPageSetupObject.SectionStart#type of section break for the specified object.
		self.twoPagesOnOne = winwordPageSetupObject.TwoPagesOnOne#True if Microsoft Word prints the specified document two pages per sheet.
		self.verticalAlignment = winwordPageSetupObject.VerticalAlignment# vertical alignment of text on each page in a document or section.
		self.gutter = winwordPageSetupObject.Gutter
		self.gutterPos = winwordPageSetupObject.GutterPos
		self.gutterStyle = winwordPageSetupObject.GutterStyle
		self.linesPage = winwordPageSetupObject.LinesPage
	
	def  isMultipleTextColumn (self):
		count = self.textColumns.Count
		return ( count  > 1 ) and (count !=wdUndefined ) and  self.activeDocument.isPrintView ()
	
	def getColumnTextStyleInfos(self, textColumns, indent= ""):
		pointsToDefaultUnits  = self.activeDocument.application.pointsToDefaultUnits 
		i = 1
		msg = ""
		while i <= textColumns.Count:
			width= pointsToDefaultUnits (textColumns (i).Width)
			# Translators: column information.
			text= _("Column {index}: {wide} wide").format(index = str(i), wide = width)
			try:
				spaceAfter = pointsToDefaultUnits (textColumns (i).SpaceAfter)
				# Translators: column information.
				text = text +" " +_("with %s spacing after")%spaceAfter
			except:
				pass
			textList.append(indent + text)
			i = i + 1
		return textList

	
	def getColumnTextInfos(self, indent = ""):
		textList = []
		# Translators: title of section's text column paragraph.
		msg  = _("Text columns:")
		textList.append(indent + msg)
		curIndent = indent +"\t"
		textColumns = self.winwordPageSetupObject.TextColumns
		count = textColumns.Count
		# Translators: text to indicate text column index.
		msg = _("Number of Text columns: %s" )%count
		textList.append(curIndent + msg)
		lineBetween  = _("yes") if textColumns.LineBetween else _("no")
		# Translators: text to indicate lines between text columns.
		msg = _("Line between columns: %s") %lineBetween
		textList.append(curIndent + msg)
		if textColumns.EvenlySpaced:
			pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
			# Translators: text to indicate  text column style
			msg = _("Style: columns' wide = {wide} (evenly spaced),  gap between columns = {spacing}") .format(wide = pointsToDefaultUnits(textColumns.Width), spacing = pointsToDefaultUnits (textColumns.Spacing) )
			textList.append(curIndent + msg)
		else:
			# Translators: text to indicate text column style.
			msg = _("Style:  various column wide")
			textList.append(curIndent + msg)
			# Translators: text to indicate text column style.
			textList.extend(self.getColumnTextStyleInfos(textColumns, curIndent))
		return textList
	def getGutterInfos(self, indent =""):
		gutterPosDescriptions = {
			wdGutterPosLeft : _("on the left side"),
			wdGutterPosRight: _("on the right side"),
			wdGutterPosTop: _("on the top side"),
			}

		textList = []
		pointsToDefaultUnits = self.activeDocument.application.pointsToDefaultUnits
		pos = gutterPosDescriptions[self.gutterPos]
		size = pointsToDefaultUnits(self.gutter)
		# Translators: gutter information.
		text = _("Gutter: {pos}, {size} large").format(pos = pos, size = size)
		textList.append(indent+text)
		return textList
	def getMarginInfos(self, indent = ""):
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
		textList.append(indent +text)
		curIndent = indent + "\t"
		if pageSetup.MirrorMargins:
			# Translators: miror inside margin information:
			text1 = _("Inside: %s") %leftMargin
			#  Translators: miror outside margin information.
			text2 = _("Outside: %s") %rightMargin
		else:
			# Translators: left margin information.
			text1 = _("Left: %s") %leftMargin
			# translators: right margin information.
			text2 = _("Right: %s") %rightMargin
		textList.append(curIndent+ text1)
		textList.append(curIndent + text2)
		# Translators: top margin information.
		text = _("Top: %s") %topMargin
		textList.append(curIndent + text)
		# Translators: bottom margin information.
		text = _("Bottom: %s")%bottomMargin
		textList.append(curIndent + text)
		# Translators: orientation information.
		orientation = self.getOrientation()
		if orientation != "":
			text = _("Orientation: %s")%orientation
			textList.append(curIndent+ text)
		twoPagesOnOneText = _("yes") if self.twoPagesOnOne  else _("no")
		# Translators: self.twoPagesOnOne  information.
		text = _("Print two pages per sheet: %s")%twoPagesOnOneText
		textList.append(curIndent+text)
		if self.gutter:
			textList.extend(self.getGutterInfos(curIndent))


		return textList

	def getOrientation(self):
		wdOrientationDescriptions = {
			wdOrientLandscape :  _("Landscape"),
			wdOrientPortrait: _("Portrait"),
		}
		try:
			return wdOrientationDescriptions[self.orientation]
		except:
			return ""
	def getVerticalAlignment(self):
		verticalAlignmentDescriptions = {
			wdAlignVerticalTop: _("top"),
			wdAlignVerticalCenter : _("centered"),
			wdAlignVerticalJustify : _("justified"),
			wdAlignVerticalBottom : _("bottom"),
			}
		try:
			return verticalAlignmentDescriptions[self.verticalAlignment]
		except:
			return ""
	def getDispositionInfos(self, indent = ""):
		sectionStartDescriptions = {
			wdSectionContinuous : _("Continuous"),
			wdSectionEvenPage:_("Even pages "),
			wdSectionNewColumn:_("New column"),
			wdSectionNewPage: _("New page"),
			wdSectionOddPage :_("Odd pages"),
			}
		pointsToDefaultUnits  = self.activeDocument.application.pointsToDefaultUnits 
		textList = []
		# Translators: title of disposition information .
		text = _("Disposition:")
		textList.append(indent + text)
		curIndent = indent + "\t"
		# Translators: section startinformation.
		text = _("Sectionstart: %s")%sectionStartDescriptions[self.sectionStart]
		textList.append(curIndent+text)
		differentFirstPageText = _("yes") if self.differentFirstPageHeaderFooter else _("no")
		# Translators: different First Page  information.
		text = _("Different first page: %s")%differentFirstPageText
		textList.append(curIndent+text)
		oddAndEvenPagesText = _("yes") if self.oddAndEvenPagesHeaderFooter  else _("no")# Translators: odd And Even Pages Header Footer  information.
		# Translators: different odd and even page informations.
		text = _("Different Odd and even pages: %s")%oddAndEvenPagesText 
		textList.append(curIndent+text)			
		# Translators:  distance between  footer and bottom information.
		text = _("Distance between footer and bottom: %s")%pointsToDefaultUnits(self.footerDistance )
		textList.append(curIndent + text)
		# Translators: distance between header and top information.
		text = _("Distance between header and top: %s")%pointsToDefaultUnits (self.headerDistance )
		textList.append(curIndent+text)
		verticalAlignmentText = self.getVerticalAlignment()
		if verticalAlignmentText != "":
			# Translators: vertical alignment information.
			text = _("Vertical alignment: %s")%verticalAlignmentText
			textList.append(curIndent+text)

		return textList
	
	def getPaperSize(self):
		paperSizeDescriptions = {
			wdPaper10x14: _(" 10 inches wide, 14 inches long"),
			wdPaper11x17 : _("Legal 11 inches wide, 17 inches long"),
			wdPaperA3: _("A3 dimensions"),
			wdPaperA4: _("A4 dimensions"),
			wdPaperA4Small : _("Small A4 dimensions"),
			wdPaperA5 : _("A5 dimensions"),
			wdPaperB4 : _("B4 dimensions"),
			wdPaperB5: _("B5 dimensions"),
			wdPaperCSheet: _("C sheet dimensions"),
			wdPaperCustom: _("Custom paper size"),
			wdPaperDSheet : _("D sheet dimensions"),
			wdPaperEnvelope10 : _("Legal envelope, size 10"),
			wdPaperEnvelope11 :_("Envelope, size 11"),
			wdPaperEnvelope12 : _("Envelope, size 12"),
			wdPaperEnvelope14 : _("Envelope, size 14"),
			wdPaperEnvelope9 :  _("Envelope, size 9"),
			wdPaperEnvelopeB4 : _("B4 envelope"),
			wdPaperEnvelopeB5 :  _("B5 envelope"),
			wdPaperEnvelopeB6 : _("B6 envelope"),
			wdPaperEnvelopeC3 : _("C3 envelope"),
			wdPaperEnvelopeC4 : _("C4 envelope"),
			wdPaperEnvelopeC5 : _("C5 envelope"),
			wdPaperEnvelopeC6 : _("C6 envelope"),
			wdPaperEnvelopeC65 : _("C65 envelope"),
			wdPaperEnvelopeDL : _("DL envelope"),
			wdPaperEnvelopeItaly : _("Italian envelope"),
			wdPaperEnvelopeMonarch : _("Monarch envelope"),
			wdPaperEnvelopePersonal : _("Personal envelope"),
			wdPaperESheet : _("E sheet dimensions"),
			wdPaperExecutive : _("Executive dimensions"),
			wdPaperFanfoldLegalGerman : _("German legal fanfold dimensions"),
			wdPaperFanfoldStdGerman : _("German standard fanfold dimensions"),
			wdPaperFanfoldUS : _("United States fanfold dimensions"),
			wdPaperFolio :_("Folio dimensions"),
			wdPaperLedger :_("Ledger dimensions"),
			wdPaperLegal : _("Legal dimensions"),
			wdPaperLetter : _("Letter dimensions"),
			wdPaperLetterSmall : _("Small letter dimensions"),
			wdPaperNote : _("Note dimensions"),
			wdPaperQuarto : _("Quarto dimensions"),
			wdPaperStatement: _("Statement dimensions"),
			wdPaperTabloid : _("Tabloid dimensions"),
			}
		try:
			return paperSizeDescriptions[self.paperSize]
		except:
			return ""


	def getPaperInfos(self, indent = ""):
		pointsToDefaultUnits  = self.activeDocument.application.pointsToDefaultUnits 
		textList = []
		# Translators: paper information title.
		text = _("Paper:")
		textList.append(indent + text)
		curIndent = indent +"\t"
		# Translators: paper size information.
		text = _("Paper size: %s")%self.getPaperSize()
		textList.append(curIndent +text)
		# Translators: # height and width page information.
		text = _("Page's size: {height}height, {width} width").format(height = pointsToDefaultUnits(self.pageHeight), width = pointsToDefaultUnits(self.pageWidth))
		textList.append(curIndent+text)
		return textList
	def getInfos(self, indent):
		pointsToDefaultUnits  = self.activeDocument.application.pointsToDefaultUnits 
		textList = []
		curIndent = indent
		# disposition informations
		textList.extend(self.getDispositionInfos(curIndent))
		# margins informations
		textList.extend(self.getMarginInfos(indent))
		# paper informations
		textList.extend(self.getPaperInfos(curIndent))
		if self.isMultipleTextColumn ():
			textList.extend(self.getColumnTextInfos(indent))
		return textList
		
	
class Section(object):
	def __init__(self,sectionsCollection, winwordSectionObject):
		self.sectionsCollection = sectionsCollection
		self.activeDocument = self.sectionsCollection.activeDocument
		self.winwordSectionObject = winwordSectionObject
		self.pageSetup = winwordSectionObject.PageSetup
		self.protectedForForms = winwordSectionObject.ProtectedForForms
		
	def getSectionInfos( self, indent = ""):
		textList = []
		protectedText = _("yes") if self.protectedForForms else _("no")
		# Translators: protection forforms information.
		text = _("Text modification  only in form fields: %s")%protectedText
		textList.append(indent+ text)
		# Translators:  section's margin informations.
		msg_margins =  _("{indent}Section' margins:\n{indent}\tLeft: {left}\n{indent}\tRight: {right}\n{indent}\tTop: {top}\n{indent}\tBottom: {bottom}")
		# Translators:  section's miror margin informations
		msg_mirrorMargins= _("{indent}Section's argins:\n{indent}\tMirror Margins:\n{indent}\tinside: {left}\n{indent}\tOutside: {right}\n{indent}\ttop: {top}\n{indent}\tbottom {bottom}")
		pageSetup = self.winwordSectionObject.PageSetup
		headers = self.winwordSectionObject.Headers
		footers = self.winwordSectionObject.Footers 
		# Translators:  title of section paragraph.
		text= _("Section {index} of {count}:") .format(index = str( self.winwordSectionObject.index), count = str(self.winwordSectionObject.Parent.Sections.Count))
		textList.append(indent + text)
		curIndent = indent +"\t"
		if  self.pageSetup.DifferentFirstPageHeaderFooter:
			if  headers (wdHeaderFooterFirstPage  ).exists :
				headerFooter = HeaderFooter(headers (wdHeaderFooterFirstPage  ))
				text = headerFooter.getRangeText()
				if text != "":
					alignment = headerFooter.getPageNumberAlignment()
					# Translators: first page header information
					msg= curIndent + _("First Page Header: {text}, {alignment}").format(text = text, aligment = alignment)
					textList.append(curIndent + msg)
			if  footers (wdHeaderFooterFirstPage  ).exists:
				headerFooter = HeaderFooter(footers (wdHeaderFooterFirstPage))
				text = headerFooter.getRangeText()
				if text  != "":
					alignment = headerFooter.getPageNumberAllignment()
					# Translators: first page footer informations.
					msg= _("First Page Footer: {text}, {alignment}").format(text =text, alignment = alignment)
					textList.append(curIndent + msg)
		else:
			if pageSetup.OddAndEvenPagesHeaderFooter and isEvenNumber (self.winwordSectionObject.Range.Information(wdActiveEndAdjustedPageNumber  )):
				if headers (wdHeaderFooterEvenPages  ).exists:
					headerFooter = HeaderFooter(headers (wdHeaderFooterEvenPages))
					text = headerFooter.getRangeText()
					if  text != "":
						alignment = headerFooter.getPageNumberAlignment()
						# Translators:  event page header information
						msg= _("Even Page Header: {text}, {alignment}").format(text = text,alignment = alignment)
						textList.append(curIndent + msg)
				if  footers (wdHeaderFooterEvenPages  ).exists:
					headerFooter = HeaderFooter(footers (wdHeaderFooterEvenPages))
					text = headerFooter.getRangeText()
					if text != "":
						alignment = headerFooter.getPageNumberAlignment()
						# Translators: Even Page Footer informations
						msg= _("Even Page Footer: {text}, {alignment}").format(text = text, alignment = alignment)
						textList.append(curIndent + msg)
			else:
				headerFooter = HeaderFooter(headers (wdHeaderFooterPrimary))
				text = headerFooter.getRangeText()
				if text  != "":
					aligment = headerFooter.getPageNumberAlignment()
					# Translators: page header informations
					msg= curIndent + _("Page header: {text}, {alignment}").format(text = text, aligment = alignment)
					textList.append(curIndent +msg)
				headerFooter = HeaderFooter(footers (wdHeaderFooterPrimary))
				text = headerFooter.getRangeText()
				if text != "":
					alignment = headerFooter.getPageNumberAlignment()
					# Translators: page footer informations.
					msg= _("Page Footer: {text}, {alignment}").format( text = text, alignment = alignment)
					textList.append(curIndent + msg)
		pageSetup = PageSetup(self.activeDocument, self.pageSetup)
		textList.extend(pageSetup.getInfos(curIndent))
		if not self.activeDocument.isProtected() and  self.winwordSectionObject.Borders.Enable:
			textList.append(Borders(self.activeDocument, self.winwordSectionObject.borders).getBordersDescription(curIndent))

		return textList


class Sections(object):
	def __init__(self,activeDocument):
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
	def __init__(self,activeDocument):	
		self.activeDocument = activeDocument
		self.winwordSelectionObject = activeDocument.winwordSelectionObject
	
	def  isMultipleTextColumn (self):
		textColumnCount = self.winwordSelectionObject.PageSetup.Textcolumns.Count
		return ( textColumnCount  > 1 ) and (textColumnCount !=9999999) and ( self.activeDocument.view.Type  == wdPrintView ) 
	
	def getRangeTextColumnNumber ( self):
		
		if  self.inTable() or not self.activeDocument.isPrintView   or ( self.winwordSelectionObject.storytype != wdMainTextStory ) :
			return 0
		textColumns = self.winwordSelectionObject.PageSetup.TextColumns
		try:
			textColumnsCount = textColumns.Count
		except:
			textColumnsCount = 0
		if  textColumnsCount <= 1:
			return 0
		position = self.winwordSelectionObject.Information(wdHorizontalPositionRelativeToPage ) - self.winwordSelectionObject.PageSetup.Leftmargin
		if  textColumns.EvenlySpaced == -1 :
			width = textColumns.Width + textColumns.Spacing
			return ( int( position / width ) + 1 )
		else:
			width = 0
			for col in range(1, textColumnsCount+1):   
				textColumn = textColumns (col)
				width =  width + textColumn.Width  + textColumn.SpaceAfter
				if  position < width :
					return col
		return textColumnsCount
	def inTable(self):
		return self.winwordSelectionObject.Information(wdWithInTable )
	def  getCurrentTable (self ):
		r = self.winwordSelectionObject.Range
		r.Collapse()
		r.Expand ( wdTable)
		return r.Tables(1)
	
	def getPositionInfos(self):
		selection = self.winwordSelectionObject
		page = selection.Information(wdActiveEndPageNumber )
		textList = []
		# Translators: title of position paragraph.
		text = _("Position in the document:")
		textList.append(text)
		if self.inTable():
			# Translators: position informations.
			text = _("In section {section}, page {page}") .format(section = selection.Sections(1).index, page = page)
			textList.append("\t" + text)
			tableIndex = self.activeDocument.winwordDocumentObject.Range(0, selection.Tables(1).Range.End).Tables.Count
			table = self.getCurrentTable()
			uniform = table.Uniform
			# Translators: table uniform information
			text = _("uniform") if uniform else _("non-uniform")
			# Translators: In table informations.
			text = _("In a {uniform} table") .format(uniform = text)
			textList.append("\t" +text)
			cell = selection.Cells[0]
			(rowIndex, colIndex) = (cell.rowIndex, cell.columnIndex)
			# Translators: cell positions in table.
			text = _(" Cell row {row} , column {col} of table {tableIndex}") .format(row = rowIndex , col = colIndex,  tableIndex = tableIndex)
			textList.append("\t" +text)
		else:
			textColumnIndex = self.getRangeTextColumnNumber ()
			if textColumnIndex:
				# Translators:  in section and column position.
				text = _("In section {section}, text column {textColumn}").format(section = selection.Sections(1).index, textColumn = textColumnIndex)
				textList.append("\t" +text)
			else:
				# Translators: just in section.
				text = _("In section {%s") %selection.Sections(1).index
				textList.append("\t" + text)
			line = selection.information(wdFirstCharacterLineNumber )
			column = selection.information(wdFirstCharacterColumnNumber )
			# Translators: indicate  line and column position in the page.
			text = _("On line {line} of the page {page}, column {col}") .format(line = line, page = page, col = column)
			textList.append("\t" +text)
		return textList

		
class ActiveDocument(object):
	def getColorName(self, color):
		colorName =self.focus.winwordColorToNVDAColor(color)
		return colorName
	def __init__(self,focus):
		self.focus = focus
		self.application = Application(self.focus.WinwordApplicationObject)
		self.winwordDocumentObject = focus.WinwordDocumentObject
		self.winwordSelectionObject = focus.WinwordSelectionObject
		self._getMainProperties()

	def isPrintView (self):
		view = self.winwordDocumentObject.ActiveWindow.View
		return view.Type  == wdPrintView 
	
	def _getMainProperties(self):
		doc = self.winwordDocumentObject
		self.name = doc.Name
		self.readOnly = doc.ReadOnly
		self.protectionType = doc.ProtectionType
		self.inTable =  self.winwordSelectionObject.information(wdWithInTable )
	
	def isProtected (self):
		return self.protectionType >= 0
	
	def  getDocumentProtection (self):
		protectionTypeToText = {
			wdAllowOnlyRevisions  : _("Allow only revisions"),
			wdAllowOnlyComments  : _("Allow only comments"),
			wdAllowOnlyFormFields  : _("Allow only form fields"),
			wdAllowOnlyReading  : _("Read-only"),
			}
		try:
			return protectionTypeToText[self.protectionType]
		except:
						return  _("No protection")

	def isProtectedForm(self):
		return self.protectionType == wdAllowOnlyFormFields
	
	def _getPositionInfos(self):
		selection = Selection(self)
		return selection.getPositionInfos()
	def getStatistics(self):
		textList = []
		# Translators: title of statistics paragraph.
		text  = _("Statistics:")
		textList.append(text)
		doc = self.winwordDocumentObject
		# Translators: text to indicate number of pages
		text = _("Pages: %s") %doc.ComputeStatistics(wdStatisticPages) 
		textList.append("\t" + text)
		# Translators:  text to indicate number of lines.
		text = _("Lines: %s") %doc.ComputeStatistics(wdStatisticLines) 
		textList.append("\t" +text)
		# Translators: text to indicate number of words
		text = _("Words: %s") %doc.ComputeStatistics(wdStatisticWords) 
		textList.append("\t" +text)
		# Translators: text to indicate number of characters
		text = _("Characters: %s") %doc.ComputeStatistics(wdStatisticCharacters) 
		textList.append("\t" +text)
		# Translators: text to indicate number of paragraphs
		text = _("Paragraphs: %s") %doc.ComputeStatistics(wdStatisticParagraphs) 
		textList.append("\t" +text)
		if doc.Sections.Count:
			# Translators: text to indicate number of sections.
			text = _("Sections: %s") %doc.Sections.Count
			textList.append("\t" +text)
		if doc.Comments.Count:
			# Translators: text to indicate number of comments.
			text = _("Comments: %s") %doc.Comments.Count
			textList.append("\t" +text)
		if doc.Tables.Count:
			# Translators:text to indicate number of tables.
			text = _("Tables: %s") %doc.Tables.Count
			textList.append("\t" +text)
		if doc.Shapes.Count:
			# Translators: text to indicate number of shape objects.
			text = _("Shape objects: %s") %doc.Shapes.Count
			textList.append("\t" +text)
		if doc.InLineShapes.Count:
			# Translators: text to indicate number of inLine shape objects.
			text = _("InLineShape objects: %s") %doc.InLineShapes.Count
			textList.append("\t" +text)
		if doc.Endnotes.Count:
			# Translators: text to indicate number of endnotes .
			text = _("Endnotes: %s") %doc.Endnotes.Count
			textList.append("\t" +text)
		if doc.Footnotes.Count:
			#Translators: text to indicate number of footnotes.
			text = _("Footnotes: %s") %doc.Footnotes.Count
			textList.append("\t" +text)
		if doc.Fields.Count:
			# Translators: text to indicate number of fields.
			text = _("Fields: %s") %doc.Fields.Count
			textList.append("\t" +text)
		if doc.FormFields.Count:
			# Translators: text to indicate number of  formfields.
			text = _("Formfields: %s") %doc.FormFields.Count
			textList.append("\t" +text)
		if doc.Frames.Count:
			# Translators: text to indicate number of frames.
			text = _("Frames: %s") %doc.Frames.Count
			textList.append("\t" + text)
		if doc.Hyperlinks.Count:
			# Translators: text to indicate number of hyperlinks.
			text = _("Hyperlinks: %s") %doc.Hyperlinks.Count
			textList.append("\t" +text)
		if doc.Lists.Count:
			# Translators: text to indicate number of lists.
			text = _("Lists: %s") %doc.Lists.Count
			textList.append("\t" +text)
		if doc.Bookmarks.Count:
			# Translators: text to indicate number of bookmarks.
			text = _("Bookmarks: %s") %doc.Bookmarks.Count
			textList.append("\t" +text)
		if doc.SpellingErrors.Count:
			# Translators: text to indicate number of spelling errors.
			text = _("Spelling errors: %s") %doc.SpellingErrors.Count
			textList.append("\t" +text)
		if  False and doc.GrammaticalErrors.Count:
			# Translators: text to indicate number of grammatical errors.
			text = _("Grammatical errors: %s") %doc.GrammaticalErrors.Count
			textList.append("\t" + text)
		return textList
	def getDocumentProperties(self):
		documentProperties = [("Author", _("Author")),
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
		text = _("File name: %s") %self.name
		textList.append("\t" + text)
		for propertyName, propertyLabel in documentProperties:
			value = self.winwordDocumentObject.BuiltInDocumentProperties(propertyName).Value()
			if value:
				text = "%s: %s"%(propertyLabel,value)
				textList.append("\t" +text)
		text = _("Protection: %s") %self.getDocumentProtection ()
		textList.append("\t" +text)
		return textList
	def getOptionsInfos(self):
		textList = []
		# Translators: title of options paragraph.
		text = _("Specific options:")
		textList.append(text)
		# Translators: text to indicate check spelling as you type.
		text = _("Check spelling as you type: %s") %( _("Yes") if self.application.isCheckSpellingAsYouTypeEnabled() else _("No"))
		textList.append("\t" + text)
		# Translators: text to indicate check grammar as you type.
		text = _("Check grammar as you type: %s") %(_("Yes") if self.application.isCheckGrammarAsYouTypeEnabled() else _("No"))
		textList.append("\t" + text)
		checkGrammarWithSpellingText = _("yes") if self.application.isCheckGrammarWithSpellingEnabled() else _("no")
		# Translators: check grammar with speeling option information.
		text = _("Check grammar with spelling: %s") %checkGrammarWithSpellingText
		textList.append(text)
		# Translators: text to indicate rack revision is activated.
		text = _("Track revision: %s") %(_("Yes") if self.winwordDocumentObject.TrackRevisions else _("No"))
		textList.append("\t" +text)
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
		if tables.Count == 0: return textList
		# Translators: title of tables informations
		text = _("Document's tables:")
		textList.append(text)
		curIndent = "\t"
		count = tables.Count
		for i in  range(1, count+1):
			# Translators: table information
			text = curIndent + _("Table {index} of {count}:").format(index = str(i), count = str(count))
			textList.append(text)
			table = Table(self,tables(i))
			text = table.getInfos(curIndent+"\t")
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

import api
def getColorDescription(color, colorIndex = None):
	colorNames = {
		wdColorAqua : _("Aqua"), # msg266_L
		wdColorBlack : _("Black"), # msg267_L
		wdColorBlue : _("Blue"), # msg268_L
		wdColorBlueGray : _("BlueGray"), # msg269_L
		wdColorBrightGreen : _("BrightGreen"), # msg270_L
		wdColorBrown : _("Brown"), # msg271_L
		wdColorDarkBlue : _("DarkBlue"), # msg272_L
		wdColorDarkGreen : _("DarkGreen"), # msg273_L
		wdColorDarkRed : _("DarkRed"), # msg274_L
		wdColorDarkTeal : _("DarkTeal"), # msg275_L
		wdColorDarkYellow : _("DarkYellow"), # msg276_L
		wdColorGold : _("Gold"),# msg277_L
		wdColorGray05 : _("Gray05"), # msg278_L
		wdColorGray10 : _("Gray10"), # msg279_L
		wdColorGray125 : _("Gray125"), # msg280_L
		wdColorGray15 : _("Gray15"), # msg281_L
		wdColorGray20 : _("Gray20"), # msg282_L
		wdColorGray25 : _("Gray25"), # msg283_L
		wdColorGray30 : _("Gray30"), # msg284_L
		wdColorGray35 : _("Gray35"), # msg285_L
		wdColorGray375 : _("Gray375"), # msg286_L
		wdColorGray40 : _("Gray40"), # msg287_L
		wdColorGray45 : _("Gray45"), # msg288_L
		wdColorGray50 : _("Gray50"), # msg289_L
		wdColorGray55 : _("Gray55"), # msg290_L
		wdColorGray60 : _("Gray60"), # msg291_L
		wdColorGray625 : _("Gray625"), # msg292_L
		wdColorGray65 : _("Gray65"), # msg293_L
		wdColorGray70 : _("Gray70"), # msg294_L
		wdColorGray75 : _("Gray75"), # msg295_L
		wdColorGray80 : _("Gray80"), # msg296_L
		wdColorGray85 : _("Gray85"), # msg297_L
		wdColorGray875 : _("Gray875"), # msg298_L
		wdColorGray90 : _("Gray90"), # msg299_L
		wdColorGray95 : _("Gray95"), # msg300_L
		wdColorGreen : _("Green"), # msg301_L
		wdColorIndigo : _("Indigo"), # msg302_L
		wdColorLavender : _("Lavender"), # msg303_L
		wdColorLightBlue : _("LightBlue"), # msg304_L
		wdColorLightGreen : _("LightGreen"), # msg305_L
		wdColorLightOrange : _("LightOrange"), # msg306_L
		wdColorLightTurquoise : _("LightTurquoise"), # msg307_L
		wdColorLightYellow : _("LightYellow"), # msg308_L
		wdColorLime : _("Lime"), # msg309_L
		wdColorOliveGreen : _("OliveGreen"), # msg310_L
		wdColorOrange : _("Orange"), # msg311_L
		wdColorPaleBlue : _("PaleBlue"), # msg312_L
		wdColorPink : _("Pink"), # msg313_L
		wdColorPlum : _("Plum"), # msg314_L
		wdColorRed : _("Red"), # msg315_L
		wdColorRose : _("Rose"), # msg316_L
		wdColorSeaGreen : _("SeaGreen"), # msg317_L
		wdColorSkyBlue : _("SkyBlue"), # msg318_L
		wdColorTan : _("Tan"), # msg319_L
		wdColorTeal : _("Teal"), # msg320_L
		wdColorTurquoise : _("Turquoise"), # msg321_L
		wdColorViolet : _("Violet"), # msg322_L
		wdColorWhite : _("White"), # msg323_L
		wdColorYellow : _("Yellow"), # msg324_L
		}
	wdColorIndex2wdColor = {
		wdColorIndexBlack : wdColorBlack ,
		wdColorIndexBlue : wdColorBlue ,
		wdColorIndexBrightGreen : wdColorBrightGreen ,
		wdColorIndexDarkBlue : wdColorDarkBlue,
		wdColorIndexDarkRed : wdColorDarkRed ,
		wdColorIndexDarkYellow : wdColorDarkYellow ,
		wdColorIndexGray25 : wdColorGray25 ,
		wdColorIndexGray50 : wdColorGray50 ,
		wdColorIndexGreen : wdColorGreen , 
		wdColorIndexPink : wdColorPink ,
		wdColorIndexRed : wdColorRed ,
		wdColorIndexTeal : wdColorTeal ,
		wdColorIndexTurquoise : wdColorTurquoise ,
		wdColorIndexViolet : wdColorViolet ,
		wdColorIndexWhite : wdColorWhite ,
		wdColorIndexYellow : wdColorYellow ,
		}
	obj = api.getFocusObject()
	colorName =obj.winwordColorToNVDAColor(color)
	return colorName
	try:
		return colorNames[color]
	except:
		pass
	if colorIndex is not None:
		try:
			return colorNames[wdColorIndex2wdColor[colorIndex]]
		except:
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
		r = doc.range (self.start, self.start)
		self.line = r.information(wdFirstCharacterLineNumber )
		self.page = r.Information(wdActiveEndPageNumber )
	
	def isUniform(self):
		return self.winwordTableObject.Uniform
	def isNested(self):
		return self.nestingLevel  >1
	
	def getInfos(self, indent = ""):
		pointsToDefaultUnits  = self.activeDocument.application.pointsToDefaultUnits 
		textList = []
		uniformText = _("uniform") if self.isUniform() else _("non-uniform")
		nestedText = _(" nested") if self.isNested() else ""
		text = _("table {nested} {uniform} of {rows} rows, {columns} columns").format(nested = nestedText,uniform = uniformText, rows = self.rowsCount, columns = self.columnsCount)
		textList.append(indent + text)
		text = _("Localized at {page} page , {line} line").format(page = self.page, line = self.line)
		textList.append(indent + text)
		if self.title and self.title != "":
			# Translators: title of table.
			text = _("Title: %s")%self.title
			textList.append(indent+text)
		if self.description and self.description != "":
			# Translators: table's description.
			text = _("Description: %s")%self.description
			textList.append(indent + text)
		allowAutoFit = _("yes") if self.allowAutoFit else _("no")
		# Translators: allow auto fit information.
		text = _("Automatic resize to fit content: %s")%allowAutoFit
		textList.append(indent+text)
		# Translators: top paddding information.
		text = _("Top padding: %s")%pointsToDefaultUnits(self.topPadding)
		textList.append(indent+text)
		# Translators: bottom padding information.
		text = _("Bottom padding: %s")%pointsToDefaultUnits  (self.bottomPadding)
		textList.append(indent + text)
		# Translators: spacing between cells.
		text = _("Spacing between cells: %s")%pointsToDefaultUnits  (self.spacing)
		textList.append(indent + text)
		if self.winwordTableObject.Borders.Enable:
			textList.append(Borders(self.activeDocument, self.winwordTableObject.borders).getBordersDescription(indent))
		if self.tablesCount:
			# Translators: number of tables contained in this table.
			text = _("Contains %s tables")%self.tablesCount
			textList.append(indent +text)
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
		
	def getBordersDescription(self, indent = ""):
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
		if (leftBorder.Visible 
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
			#color = getColorDescription(color = topBorder.Color)
			color = self.activeDocument.getColorName(topBorder.Color)
			lineStyleText = border.getLineStyle()
			lineWidthText = border.getLineWidth()
			# check if border has art style
			try:
				artStyle = topBorder.ArtStyle
			except:
				artStyle = False
			
			if artStyle:
				artStyleText = border.getArtStyle()
				msg = _("surrounding {color} {lineStyle} with width of {lineWidth} and art style {artStyle}")
				description.append(msg.format(color = color, lineStyle = lineStyleText, lineWidth = lineWidthText, artStyle = artStyleText))
			else:
				msg = _("surrounding {color} {lineStyle} with width of {lineWidth}")
				description.append(msg.format(color = color,  lineStyle = lineStyleText, lineWidth = lineWidthText))
			return " ".join(description)
		# the borders are not uniform:
		foundVisibleBorder  = False
		curIndent = indent +"\t"
		for index in self._borderNames:
			if borderCollection(index).Visible:
				foundVisibleBorder = True
				border = borderCollection(index)
				b =Border(border)
				name = self.getBorderName(index)
				#color = getColorDescription(color = border.color)
				color = self.activeDocument.getColorName(border.Color)
				lineStyleText = b.getLineStyle()
				lineWidthText = b.getLineWidth()
				try:
					artStyle = border.ArtStyle
				except:
					artStyle = False
				if artStyle:
					artStyleText = b.getArtStyle()
					msg= _("{name}= {color} {lineStyle} with width of {lineWidth} and art style {artStyle}")
					text = msg.format(name = name, color = color, lineStyle = lineStyleText, lineWidth = lineWidthText, artStyle = artStyleText)
				else:
					msg= _("{name}= {color} {lineStyle} with width of {lineWidth}")
					text =msg.format(name = name, color = color, lineStyle = lineStyleText, lineWidth = lineWidthText)
				description.append(curIndent + text)
		if foundVisibleBorder :
			return "\n".join(description)
		# Translators: no border is visible 
			text = _("None")
			description.append (text)
			return " ".join(description)


class Border(object):
	def __init__(self, winwordBorderObject):
		self.winwordBorderObject = winwordBorderObject
	
	def getLineStyle(self):
		lineStyleDescriptions = {
			wdLineStyleDashDot : _("Dash Dot"), #msg67_L
			wdLineStyleDashDotDot : _("Dash Dot Dot"), #msg68_L
			wdLineStyleDashDotStroked :_("Dash Dot Stroked"), #msg69_L
			wdLineStyleDashLargeGap : _("Dash Large Gap"), #msg70_L
			wdLineStyleDashSmallGap : _("Dash Small Gap"), #msg71_L
			#wdLineStyleDot : _("Dotted"), #msg72_L
			wdLineStyleDouble : _("Double"), #msg73_L
			wdLineStyleDoubleWavy : _("Double Wavy"), #msg74_L
			wdLineStyleEmboss3D : _("Emboss 3D"), #msg75_L
			wdLineStyleEngrave3D : _("Engrave 3D"), #msg76_L
			wdLineStyleInset : _("Inset"), #msg77_L
			wdLineStyleNone : _("None"), #msg78_L
			wdLineStyleOutset : _("Outset"), #msg79_L
			wdLineStyleSingle : _("Single"), #msg80_L
			wdLineStyleSingleWavy : _("Single Wavy"), #msg81_L
			wdLineStyleThickThinLargeGap : _("Thick Thin Large Gap"), #msg82_L
			wdLineStyleThickThinMedGap : _("Thick Thin Medium Gap"), #msg83_L
			wdLineStyleThickThinSmallGap : _("Thick Thin Small Gap"), #msg84_L
			wdLineStyleThinThickLargeGap : _("Thin Thick Large Gap"), #msg85_L
			wdLineStyleThinThickMedGap : _("Thin Thick Medium Gap"), #msg86_L
			wdLineStyleThinThickSmallGap : _("Thin Thick Small Gap"), #msg87_L
			wdLineStyleThinThickThinLargeGap : _("Thin Thick Thin Large Gap"), #msg88_L
			wdLineStyleThinThickThinMedGap : _("Thin Thick Thin Medium Gap"), #msg89_L
			wdLineStyleThinThickThinSmallGap : _("Thin Thick Thin Small Gap"), #msg90_L
			wdLineStyleTriple : _("Triple"), #msg91_L
			}
		lineStyle = self.winwordBorderObject.LineStyle
		try:
			desc= lineStyleDescriptions[lineStyle]
		except:
			desc = _("mixed") #msg92_L
		return _("%s line")%desc#msgBorderLine1,descr)
	
	def getLineWidth(self):
		lineWidthDescriptions = {
			wdLineWidth025pt  : _("0.25 points"), #msg257_L
			wdLineWidth050pt  : _("0.5 points"), #msg258_L
			wdLineWidth075pt  : _("0.75 points"), #msg259_L
			wdLineWidth100pt  : _("1 point"), #msg260_L
			wdLineWidth150pt  : _("1.5 points"), #msg261_L
			wdLineWidth225pt  : _("2.25 points"), #msg262_L
			wdLineWidth300pt  : _("3 points"), #msg263_L
			wdLineWidth450pt  : _("4.5 points"), #msg264_L
			wdLineWidth600pt  : _("6 points"), #msg265_L
			-1:   _("custom width"),#msgCustomWidth 
			}
		lineWidth = self.winwordBorderObject.LineWidth
		try:
			desc = lineWidthDescriptions[lineWidth]
		except:
			desc = ""
		return desc
	
	def getArtStyle(self):
		artStyleDescriptions = {
			wdArtApples  : _("Apples"),#msg93_L
			wdArtArchedScallops  : _("ArchedScallops"),#msg94_L
			wdArtBabyPacifier  : _("BabyPacifier"),#msg95_L
			wdArtBabyRattle  : _("BabyRattle"),#msg96_L
			wdArtBalloons3Colors  : _("Balloons3Colors"),#msg97_L
			wdArtBalloonsHotAir  : _("BalloonsHotAir"),#msg98_L
			wdArtBasicBlackDashes  : _("BasicBlackDashes"),#msg99_L
			wdArtBasicBlackDots  : _("BasicBlackDots"),#msg100_L
			wdArtBasicBlackSquares  : _("BasicBlackSquares"),#msg101_L
			wdArtBasicThinLines  : _("BasicThinLines"), #msg102_L
			wdArtBasicWhiteDashes  : _("BasicWhiteDashes"), #msg103_L
			wdArtBasicWhiteDots  : _("BasicWhiteDots"), #msg104_L
			wdArtBasicWhiteSquares  : _("BasicWhiteSquares"), #msg105_L
			wdArtBasicWideInline  : _("BasicWideInline"), #msg106_L
			wdArtBasicWideMidline  : _("BasicWideMidline"), #msg107_L
			wdArtBasicWideOutline  : _("BasicWideOutline"), #msg108_L
			wdArtBats  : _("Bats"), #msg109_L
			wdArtBirds  : _("Birds"), #msg110_L
			wdArtBirdsFlight  : _("BirdsFlight"), #msg111_L
			wdArtCabins  : _("Cabins"), #msg112_L
			wdArtCakeSlice  : _("CakeSlice"), #msg113_L
			wdArtCandyCorn  : _("CandyCorn"), #msg114_L
			wdArtCelticKnotwork  : _("CelticKnotwork"), #msg115_L
			wdArtCertificateBanner  : _("CertificateBanner"), #msg116_L
			wdArtChainLink  : _("ChainLink"), #msg117_L
			wdArtChampagneBottle  : _("ChampagneBottle"), #msg118_L
			wdArtCheckedBarBlack  : _("CheckedBarBlack"), #msg119_L
			wdArtCheckedBarColor  : _("CheckedBarColor"), #msg120_L
			wdArtCheckered  : _("Checkered"), #msg121_L
			wdArtChristmasTree  : _("ChristmasTree"), #msg122_L
			wdArtCirclesLines  : _("CirclesLines"), #msg123_L
			wdArtCirclesRectangles  : _("CirclesRectangles"), #msg124_L
			wdArtClassicalWave  : _("ClassicalWave"), #msg125_L
			wdArtClocks  : _("Clocks"), #msg126_L
			wdArtCompass  : _("Compass"), #msg127_L
			wdArtConfetti  : _("Confetti"), #msg128_L
			wdArtConfettiGrays  : _("ConfettiGrays"), #msg129_L
			wdArtConfettiOutline  : _("ConfettiOutline"), #msg130_L
			wdArtConfettiStreamers  : _("ConfettiStreamers"), #msg131_L
			wdArtConfettiWhite  : _("ConfettiWhite"), #msg132_L
			wdArtCornerTriangles  : _("CornerTriangles"), #msg133_L
			wdArtCouponCutoutDashes  : _("CouponCutoutDashes"), #msg134_L
			wdArtCouponCutoutDots  : _("CouponCutoutDots"), #msg135_L
			wdArtCrazyMaze  : _("CrazyMaze"), #msg136_L
			wdArtCreaturesButterfly  : _("CreaturesButterfly"), #msg137_L
			wdArtCreaturesFish  : _("CreaturesFish"), #msg138_L
			wdArtCreaturesInsects  : _("CreaturesInsects"), #msg139_L
			wdArtCreaturesLadyBug  : _("CreaturesLadyBug"), #msg140_L
			wdArtCrossStitch  : _("CrossStitch"), #msg141_L
			wdArtCup  : _("Cup"), #msg142_L
			wdArtDecoArch  : _("DecoArch"), #msg143_L
			wdArtDecoArchColor  : _("DecoArchColor"), #msg144_L
			wdArtDecoBlocks  : _("DecoBlocks"), #msg145_L
			wdArtDiamondsGray  : _("DiamondsGray"), #msg146_L
			wdArtDoubleD  : _("DoubleD"), #msg147_L
			wdArtDoubleDiamonds  : _("DoubleDiamonds"), #msg148_L
			wdArtEarth1  : _("Earth1"), #msg149_L
			wdArtEarth2  : _("Earth2"), #msg150_L
			wdArtEclipsingSquares1  : _("EclipsingSquares1"), #msg151_L
			wdArtEclipsingSquares2  : _("EclipsingSquares2"), #msg152_L
			wdArtEggsBlack  : _("EggsBlack"), #msg153_L
			wdArtFans  : _("Fans"), #msg154_L
			wdArtFilm  : _("Film"), #msg155_L
			wdArtFirecrackers  : _("Firecrackers"), #msg156_L
			wdArtFlowersBlockPrint  : _("FlowersBlockPrint"), #msg157_L
			wdArtFlowersDaisies  : _("FlowersDaisies"), #msg158_L
			wdArtFlowersModern1  : _("FlowersModern1"), #msg159_L
			wdArtFlowersModern2  : _("FlowersModern2"), #msg160_L
			wdArtFlowersPansy  : _("FlowersPansy"), #msg161_L
			wdArtFlowersRedRose  : _("FlowersRedRose"), #msg162_L
			wdArtFlowersRoses  : _("FlowersRoses"), #msg163_L
			wdArtFlowersTeacup  : _("FlowersTeacup"), #msg164_L
			wdArtFlowersTiny  : _("FlowersTiny"), #msg165_L
			wdArtGems  : _("Gems"), #msg166_L
			wdArtGingerbreadMan  : _("GingerbreadMan"), #msg167_L
			wdArtGradient  : _("Gradient"), #msg168_L
			wdArtHandmade1  : _("Handmade1"), #msg169_L
			wdArtHandmade2  : _("Handmade2"), #msg170_L
			wdArtHeartBalloon  : _("HeartBalloon"), #msg171_L
			wdArtHeartGray  : _("HeartGray"), #msg172_L
			wdArtHearts  : _("Hearts"), #msg173_L
			wdArtHeebieJeebies  : _("HeebieJeebies"), #msg174_L
			wdArtHolly  : _("Holly"), #msg175_L
			wdArtHouseFunky  : _("HouseFunky"), #msg176_L
			wdArtHypnotic  : _("Hypnotic"), #msg177_L
			wdArtIceCreamCones  : _("IceCreamCones"), #msg178_L
			wdArtLightBulb  : _("LightBulb"), #msg179_L
			wdArtLightning1  : _("Lightning1"), #msg180_L
			wdArtLightning2  : _("Lightning2"), #msg181_L
			wdArtMapleLeaf  : _("MapleLeaf"), #msg182_L
			wdArtMapleMuffins  : _("MapleMuffins"), #msg183_L
			wdArtMapPins  : _("MapPins"), #msg184_L
			wdArtMarquee  : _("Marquee"), #msg185_L
			wdArtMarqueeToothed  : _("MarqueeToothed"), #msg186_L
			wdArtMoons  : _("Moons"), #msg187_L
			wdArtMosaic  : _("Mosaic"), #msg188_L
			wdArtMusicNotes  : _("MusicNotes"), #msg189_L
			wdArtNorthwest  : _("Northwest"), #msg190_L
			wdArtOvals  : _("Ovals"), #msg191_L
			wdArtPackages  : _("Packages"), #msg192_L
			wdArtPalmsBlack  : _("PalmsBlack"), #msg193_L
			wdArtPalmsColor  : _("PalmsColor"), #msg194_L
			wdArtPaperClips  : _("PaperClips"), #msg195_L
			wdArtPapyrus  : _("Papyrus"), #msg196_L
			wdArtPartyFavor  : _("PartyFavor"), #msg197_L
			wdArtPartyGlass  : _("PartyGlass"), #msg198_L
			wdArtPencils  : _("Pencils"), #msg199_L
			wdArtPeople  : _("People"), #msg200_L
			wdArtPeopleHats  : _("PeopleHats"), #msg201_L
			wdArtPeopleWaving  : _("PeopleWaving"), #msg202_L
			wdArtPoinsettias  : _("Poinsettias"), #msg203_L
			wdArtPostageStamp  : _("PostageStamp"), #msg204_L
			wdArtPumpkin1  : _("Pumpkin1"), #msg205_L
			wdArtPushPinNote1  : _("PushPinNote1"), #msg206_L
			wdArtPushPinNote2  : _("PushPinNote2"), #msg207_L
			wdArtPyramids  : _("Pyramids"), #msg208_L
			wdArtPyramidsAbove  : _("PyramidsAbove"), #msg209_L
			wdArtQuadrants  : _("Quadrants"), #msg210_L
			wdArtRings  : _("Rings"), #msg211_L
			wdArtSafari  : _("Safari"), #msg212_L
			wdArtSawtooth  : _("Sawtooth"), #msg213_L
			wdArtSawtoothGray  : _("SawtoothGray"), #msg214_L
			wdArtScaredCat  : _("ScaredCat"), #msg215_L
			wdArtSeattle  : _("Seattle"), #msg216_L
			wdArtShadowedSquares  : _("ShadowedSquares"), #msg217_L
			wdArtSharksTeeth  : _("SharksTeeth"), #msg218_L
			wdArtShorebirdTracks  : _("ShorebirdTracks"), #msg219_L
			wdArtSkyrocket  : _("Skyrocket"), #msg220_L
			wdArtSnowflakeFancy  : _("SnowflakeFancy"), #msg221_L
			wdArtSnowflakes  : _("Snowflakes"), #msg222_L
			wdArtSombrero  : _("Sombrero"), #msg223_L
			wdArtSouthwest  : _("Southwest"), #msg224_L
			wdArtStars  : _("Stars"), #msg225_L
			wdArtStars3D  : _("Stars3D"), #msg226_L
			wdArtStarsBlack  : _("StarsBlack"), #msg227_L
			wdArtStarsShadowed  : _("StarsShadowed"), #msg228_L
			wdArtStarsTop  : _("StarsTop"), #msg229_L
			wdArtSun  : _("Sun"), #msg230_L
			wdArtSwirligig  : _("Swirligig"), #msg231_L
			wdArtTornPaper  : _("TornPaper"), #msg232_L
			wdArtTornPaperBlack  : _("TornPaperBlack"), #msg233_L
			wdArtTrees  : _("Trees"), #msg234_L
			wdArtTriangleParty  : _("TriangleParty"), #msg235_L
			wdArtTriangles  : _("Triangles"),#msg236_L
			wdArtTribal1  : _("Tribal1"), #msg237_L
			wdArtTribal2  : _("Tribal2"), #msg238_L
			wdArtTribal3  : _("Tribal3"), #msg239_L
			wdArtTribal4  : _("Tribal4"), #msg240_L
			wdArtTribal5  : _("Tribal5"), #msg241_L
			wdArtTribal6  : _("Tribal6"), #msg242_L
			wdArtTwistedLines1  : _("TwistedLines1"), #msg243_L
			wdArtTwistedLines2  : _("TwistedLines2"), #msg244_L
			wdArtVine  : _("Vine"), #msg245_L
			wdArtWaveline  : _("Waveline"), #msg246_L
			wdArtWeavingAngles  : _("WeavingAngles"), #msg247_L
			wdArtWeavingBraid  : _("WeavingBraid"), #msg248_L
			wdArtWeavingRibbon  : _("WeavingRibbon"), #msg249_L
			wdArtWeavingStrips  : _("WeavingStrips"), #msg250_L
			wdArtWhiteFlowers  : _("WhiteFlowers"), #msg251_L
			wdArtWoodwork  : _("Woodwork"), #msg252_L
			wdArtXIllusions  : _("XIllusions"), #msg253_L
			wdArtZanyTriangles  : _("ZanyTriangles"), #msg254_L
			wdArtZigZag  : _("ZigZag"), #msg255_L
			wdArtZigZagStitch  : _("ZigZagStitch"), #msg256_L
			}
		
		artStyle = self.winwordBorderObject.ArtStyle
		if artStyle == 0:
			return ""
		try:
			desc = artStyleDescriptions[artStyle]
		except:
			desc = ""
		return desc

