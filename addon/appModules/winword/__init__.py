# appModules\winword\__init__.py
#A part of wordAccessEnhancement add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

import addonHandler
addonHandler.initTranslation()
import appModuleHandler
import globalVars
from logHandler import log
import time
import scriptHandler
import api
import os
import wx
import ui
import tones
import eventHandler
import queueHandler
import keyboardHandler
import textInfos
import sayAllHandler
import controlTypes
import speech
import braille
from NVDAObjects.window.winword import wdParagraph, wdSentence,wdActiveEndPageNumber , wdFirstCharacterLineNumber ,WordDocumentTextInfo
import NVDAObjects.IAccessible.winword
import inputCore
import browseMode
import winUser
import oleacc
from comtypes import COMError, GUID, BSTR
import comtypes.client
import comtypes.automation
# appmodule module import
from .ww_wdConst import wdHorizontalPositionRelativeToPage , wdVerticalPositionRelativeToPage ,wdFirstCharacterColumnNumber , wdWord
from . import ww_comments
from . import ww_headings
from . import ww_tables
from . import ww_choice
from .ww_comments import Comments
from .ww_revisions import Revisions
from .ww_browseMode import *
from . import ww_document
import sys
_curAddon = addonHandler.getCodeAddon()
path = os.path.join(_curAddon.path, "shared")
sys.path.append(path)
from ww_utils  import printDebug, maximizeWindow, toggleDebugFlag
from ww_informationDialog import InformationDialog
from ww_NVDAStrings import NVDAString
from ww_py3Compatibility import _unicode
from ww_addonConfigManager import _addonConfigManager 
del sys.path[-1]


_addonSummary = _unicode(_curAddon.manifest['summary'])
_scriptCategory =  _addonSummary

# global timer
GB_scriptTimer = None
# maximum delay for waiting new script call
_delay = 250

def stopScriptTimer():
	global GB_scriptTimer
	if GB_scriptTimer != None:
		GB_scriptTimer.Stop()
		GB_scriptTimer = None
		
class MainWordDocumentEx(NVDAObjects.IAccessible.winword.WordDocument):
	scriptCategory = _scriptCategory
	treeInterceptorClass=WordDocumentTreeInterceptorEx
	shouldCreateTreeInterceptor=False
	TextInfo=WordDocumentTextInfo
	
	_commonGestures =  {
		"toggleLayerMode" : ("kb:nvda+e",),
		"toggleReportAllCellsFlag": ("kb:windows+alt+space",),
		"report_location": ("kb:alt+numpaddelete","kb(laptop):alt+delete"),
		# word shortcuts
		# move sentence by sentence
		"nextSentence": ("kb:alt+downArrow",),
		"previousSentence": ("kb:alt+upArrow" , ),
		# report element text
		"reportCurrentEndNoteOrFootNote" : ("kb:windows+alt+n", ),
		"reportCurrentRevision": ("kb:windows+alt+m",),
		"insertComment": ("kb:windows+alt+f2",),
		"reportDocumentInformations": ("kb:windows+alt+f1",),
		"makeChoice": ("kb:windows+alt+f5",),
		# tables
		# report table element
		"reportCurrentHeadersEx": ("kb:windows+alt+h",),
		"reportCurrentCell": ("kb:windows+alt+j",),
		"report_currentRow": ("kb:windows+Alt+;",),
		"report_currentColumn": ("kb:windows+Alt+,",),
		"reportFirstCellOfColumn": ("kb:windows+alt+pageUp",),
		"reportLastCellOfColumn": ("kb:windows+alt+pageDown",),
		"reportFirstCellOfRow": ("kb:windows+alt+home",),
		"reportLastCellOfRow": ("kb:windows+alt+end",),
		"reportNextCellInRow": ("kb:windows+alt+rightArrow",),
		"reportPreviousCellInRow": ("kb:windows+alt+leftArrow",),
		"reportNextCellInColumn": ("kb:windows+alt+downArrow",),
		"reportPreviousCellInColumn": ("kb:windows+alt+upArrow",),
		# move to  table element and say all cells of row or column
		"moveToAndReportNextColumn": None,
		"moveToAndReportPreviousColumn": None,
		"moveToAndReportNextRow": None,
		"moveToAndReportPreviousRow": None,
		"moveToStartOfRowAndReportColumn": None,
		"moveToEndOfRowAndReportColumn": None,
		"moveToStartOfColumnAndReportRow": None,
		"moveToEndOfColumnAndReportRow": None,
		}

	def initOverlayClass(self):
		printDebug ("MainWordDocumentEx InitOverlayClass")
		self.postOverlayClassInitGestureMap = self._gestureMap.copy()
		if self.appModule.layerMode is None:
			self.appModule.layerMode = False
		self._updateGestureBinding()
	
	def event_focusEntered(self):
		printDebug("MainWordDocumentEx: event_focusEntered")
		super( MainWordDocumentEx, self).event_focusEntered()
	def event_gainFocus(self):
		printDebug("MainWordDocumentEx: event_gainFocus")
		super( MainWordDocumentEx, self).event_gainFocus()
		globalVars.lastWinwordObject = self
		if self.appModule.layerMode:
			speech.speakMessage(_("Table layer mode on"))

	def _updateGestureBinding(self):
		printDebug ("updateGestureBinding: %s"%self.appModule.layerMode)
		self._gestureMap.clear()
		self._gestureMap.update(self.postOverlayClassInitGestureMap)
		for script in self._commonGestures:
			gests = self._commonGestures[script]
			if gests is None: continue
			for gest in gests:
				self.bindGesture(gest, script)
		if self.appModule.layerMode:
			gestures = self._layerGestures
			oldGestures = self._baseGestures
			scriptCategory = _scriptCategory + _(" (Table layer mode)")
		else:
			gestures = self._baseGestures
			oldGestures = self._layerGestures
			scriptCategory = _scriptCategory
		
		for script in gestures:
			gests = gestures.get(script)
			if gests is not None:
				for gest in gests:
					self.bindGesture(gest, script)

		# remove gest and doc of previous binding
		for script in oldGestures:
			scr = "scrip_%s"%script
			scriptFunc = getattr(self, "script_%s"%script)
			scriptFunc.__func__.__doc__ = None
		
		# set doc of new scripts
		for script in gestures:
			docScriptName = script.replace("layer_", "")
			scr = "scrip_%s"%script
			doc = "_%s"%docScriptName
			scriptFunc = getattr(self, "script_%s"%script)
			docFunc = getattr(self, doc)
			scriptFunc.__func__.__doc__ = docFunc.__doc__
			scriptFunc.__func__.category = scriptCategory

	def _reportLayerModeState(self, state):
		if state:
			speech.speakMessage(_("Table layer mode on"))
		else:
			speech.speakMessage(_("Table layer mode off"))
	def script_toggleLayerMode(self, gesture):
		self.appModule.layerMode = not self.appModule.layerMode
		self._reportLayerModeState(self.appModule.layerMode )
		self._updateGestureBinding()

	
	script_toggleLayerMode.__doc__ = _("Toggle table layer mode")
	def _exitLayerMode(self):
		self.appModule.layerMode = False
		self._reportLayerModeState(self.appModule.layerMode )
		self._updateLayerMode()
	
	def script_toggleReportAllCellsFlag(self, gesture):
		self.appModule.reportAllCellsFlag = not self.appModule.reportAllCellsFlag
		if self.appModule.reportAllCellsFlag:
			# Translators: message to user to indicate  report all cells mode is activated.
			msg = _("Report all Cells")
		else:
			# Translators: message to user to indicate report all cells mode is desactivated.
			msg = _("Do not report all cells")
		speech.speakMessage(msg)
	script_toggleReportAllCellsFlag.__doc__ = _("activate or desactivate the report of all cells of row or column")
	
	def inTable(self, withMessage = False):
		doc = self.WinwordDocumentObject
		selection = self.WinwordSelectionObject
		selection.Collapse()
		r = doc.range (selection.Start, selection.End)
		inTable = False
		try:
			if r.Tables.Count == 1 and r.Cells.Count == 1:
				inTable = True
		except:
			pass
		
		if not inTable and withMessage:
			# Translators: The message reported when a user attempts to use a table movement command.
			# when the cursor is not withnin a table.
			ui.message(_("Not in table"))
		return inTable
	def getReferenceAtFocus(self, fieldType):
		info=self.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_CHARACTER)
		fields=info.getTextWithFields(formatConfig={'reportComments':True})
		for field in reversed(fields):
			if isinstance(field,textInfos.FieldCommand) and isinstance(field.field,textInfos.FormatField): 
				reference=field.field.get(fieldType)
				if reference:
					return reference
		return None
	
	def reportCurrentElement(self, elementClass):
		col = elementClass(None,self, "focus")
		col.reportElements()
	def _getPageAndLineNumber(self):
		(start, end) = (self.WinwordSelectionObject.Start, self.WinwordSelectionObject.End)
		try:
			r = self.WinwordDocumentObject.range (start, start)
		except:
			log.warning ("WordDocumentEx getPageAndLineNumber:  cannot get rangeat start")
			return None
		if self.inTable():
			return None

		line = r.information(wdFirstCharacterLineNumber )
		page = r.Information(wdActiveEndPageNumber )
		return (page, line)
	def playSoundOnSkippedParagraph(self):

		option = _addonConfigManager .togglePlaySoundOnSkippedParagraphOption(False)
		if option:
			tones.beep(400,12)
	
	def _moveToNextOrPriorElement(self,direction , type ="paragraph"):
		def move(wdUnit, direction):
			info=self.makeTextInfo(textInfos.POSITION_CARET)
			if not info:
				return
			# #4375: can't use self.move here as it may check document.chracters.count which can take for ever on large documents.
			info._rangeObj.move(wdUnit, direction)
			info.updateCaret()
		
		stopScriptTimer()
		res = self._getPageAndLineNumber()
		if not res:
			# no document or not in text
			return
		if res == (1,1) and direction == -1:
			# top of document
			return
		if type == "paragraph":
			unit = textInfos.UNIT_PARAGRAPH
			wdUnit = wdParagraph
		else:
			unit = textInfos.UNIT_SENTENCE
			wdUnit = wdSentence
		move(wdUnit, direction)
		option = _addonConfigManager.toggleSkipEmptyParagraphsOption(False)
		if type != "paragraph" or not option:
			return
		# we skip a maximum  of empty paragraph
		i= 100
		playSound = False
		while i:
			i = i -1
			res = self._getPageAndLineNumber()
			if not res:
				return
			if res == (1,1):
				return
			try:
				info=self.makeTextInfo(textInfos.POSITION_CARET)
			except:
				return
			info.expand(unit)
			text = info.text.strip()
			info.collapse()
			if len(text) != 0:
				if playSound:
					self.playSoundOnSkippedParagraph()
				return
			move(wdUnit, direction)
			playSound = True
	
	def script_nextSentence(self,gesture):
		self._moveToNextOrPriorElement(1, "sentence")
		self._caretScriptPostMovedHelper(textInfos.UNIT_SENTENCE,gesture,None)
	
	def script_previousSentence(self,gesture):
		self._moveToNextOrPriorElement(-1, "sentence")
		self._caretScriptPostMovedHelper(textInfos.UNIT_SENTENCE,gesture,None)
	# Translators: a description for a script.
	script_previousSentence.__doc__ = _("Move to the previous sentence")
	script_previousSentence.category = _scriptCategory
	script_previousSentence.resumeSayAllMode=sayAllHandler.CURSOR_CARET
	# Translators: a description for a script.
	script_nextSentence.__doc__ = _("Move to the next sentence")
	script_nextSentence.category = _scriptCategory
	script_nextSentence.resumeSayAllMode=sayAllHandler.CURSOR_CARET	
	
	def script_nextParagraph(self,gesture):
		self._moveToNextOrPriorElement(1, "paragraph")
		self._caretScriptPostMovedHelper(textInfos.UNIT_PARAGRAPH,gesture,None)
	
	def script_previousParagraph(self,gesture):
		self._moveToNextOrPriorElement(-1, "paragraph")
		self._caretScriptPostMovedHelper(textInfos.UNIT_PARAGRAPH,gesture,None)

	script_previousParagraph.resumeSayAllMode=sayAllHandler.CURSOR_CARET
	# Translators: a description for a script.
	script_previousParagraph.__doc__ = _("Move to the previous paragraph( skip empty paragraphs if SkipEmptyParagraph flag is set to on) ")
	script_previousParagraph.category = _scriptCategory
	script_nextParagraph.resumeSayAllMode=sayAllHandler.CURSOR_CARET
	# Translators: a description for a script.
	script_nextParagraph.__doc__ = _("Move to the next paragraph( skip empty paragraphs if SkipEmptyParagraph flag is set to on) ")
	script_nextParagraph.category = _scriptCategory
	
	def script_insertComment(self,gesture):
		printDebug ("insert comment")
		stopScriptTimer()
		winwordApplication = self._get_WinwordApplicationObject()
		if winwordApplication :
			wx.CallAfter(ww_comments.insertComment, winwordApplication)
		else:
			log.warning(" No winword application")
	# Translators: a description for a script.
	script_insertComment.__doc__ =_("Insert a comment where the System caret is located.")			
	def script_reportCurrentEndNoteOrFootNote(self, gesture):
		stopScriptTimer()
		info=self.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_CHARACTER)
		fields=info.getTextWithFields(formatConfig={'reportComments':True})
		for field in reversed(fields):
			if not(isinstance(field,textInfos.FieldCommand) and isinstance(field.field,textInfos.ControlField)): 
				continue
			role=field.field.get('role')
			start = self.WinwordDocumentObject.content.start
			end = self.WinwordDocumentObject.content.end
			range=self.WinwordDocumentObject.range(start,end)
			if role == controlTypes.ROLE_FOOTNOTE:
				val =field.field.get('value')
				try:
					col = range.footNotes
					text=col[int(val)].range.text
					if text == None or len(text) == 0:
						# Translators: message to the user to say  empty .
						text = _("empty")
					ui.message(text)
					return
				except :
					break
			if role == controlTypes.ROLE_ENDNOTE:
				val =field.field.get('value')
				try:
					text=range.EndNotes[int(val)].range.text
					if text == None or len(text) == 0:
						# Translators: message to the user to say  empty .
						text = _("empty")
					ui.message(text)
					return
				except :
					printDebug ("endnotes")
					break
				
				return
		# Translators: a message when there is no endnoe or footnote to report in Microsoft Word.
		ui.message(_("no endnote or footnote"))
	# Translators: a description for a script.
	script_reportCurrentEndNoteOrFootNote.__doc__=_("Reports the text of the endnote or footnote where the System caret is located.") 		
	def script_reportCurrentRevision(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self.reportCurrentElement, Revisions)
	# Translators: a description for a script.
	script_reportCurrentRevision.__doc__=_("Reports the text of the revision where the System caret is located.")	
	def script_makeChoice(self, gesture, ):
		printDebug ("_makeChoice")
		stopScriptTimer()
		wx.CallAfter(ww_choice.ChoiceDialog.run, self)

	# Translators: a description for a script.
	script_makeChoice.__doc__ = _("Display a dialog  to get word object informations in a range")	
	def script_report_location(self, gesture):
		stopScriptTimer()
		doc = self.WinwordDocumentObject
		(start, end) = (self.WinwordSelectionObject.Start, self.WinwordSelectionObject.End)
		r = doc.range (start, start)
		left= self.getLocalizedMeasurementTextForPointSize (r.information(wdHorizontalPositionRelativeToPage))
		top=self.getLocalizedMeasurementTextForPointSize( r.information(wdVerticalPositionRelativeToPage))
		# Translators:  message in report location script.
		position = _("at {0} from left edge and {1} from top edge of the page").format(left, top)
		if start == end:
			if self.inTable():
				cell = r.Cells[0]
				(row, col) = (cell.rowIndex, cell.columnIndex)
				# Translators: message in report location script.
				location = _("You are at the row {row}, column {column}  of table") .format(row = row, column = col)
				text = "%s, %s." %(location, position)
			else:
				line = r.information(wdFirstCharacterLineNumber )
				column = r.information(wdFirstCharacterColumnNumber )
				page = r.Information(wdActiveEndPageNumber )
				location = _("you are at Line {line}, column {column} of page {page}") .format(line = line, column = column, page= page)
				text = "%s, %s." %(location, position)
		else:
			line = r.information(wdFirstCharacterLineNumber )
			column = r.information(wdFirstCharacterColumnNumber )
			page = r.Information(wdActiveEndPageNumber )
			location = _("Selection begins at line {line}, column {column} of page {page}").format(line =line, column = column, page= page)
			text = "%s, %s." %(location, position)
			r = doc.range (end-1, end)
			line = r.information(wdFirstCharacterLineNumber )
			column = r.information(wdFirstCharacterColumnNumber )
			page = r.Information(wdActiveEndPageNumber )
			left= self.getLocalizedMeasurementTextForPointSize (r.information(wdHorizontalPositionRelativeToPage))
			top=self.getLocalizedMeasurementTextForPointSize( r.information(wdVerticalPositionRelativeToPage))
			position = _("at {0} from left edge and {1} from top edge of the page.").format(left, top)
			location = _("It ends at line {line}, column {column} of page {page}").format(line = line, column = column, page= page)
			text = text + "\r\n" + "%s, %s" %(location, position)
		count = scriptHandler.getLastScriptRepeatCount()
		if count > 0:
			# Translators: this is the title of informationdialog box to show appModule informations.
			dialogTitle = _("Location information")
			InformationDialog.run(None, dialogTitle, "", text)
		else:
			ui.message(text)
		

	
	# Translators: a description for a script.
	script_report_location.__doc__ = _("Report the  page number and line number relative to the page. Twice: display this information")
	
	def script_reportCurrentHeadersEx(self,gesture):
		stopScriptTimer()
		if not self.inTable(True):
			return
		cell=self.WinwordSelectionObject.cells[1]
		# Translators: message to say text is empty.
		emptyText = _("Empty")
		rowText=self.fetchAssociatedHeaderCellText(cell,False) 
		if rowText == None:
			# Translators: message to say no row header.
			rowText = _("No row header")
		elif len(rowText) == 2:
			# Translators: message to report row header.
			rowText = _("Row header: %s") %emptyText
		else:
			# Translators: message to report row header.
			rowText = _("Row header: %s") %rowText
		columnText=self.fetchAssociatedHeaderCellText(cell,True)
		if columnText == None:
			# Translators: message to report no column header.
			columnText = _("No column header")
		elif len(columnText) == 2 :
			# Translators: message to report column header.
			columnText = _("Column header: %s") % emptyText
		else:
			# Translators: message to report column header.
			columnText = _("Column header: %s") %columnText
		ui.message("%s , %s" %(rowText, columnText))
	# Translators: a description for a script.
	script_reportCurrentHeadersEx.__doc__ = _("Table: report current headers")
	
	def script_reportDocumentInformations(self, gesture):
		activeDocument= ww_document.ActiveDocument(self)
		wx.CallAfter(activeDocument.reportDocumentInformations)
	
	# Translators: a description for a script.
	script_reportDocumentInformations.__doc__ = _("Display document's informations")


		
	def _reportTableElement(self, elementType = "cell",position = "current", reportAllCells = None):
		printDebug ("reportTableElement: elementType = %s,position= %s"%(elementType, position))
		stopScriptTimer()
		if not self.inTable(True):
			return
		
		self.doc = self.WinwordDocumentObject
		(start, end) = (self.WinwordSelectionObject.Start, self.WinwordSelectionObject.End)
		r = self.doc.range (start, start)
		cell = ww_tables.Cell(self, r.Cells[0])
		table = ww_tables.Table(self, r.Tables[0])
		table.sayElement( elementType, position,cell, self.appModule.reportAllCellsFlag if reportAllCells is None  else reportAllCells)
	
	def script_reportCurrentCell(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement,elementType = "cell", position = "current", reportAllCells = False)
	script_reportCurrentCell.__doc__ = _("Table: report current cell")
	def script_report_currentRow(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement,elementType = "row", position = "current")
	# Translators: a description for a script.
	script_report_currentRow.__doc__ = _("Table: report all cells of current  row")
	def script_report_currentColumn(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement,elementType = "column", position = "current")

	# Translators: a description for a script.
	script_report_currentColumn.__doc__ = _("Table: report all cells  of current column")
	
	def script_reportNextCellInColumn(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement, elementType = "cell", position = "nextInColumn")
	# Translators: a description for a script.
	script_reportNextCellInColumn.__doc__ = _("Table: Report the nextt cell in the column")
	
	def script_reportPreviousCellInColumn (self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement, elementType = "cell", position = "previousInColumn")
	# Translators: a description for a script.
	script_reportPreviousCellInColumn.__doc__ = _("Table: report the previous cell in the column")
	
	def script_reportNextCellInRow(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement, elementType = "cell", position = "nextInRow")
	# Translators: a description for a script.
	script_reportNextCellInRow.__doc__ = _("Table: report the next cell in the  row")
	
	def script_reportPreviousCellInRow (self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement, elementType = "cell", position = "previousInRow")
	# Translators: a description for a script.
	script_reportPreviousCellInRow.__doc__ = _("Table: report the previous cell in the row")
	
	def script_reportFirstCellOfColumn(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement, elementType = "cell", position = "firstInColumn")
	# Translators: a description for a script.
	script_reportFirstCellOfColumn.__doc__ = _("Table: report the first cell of the  column")
	
	def script_reportLastCellOfColumn (self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement, elementType = "cell", position = "lastInColumn")
	# Translators: a description for a script.
	script_reportLastCellOfColumn.__doc__ = _("Table: report the last cell of the column")


	def script_reportFirstCellOfRow(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement, elementType = "cell", position = "firstInRow")
	# Translators: a description for a script.
	script_reportFirstCellOfRow.__doc__ = _("Table: report the first cell of the  row")
	
	def script_reportLastCellOfRow (self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._reportTableElement, elementType = "cell", position = "lastInRow")
	# Translators: a description for a script.
	script_reportLastCellOfRow.__doc__ = _("Table: report the last cell of the row")

	def _moveToTableElement(self, position = "current", reportRow = None):
		printDebug ("_moveToTableElement: position= %s, reportRow= %s"%(position, reportRow))
		stopScriptTimer()
		if not self.inTable(True):
			return
		
		self.doc = self.WinwordDocumentObject
		(start, end) = (self.WinwordSelectionObject.Start, self.WinwordSelectionObject.End)
		r = self.doc.range (start, start)
		cell = ww_tables.Cell(self, r.Cells[0])
		table = ww_tables.Table(self, r.Tables[0])
		cell = table.moveToCell( position, cell)
		if cell is None:
			return
		moveInRow = table.getMoveInRow(position)
		if  position== "firstInRow":
			# Translators: a message to say moving to start of row.
			ui.message(_("Start of row {0}") .format(cell.rowIndex) if cell != None else "")
			# Translators: a message to report column number.
			msg = _(" column {0}").format(cell.columnIndex)
		elif  position == "lastInRow":
			# Translators: message to say moving to end of row.
			ui.message( _("End of row {0}") .format(cell.rowIndex) if cell != None else "")
			# Translators: message to report row number.
			msg = _(" column {0}").format(cell.columnIndex)
		elif position == "firstInColumn":
			# Translators: message to say moving to start of column.
			ui.message(_("Start of column {0}") .format( cell.columnIndex) if cell != None else "")
			# Translators: message to report row number.
			msg = _(" row {0}").format(cell.rowIndex)
		elif position == "lastInColumn":
			# Translators: message to say moving to end of column
			ui.message(_("End of column {0}") .format(cell.columnIndex) if cell != None else "")
			# Translators: message to report row number.
			msg = _(" row {0}").format(cell.rowIndex)
		else:
			msg = None
			if moveInRow is True:
				# Translators: message to report column number.
				msg = _("column %s") %cell.columnIndex
			elif moveInRow is False:
				# Translators: message to report row number.
				msg = _("row %s") %cell.rowIndex
		info=self.makeTextInfo(textInfos.POSITION_CARET)
		info.updateCaret()
		if moveInRow is not None and (self.appModule.reportAllCellsFlag or reportRow is not None):
			elementType = "column" if moveInRow else "row"
			table.sayElement( elementType, "current", cell)
			return
		if msg is not None:
			ui.message(msg)
		cell.sayText(self, moveInRow)
	
	def _moveToNextRow(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "nextInColumn")
	
	# Translators: a description for a script.
	_moveToNextRow.__doc__ = _("Table: move to next row")
	
	def _moveToPreviousRow(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "previousInColumn")
	# Translators: a description for a script.
	_moveToPreviousRow.__doc__ = _("Table: move to previous  row")
	
	def _moveToNextColumn(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "nextInRow")
	# Translators: a description for a script.
	_moveToNextColumn.__doc__ = _("Table: move to next column")
	
	def _moveToPreviousColumn(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "previousInRow")
	# Translators: a description for a script.
	_moveToPreviousColumn.__doc__ = _("Table: move to  previous column ")
	
	def _moveToFirstColumn(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "firstInRow")
	# Translators: a description for a script.
	_moveToFirstColumn.__doc__ = _("Table: move to the first cell of row")
	
	def _moveToLastColumn(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "lastInRow")
	# Translators: a description for a script.
	_moveToLastColumn.__doc__ = _("Table: move to  the last cell of row")
	
	def _moveToFirstRow(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "firstInColumn")
	# Translators: a description for a script.
	_moveToFirstRow.__doc__ = _("Table: move to the first cell of column")
	
	def _moveToLastRow(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement,  position = "lastInColumn")
	# Translators: a description for a script.
	_moveToLastRow.__doc__ = _("Table: move to  the last cell of column ")



	def script_moveToAndReportNextRow(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "nextInColumn", reportRow= True)
	# Translators: a description for a script.
	script_moveToAndReportNextRow.__doc__ = _("Table: move to next row and report all cells of row")
	
	def script_moveToAndReportPreviousRow(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "previousInColumn", reportRow= True)
	# Translators: a description for a script.
	script_moveToAndReportPreviousRow.__doc__ = _("Table: move to prior row and report all cells of row")
	
	def script_moveToAndReportNextColumn(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "nextInRow", reportRow= False)
	# Translators: a description for a script.
	script_moveToAndReportNextColumn.__doc__ = _("Table: move to next columnand report all cells of column")
	
	def script_moveToAndReportPreviousColumn(self,gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "previousInRow", reportRow= False)
	# Translators: a description for a script.
	script_moveToAndReportPreviousColumn.__doc__ = _("Table: move to  prior column  and report all cells of column")
	
	def script_moveToStartOfColumnAndReportRow(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "firstInColumn", reportRow= True)
	# Translators: a description for a script.
	script_moveToStartOfColumnAndReportRow.__ = _("Table: move to first cell of column and report all cells of row")
	
	def script_moveToEndOfColumnAndReportRow(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "lastInColumn", reportRow= True)
	# Translators: a description for a script.
	script_moveToEndOfColumnAndReportRow.__doc__ = _("Table: move to last cell of column and report all cells of row")
	
	def script_moveToStartOfRowAndReportColumn (self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "firstInRow", reportRow= False)
	# Translators: a description for a script.
	script_moveToStartOfRowAndReportColumn.__doc__ = _("Table: move to first cell of the row and report all cells of column")
	
	def script_moveToEndOfRowAndReportColumn (self, gesture):
		stopScriptTimer()
		wx.CallAfter(self._moveToTableElement, position = "lastInRow", reportRow= False)
	# Translators: a description for a script.
	script_moveToEndOfRowAndReportColumn.__doc__ = _("Table: mMove to last cell of the row and report all cells of column")
	

class WordDocumentEx(MainWordDocumentEx):

	_baseGestures =  {
				# move to element
		"moveToNextColumn": ("kb:windows+control+rightArrow",),
		"moveToPreviousColumn": ("kb:windows+control+leftArrow",),
		"moveToNextRow": ("kb:windows+control+downArrow",),
		"moveToPreviousRow": ("kb:windows+control+upArrow",),
		"moveToFirstColumn": ("kb:windows+control+home",),
		"moveToLastColumn": ("kb:windows+control+end",),
		"moveToFirstRow": ("kb:windows+control+pageUp",),
		"moveToLastRow": ("kb:windows+control+pageDown",),
		}
	_layerGestures = {
				# move to elements
		"layer_moveToNextColumn": ("kb:rightArrow",),
		"layer_moveToPreviousColumn": ("kb:leftArrow",),
		"layer_moveToNextRow": ("kb:downArrow",),
		"layer_moveToPreviousRow": ("kb:upArrow",),
		"layer_moveToFirstColumn": ("kb:home",),
		"layer_moveToLastColumn": ("kb:end",),
		"layer_moveToFirstRow": ("kb:pageUp",),
		"layer_moveToLastRow": ("kb:pageDown",),
		}
	
	#def script_insertComment(self,gesture):
		#printDebug ("script_insert comment")
		#self._insertComment(gesture)



	def script_moveToNextRow(self,gesture):
		self._moveToNextRow(gesture)
	
	def script_moveToPreviousRow(self,gesture):
		self._moveToPreviousRow(gesture)

	
	def script_moveToNextColumn(self,gesture):
		self._moveToNextColumn(gesture)

	
	def script_moveToPreviousColumn(self,gesture):
		self._moveToPreviousColumn(gesture)

	def script_moveToFirstColumn(self,gesture):
		self._moveToFirstColumn(gesture)
	
	def script_moveToLastColumn(self,gesture):
		self._moveToLastColumn(gesture)


	def script_moveToFirstRow(self,gesture):
		self._moveToFirstRow(gesture)
	
	def script_moveToLastRow(self,gesture):
		self._moveToLastRow(gesture)
		
		# scripts when layer mode is on
	def script_layer_moveToNextRow(self,gesture):
		self._moveToNextRow(gesture)
	
	def script_layer_moveToPreviousRow(self,gesture):
		self._moveToPreviousRow(gesture)
	
	def script_layer_moveToNextColumn(self,gesture):
		self._moveToNextColumn(gesture)
	
	def script_layer_moveToPreviousColumn(self,gesture):
		self._moveToPreviousColumn(gesture)
	
	def script_layer_moveToFirstColumn(self,gesture):
		self._moveToFirstColumn(gesture)
	
	def script_layer_moveToLastColumn(self,gesture):
		self._moveToLastColumn(gesture)
	
	def script_layer_moveToFirstRow(self,gesture):
		self._moveToFirstRow(gesture)
	
	def script_layer_moveToLastRow(self,gesture):
		self._moveToLastRow(gesture)


	
	
	


class AppModule(appModuleHandler.AppModule):
	layerMode = None
	reportAllCellsFlag = False
	
	def __init__(self, *args, **kwargs):
		printDebug ("word appmodule init")
		super(AppModule, self).__init__(*args, **kwargs)
		# configuration load
		self.hasFocus = False
	
	def terminate (self):
		if hasattr(self, "checkObjectsTimer") and self.checkObjectsTimer is not None:
			self.checkObjectsTimer.Stop()
			self.checkObjectsTimer = None
		super(AppModule, self).terminate()	
	
	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		#printDebug ("Word Appmodule ChooseNVDAOverlayClass : %s, %s, clsList:%s" %(obj.role, obj.windowClassName,clsList))

		if NVDAObjects.IAccessible.winword.WordDocument in clsList:
			clsList.insert(0, WordDocumentEx)
			clsList.remove(NVDAObjects.IAccessible.winword.WordDocument)
		
	
	def event_appModule_gainFocus(self):
		printDebug ("Word: event_appModuleGainFocus")
		self.hasFocus = True
	def event_appModule_loseFocus(self):
		printDebug ("Word: event_appModuleLoseFocus")
		self.hasFocus = False

	
	def event_foreground(self, obj, nextHandler):
		printDebug ("word: event_foreground")
		if obj.windowClassName == "OpusApp":
			self.opusApp = obj
		if obj.windowClassName in ["_WwG"]:
			maximizeWindow(obj.windowHandle)
		nextHandler()
	
	
	def event_gainFocus(self, obj, nextHandler):
		printDebug ("Word: event_gainFocus: %s, %s" %(obj.role, obj.windowClassName))
		if not hasattr(self, "WinwordWindowObject"):
			try:
				self.WinwordWindowObject = obj.WinwordWindowObject
				self.WinwordVersion = obj.WinwordVersion
			except:
				pass
		if not self.hasFocus:
			nextHandler()
			return
		if obj.windowClassName == "OpusApp":
			# to suppress double announce of document window title
			return
		#for spelling and grammar ending check window
		if obj.role == controlTypes.ROLE_BUTTON and obj.name.lower() == "ok":
			foreground = api.getForegroundObject()
			if foreground.windowClassName == "#32770" and foreground.name == "Microsoft Word":
				lastChild = foreground.getChild(foreground.childCount-1)
				if lastChild.windowClassName == "MSOUNISTAT":
					speech.speakMessage(foreground.description)
		nextHandler()
	def isSupportedVersion(self, obj = None):
		if obj is None: obj = self
		if hasattr(obj, "WinwordVersion") and obj.WinwordVersion in [15.0, 16.0]:
			return True
		return False
	
	def event_typedCharacter(self, obj, nextHandler, ch):
		nextHandler()
		if self.isSupportedVersion():
			# only for office 2013 and 2016
			from .ww_spellingChecker import SpellingChecker
			sc = SpellingChecker(obj, self.WinwordVersion)
			if sc.inSpellingChecker():
				if (ch > "a" and ch < "z"
					or ord(ch) in [wx.WXK_SPACE, wx.WXK_RETURN] and obj.role == controlTypes.ROLE_BUTTON):
					self.script_spellingCheckerHelper(None)
	
	def script_toggleSkipEmptyParagraphsOption(self, gesture):
		if _addonConfigManager.toggleSkipEmptyParagraphsOption():
			# Translators: message to report skipping of empty paragraph when moving by paragraph.
			speech.speakMessage( _("Skip empty paragraphs"))
		else:
			# Translators: message to report no skipping empty paragraph when moving by paragraph.
			speech.speakMessage(_("Don't skip empty paragraphs"))
	# Translators: a description for a script.
	script_toggleSkipEmptyParagraphsOption.__doc__ = _("Toggle on or off the option to skip empty paragraphs")
	script_toggleSkipEmptyParagraphsOption.category = _scriptCategory
	
	def script_EscapeKey(self,gesture):
		if self.layerMode:
			focus = api.getFocusObject()
			try:
				focus._exitLayerMode()
			except:
				pass
		if not self.isSupportedVersion():
			gesture.send()
			return
		from .ww_spellingChecker import SpellingChecker
		focus = api.getFocusObject()
		sc = SpellingChecker(focus, self.WinwordVersion)
		if not sc.inSpellingChecker():
			gesture.send()
			return
		gesture.send()
		time.sleep(0.1)
		api.processPendingEvents()
		speech.cancelSpeech()
		speech.speakObject(api.getFocusObject(), controlTypes.REASON_FOCUS)		
	
	def script_f7KeyStroke(self, gesture):
		def verify(oldSpeechMode):
			api.processPendingEvents()
			speech.cancelSpeech()
			speech.speechMode =oldSpeechMode
			focus = api.getFocusObject()
			from .ww_spellingChecker import SpellingChecker
			sc = SpellingChecker(focus, self.WinwordVersion)
			if not sc.inSpellingChecker():
				return
			if focus.role == controlTypes.ROLE_PANE:
				# focus on the pane not  not on an object of the pane
				# Translators: message to ask user to hit tab key.
				queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, _("Hit tab to move focus in the spelling checker pane"))
			else:
				sc.sayErrorAndSuggestion(title = True, spell = False)
				queueHandler.queueFunction(queueHandler.eventQueue,speech.speakObject, focus, controlTypes.REASON_FOCUS)
		
		stopScriptTimer()
		if not self.isSupportedVersion():
			gesture.send()
			return
		focus =api.getFocusObject()
		from .ww_spellingChecker import SpellingChecker
		sc = SpellingChecker(focus, self.WinwordVersion)
		if not sc.inSpellingChecker():
			# moving to spelling checker
			oldSpeechMode = speech.speechMode
			speech.speechMode = speech.speechMode_off
			gesture.send()
			core.callLater(500, verify, oldSpeechMode)
			return
		# moving to document selection
		winwordWindowObject = self.WinwordWindowObject
		selection = winwordWindowObject.Selection
		selection.Characters(1).GoTo()
		time.sleep(0.1)
		api.processPendingEvents()
		speech.cancelSpeech()
		speech.speakObject(api.getFocusObject(), controlTypes.REASON_FOCUS)

	
	def script_spellingCheckerHelper(self, gesture):
		global GB_scriptTimer
		if not self.isSupportedVersion():
			# Translators: message to the user when word version is not supported.
			speech.speakMessage(_("Not available for this Word version"))
			return
		stopScriptTimer()
		focus =api.getFocusObject()
		from .ww_spellingChecker import SpellingChecker
		sc = SpellingChecker(focus, self.WinwordVersion)
		if not sc.inSpellingChecker():
			# Translators: message to  indicate the focus is not in spellAndGrammar checker.
			queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, _("You are Not in the spelling checker")) 
			return
		if focus.role == controlTypes.ROLE_PANE:
			# focus on the pane not  not on an object of the pane
			# Translators: message to ask user to hit tab key.
			queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, _("Hit tab to move focus in the spelling checker pane"))
			return
		count = scriptHandler.getLastScriptRepeatCount()
		count = scriptHandler.getLastScriptRepeatCount() 
		if count == 0:
			GB_scriptTimer = core.callLater(_delay, sc.sayErrorAndSuggestion,  spell = False)
		elif count == 1:
			GB_scriptTimer = core.callLater(_delay, sc.sayErrorAndSuggestion,spell = True)
		else:
			wx.CallAfter(sc.sayHelpText)
	# Translators: a description for a script.
	script_spellingCheckerHelper.__doc__ = _("Report error and suggestion displayed by the spelling checker.Twice:spell them. Third: say help text")
	script_spellingCheckerHelper.category = _scriptCategory
	
	def sayCurrentSentence(self):
		winwordWindowObject = self.WinwordWindowObject
		selection = winwordWindowObject.Selection
		queueHandler.queueFunction(queueHandler.eventQueue,ui.message, selection.Sentences(1).Text)

	def script_reportCurrentSentence(self, gesture):
		
		stopScriptTimer()
		wx.CallAfter(self.sayCurrentSentence)
	# Translators: a description for a script.
	script_reportCurrentSentence.__doc__ = _("Report current  sentence  ")
	script_reportCurrentSentence.category = _scriptCategory
	
	def script_toggleDebugFlag(self,gesture):
		if toggleDebugFlag():
			speech.speakMessage("Word debug On")
		else:
			speech.speakMessage("Word debug off")
	def script_testWord(self, gesture):
		print ("testWord")
		ui.message("test word")
	
	__gestures ={
		# for spelling checker
		"kb:f7": "f7KeyStroke",
		"kb:escape": "EscapeKey",
		"kb:nvda+shift+f7" : "spellingCheckerHelper",
		"kb:nvda+control+f7": "reportCurrentSentence",
		# for empty paragraph
		"kb:windows+alt+f4": "toggleSkipEmptyParagraphsOption",
		"kb:windows+alt+f9": "toggleDebugFlag",
		"kb:alt+control+f10":"testWord",

	}
