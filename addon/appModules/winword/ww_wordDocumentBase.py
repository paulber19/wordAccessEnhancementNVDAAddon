# appModules\winword\wordDocumentBase.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2020-2021 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import globalVars
from logHandler import log
import scriptHandler
from scriptHandler import script
import gui
import queueHandler
import os
import wx
import ui
import tones
import textInfos
try:
	# for nvda version < 2021.1
	from sayAllHandler import CURSOR_CARET
except (AttributeError, ImportError):
	from speech.sayAll import CURSOR
	CURSOR_CARET =  CURSOR.CARET
import speech
import braille
import NVDAObjects.window.winword
from NVDAObjects.window.winword import wdParagraph, wdSentence,\
	wdActiveEndPageNumber, wdFirstCharacterLineNumber
import NVDAObjects.IAccessible.winword
from .ww_wdConst import wdHorizontalPositionRelativeToPage,\
	wdVerticalPositionRelativeToPage, wdFirstCharacterColumnNumber
from . import ww_tables
from . import ww_choice
from .ww_revisions import Revisions
from .ww_browseMode import *  # noqa:F403
from . import ww_document
from .ww_scriptTimer import stopScriptTimer
import sys
try:
	# for nvda version >= 2021.2
	from controlTypes.role import Role
	ROLE_FOOTNOTE = Role.FOOTNOTE
	ROLE_ENDNOTE = Role.ENDNOTE
	from controlTypes.outputReason import OutputReason
	REASON_CARET = OutputReason.CARET
	REASON_FOCUS = OutputReason.FOCUS
except ImportError:
	from controlTypes import (
	ROLE_FOOTNOTE, ROLE_ENDNOTE,
	)
	try:
		# for nvda version == 2021.1
		from controlTypes import OutputReason
		REASON_CARET = OutputReason.CARET
		REASON_FOCUS = OutputReason.FOCUS
	except AttributeError:
		# fornvda version <  2020.1
		REASON_CARET = controlTypes.REASON_CARET
		REASON_FOCUS = controlTypes.REASON_FOCUS



_curAddon = addonHandler.getCodeAddon()
debugToolsPath = os.path.join(_curAddon.path, "debugTools")
sys.path.append(debugToolsPath)
try:
	from appModuleDebug import printDebug, toggleDebugFlag
except ImportError:
	def printDebug(msg): return
	def toggleDebugFlag(): return
del sys.path[-1]

sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_informationDialog import InformationDialog  # noqa:E402
from ww_utils import makeAddonWindowTitle, isOpened
from ww_addonConfigManager import _addonConfigManager  # noqa:E402

del sys.path[-1]

addonHandler.initTranslation()
_addonSummary = _curAddon.manifest['summary']
_scriptCategory = _addonSummary


class ScriptsForTable(NVDAObjects.NVDAObject):
	scriptCategory = _scriptCategory
	_commonGestures = {
	# report table element
		"reportCurrentHeadersEx": ("kb:windows+alt+h",),
		"reportCurrentCell": ("kb:windows+alt+j",),
		"report_currentRow": ("kb:windows+Alt+k",),
		"report_currentColumn": ("kb:windows+Alt+l",),
		"reportFirstCellOfColumn": ("kb:windows+alt+pageUp",),
		"reportLastCellOfColumn": ("kb:windows+alt+pageDown",),
		"reportFirstCellOfRow": ("kb:windows+alt+home",),
		"reportLastCellOfRow": ("kb:windows+alt+end",),
		"reportNextCellInRow": ("kb:windows+alt+rightArrow",),
		"reportPreviousCellInRow": ("kb:windows+alt+leftArrow",),
		"reportNextCellInColumn": ("kb:windows+alt+downArrow",),
		"reportPreviousCellInColumn": ("kb:windows+alt+upArrow",),
		# move to table element and say all cells of row or column
		"moveToAndReportNextColumn": None,
		"moveToAndReportPreviousColumn": None,
		"moveToAndReportNextRow": None,
		"moveToAndReportPreviousRow": None,
		"moveToStartOfRowAndReportColumn": None,
		"moveToEndOfRowAndReportColumn": None,
		"moveToStartOfColumnAndReportRow": None,
		"moveToEndOfColumnAndReportRow": None,
		}

	_baseGestures = {
		"moveToNextColumn": ("kb:windows+control+rightArrow",),
		"moveToPreviousColumn": ("kb:windows+control+leftArrow",),
		"moveToNextRow": ("kb:windows+control+downArrow",),
		"moveToPreviousRow": ("kb:windows+control+upArrow",),
		"moveToFirstColumn": ("kb:windows+control+home",),
		"moveToLastColumn": ("kb:windows+control+end",),
		"moveToFirstRow": ("kb:windows+control+pageUp",),
		"moveToLastRow": ("kb:windows+control+pageDown",),
		"moveToFirstCellOfTable": ("kb:windows+control+k",),
		"moveToLastCellOfTable": ("kb:windows+control+l",),
		}
	_layerGestures = {
		"layer_moveToNextColumn": ("kb:rightArrow",),
		"layer_moveToPreviousColumn": ("kb:leftArrow",),
		"layer_moveToNextRow": ("kb:downArrow",),
		"layer_moveToPreviousRow": ("kb:upArrow",),
		"layer_moveToFirstColumn": ("kb:home",),
		"layer_moveToLastColumn": ("kb:end",),
		"layer_moveToFirstRow": ("kb:pageUp",),
		"layer_moveToLastRow": ("kb:pageDown",),
		"layer_moveToFirstCellOfTable": ("kb:k",),
		"layer_moveToLastCellOfTable": ("kb:l",),
		}

	def _moveInTable(self, row=True, forward=True):
		res = super(ScriptsForTable, self)._moveInTable(row, forward)
		if res:
			# selection of the cell for copy with "control+c"
			start = self.WinwordSelectionObject.Start
			end = self.WinwordSelectionObject.End
			r = self.WinwordDocumentObject.range(start, start)
			if start == end:
				if self.inTable():
					cell = r.Cells[0]
					cell.Select()
		return res

	def _moveToTableElement(self, position="current", reportRow=None):
		printDebug("_moveToTableElement: position= %s, reportRow= %s" % (
			position, reportRow))
		stopScriptTimer()
		if not self.inTable(True):
			return
		self.doc = self.WinwordDocumentObject
		start = self.WinwordSelectionObject.Start
		r = self.doc.range(start, start)
		cell = ww_tables.Cell(self, r.Cells[0])
		table = ww_tables.Table(self, r.Tables[0])
		cell = table.moveToCell(position, cell)
		if cell is None:
			return
		moveInRow = table.getMoveInRow(position)
		if config.conf["documentFormatting"]["reportTableCellCoords"]:
			# Translators: a message to report row number.
			msgRowIndex = _(" row %s") % cell.rowIndex
			# Translators: a message to report column number.
			msgColumnIndex = _(" column %s") % cell.columnIndex
			# Translators: message to report row and column number.
			msgRowColumnIndex = _(" row {rowIndex}, column {columnIndex}").format(
				rowIndex=cell.rowIndex, columnIndex=cell.columnIndex)
		else:
			msgRowIndex = msgColumnIndex = msgRowColumnIndex = None
		if position == "firstInRow":
			# Translators: a message to say moving to start of row.
			ui.message(_("Start of row {0}") .format(
				cell.rowIndex) if cell is not None else "")
			msg = msgColumnIndex
		elif position == "lastInRow":
			# Translators: message to say moving to end of row.
			ui.message(_("End of row {0}") .format(
				cell.rowIndex) if cell is not None else "")
			msg = msgColumnIndex
		elif position == "firstInColumn":
			# Translators: message to say moving to start of column.
			ui.message(_("Start of column {0}") .format(
				cell.columnIndex) if cell is not None else "")
			msg = msgRowIndex
		elif position == "lastInColumn":
			# Translators: message to say moving to end of column
			ui.message(_("End of column {0}") .format(
				cell.columnIndex) if cell is not None else "")
			msg = msgRowIndex
		elif position in ["firstCellOfTable", "lastCellOfTable"]:
			msg = msgRowColumnIndex
		else:
			msg = None
			if moveInRow is True:
				msg = msgColumnIndex
			elif moveInRow is False:
				msg = msgRowIndex
		info = self.makeTextInfo(textInfos.POSITION_CARET)
		info.updateCaret()
		if moveInRow is not None\
			and (self.appModule.reportAllCellsFlag or reportRow is not None):
			elementType = "column" if moveInRow else "row"
			table.sayElement(elementType, "current", cell)
			# selection of the current cell for copy by "control+c"
			self.WinwordSelectionObject.SelectCell()
			return
		if msg is not None:
			ui.message(msg)
		reportColumnHeader = True\
			if moveInRow is not None or position == "lastCellOfTable" else False
		cell.sayText(self, reportColumnHeader)
		# selection of the current cell for copy by "control+c"
		cell.select()

	def _moveToNextRow(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="nextInColumn")

	# Translators: a description for a script.
	_moveToNextRow.__doc__ = _("Table: move to next row")

	def _moveToPreviousRow(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="previousInColumn")
	# Translators: a description for a script.
	_moveToPreviousRow.__doc__ = _("Table: move to previous row")

	def _moveToNextColumn(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="nextInRow")
	# Translators: a description for a script.
	_moveToNextColumn.__doc__ = _("Table: move to next column")

	def _moveToPreviousColumn(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="previousInRow")
	# Translators: a description for a script.
	_moveToPreviousColumn.__doc__ = _("Table: move to previous column ")

	def _moveToFirstColumn(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="firstInRow")
	# Translators: a description for a script.
	_moveToFirstColumn.__doc__ = _("Table: move to the first cell of row")

	def _moveToLastColumn(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="lastInRow")
	# Translators: a description for a script.
	_moveToLastColumn.__doc__ = _("Table: move to the last cell of row")

	def _moveToFirstRow(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="firstInColumn")
	# Translators: a description for a script.
	_moveToFirstRow.__doc__ = _("Table: move to the first cell of column")

	def _moveToLastRow(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="lastInColumn")
	# Translators: a description for a script.
	_moveToLastRow.__doc__ = _("Table: move to the last cell of column ")

	def _moveToFirstCellOfTable(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="firstCellOfTable")
	# Translators: a description for a script.
	_moveToFirstCellOfTable.__doc__ = _("Table: move to first cell of table")

	def _moveToLastCellOfTable(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="lastCellOfTable")
	# Translators: a description for a script.
	_moveToLastCellOfTable.__doc__ = _("Table: move to last cell of table")


	# scripts when layer mode is off
	
	def script_moveToNextRow(self, gesture):
		self._moveToNextRow(gesture)

	def script_moveToPreviousRow(self, gesture):
		self._moveToPreviousRow(gesture)

	def script_moveToNextColumn(self, gesture):
		self._moveToNextColumn(gesture)

	def script_moveToPreviousColumn(self, gesture):
		self._moveToPreviousColumn(gesture)

	def script_moveToFirstColumn(self, gesture):
		self._moveToFirstColumn(gesture)

	def script_moveToLastColumn(self, gesture):
		self._moveToLastColumn(gesture)

	def script_moveToFirstRow(self, gesture):
		self._moveToFirstRow(gesture)

	def script_moveToLastRow(self, gesture):
		self._moveToLastRow(gesture)

	def script_moveToFirstCellOfTable(self, gesture):
		self._moveToFirstCellOfTable(gesture)

	def script_moveToLastCellOfTable(self, gesture):
		self._moveToLastCellOfTable(gesture)

	# scripts when layer mode is on
	def script_layer_moveToNextRow(self, gesture):
		self._moveToNextRow(gesture)

	def script_layer_moveToPreviousRow(self, gesture):
		self._moveToPreviousRow(gesture)

	def script_layer_moveToNextColumn(self, gesture):
		self._moveToNextColumn(gesture)

	def script_layer_moveToPreviousColumn(self, gesture):
		self._moveToPreviousColumn(gesture)

	def script_layer_moveToFirstColumn(self, gesture):
		self._moveToFirstColumn(gesture)

	def script_layer_moveToLastColumn(self, gesture):
		self._moveToLastColumn(gesture)

	def script_layer_moveToFirstRow(self, gesture):
		self._moveToFirstRow(gesture)

	def script_layer_moveToLastRow(self, gesture):
		self._moveToLastRow(gesture)

	def script_layer_moveToFirstCellOfTable(self, gesture):
		self._moveToFirstCellOfTable(gesture)

	def script_layer_moveToLastCellOfTable(self, gesture):
		self._moveToLastCellOfTable(gesture)


	def _reportTableElement(
		self, elementType="cell", position="current", reportAllCells=None):
		printDebug("reportTableElement: elementType = %s,position= %s" % (
			elementType, position)
			)
		stopScriptTimer()
		if not self.inTable(True):
			return
		self.doc = self.WinwordDocumentObject
		start = self.WinwordSelectionObject.Start
		r = self.doc.range(start, start)
		cell = ww_tables.Cell(self, r.Cells[0])
		table = ww_tables.Table(self, r.Tables[0])
		table.sayElement(
			elementType,
			position,
			cell,
			self.appModule.reportAllCellsFlag if reportAllCells is None else reportAllCells)  # noqa:E501

	@script(
		# Translators: Input help mode message for report current headers command.
		description=_("Table: report current headers"),
	)
	def script_reportCurrentHeadersEx(self, gesture):
		stopScriptTimer()
		if not self.inTable(True):
			return
		cell = self.WinwordSelectionObject.cells[1]
		# Translators: message to say text is empty.
		emptyText = _("Empty")
		rowText = self.fetchAssociatedHeaderCellText(cell, False)
		if rowText is None:
			# Translators: message to say no row header.
			rowText = _("No row header")
		elif len(rowText) == 2:
			# Translators: message to report row header.
			rowText = _("Row header: %s") % emptyText
		else:
			# Translators: message to report row header.
			rowText = _("Row header: %s") % rowText
		columnText = self.fetchAssociatedHeaderCellText(cell, True)
		if columnText is None:
			# Translators: message to report no column header.
			columnText = _("No column header")
		elif len(columnText) == 2:
			# Translators: message to report column header.
			columnText = _("Column header: %s") % emptyText
		else:
			# Translators: message to report column header.
			columnText = _("Column header: %s") % columnText
		ui.message("%s , %s" % (rowText, columnText))

	@script(
		# Translators: Input help mode message for report current cell command.
		description=_("Table: report current cell"),
	)
	def script_reportCurrentCell(self, gesture):
		wx.CallAfter(
			self._reportTableElement,
			elementType="cell",
			position="current",
			reportAllCells=False)

	@script(
		# Translators: Input help mode message for report current row command.
		description=_("Table: report all cells of current row"),
	)
	def script_report_currentRow(self, gesture):
		wx.CallAfter(self._reportTableElement, elementType="row", position="current")

	@script(
		# Translators: Input help mode message for report current column command.
		description=_("Table: report all cells of current column"),
	)
	def script_report_currentColumn(self, gesture):
		wx.CallAfter(
			self._reportTableElement, elementType="column", position="current")

	@script(
		# Translators: Input help mode message for report next cell in column command.
		description=_("Table: Report the nextt cell in the column"),
	)
	def script_reportNextCellInColumn(self, gesture):
		wx.CallAfter(
			self._reportTableElement, elementType="cell", position="nextInColumn")

	@script(
		# Translators: Input help mode message for report previous cell in column command.
		description=_("Table: report the previous cell in the column"),
	)
	def script_reportPreviousCellInColumn(self, gesture):
		wx.CallAfter(
			self._reportTableElement, elementType="cell", position="previousInColumn")

	@script(
		# Translators: Input help mode message for report next cell in row command.
		description=_("Table: report the next cell in the row"),
	)
	def script_reportNextCellInRow(self, gesture):
		wx.CallAfter(
			self._reportTableElement, elementType="cell", position="nextInRow")

	@script(
		# Translators: Input help mode message for report previous cell in row command.
		description=_("Table: report the previous cell in the row"),
	)
	def script_reportPreviousCellInRow(self, gesture):
		wx.CallAfter(
			self._reportTableElement, elementType="cell", position="previousInRow")

	@script(
		# Translators: Input help mode message for report first cell of column command.
		description=_("Table: report the first cell of the column"),
	)
	def script_reportFirstCellOfColumn(self, gesture):
		wx.CallAfter(
			self._reportTableElement, elementType="cell", position="firstInColumn")

	@script(
		# Translators: Input help mode message for report last cell of column command.
		description=_("Table: report the last cell of the column")
	)
	def script_reportLastCellOfColumn(self, gesture):  # noqa:E501
		wx.CallAfter(
			self._reportTableElement, elementType="cell", position="lastInColumn")

	@script(
		# Translators: Input help mode message for report first cell of row command.
		description=_("Table: report the first cell of the row"),
	)
	def script_reportFirstCellOfRow(self, gesture):
		wx.CallAfter(
			self._reportTableElement, elementType="cell", position="firstInRow")

	@script(
		# Translators: Input help mode message for report last cell of row command.
		description=_("Table: report the last cell of the row"),
	)
	def script_reportLastCellOfRow(self, gesture):
		wx.CallAfter(
			self._reportTableElement, elementType="cell", position="lastInRow")


	@script(
		# Translators: Input help mode message for move to and report next row command.
		description=_("Table: move to next row and report all cells of row"),
	)
	def script_moveToAndReportNextRow(self, gesture):
		wx.CallAfter(
			self._moveToTableElement, position="nextInColumn", reportRow=True)

	@script(
		# Translators: Input help mode message for move to and report Previous row command.
		description=_("Table: move to previous row and report all cells of row"),
	)
	def script_moveToAndReportPreviousRow(self, gesture):
		wx.CallAfter(
			self._moveToTableElement, position="previousInColumn", reportRow=True)

	@script(
		# Translators: Input help mode message for move to and report next column command.
		description=_("Table: move to next column and report all cells of column"),
	)
	def script_moveToAndReportNextColumn(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="nextInRow", reportRow=False)

	@script(
		# Translators: Input help mode message for move to and report previous column command.
		description=_("Table: move to prior column and report all cells of column"),
	)
	def script_moveToAndReportPreviousColumn(self, gesture):
		wx.CallAfter(
			self._moveToTableElement, position="previousInRow", reportRow=False)

	@script(
		# Translators: Input help mode message for move to start ofcolumn and report row command.
		description=_(
			"Table: move to first cell of column and report all cells of row"
		),
	)
	def script_moveToStartOfColumnAndReportRow(self, gesture):
		wx.CallAfter(
			self._moveToTableElement, position="firstInColumn", reportRow=True)

	@script(
		# Translators: Input help mode message for move to end of column and report row command.
		description=_(
			"Table: move to last cell of column and report all cells of row"
		),
	)
	def script_moveToEndOfColumnAndReportRow(self, gesture):
		wx.CallAfter(
			self._moveToTableElement, position="lastInColumn", reportRow=True)

	@script(
		# Translators: Input help mode message for move to start of row and report column command.
		description=_(
			"Table: move to first cell of the row and report all cells of column"
		),
	)
	def script_moveToStartOfRowAndReportColumn(self, gesture):
		wx.CallAfter(
			self._moveToTableElement, position="firstInRow", reportRow=False)

	@script(
		# Translators: Input help mode message for move to end of row and report column command.
		description=_(
			"Table: move to last cell of the row and report all cells of column"
		),
	)
	def script_moveToEndOfRowAndReportColumn(self, gesture):
		wx.CallAfter(self._moveToTableElement, position="lastInRow", reportRow=False)



class WordDocument(ScriptsForTable, NVDAObjects.NVDAObject):
	scriptCategory = _scriptCategory
	_mainGestures = {
		"toggleLayerMode": ("kb:nvda+e",),
		"EscapeKey": ("kb:escape",),
		"toggleReportAllCellsFlag": ("kb:windows+alt+space",),
		"reportDocumentInformations": ("kb:windows+alt+f1",),
		"insertElement": ("kb:windows+alt+f2",),
		"toggleAutomaticReading": ("kb:windows+alt+f3",),
		"makeChoice": ("kb:windows+alt+f5",),
		"report_location": ("kb:alt+numpaddelete", "kb(laptop):alt+delete"),
		# word shortcuts
		"tab": ("kb:tab",),
		# move sentence by sentence
		"nextSentence": ("kb:alt+downArrow",),
		"previousSentence": ("kb:alt+upArrow", ),
		# report element text
		"reportCurrentEndNoteOrFootNote": ("kb:windows+alt+n", ),
		"reportCurrentRevision": ("kb:windows+alt+m",),
				}

	def initOverlayClass(self):
		printDebug("WordDocument InitOverlayClass")
		self.postOverlayClassInitGestureMap = self._gestureMap.copy()
		if self.appModule.layerMode is None:
			self.appModule.layerMode = False
		self._updateGestureBinding()

	def event_gainFocus(self):
		printDebug("WordDocument: event_gainFocus")
		super(WordDocument, self).event_gainFocus()
		globalVars.lastWinwordObject = self
		if self.appModule.layerMode:
			ui.message(_("Table layer mode on"))

	def _updateGestureBinding(self):
		printDebug("updateGestureBinding: %s" % self.appModule.layerMode)
		self._gestureMap.clear()
		self._gestureMap.update(self.postOverlayClassInitGestureMap)
		tempGestures = self._mainGestures.copy()
		tempGestures.update(self._commonGestures)
		for scr in tempGestures:
			gests = tempGestures[scr]
			if gests is None:
				continue
			for gest in gests:
				self.bindGesture(gest, scr)
		if self.appModule.layerMode:
			gestures = self._layerGestures
			oldGestures = self._baseGestures
			scriptCategory = _scriptCategory + _(" (Table layer mode)")
		else:
			gestures = self._baseGestures
			oldGestures = self._layerGestures
			scriptCategory = _scriptCategory
		for scr in gestures:
			gests = gestures.get(scr)
			if gests is not None:
				for gest in gests:
					self.bindGesture(gest, scr)

		# remove gest and doc of previous binding
		for scr in oldGestures:
			scriptFunc = getattr(self, "script_%s" % scr)
			scriptFunc.__func__.__doc__ = None
		# set doc of new scripts
		for scr in gestures:
			docScriptName = scr.replace("layer_", "")
			doc = "_%s" % docScriptName
			scriptFunc = getattr(self, "script_%s" % scr)
			docFunc = getattr(self, doc)
			scriptFunc.__func__.__doc__ = docFunc.__doc__
			scriptFunc.__func__.category = scriptCategory

	def script_EscapeKey(self, gesture):
		stopScriptTimer()
		if self.appModule.layerMode:
			self._exitLayerMode()
			return
		gesture.send()

	def _reportLayerModeState(self, state):
		if state:
			ui.message(_("Table layer mode on"))
		else:
			ui.message(_("Table layer mode off"))

	@script(
		# Translators: Input help mode message for toggle layer mode command.
		description=_("Toggle table layer mode"),
	)
	def script_toggleLayerMode(self, gesture):
		stopScriptTimer()
		self.appModule.layerMode = not self.appModule.layerMode
		self._reportLayerModeState(self.appModule.layerMode)
		self._updateGestureBinding()

	def _exitLayerMode(self):
		self.appModule.layerMode = False
		self._reportLayerModeState(self.appModule.layerMode)
		self._updateGestureBinding()

	@script(
		# Translators: Input help mode message for toggle report all cells flag command.
		description=_(
			"activate or desactivate the report"
			" of all cells of row or column"),
	)
	def script_toggleReportAllCellsFlag(self, gesture):
		stopScriptTimer()
		self.appModule.reportAllCellsFlag = not self.appModule.reportAllCellsFlag
		if self.appModule.reportAllCellsFlag:
			# Translators: message to user to indicate
			# report all cells mode is activated.
			msg = _("Report all Cells")
		else:
			# Translators: message to user to indicate
			# report all cells mode is desactivated.
			msg = _("Do not report all cells")
		ui.message(msg)

	@script(
		# Translators: Input help mode message for toggle automatic reading command.
		description=_("activate or desactivate automatic reading"),
	)
	def script_toggleAutomaticReading(self, gesture):
		stopScriptTimer()
		if _addonConfigManager.toggleAutomaticReadingOption():
			# Translators: message to user to indicate automatic reading enabled.
			msg = _("Automatic reading enabled")
		else:
			# Translators: message to user to indicate automatic reading disabled.
			msg = _("Automatic reading disabled")
		ui.message(msg)

	def inTable(self, withMessage=False):
		doc = self.WinwordDocumentObject
		selection = self.WinwordSelectionObject
		inTable = selection.Information(wdWithInTable) 
		if not inTable and withMessage:
			# Translators: The message reported
			# when a user attempts to use a table movement command.
			# when the cursor is not withnin a table.
			ui.message(_("Not in table"))
		return inTable

	def getReferenceAtFocus(self, fieldType):
		info = self.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_CHARACTER)
		fields = info.getTextWithFields(formatConfig={'reportComments': True})
		for field in reversed(fields):
			if isinstance(field, textInfos.FieldCommand)\
				and isinstance(field.field, textInfos.FormatField):
				reference = field.field.get(fieldType)
				if reference:
					return reference
		return None

	def reportCurrentElement(self, elementClass):
		print ("reportCurrentElement: %s"%elementClass)
		col = elementClass(None, self, "focus")
		col.reportElements()

	def _getPageAndLineNumber(self):
		start = self.WinwordSelectionObject.Start
		try:
			r = self.WinwordDocumentObject.range(start, start)
		except:  # noqa:E722
			log.warning("WordDocumentEx getPageAndLineNumber: cannot get rangeat start")
			return None
		if self.inTable():
			return None
		line = r.information(wdFirstCharacterLineNumber)
		page = r.Information(wdActiveEndPageNumber)
		return (page, line)

	def _getPosition(self):
		start = self.WinwordSelectionObject.Start
		try:
			r = self.WinwordDocumentObject.range(start, start)
		except:  # noqa:E722
			log.warning("WordDocumentEx getPageAndLineNumber: cannot get rangeat start")
			return None
		if self.inTable():
			return None

		line = r.information(wdFirstCharacterLineNumber)
		column = r.information(wdFirstCharacterColumnNumber)
		page = r.Information(wdActiveEndPageNumber)
		return (page, line, column)

	def playSoundOnSkippedParagraph(self):

		option = _addonConfigManager .togglePlaySoundOnSkippedParagraphOption(False)
		if option:
			tones.beep(400, 12)

	def _moveToNextOrPriorElement(self, direction, type="paragraph"):

		def move(wdUnit, direction, curPosition):
			info = self.makeTextInfo(textInfos.POSITION_CARET)
			if not info:
				return
			# #4375: can't use self.move here as it may check document.
				# chracters.count which can take for ever on large documents.
			info._rangeObj.move(wdUnit, direction)
			info.updateCaret()
			position = self._getPosition()
			return position
		stopScriptTimer()
		position = self._getPosition()
		if not position:
			# no document or not in text
			return False
		if type == "paragraph":
			unit = textInfos.UNIT_PARAGRAPH
			wdUnit = wdParagraph
			msgNoOther = _("No other paragraph")
		else:
			unit = textInfos.UNIT_SENTENCE
			wdUnit = wdSentence
			msgNoOther = _("No other sentence")
			doc = self.WinwordDocumentObject
			selection = self.WinwordSelectionObject
			start = selection.Start
			end = doc.storyRanges[1].End
			r = doc.range(start, end)
			sentences = r.Sentences
			if direction == 1 and sentences.Count == 1:
				ui.message(_("No other sentence"))
				return False
		oldPosition = position
		position = move(wdUnit, direction, position)
		if position == oldPosition:
			ui.message(msgNoOther)
			return False
		option = _addonConfigManager.toggleSkipEmptyParagraphsOption(False)
		if type != "paragraph" or not option:
			return True
		# we skip a maximum of empty paragraph
		i = 100
		playSound = False
		while i:
			i = i-1
			try:
				info = self.makeTextInfo(textInfos.POSITION_CARET)
			except:  # noqa:E722
				return False
			info.expand(unit)
			text = info.text.strip()
			info.collapse()
			if len(text) != 0:
				if playSound:
					self.playSoundOnSkippedParagraph()
				return True
# move to next paragraph
			oldPosition = position
			position = move(wdUnit, direction, position)
			if not position or position == oldPosition:
				ui.message(msgNoOther)
				if playSound:
					self.playSoundOnSkippedParagraph()
				return False
			playSound = True

	@script(
		# Translators: Input help mode message for next sentence command.
		description=_("Move to the next sentence"),
	)
	def script_nextSentence(self, gesture):
		if self._moveToNextOrPriorElement(1, "sentence"):
			self._caretScriptPostMovedHelper(textInfos.UNIT_SENTENCE, gesture, None)
	script_nextSentence.resumeSayAllMode = CURSOR_CARET

	@script(
		# Translators: Input help mode message for previous sentence command.
		description=_("Move to the previous sentence"),
	)
	def script_previousSentence(self, gesture):
		if self._moveToNextOrPriorElement(-1, "sentence"):
			self._caretScriptPostMovedHelper(textInfos.UNIT_SENTENCE, gesture, None)
	script_previousSentence.resumeSayAllMode = CURSOR_CARET

	@script(
		# Translators: Input help mode message for next Paragraph command.
		description=_(
			"Move to the next paragraph "
			"""(skip empty paragraphs if "Skip Empty Paragraph" option is set to on) """
		),
	)
	def script_nextParagraph(self, gesture):
		if self._moveToNextOrPriorElement(1, "paragraph"):
			self._caretScriptPostMovedHelper(textInfos.UNIT_PARAGRAPH, gesture, None)
	script_nextParagraph.resumeSayAllMode = CURSOR_CARET

	@script(
		# Translators: Input help mode message for previous Paragraph command.
		description=_(
			"Move to the previous paragraph "
			"""(skip empty paragraphs if "Skip Empty Paragraph" option is set to on) """
		),
	)
	def script_previousParagraph(self, gesture):
		if self._moveToNextOrPriorElement(-1, "paragraph"):
			self._caretScriptPostMovedHelper(textInfos.UNIT_PARAGRAPH, gesture, None)
	script_previousParagraph.resumeSayAllMode = CURSOR_CARET

	@script(
		# Translators: Input help mode message for insert comment command.
		description=_("Insert an element (comment, footnote, endnote) where the System caret is located."),
	)
	def script_insertElement(self, gesture):
		printDebug("insert comment")
		stopScriptTimer()
		winwordApplication = self._get_WinwordApplicationObject()
		if winwordApplication:
			wx.CallAfter(InsertElementDialog.run, winwordApplication)
		else:
			log.warning(" No winword application")

	@script(
		# Translators: Input help mode message for report current endNote or footNote command.
		description=_(
			"Reports the text of the "
			"endnote or footnote where the System caret is located."),
	)
	def script_reportCurrentEndNoteOrFootNote(self, gesture):
		stopScriptTimer()
		info = self.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_CHARACTER)
		fields = info.getTextWithFields(formatConfig={'reportComments': True})
		for field in reversed(fields):
			if not(
				isinstance(field, textInfos.FieldCommand)
				and isinstance(field.field, textInfos.ControlField)):
				continue
			role = field.field.get('role')
			start = self.WinwordDocumentObject.content.start
			end = self.WinwordDocumentObject.content.end
			range = self.WinwordDocumentObject.range(start, end)
			if role == ROLE_FOOTNOTE:
				val = field.field.get('value')
				try:
					col = range.footNotes
					text = col[int(val)].range.text
					if text is None or len(text) == 0:
						# Translators: message to the user to say empty .
						text = _("empty")
					ui.message(text)
					return
				except:  # noqa:E722
					break
			if role == ROLE_ENDNOTE:
				val = field.field.get('value')
				try:
					text = range.EndNotes[int(val)].range.text
					if text is None or len(text) == 0:
						# Translators: message to the user to say empty .
						text = _("empty")
					ui.message(text)
					return
				except:  # noqa:E722
					printDebug("endnotes")
					break
				return
		# Translators: a message when there is no endnoe or footnote
				# to report in Microsoft Word.
		ui.message(_("no endnote or footnote"))

	@script(
		# Translators: Input help mode message for report current revision command.
		description=_(
			"Reports the text of the revision where the System caret is located."
		),
	)
	def script_reportCurrentRevision(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self.reportCurrentElement, Revisions)

	@script(
		# Translators: Input help mode message for make choice command.
		description=_(
			"Display a dialog to get word object informations in a range"
		),
	)
	def script_makeChoice(self, gesture, ):
		printDebug("_makeChoice")
		stopScriptTimer()
		wx.CallAfter(ww_choice.ChoiceDialog.run, self)

	@script(
		# Translators: Input help mode message for report location command.
		description=_(
			"Report the page number and line number relative to the page."
			" Twice: display this information"
		),
	)
	def script_report_location(self, gesture):
		stopScriptTimer()
		doc = self.WinwordDocumentObject
		start = self.WinwordSelectionObject.Start
		end = self.WinwordSelectionObject.End
		r = doc.range(start, start)
		left = self.getLocalizedMeasurementTextForPointSize(
			r.information(wdHorizontalPositionRelativeToPage))
		top = self.getLocalizedMeasurementTextForPointSize(
			r.information(wdVerticalPositionRelativeToPage))
		# Translators: message in report location script.
		position = _("at {0} from left edge and {1} from top edge of the page")
		position = position.format(left, top)
		if start == end:
			if self.inTable():
				cell = r.Cells[0]
				(row, col) = (cell.rowIndex, cell.columnIndex)
				# Translators: message in report location script.
				msg = _("You are at the row {row}, column {column} of table")
				location = msg.format(row=row, column=col)
				text = "%s, %s." % (location, position)
			else:
				line = r.information(wdFirstCharacterLineNumber)
				column = r.information(wdFirstCharacterColumnNumber)
				page = r.Information(wdActiveEndPageNumber)
				# Translators: message to user to report location in page.
				msg = _("you are at Line {line}, column {column} of page {page}")
				location = msg.format(line=line, column=column, page=page)
				text = "%s, %s." % (location, position)
		else:
			line = r.information(wdFirstCharacterLineNumber)
			column = r.information(wdFirstCharacterColumnNumber)
			page = r.Information(wdActiveEndPageNumber)
			# Translators: message to user
			msg = _("Selection begins at line {line}, column {column} of page {page}")
			location = msg.format(line=line, column=column, page=page)
			text = "%s, %s." % (location, position)
			r = doc.range(end-1, end)
			line = r.information(wdFirstCharacterLineNumber)
			column = r.information(wdFirstCharacterColumnNumber)
			page = r.Information(wdActiveEndPageNumber)
			left = self.getLocalizedMeasurementTextForPointSize(
				r.information(wdHorizontalPositionRelativeToPage))
			top = self.getLocalizedMeasurementTextForPointSize(
				r.information(wdVerticalPositionRelativeToPage))
			# Translators: text to report position.
			position = _("at {0} from left edge and {1} from top edge of the page.")
			position = position.format(left, top)
			# Translators: text to report location.
			location = _("It ends at line {line}, column {column} of page {page}")
			location = location.format(line=line, column=column, page=page)
			text = text + "\r\n" + "%s, %s" % (location, position)
		count = scriptHandler.getLastScriptRepeatCount()
		if count > 0:
			# Translators: this is the title of informationdialog box
			#  to show location informations.
			dialogTitle = _("Location information")
			InformationDialog.run(None, dialogTitle, "", text)
		else:
			ui.message(text)

	@script(
		# Translators: Input help mode message for report document informations command.
		description=_("Display document's informations"),
	)
	def script_reportDocumentInformations(self, gesture):
		stopScriptTimer()
		activeDocument = ww_document.ActiveDocument(self)
		wx.CallAfter(activeDocument.reportDocumentInformations)

	
	def script_tab(self, gesture):
		stopScriptTimer()
		"""
		A script for the tab key which:
		* if in a table, announces the newly selected cell or new cell where the caret is, 
		and announce if caret is on first or last cell of table.
		* If not in a table, announces the distance of the caret from the left edge of the document, and any remaining text on that line.
		"""

		gesture.send()
		selectionObj = self.WinwordSelectionObject
		inTable = self.inTable()
		info = self.makeTextInfo(textInfos.POSITION_SELECTION)
		isCollapsed = info.isCollapsed
		if inTable and isCollapsed:
			info.expand(textInfos.UNIT_PARAGRAPH)
			isCollapsed = info.isCollapsed
		if not isCollapsed:
			self.doc = self.WinwordDocumentObject
			start = self.WinwordSelectionObject.Start
			r = self.doc.range(start, start)
			cell = ww_tables.Cell(self, r.Cells[0])
			table = ww_tables.Table(self, r.Tables[0])
			if table.isLastCellOfTable(cell):
				ui.message(_("Last cell"))
			elif table.isFirstCellOfTable(cell):
				ui.message(_("First cell"))
			speech.speakTextInfo(info, reason=REASON_FOCUS)
			braille.handler.handleCaretMove(self)
		if selectionObj and isCollapsed:
			offset = selectionObj.information(wdHorizontalPositionRelativeToPage)
			msg = self.getLocalizedMeasurementTextForPointSize(offset)
			ui.message(msg)
			if selectionObj.paragraphs[1].range.start == selectionObj.start:
				info.expand(textInfos.UNIT_LINE)
				speech.speakTextInfo(
					info,
					unit=textInfos.UNIT_LINE,
					reason=REASON_CARET
				)

	def _get_ElementsListDialog(self):
		from .ww_elementsListDialog import ElementsListDialog
		return ElementsListDialog




class InsertElementDialog(wx.Dialog):
	_instance = None
	title = None

	def __new__(cls, *args, **kwargs):
		if InsertElementDialog._instance is not None:
			return InsertElementDialog._instance
		return wx.Dialog.__new__(cls)

	def __init__(self, parent, wordApp):
		if InsertElementDialog._instance is not None:
			return
		InsertElementDialog._instance = self
		# Translators: This is the title of the InsertElementDialog dialog.
		dialogTitle = _("Insert element")
		title = InsertElementDialog.title = makeAddonWindowTitle(dialogTitle)
		super(InsertElementDialog, self).__init__(parent, -1, title)
		self.wordApp = wordApp
		self.doGui()

	def doGui(self):
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		sHelper = gui.guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		# Translators: This is a label appearing on Insert Element dialog.
		labelText = _("Element's &type:")
		elementTypeList = [_("Comment"), _("Footnote"), _("Endnote")]
		self.elementTypeListBox = sHelper.addLabeledControl(
			labelText, wx.ListBox, choices=elementTypeList)
		self.elementTypeListBox.SetSelection(0)
		bHelper = sHelper.addDialogDismissButtons(
			gui.guiHelper.ButtonHelper(wx.HORIZONTAL))
		insertButton = bHelper.addButton(
			self,
			# Translators: This is a label of a button appearing
			# on Insert Element Dialog.
			label=_("&Insert")
		)
		closeButton = bHelper.addButton(self, id=wx.ID_CLOSE)
		mainSizer.Add(
			sHelper.sizer,
			border=gui.guiHelper.BORDER_FOR_DIALOGS,
			flag=wx.ALL)
		mainSizer.Fit(self)
		self.SetSizer(mainSizer)
		# events
		insertButton.Bind(wx.EVT_BUTTON, self.onInsertButton)
		closeButton.Bind(wx.EVT_BUTTON, lambda evt: self.Destroy())
		self.SetEscapeId(wx.ID_CLOSE)
		insertButton.SetDefault()
		self.elementTypeListBox.SetFocus()

	def Destroy(self):
		InsertElementDialog._instance = None
		super(InsertElementDialog, self).Destroy()

	def onInsertButton(self, evt):
		from .ww_comments import Comments
		from .ww_footnotes import Footnotes
		from .ww_endnotes import Endnotes
		classElements = (Comments, Footnotes, Endnotes)
		index = self.elementTypeListBox.GetSelection()
		classElement = classElements[index]
		with wx.TextEntryDialog(
			self,
			classElement.entryDialogStrings["entryBoxLabel"],
			classElement.entryDialogStrings["insertDialogTitle"],
			value="",
			style=wx.TextEntryDialogStyle | wx.TE_MULTILINE
			) as entryDialog:
			if entryDialog.ShowModal() != wx.ID_OK:
				return
			text = entryDialog.Value
			if len(text) == 0:
				return

			wx.CallLater(400, classElements[index].insert, self.wordApp, text)
			self.Close()
		evt.Skip()

	@classmethod
	def run(cls, wordApp):
		if isOpened(cls):
			return
		gui.mainFrame.prePopup()
		d = cls(gui.mainFrame, wordApp)
		d.CentreOnScreen()
		d.ShowModal()
		gui.mainFrame.postPopup()
