# appModules\winword\ww_tables.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
from logHandler import log
import config
import ui
import speech
import time
from .ww_wdConst import wdGoToTable
from .ww_collection import Collection, CollectionElement, ReportDialog
import sys
import os
_curAddon = addonHandler.getCodeAddon()
debugToolsPath = os.path.join(_curAddon.path, "debugTools")
sys.path.append(debugToolsPath)
try:
	from appModuleDebug import printDebug, toggleDebugFlag
except ImportError:
	def printDebug(msg): return
	def toggleDebugFlag(): return
del sys.path[-1]

addonHandler.initTranslation()


class Table(CollectionElement):
	_rowPositions = {
		"nextInRow": "next",
		"previousInRow": "previous",
		"firstInRow": "first", "lastInRow": "last",
		"current": "current"}
	_columnPositions = {
		"nextInColumn": "next",
		"previousInColumn": "previous",
		"firstInColumn": "first",
		"lastInColumn": "last",
		"current": "current"}

	def __init__(self, parent, tableObj):
		super(Table, self).__init__(parent, tableObj)
		self.columnsCount = tableObj.Columns.Count
		self.rowsCount = tableObj.Rows.Count
		self.range = tableObj.Range
		self.title = ""
		self.uniform = tableObj.uniform
		self.start = tableObj.range.Start
		self.setLineAndPageNumber()

	def formatInfos(self):
		sInfo = _("""Page {page}, line {line}
Title: {title}
Number of rows: {rowsCount}
Number of columns: {columnsCount}
""")
		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(
			page=self.page,
			line=self.line,
			title=self.title,
			rowsCount=self.rowsCount,
			columnsCount=self.columnsCount)

	def getMoveInRow(self, position):
		if position == "current":
			return None
		if position in self._rowPositions:
			return True
		elif position in self._columnPositions:
			return False
		return None

	def _getCellIndexInCollection(self, cell, collection):
		for i in range(0, len(collection)):
			element = collection[i]
			if (element.rowIndex == cell.rowIndex)\
				and (element.columnIndex == cell.columnIndex):
				return i
		return None

	def moveToCell(self, position, curCell):
		(rowIndex, columnIndex) = (curCell.rowIndex, curCell.columnIndex)
		printDebug("moveToCell: position = %s, rowIndex: %s, columnIndex= %s" % (position, rowIndex, columnIndex))  # noqa:E501
		if position in self._rowPositions:
			cells = self._getCellsOfRow(rowIndex)
			newCell = self._getCellInCollection(
				cells,
				self._rowPositions[position],
				self._getCellIndexInCollection(curCell, cells))
			if newCell is None or newCell.columnIndex == columnIndex:
				ui.message(_("Row's limit"))
				newCell = None
		elif position in self._columnPositions:
			cells = self._getCellsOfColumn(columnIndex)
			newCell = self._getCellInCollection(
				cells,
				self._columnPositions[position],
				self._getCellIndexInCollection(curCell, cells))
			if newCell is None or newCell.rowIndex == rowIndex:
				ui.message(_("Column's limit"))
				newCell = None
		elif position == "firstCellOfTable":
			cells = self._getCellsOfRow(1)
			newCell = self._getCellInCollection(cells, "first", None)
			if newCell is None or (
				(newCell.columnIndex == columnIndex)
				and (newCell.rowIndex == rowIndex)):
				ui.message(_("Table's start"))
				newCell = None
		elif position == "lastCellOfTable":
			cells = self._getCellsOfRow(self.obj.Rows.Count)
			newCell = self._getCellInCollection(cells, "last", None)
			if newCell is None or (
				(newCell.columnIndex == columnIndex)
				and (newCell.rowIndex == rowIndex)):
				ui.message(_("Table's end"))
				newCell = None
		else:
			# error
			log.warning("moveToCell error: %s, %s" % (type, whichOne))
			return None
		if newCell is None:
			return None
		newCell.goTo()
		if position in ["nextInRow", "nextInColumn", "lastInRow", "lastInColumn", "lastCellOfTable"] and self.isLastCellOfTable(newCell):  # noqa:E501
			speech.speakMessage(_("Last cell"))
		elif position in ["previousInRow", "previousInColumn", "firstInRow", "firstInColumn", "firstCellOfTable"] and self.isFirstCellOfTable(newCell):  # noqa:E501
			speech.speakMessage(_("First cell"))
		return newCell

	def _getCellInCollection(self, cells, position, index):
		printDebug("_getCellInCollection: position= %s" % position)
		try:
			if position == "previous":
				cell = cells[index-1 if index > 0 else None]
			elif position == "next":
				cell = cells[index+1 if index+1 < len(cells) else None]
			elif position == "first":
				cell = cells[0]
			elif position == "last":
				cell = cells[-1]
			elif position == "current":
				cell = cells[index]
			else:
				# error
				cell = None
		except:  # noqa:E722
			cell = None
		return cell

	def _getCellsOfRow(self, rowIndex):
		cells = []
		for i in range(1, self.columnsCount+1):
			try:
				cell = self.obj.cell(rowIndex, i)

			except:  # noqa:E722
				continue
			cells.append(Cell(self, cell))

		return cells

	def _getCellsOfColumn(self, columnIndex):
		cells = []
		for i in range(1, self.rowsCount+1):
			try:
				cell = self.obj.cell(i, columnIndex)
			except:  # noqa:E722
				continue
			cells.append(Cell(self, cell))
		return cells

	def sayElement(
		self, elementType, position, currentCell, reportAllCells=False):
		printDebug("sayElement: %s, %s, reportAllCells = %s" % (elementType, position, reportAllCells))  # noqa:E501
		if elementType == "row":
			self.sayRow(self._rowPositions[position], currentCell)
		elif elementType == "column":
			self.sayColumn(self._columnPositions[position], currentCell)
		elif elementType == "cell":
			if reportAllCells:
				if position in self._rowPositions:
					pos = self._rowPositions[position]
					self.sayColumn(pos, currentCell)
				elif position in self._columnPositions:
					pos = self._columnPositions[position]
					self.sayRow(pos, currentCell)
			else:
				self.sayCell(position, currentCell)
		else:
			# error
			log.error("SayElement invalid parameters %s" % type)

	def sayCell(self, position="current", currentCell=None, columnHeader=None):
		printDebug("sayCell: position= %s, columnHeader = %s, row= %s, column=%s" % (position, columnHeader, currentCell.rowIndex, currentCell.columnIndex))  # noqa:E501
		newCell = None
		row = None
		if position == "current":
			currentCell.sayText(self.parent, reportColumnHeader=columnHeader)
			return
		if position in self._rowPositions:
			cells = self._getCellsOfRow(currentCell.rowIndex)
			row = True
			pos = self._rowPositions[position]
		elif position in self._columnPositions:
			cells = self._getCellsOfColumn(currentCell.columnIndex)
			row = False
			pos = self._columnPositions[position]
		else:
			# error
			log.error("error, sayCell invalid parameter %s" % whichOne)
			return
		if newCell is None:
			index = self._getCellIndexInCollection(currentCell, cells)
			newCell = self._getCellInCollection(cells, pos, index)
			if newCell is None:
				if row is None:
					pass
				elif row:
					ui.message(_("Row's limit"))
				else:
					ui.message(_("Column's limit"))
				return
			if row is True:
				ui.message(_("column %s") % newCell.columnIndex)
			elif row is False:
				ui.message(_("row %s") % newCell.rowIndex)
			newCell.sayText(
				self.parent,
				reportColumnHeader=row if columnHeader is None else columnHeader)

	def sayRow(self, position, currentCell):
		cells = self._getCellsOfColumn(currentCell.columnIndex)
		index = self._getCellIndexInCollection(currentCell, cells)
		cell = self._getCellInCollection(cells, position, index)
		if cell is None:
			ui.message(_("Column's limit"))
			return
		cells = self._getCellsOfRow(cell.rowIndex)
		ui.message(_("row %s") % cell.rowIndex)
		for cell in cells:
			time.sleep(0.01)
			self.sayCell("current", cell, columnHeader=True)

	def sayColumn(self, position, currentCell):
		cells = self._getCellsOfRow(currentCell.rowIndex)
		index = self._getCellIndexInCollection(currentCell, cells)
		cell = self._getCellInCollection(cells, position, index)
		if cell is None:
			ui.message(_("Row's limit"))
			return
		ui.message(_("column %s") % cell.columnIndex)
		cells = self._getCellsOfColumn(cell.columnIndex)
		for cell in cells:
			time.sleep(0.01)
			self.sayCell("current", cell, columnHeader=False)

	def isLastCellOfTable(self, currentCell=None):
		if currentCell.rowIndex != self.obj.Rows.Count:
			return False
		cells = self._getCellsOfRow(currentCell.rowIndex)
		pos = self._rowPositions["nextInRow"]
		index = self._getCellIndexInCollection(currentCell, cells)
		newCell = self._getCellInCollection(cells, pos, index)
		if newCell:
			return False
		cells = self._getCellsOfColumn(currentCell.columnIndex)
		pos = self._columnPositions["nextInColumn"]
		index = self._getCellIndexInCollection(currentCell, cells)
		newCell = self._getCellInCollection(cells, pos, index)
		if newCell:
			return False
		return True

	def isFirstCellOfTable(self, currentCell):
		if currentCell.rowIndex != 1:
			return False
		cells = self._getCellsOfRow(currentCell.rowIndex)
		pos = self._rowPositions["previousInRow"]
		index = self._getCellIndexInCollection(currentCell, cells)
		newCell = self._getCellInCollection(cells, pos, index)
		if newCell:
			return False
		cells = self._getCellsOfColumn(currentCell.columnIndex)
		pos = self._columnPositions["previousInColumn"]
		index = self._getCellIndexInCollection(currentCell, cells)
		newCell = self._getCellInCollection(cells, pos, index)
		if newCell:
			return False
		return True

	def isUniform(self):
		return self.uniform


class Tables(Collection):
	_propertyName = (("Tables", Table),)
	_name = (_("Table"), _("Tables"))
	_wdGoToItem = wdGoToTable

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = TablesDialog
		self.noElement = _("No table")
		super(Tables, self).__init__(parent, focus)


class TablesDialog(ReportDialog):
	def __init__(self, parent, obj):
		super(TablesDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		self.lcLabel = _("Tables")
		self.lcColumns = (
			(_("Table"), 100),
			(_("Location"), 150),
			(_("Title"), 300),
			(_("Number of rows"), 100),
			(_("Number of columns"), 100)
			)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]

		self.lcSize = (lcWidth, self._defaultLCWidth)

		self.buttons = (
			(100, _("&Go to"), self.goTo),
			(101, _("&Delete"), self.delete),
		)
		self.tc1 = None
		self.tc2 = None

	def get_lcColumnsDatas(self, element):
		location = (
			_("Page {page}, line {line}")).format(page=element.page, line=element.line)
		index = self.collection.index(element)+1

		datas = (
			index,
			location,
			element.title,
			element.rowsCount,
			element.columnsCount)
		return datas


class Cell(CollectionElement):
	def __init__(self, parent, cellObj):
		super(Cell, self).__init__(parent, cellObj)
		self.wordDocument = self.parent.parent
		self.columnIndex = cellObj.ColumnIndex
		self.rowIndex = cellObj.RowIndex
		self.range = cellObj.Range
		self.start = self.range.Start

	def sayText(self, wordDocument, reportColumnHeader=None):
		printDebug("cell setText: columnHeader= %s" % reportColumnHeader)
		headerText = ""
		reportTableHeadersFlag = config.conf['documentFormatting']["reportTableHeaders"]  # noqa:E501
		if reportTableHeadersFlag:
			if reportColumnHeader is not None:
				headerText = wordDocument.fetchAssociatedHeaderCellText(
					self.obj,
					reportColumnHeader)
				if headerText is not None:
					ui.message(headerText)
			else:
				columnHeaderText = wordDocument.fetchAssociatedHeaderCellText(
					self.obj,
					columnHeader=True)
				rowHeaderText = wordDocument.fetchAssociatedHeaderCellText(
					self.obj,
					columnHeader=False)
				if rowHeaderText is not None:
					ui.message(rowHeaderText)
				if columnHeaderText is not None:
					ui.message(columnHeaderText)
		text = self.obj.Range.Text
		text = text[:-2]
		if text != "":
			ui.message(text)
		else:
			ui.message(_("empty"))

	def select(self):
		self.obj.Select()
