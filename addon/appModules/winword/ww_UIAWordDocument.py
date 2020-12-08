# appModules\winword\ww_UIAWordDocument.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2020 paulber19
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

import addonHandler
import os
import textInfos
import UIAHandler
from . import ww_wordDocumentBase
import NVDAObjects.UIA.wordDocument
from comtypes import COMError
from tableUtils import HeaderCellTracker
from .ww_UIABrowseMode import UIAWordBrowseModeDocument
from versionInfo import version_year, version_major
import sys
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
from ww_NVDAStrings import NVDAString  # noqa:E402
del sys.path[-1]

addonHandler.initTranslation()
"""
_NVDAVersion = [version_year, version_major]
if _NVDAVersion < [2019, 3]:
	# automatic reading not available
	WordDocumentTextInfoEx = WordDocumentTextInfo
else:
	from .ww_automaticReading import AutomaticReadingWordTextInfo
"""

_NVDAVersion = [version_year, version_major]
if _NVDAVersion < [2019, 3]:
	# automatic reading not available
	UIAWordDocumentTextInfo = NVDAObjects.UIA.wordDocument.WordDocumentTextInfo
else:
	from .ww_automaticReading import AutomaticReadingWordTextInfo

	class UIAWordDocumentTextInfo (
		AutomaticReadingWordTextInfo,
		NVDAObjects.UIA.wordDocument.WordDocumentTextInfo):
		pass


class UIAWordDocument(ww_wordDocumentBase.WordDocument, NVDAObjects.UIA.wordDocument.WordDocument):
	treeInterceptorClass = UIAWordBrowseModeDocument
	shouldCreateTreeInterceptor = False
	announceEntireNewLine = True
	TextInfo = UIAWordDocumentTextInfo
	# Microsoft Word duplicates the full title of the document on this control, which is redundant as it appears in the title of the app itself.
	name = u""

	def initOverlayClass(self):
		printDebug("UIAWordDocument InitOverlayClass")
		# textInfo.UNIT_SENTENCE does not exist for UIA,
		# so unbind next and previous sentence script
		self.removeGestureBinding("kb:alt+upArrow")
		self.removeGestureBinding("kb:alt+downArrow")

	# part of code from WordDocument class of IAccessible.winword file
	def populateHeaderCellTrackerFromHeaderRows(self, headerCellTracker, table):
		rows = table.rows
		numHeaderRows = 0
		for rowIndex in range(rows.count):
			try:
				row = rows.item(rowIndex+1)
			except COMError:
				break
			try:
				headingFormat = row.headingFormat
			except (COMError, AttributeError, NameError):
				headingFormat = 0
			if headingFormat == -1:  # is a header row
				numHeaderRows += 1
			else:
				break
		if numHeaderRows > 0:
			headerCellTracker.addHeaderCellInfo(rowNumber=1, columnNumber=1, rowSpan=numHeaderRows, isColumnHeader=True, isRowHeader=False)

	def populateHeaderCellTrackerFromBookmarks(self, headerCellTracker, bookmarks):
		for x in bookmarks:
			name = x.name
			lowerName = name.lower()
			isColumnHeader = isRowHeader = False
			if lowerName.startswith('title'):
				isColumnHeader = isRowHeader = True
			elif lowerName.startswith('columntitle'):
				isColumnHeader = True
			elif lowerName.startswith('rowtitle'):
				isRowHeader = True
			else:
				continue
			try:
				headerCell = x.range.cells.item(1)
			except COMError:
				continue
			headerCellTracker.addHeaderCellInfo(rowNumber=headerCell.rowIndex, columnNumber=headerCell.columnIndex, name=name, isColumnHeader=isColumnHeader, isRowHeader=isRowHeader)

	_curHeaderCellTrackerTable = None
	_curHeaderCellTracker = None

	def getHeaderCellTrackerForTable(self, table):
		tableRange = table.range
		# Sometimes there is a valid reference in _curHeaderCellTrackerTable,
		# but we get a COMError when accessing the range (#6827)
		try:
			tableRangesEqual = tableRange.isEqual(self._curHeaderCellTrackerTable.range)
		except (COMError, AttributeError):
			tableRangesEqual = False
		if not tableRangesEqual:
			self._curHeaderCellTracker = HeaderCellTracker()
			self.populateHeaderCellTrackerFromBookmarks(self._curHeaderCellTracker, tableRange.bookmarks)
			self.populateHeaderCellTrackerFromHeaderRows(self._curHeaderCellTracker, table)
			self._curHeaderCellTrackerTable = table
		return self._curHeaderCellTracker

	def setAsHeaderCell(self, cell, isColumnHeader=False, isRowHeader=False):
		rowNumber = cell.rowIndex
		columnNumber = cell.columnIndex
		headerCellTracker = self.getHeaderCellTrackerForTable(cell.range.tables[1])
		oldInfo = headerCellTracker.getHeaderCellInfoAt(rowNumber, columnNumber)
		if oldInfo:
			if isColumnHeader and not oldInfo.isColumnHeader:
				oldInfo.isColumnHeader = True
			elif isRowHeader and not oldInfo.isRowHeader:
				oldInfo.isRowHeader = True
			else:
				return False
			isColumnHeader = oldInfo.isColumnHeader
			isRowHeader = oldInfo.isRowHeader
		if isColumnHeader and isRowHeader:
			name = "Title_"
		elif isRowHeader:
			name = "RowTitle_"
		elif isColumnHeader:
			name = "ColumnTitle_"
		else:
			raise ValueError("One or both of isColumnHeader or isRowHeader must be True")
		name += uuid.uuid4().hex
		if oldInfo:
			self.WinwordDocumentObject.bookmarks[oldInfo.name].delete()
			oldInfo.name = name
		else:
			headerCellTracker.addHeaderCellInfo(rowNumber=rowNumber, columnNumber=columnNumber, name=name, isColumnHeader=isColumnHeader, isRowHeader=isRowHeader)
		self.WinwordDocumentObject.bookmarks.add(name, cell.range)
		return True

	def forgetHeaderCell(self, cell, isColumnHeader=False, isRowHeader=False):
		rowNumber = cell.rowIndex
		columnNumber = cell.columnIndex
		if not isColumnHeader and not isRowHeader:
			return False
		headerCellTracker = self.getHeaderCellTrackerForTable(cell.range.tables[1])
		info = headerCellTracker.getHeaderCellInfoAt(rowNumber, columnNumber)
		if not info or not hasattr(info, 'name'):
			return False
		if isColumnHeader and info.isColumnHeader:
			info.isColumnHeader = False
		elif isRowHeader and info.isRowHeader:
			info.isRowHeader = False
		else:
			return False
		headerCellTracker.removeHeaderCellInfo(info)
		self.WinwordDocumentObject.bookmarks(info.name).delete()
		if info.isColumnHeader or info.isRowHeader:
			self.setAsHeaderCell(cell, isColumnHeader=info.isColumnHeader, isRowHeader=info.isRowHeader)
		return True

	def fetchAssociatedHeaderCellText(self, cell, columnHeader=False):
		table = cell.range.tables[1]
		rowNumber = cell.rowIndex
		columnNumber = cell.columnIndex
		headerCellTracker = self.getHeaderCellTrackerForTable(table)
		for info in headerCellTracker.iterPossibleHeaderCellInfosFor(rowNumber, columnNumber, columnHeader=columnHeader):
			textList = []
			if columnHeader:
				for headerRowNumber in range(info.rowNumber, info.rowNumber + info.rowSpan):
					tempColumnNumber = columnNumber
					while tempColumnNumber >= 1:
						try:
							headerCell = table.cell(headerRowNumber, tempColumnNumber)
						except COMError:
							tempColumnNumber -= 1
							continue
						break
					textList.append(headerCell.range.text)
			else:
				for headerColumnNumber in range(info.columnNumber, info.columnNumber + info.colSpan):
					tempRowNumber = rowNumber
					while tempRowNumber >= 1:
						try:
							headerCell = table.cell(tempRowNumber, headerColumnNumber)
						except COMError:
							tempRowNumber -= 1
							continue
						break
					textList.append(headerCell.range.text)
			text = " ".join(textList)
			if text:
				return text

	#  end  of code from WordDocument class of IAccessible.winword file
	# to fix bug of NVDA script: "Comment" must be in Word language.
	# we suppose that NVDA and Word are in same language.
	def script_reportCurrentComment(self, gesture):
		caretInfo = self.makeTextInfo(textInfos.POSITION_CARET)
		caretInfo.expand(textInfos.UNIT_CHARACTER)
		val = caretInfo._rangeObj.getAttributeValue(UIAHandler.UIA_AnnotationObjectsAttributeId)
		if not val:
			return
		try:
			UIAElementArray = val.QueryInterface(UIAHandler.IUIAutomationElementArray)
		except COMError:
			return
		for index in range(UIAElementArray.length):
			UIAElement = UIAElementArray.getElement(index)
			UIAElement = UIAElement.buildUpdatedCache(UIAHandler.handler.baseCacheRequest)
			obj = UIA(UIAElement=UIAElement)
			if not obj.parent or obj.parent.name != NVDAString('Comment'):
				continue
			comment = obj.makeTextInfo(textInfos.POSITION_ALL).text
			dateObj = obj.previous
			date = dateObj.name
			authorObj = dateObj.previous
			author = authorObj.name
			# Translators: The message reported for a comment in Microsoft Word
			ui.message(_("{comment} by {author} on {date}").format(
				comment=comment, date=date, author=author))
			return
	# Translators: a description for a script
	script_reportCurrentComment.__doc__ = NVDAString("Reports the text of the comment where the System caret is located.")
