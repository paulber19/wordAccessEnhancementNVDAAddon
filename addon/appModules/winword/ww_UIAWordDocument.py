# appModules\winword\ww_UIAWordDocument.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2020-2024 paulber19, Abdel
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

import addonHandler
import os
from . import ww_wordDocumentBase
from logHandler import log
import oleacc
import NVDAObjects
import comtypes
import winUser
import config
import controlTypes
from typing import Optional, Dict
from NVDAObjects.IAccessible import getNVDAObjectFromEvent
import UIAHandler
import textInfos
import NVDAObjects.UIA.wordDocument
from .automaticReading import AutomaticReadingWordTextInfo
import uuid
from comtypes import COMError
from tableUtils import HeaderCellTracker
from .ww_UIABrowseMode import UIAWordBrowseModeDocument
import sys
_curAddon = addonHandler.getCodeAddon()
debugToolsPath = os.path.join(_curAddon.path, "debugTools")
sys.path.append(debugToolsPath)
try:
	from appModuleDebug import printDebug, toggleDebugFlag
except ImportError:

	def printDebug(msg):
		return

	def toggleDebugFlag():
		return
del sys.path[-1]
sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_addonConfigManager import _addonConfigManager
del sys.path[-1]

addonHandler.initTranslation()


def getIAccessibleTextWithFields(textInfo, unit, formatConfig: Optional[Dict] = None):
	"""
	Function added by Abdel to retrieve the fields returned by the getTextWithFields method,
	in IAccessible mode, in order to be able to use them in UIA mode.
	@param textInfo: textInfos.TextInfo object of 'item.
	@textInfo type: textInfos.TextInfo.
	@param unit: Text range scope.
	@unit type: str.
	@param formatConfig: An optional dictionary,
		indicating the configuration of the presentation format of the fields to check.
	@type formatConfig: dict.
	@returns: The items returned by the getTextWithFields method of IAccessible mode,
		respecting the scope of the text range.
	@rtype: textInfos.TextInfo.TextWithFieldsT.
	"""
	formatConfig = formatConfig if formatConfig else config.conf["documentFormatting"]
	startObj = textInfo.NVDAObjectAtStart
	obj = getNVDAObjectFromEvent(startObj.windowHandle, winUser.OBJID_CLIENT, winUser.CHILDID_SELF)
	info = obj.makeTextInfo(textInfos.POSITION_CARET)
	info.expand(unit)
	return info.getTextWithFields(formatConfig=formatConfig)

_UIAUnits = [x for x in UIAHandler.NVDAUnitsToUIAUnits]


class UIAWordDocumentTextInfo (
	AutomaticReadingWordTextInfo,
	NVDAObjects.UIA.wordDocument.WordDocumentTextInfo
):

	def _get_currentUnit(self):
		info = self.copy()
		info.collapse()
		units = _UIAUnits
		for unit in units:
			info.expand(unit)
			if (
				info.text == self.text
				and info.compareEndPoints(self, "startToStart") == 0
				and info.compareEndPoints(self, "endToEnd") == 0
			):
				return unit
		return textInfos.UNIT_CHARACTER

	def getTextWithFields(self, formatConfig: Optional[Dict] = None) -> textInfos.TextInfo.TextWithFieldsT:
		items = super(UIAWordDocumentTextInfo, self).getTextWithFields(formatConfig=formatConfig)
		if not _addonConfigManager.toggleAutomaticReadingOption(False):
			return items

		IAccessibleFields = getIAccessibleTextWithFields(
			textInfo=self,
			unit=self.currentUnit,
			formatConfig=formatConfig
		)
		IAccessibleComment  = None
		for field in IAccessibleFields:
			if isinstance(field, textInfos.FieldCommand)\
				and isinstance(field.field, textInfos.FormatField):
				IAccessibleComment = field.field.get('comment')
				break

		IAccessibleFootnotes = [
			IAccessibleFields.index(x)
			for x in IAccessibleFields if hasattr(
				x, "field") and x.field and x.field.get("role") == controlTypes.Role.FOOTNOTE
		]
		IAccessibleEndnotes = [
			IAccessibleFields.index(x)
			for x in IAccessibleFields if hasattr(
				x, "field") and x.field and x.field.get("role") == controlTypes.Role.ENDNOTE
		]
		for index in range(len(items)):
			if isinstance(items[index], textInfos.FieldCommand)\
				and isinstance(items[index].field, textInfos.FormatField):
				UIAComment = items[index].field.get('comment', None)
				if UIAComment and IAccessibleComment:
					items[index].field["extendedComment"] = IAccessibleComment

			if hasattr(items[index], "field") and items[index].field.get("extendedRole") == controlTypes.Role.FOOTNOTE:
				if len(IAccessibleFootnotes) > 0:
					footnote = IAccessibleFootnotes.pop(0)
					items[index].field["value"] = IAccessibleFields[footnote].field.get("value")
			elif (
				hasattr(items[index], "field")
				and items[index].field.get("extendedRole") == controlTypes.Role.ENDNOTE
			):
				if len(IAccessibleEndnotes) > 0:
					endnote = IAccessibleEndnotes.pop(0)
					items[index].field["value"] = IAccessibleFields[endnote].field.get("value")
		return items

	def _getControlFieldForUIAObject(self, obj, isEmbedded=False, startOfNode=False, endOfNode=False):
		UIAFields = super(UIAWordDocumentTextInfo, self)._getControlFieldForUIAObject(
			obj,
			isEmbedded=isEmbedded,
			startOfNode=startOfNode,
			endOfNode=endOfNode
		)
		unit = self.currentUnit
		info = self.copy()
		if (
			UIAFields.get("role") == controlTypes.Role.LINK
			and UIAFields.get("name") == UIAFields.get("content")
		):
			if unit in ("character", "word"):
				info.expand(unit)
			if UIAHandler.AnnotationType_Footnote in info._rangeObj.getAttributeValue(
				UIAHandler.UIA_AnnotationTypesAttributeId
			):
				UIAFields["extendedRole"] = controlTypes.Role.FOOTNOTE
			elif UIAHandler.AnnotationType_Endnote in info._rangeObj.getAttributeValue(
				UIAHandler.UIA_AnnotationTypesAttributeId
			):
				UIAFields["extendedRole"] = controlTypes.Role.ENDNOTE
		return UIAFields

	def _get_controlFieldNVDAObjectClass(self):
		return WordDocumentNode


class WordDocumentNode(NVDAObjects.UIA.wordDocument.WordDocumentNode):
	TextInfo = UIAWordDocumentTextInfo


class UIAWordDocument(
	WordDocumentNode,
	ww_wordDocumentBase.WordDocument,
	NVDAObjects.UIA.wordDocument.WordDocument
):
	treeInterceptorClass = UIAWordBrowseModeDocument
	shouldCreateTreeInterceptor = True
	disableAutoPassThrough = True
	announceEntireNewLine = True

	# Microsoft Word duplicates the full title of the document on this control,
	# which is redundant as it appears in the title of the app itself.
	name = u""

	def initOverlayClass(self):
		printDebug("UIAWordDocument InitOverlayClass")

	def _get_WinwordWindowObject(self):
		if not getattr(
			self,
			'_WinwordWindowObject',
			None
		):
			try:
				pDispatch = oleacc.AccessibleObjectFromWindow(
					self.documentWindowHandle,
					winUser.OBJID_NATIVEOM,
					interface=comtypes.automation.IDispatch
				)
			except (COMError, WindowsError):
				log.debugWarning(
					"Could not get MS Word object model from window %s with class %s" % (
						self.documentWindowHandle,
						winUser.getClassName(
							self.documentWindowHandle
						)
					), exc_info=True
				)
				return None
			self._WinwordWindowObject = comtypes.client.dynamic.Dispatch(pDispatch)
		return self._WinwordWindowObject

	def _get_WinwordDocumentObject(self):
		if not getattr(
			self,
			'_WinwordDocumentObject',
			None
		):
			windowObject = self.WinwordWindowObject
			if not windowObject:
				return None
			self._WinwordDocumentObject = windowObject.document
		return self._WinwordDocumentObject

	def _get_WinwordApplicationObject(self):
		if not getattr(
			self,
			'_WinwordApplicationObject',
			None
		):
			self._WinwordApplicationObject = self.WinwordWindowObject.application
		return self._WinwordApplicationObject

	def _get_WinwordSelectionObject(self):
		if not getattr(
			self,
			'_WinwordSelectionObject',
			None
		):
			windowObject = self.WinwordWindowObject
			if not windowObject:
				return None
			self._WinwordSelectionObject = windowObject.selection
		return self._WinwordSelectionObject
	"""
	def event_caret(self, ):
		from controlTypes.role import _roleLabels as roleLabels
		printDebug("UIAWordDocument: event_caret: role = %s, name = %s" % (roleLabels.get(self.role), self.name))
		#  in a table and UIA used,
		#  selected and unselected state
		#  are announced after when moving to another cell and using add-on table scripts
		#  so disable it
		#  super().event_caret()
	"""

	# part of code from WordDocument class of NVDAObjects.IAccessible.winword file
	def populateHeaderCellTrackerFromHeaderRows(self, headerCellTracker, table):
		rows = table.rows
		numHeaderRows = 0
		for rowIndex in range(rows.count):
			try:
				row = rows.item(rowIndex + 1)
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
			headerCellTracker.addHeaderCellInfo(
				rowNumber=1,
				columnNumber=1,
				rowSpan=numHeaderRows,
				isColumnHeader=True,
				isRowHeader=False)

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
			headerCellTracker.addHeaderCellInfo(
				rowNumber=headerCell.rowIndex,
				columnNumber=headerCell.columnIndex,
				name=name,
				isColumnHeader=isColumnHeader,
				isRowHeader=isRowHeader)

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
			headerCellTracker.addHeaderCellInfo(
				rowNumber=rowNumber,
				columnNumber=columnNumber,
				name=name,
				isColumnHeader=isColumnHeader,
				isRowHeader=isRowHeader)
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
		for info in headerCellTracker.iterPossibleHeaderCellInfosFor(
			rowNumber, columnNumber, columnHeader=columnHeader):
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
