# appModules\winword\ww_revisions.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.

import addonHandler
import api
import wx
import os
import sys
import ui
import time
from eventHandler import queueEvent
from .ww_collection import Collection, CollectionElement, ReportDialog
_curAddon = addonHandler.getCodeAddon()
sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_utils import myMessageBox  # noqa:E402
del sys.path[-1]

addonHandler.initTranslation()
revisionTypeText = {
	# wdNoRevision
	0: _("No revision"),
	# wdRevisionInsert
	1: _("Insertion"),
	# wdRevisionDelete
	2: _("Deletion"),
	# wdRevisionProperty
	3: _("Property changed"),
	# wdRevisionParagraphNumber
	4: _("Paragraph number changed"),
	# wdRevisionDisplayField
	5: _("Field display changed"),
	# wdRevisionReconcile
	6: _("Revision marked as reconciled conflict"),
	# wdRevisionConflict
	7: _("Revision marked as a conflict"),
	# wdRevisionStyle
	8: _("Style changed"),
	# wdRevisionReplace
	9: _("Replaced"),
	# wdRevisionParagraphProperty
	10: _("Paragraph property changed"),
	# wdRevisionTableProperty
	11: _("Table property changed"),
	# wdRevisionSectionProperty
	12: _("Section property changed"),
	# wdRevisionStyleDefinition
	13: _("Style definition changed"),
	# wdRevisionMovedFrom
	14: _("Content moved from"),
	# wdRevisionMovedTo
	15: _("Content moved to"),
	# wdRevisionCellInsertion
	16: _("Table cell inserted"),
	# wdRevisionCellDeletion
	17: _("Table cell deleted"),
	# wdRevisionCellMerge
	18: _("Table cells merged")
	}

class Revision(CollectionElement):
	def __init__(self, parent, item):
		super(Revision, self).__init__(parent, item)

		self.start = item.range.Start
		self.end = item.range.End
		self.text = ""
		if item.Range.Text:
			self.text = item.range.text
		self.author = item.Author
		self.date = item.date
		self.type = item.type
		self.setLineAndPageNumber()

	def get_typeText(self):
		return revisionTypeText[self.type]
	def FormatRevisionTypeAndAuthorText(self):
		from ww_addonConfigManager import _addonConfigManager
		if not _addonConfigManager.toggleReportRevisionTypeWithAuthorOption(False):
			author = ""
		elif self.author != "":
			# Translators: part of message to report author of revision
			author= _("by %s") % self.author
		revisionTypeText = self.get_typeText()
		text = "%s %s" % (revisionTypeText, author)
		return text

	def accept(self):
		self.obj.accept()

	def reject(self):
		self.obj.Reject()

	def formatInfos(self):

		sInfo = _("""Page {page}, line {line}
Author: {author}
Date: {date}
Type: {type}
Revision's ttext:
{text}
""")

		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(
			page=self.page, line=self.line,
			text=self.text, author=self.author,
			date=self.date, type=self.get_typeText())

	def reportText(self):
		ui.message(_("{type} by {author} {text}") .format(
			type=self.get_typeText(), author=self.author, text=self.text)		)


class Revisions(Collection):
	_propertyName = (("Revisions", Revision),)
	_name = (_("Revision"), _("Revisions"))
	_wdGoToItem = None

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = RevisionsDialog
		self.noElement = _("No revision")
		super(Revisions, self).__init__(parent, focus)

	def accept(self, element):
		element.accept()
		# self.objectsInCollection.remove(element.obj)
		self.collection.remove(element)

	def reject(self, element):
		element.reject()
		# self.objectsInCollection.remove(element.obj)
		self.collection.remove(element)

	def reportCollection(self, copyToClip, makeChoiceDialog):
		if not self.doc.TrackRevisions:
			ui.message(_("revision track is not activated"))
			time.sleep(2.0)
			return

		super(Revisions, self).reportCollection(copyToClip, makeChoiceDialog)


class RevisionsDialog(ReportDialog):
	def __init__(self, parent, obj):
		super(RevisionsDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		self.lcLabel = _("Revisions:")
		self.lcColumns = (
			(_("Number"), 100),
			(_("Location"), 150),
			(_("Type"), 150),
			(_("Author"), 300),
			(_("Date"), 150)
			)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]

		self.lcSize = (lcWidth, self._defaultLCWidth)

		self.buttons = (
			(100, _("&Go to"), self.goTo),
			(101, _("&Accept"), self.accept),
			(102, _("&Reject"), self.reject)
			)

		self.tc1 = {
			"label": _("Revision's text"),
			"size": (800, 200)
		}
		self.tc2 = None

	def onInputChar(self, evt):
		key = evt.GetKeyCode()
		if (key == wx.WXK_RETURN):
			self.goTo()
			return

		elif (key == wx.WXK_DELETE) or (key == wx.WXK_NUMPAD_DELETE):
			if myMessageBox(
				# Translators: text of message
				_("Are you sure you wish to reject this items?"),
				# Translators: title of message box dialog.
				_("Reject of items"),
				wx.YES_NO | wx.ICON_WARNING) != wx.YES:
				return
			self.reject()
			return

		elif (key == wx.WXK_ESCAPE):
			self.Close()
			return

		evt.Skip()

	def accept(self):
		index = self.entriesList.GetFocusedItem()
		item = self.collection[index]
		self.collectionObject.accept(item)
		self.refreshList(index)
		if len(self.collection) == 0:
			return
		self.entriesList.SetFocus()
		time.sleep(0.1)
		api.processPendingEvents()
		obj = api.getFocusObject()
		queueEvent("gainFocus", obj)

	def reject(self):
		index = self.entriesList.GetFocusedItem()
		item = self.collection[index]
		self.collectionObject.reject(item)
		self.refreshList(index)
		if len(self.collection) == 0:
			return
		self.entriesList.SetFocus()
		time.sleep(0.1)
		api.processPendingEvents()

		obj = api.getFocusObject()
		queueEvent("gainFocus", obj)

	def get_lcColumnsDatas(self, element):
		location = (_("Page {page}, line {line}")).format(
			page=element.page, line=element.line)
		typeText = element.get_typeText()
		index = self.collection.index(element)+1
		datas = (index, location, typeText, element.author, element.date)
		return datas

	def get_tc1Datas(self, element):
		return element.text
