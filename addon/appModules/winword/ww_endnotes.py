# appModules\winword\ww_endnotes.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.

import addonHandler
import textInfos
from .ww_wdConst import wdGoToEndnote
from .ww_collection import Collection, CollectionElement, ReportDialog

addonHandler.initTranslation()


class Endnote(CollectionElement):
	def __init__(self, parent, item):
		super(Endnote, self).__init__(parent, item)
		self.start = item.reference.Start
		self.end = item.reference.End
		self.reference = item.Reference
		self.text = ""
		if item.range.text:
			self.text = item.range.text
		self.setLineAndPageNumber(r)

	def formatInfos(self):
		sInfo = _("""Page {page}, line {line}
Note's ttext:
{text}
""")

		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(
			page=self.page, line=self.line, text=self.text)

	def reportText(self):
		pass


class Endnotes(Collection):
	_propertyName = (("Endnotes", Endnote),)
	_name = (_("Endnote"), _("Endnotes"))
	_wdGoToItem = wdGoToEndnote

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = EndnotesDialog
		self.noElement = _("No endnote")
		super(Endnotes, self).__init__(parent, focus)
		self.__class__._elementUnit = textInfos.UNIT_CHARACTER


class EndnotesDialog(ReportDialog):

	def __init__(self, parent, obj):
		super(EndnotesDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		self.lcLabel = _("Notes:")
		self.lcColumns = (
			(_("Number"), 100),
			(_("Location"), 150),
			)

		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]
		self.lcSize = (lcWidth, self._defaultLCWidth)
		self.buttons = (
			(100, _("&Go to"), self.goTo),
			(101, _("&Delete"), self.delete)
			)
		self.tc1 = {
			"label": _("Note's text"),
			"size": (800, 200)
		}
		self.tc2 = None

	def get_lcColumnsDatas(self, element):
		location = _("Page {page}, line {line}") .format(
			page=element.page, line=element.line)
		index = self.collection.index(element)+1
		datas = (index, location)
		return datas

	def get_tc1Datas(self, element):
		return element.text
