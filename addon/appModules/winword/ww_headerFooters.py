# appModules\winword\ww_headerFooters.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2021 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import ui

from .ww_wdConst import wdStoryValueToText
from .ww_collection import Collection, CollectionElement, ReportDialog

addonHandler.initTranslation()


class HeaderFooter(CollectionElement):
	def __init__(self, parent, item):
		super(HeaderFooter, self).__init__(parent, item)
		self.exists = item.Exists
		self.index = item.Index
		self.isHeader = item.isHeader
		self.linkToPrevious = item.LinkToPrevious
		self.pageNumbers = item.PageNumbers
		self.range = item.Range
		self.text = self.range.text
		self.shape = item.Shape

	def get_storyTypeText(self):
		return wdStoryValueToText[self.storyType]

	def formatInfos(self):
		storyTypeText = self.get_storyTypeText()
		# Translators: information of bookmark element.
		sInfo = _("""Page {page}, line {line}
Story type text: {storyTypeText}
Associated text:
{text}
""")
		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(
			page=self.page, line=self.line,
			text=self.bookmarkedText, storyTypeText=storyTypeText)

	def speakInfo(self):
		ui.message(self.name)


class HeaderFooters(Collection):
	_propertyName = (("HeaderFooters", HeaderFooter),)
	# Translators: name of collection element type.
	_name = (_("Header/Footer"), _("Headers/Footers"))
	_wdGoToItem = None

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = HeaderFootersDialog
		# Translators: text to report no header or Footer.
		self.noElement = _("No header or footer")
		super(HeaderFooters, self).__init__(parent, focus)


class HeaderFootersDialog(ReportDialog):
	def __init__(self, parent, obj):
		super(HeaderFootersDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		self.lcLabel = _("Bookmarks")
		self.lcColumns = (
			(_("Number"), 100),
			(_("Location"), 150),
			(_("Name"), 300),
			(_("Story type"), 150)
		)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]

		self.lcSize = (lcWidth, self._defaultLCWidth)

		self.buttons = (
			(100, _("&Go to"), self.goTo),
			(101, _("&Delete"), self.delete),
			(102, _("Delete &all"), self.deleteAll)
		)
		self.tc1 = {
			"label": _("Bookmark's text"),
			"size": (800, 200)
		}
		self.tc2 = None

	def get_lcColumnsDatas(self, element):
		location = (_("Page {page}, line {line}")).format(
			page=element.page, line=element.line)
		index = self.collection.index(element) + 1
		datas = (index, location, element.name, element.get_storyTypeText())
		return datas

	def get_tc1Datas(self, element):
		return element.bookmarkedText
