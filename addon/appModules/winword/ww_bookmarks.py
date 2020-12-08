# appModules\winword\ww_bookmarks.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import ui

from .ww_wdConst import wdStoryValueToText
from .ww_collection import Collection, CollectionElement, ReportDialog

addonHandler.initTranslation()


class Bookmark(CollectionElement):
	def __init__(self, parent, item):
		super(Bookmark, self).__init__(parent, item)
		self.name = ""
		if item.Name:
			self.name = item.Name
		# a value that indicates if the specified bookmark is a table column.
		self.column = item.column
		# a value that indicates if the specified bookmark is empty.
		self.empty = item.empty
		self.start = item.Start
		self.end = item.End
		self.storyType = item.StoryType
		self.bookmarkedText = ""
		if not self.empty:
			self.bookmarkedText = item.range.text

		self.rangeStart = item.range.Start
		self.rangeEnd = item.range.End
		r = self.doc.range(self.rangeStart, self.rangeStart)
		self.setLineAndPageNumber(r)

	def get_storyTypeText(self):
		return wdStoryValueToText[self.storyType]

	def formatInfos(self):
		storyTypeText = self.get_storyTypeText()
		# Translators: information of bookmark element.
		sInfo = _("""Page {page}, line {line}
Story type text: {storyTypeText}
Associated Ttext:
{text}
""")
		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(
			page=self.page, line=self.line,
			text=self.bookmarkedText, storyTypeText=storyTypeText)

	def speakInfo(self):
		ui.message(self.name)


class Bookmarks(Collection):
	_propertyName = (("Bookmarks", Bookmark),)
	# Translators: name of collection element type
	_name = (_("Bookmark"), _("Bookmarks"))
	_wdGoToItem = None  # wdGoToBookmark don't work

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = BookmarksDialog
		# Translators: text to report no bookmark.
		self.noElement = _("No bookmark")
		super(Bookmarks, self).__init__(parent, focus)


class BookmarksDialog(ReportDialog):
	def __init__(self, parent, obj):
		super(BookmarksDialog, self).__init__(parent, obj)

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
		index = self.collection.index(element)+1
		datas = (index, location, element.name, element.get_storyTypeText())
		return datas

	def get_tc1Datas(self, element):
		return element.bookmarkedText
