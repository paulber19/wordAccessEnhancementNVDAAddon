# appModules\winword\ww_headings.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import wx
import ui
from NVDAObjects.window.winword import wdGoToHeading
from .ww_collection import Collection, CollectionElement, ReportDialog

addonHandler.initTranslation()


def reportHeadingLevel(parent):
	app = parent.WinwordApplicationObject
	selection = app.selection
	doc = parent.doc = app.activeDocument
	theRange = doc.Range(selection.Start, selection.End)
	paragraphs = getattr(theRange, "paragraphs")
	if paragraphs.Count != 1:
		return
	paragraph = paragraphs[0]
	if paragraph.OutLineLevel >= 10:
		return

	heading = Heading(parent, paragraph)
	heading.sayLevel()


class Heading(CollectionElement):
	def __init__(self, parent, item):
		super(Heading, self).__init__(parent, item)
		self.text = ""
		if item.range.text:
			self.text = item.range.text
		self.start = item.range.start
		self.level = item.OutLineLevel
		self.style = item.style.nameLocal
		self.setLineAndPageNumber()

	def sayLevel(self):
		ui.message(_("heading level %s") % self.level)

	def formatInfos(self):

		sInfo = _("""	Page {page}, line {line}
	Text: {text}
	Level: {level}
	Style: {style}
	""")
		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(
			page=self.page, line=self.line,
			text=self.text, level=self.level,
			style=self.style)


class Headings(Collection):
	_propertyName = ((None, Heading),)
	_name = (_("Heading"), _("Headings"))
	_wdGoToItem = wdGoToHeading

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = HeadingsDialog
		self.noElement = _("No heading")
		super(Headings, self).__init__(parent, focus)

	def getCollectionInRange(self, theRange):
		collection = []
		i = 0
		rangeObj = theRange
		while True:
			i = i + 1
			wx.Yield()
			# stopped by user?
			if self.parent and self.parent.canceled:
				return []
			paragraph = rangeObj.Paragraphs[1]
			level = paragraph.OutLineLevel
			if level < 10:
				collection.append((paragraph, Heading))
			newRangeObj = rangeObj.gotoNext(wdGoToHeading)
			if i > 400\
				or not newRangeObj\
				or newRangeObj.start > theRange.End\
				or newRangeObj.Start == rangeObj.Start:
				break
			rangeObj = self.doc.Range(newRangeObj.Start, theRange.End)
		return collection


class HeadingsDialog(ReportDialog):
	def __init__(self, parent, obj):
		super(HeadingsDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		self.lcLabel = _("Headings")
		self.lcColumns = (
			(_("Heading"), 100),
			(_("Text"), 300),
			(_("Location"), 150),
			(_("Level"), 100),
			(_("Style"), 200)
		)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]
		self.lcSize = (lcWidth, self._defaultLCWidth)
		self.buttons = (
			(100, _("&Go to"), self.goTo),
		)
		self.tc1 = None
		self.tc2 = None

	def get_lcColumnsDatas(self, element):
		location = (_("Page {page}, line {line}")).format(
			page=element.page, line=element.line)
		index = self.collection.index(element) + 1

		datas = (index, element.text, location, element.level, element.style)

		return datas

	def onInputChar(self, evt):
		key = evt.GetKeyCode()
		if (key == wx.WXK_RETURN):
			self.goTo()
			return
		elif (key == wx.WXK_ESCAPE):
			self.Close()
			return

		evt.Skip()
