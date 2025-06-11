# appModules\winword\ww_spellingErrors.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import api
import ui
import time
from eventHandler import queueEvent
from .ww_wdConst import wdSentence
from .ww_collection import Collection, CollectionElement, ReportDialog
addonHandler.initTranslation()


class SpellingError(CollectionElement):
	def __init__(self, parent, item):
		super(SpellingError, self).__init__(parent, item)
		self.start = item.Start
		self.end = item.End
		self.text = ""
		if item.Text:
			self.text = item.Text
		r = self.doc.range(self.start, self.end)
		self.setLineAndPageNumber(r)
		r.expand(wdSentence)
		self.sentence = r.Text

	def modify(self, text):
		r = self.doc.range(self.start, self.end)
		self.obj.Text = text
		self.text = text
		r.Select()
		r.Collapse()

	def formatInfos(self):
		sInfo = _("""Page {page}, line {line}
Word: {word}
""")

		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(page=self.page, line=self.line, word=self.text)


class SpellingErrors(Collection):
	_propertyName = (("SpellingErrors", SpellingError),)
	_name = (_("Spelling error"), _("Spelling errors"))
	_wdGoToItem = None  # wdGoToSpellingError don't work

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = SpellingErrorsDialog
		self.noElement = _("No spelling error")
		wordApp = focus._WinwordWindowObject.Application
		if not wordApp.Options.CheckSpellingAsYouType:
			self.collection = None
			ui.message(_("Not available, Check spelling as you type is not activated"))
			time.sleep(3.0)
			api.processPendingEvents()
			time.sleep(0.2)
			queueEvent("gainFocus", api.getFocusObject())
			return
		super(SpellingErrors, self).__init__(parent, focus)


class SpellingErrorsDialog(ReportDialog):
	def __init__(self, parent, obj):
		super(SpellingErrorsDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		self.lcLabel = _("Errors:")
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
		)

		self.tc1 = {
			"label": _("Mispelled word"),
			"size": (800, 200)}

		self.tc2 = {
			"label": _("Involved sentence"),
			"size": (800, 200)}

	def get_lcColumnsDatas(self, element):
		location = _("Page {page}, line {line}") .format(
			page=element.page, line=element.line)
		index = self.collection.index(element) + 1
		datas = (index, location)
		return datas

	def get_tc1Datas(self, element):
		return element.text

	def get_tc2Datas(self, element):
		return element.sentence
