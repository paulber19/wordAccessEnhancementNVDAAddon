# appModules\winword\ww_footnotes.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2022 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import textInfos
from .ww_wdConst import wdGoToFootnote, wdCollapseEnd
from .ww_collection import Collection, CollectionElement, ReportDialog

addonHandler.initTranslation()


class Footnote(CollectionElement):
	def __init__(self, parent, item):
		super(Footnote, self).__init__(parent, item)
		self.start = item.reference.Start
		self.end = item.reference.End
		self.text = " "
		if item.Range.Text:
			self.text = item.Range.Text
		self.setLineAndPageNumber()

	def modifyText(self, text):
		self.obj.Range.Text = text
		self.text = text

	def formatInfos(self):
		sInfo = _("""Page {page}, line {line}
Note's Ttext:
{text}
""")

		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(
			page=self.page, line=self.line, text=self.text)


class Footnotes(Collection):
	_propertyName = (("Footnotes", Footnote),)
	_name = (_("Footnote"), _("Footnotes"))
	_wdGoToItem = wdGoToFootnote
	entryDialogStrings = {
		# Translators: text of entry text box:
		"entryBoxLabel": _("Enter the footnote's text"),
		# Translators: title of insert dialog.
		"insertDialogTitle": _("Footnote's insert"),
		# Translators: title of modify dialog.
		"modifyDialogTitle": _("Footnote's modification"),
	}

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = FootnotesDialog
		self.noElement = _("No footnote")
		super(Footnotes, self).__init__(parent, focus)
		self.__class__._elementUnit = textInfos.UNIT_CHARACTER

	@classmethod
	def insert(cls, wordApp, text):
		doc = wordApp.ActiveDocument
		selection = wordApp.Selection
		selection.Collapse(wdCollapseEnd)
		r = selection.Range
		footnoteObj = doc.Footnotes.Add(r)
		footnoteObj.range.text = text
		r.Select()


class FootnotesDialog(ReportDialog):
	collectionClass = Footnotes

	def __init__(self, parent, obj):
		super(FootnotesDialog, self).__init__(parent, obj)

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
			(101, _("&Modify"), self.modifyTC1Text),
			(102, _("&Delete"), self.delete),
			(103, _("Delete &all"), self.deleteAll),
		)
		self.tc1 = {
			"label": _("Note's text"),
			"size": (800, 200)
		}

		self.tc2 = None

	def get_lcColumnsDatas(self, element):
		location = _("Page {page}, line {line}") .format(
			page=element.page, line=element.line)
		index = self.collection.index(element) + 1
		datas = (index, location)
		return datas

	def get_tc1Datas(self, element):
		return element.text
