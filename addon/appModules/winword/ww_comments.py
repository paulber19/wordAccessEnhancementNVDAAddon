# appModules\winword\ww_comments.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2022 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import config
import textInfos
from .ww_wdConst import wdGoToComment
from .ww_collection import Collection, CollectionElement, ReportDialog

addonHandler.initTranslation()


def getReferenceAtFocus(focus):
	info = focus.makeTextInfo(textInfos.POSITION_CARET)
	info.expand(textInfos.UNIT_CHARACTER)
	formatConfig = config.conf["documentFormatting"]
	for item in formatConfig:
		formatConfig[item] = False
	formatConfig["reportComments"] = True
	fields = info.getTextWithFields(formatConfig)
	for field in reversed(fields):
		if isinstance(field, textInfos.FieldCommand)\
			and isinstance(field.field, textInfos.FormatField):
			reference = field.field.get('comment')
			if reference:
				return int(reference)
	return None


class Comment(CollectionElement):
	def __init__(self, parent, item):
		super(Comment, self).__init__(parent, item)
		self.reference = item.Reference
		self.text = ""
		if item.Range.Text:
			self.text = item.range.text
		self.author = ""
		if item.author:
			self.author = item.Author
		self.date = ""
		if item.date:
			self.date = item.date
		r = self.associatedTextRange = item.scope
		self.start = r.Start
		self.associatedText = ""
		if r.text:
			self.associatedText = r.text
		self.setLineAndPageNumber()

	def modifyText(self, text):
		self.obj.Range.Text = text
		self.text = text

	def formatInfos(self):
		sInfo = _("""Page {page}, line {line}
Author: {author}
Date: {date}
Comment's text:
{text}
Commented text:
{commentedText}
""")

		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(
			page=self.page, line=self.line,
			text=self.text, author=self.author,
			date=self.date, commentedText=self.associatedText)


class Comments(Collection):
	_propertyName = (("Comments", Comment),)
	_name = (_("Comment"), _("Comments"))
	_wdGoToItem = wdGoToComment
	entryDialogStrings = {
		# Translators: text of entry text box:
		"entryBoxLabel": _("Enter the comment's text"),
		# Translators: title of insert dialog.
		"insertDialogTitle": _("Comment's insert"),
		# Translators: title of modification dialog.
		"modifyDialogTitle": _("Comment's Modification"),
	}

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = CommentsDialog
		self.noElement = _("No comment")
		reference = getReferenceAtFocus(focus)
		if reference:
			self.reference = reference
		super(Comments, self).__init__(parent, focus)

	@classmethod
	def insert(cls, wordApp, text):
		doc = wordApp.ActiveDocument
		selection = wordApp.Selection
		comments = doc.Comments
		r = doc.Range(selection.Start, selection.End)
		comments.Add(r, text)


class CommentsDialog(ReportDialog):
	collectionClass = Comments

	def __init__(self, parent, obj):
		super(CommentsDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		self.lcLabel = _("Comments:")
		self.lcColumns = (
			(_("Number"), 100),
			(_("Location"), 150),
			(_("Author"), 300),
			(_("Date"), 150)
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
			"label": _("Comment's text"),
			"size": (800, 200)
		}
		self.tc2 = {
			"label": _("Commented text"),
			"size": (800, 200)
		}

	def get_lcColumnsDatas(self, element):

		location = (_("Page {page}, line {line}")).format(
			page=element.page, line=element.line)
		index = self.collection.index(element) + 1
		datas = (index, location, element.author, element.date)
		return datas

	def get_tc1Datas(self, element):
		return element.text

	def get_tc2Datas(self, element):
		return element.associatedText
