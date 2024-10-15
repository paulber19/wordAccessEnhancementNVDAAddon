# appModules\winword\ww_comments.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2022 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import config
import textInfos
import ui
import gui
import wx
from .ww_wdConst import wdGoToComment
from .ww_collection import Collection, CollectionElement, ReportDialog
from .ww_textUtils import askForText, askForAuthor
import sys
import os
_curAddon = addonHandler.getCodeAddon()
sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_informationDialog import InformationDialog
del sys.path[-1]
addonHandler.initTranslation()


def getReferenceAtFocus(focus):
	info = focus.makeTextInfo(textInfos.POSITION_CARET)
	info.expand(textInfos.UNIT_CHARACTER)
	formatConfig = config.conf["documentFormatting"].copy()
	for item in formatConfig:
		formatConfig[item] = False
	formatConfig["reportComments"] = True
	fields = info.getTextWithFields(formatConfig)
	for field in reversed(fields):
		if isinstance(field, textInfos.FieldCommand)\
			and isinstance(field.field, textInfos.FormatField):
			reference = field.field.get('comment')
			if reference:
				# if UIA is used, "comment" field is a boolean.
				# if "comment field is True, getTextWithFields method
				# has added a "extendedComment" field equal to the IAccessible "comment" field.
				if type(reference) is bool:
					if field.field.get('extendedComment'):
						return int(field.field.get('extendedComment'))
				else:
					return int(reference)
	return None


class Comment(CollectionElement):
	def __init__(self, parent, item):
		super(Comment, self).__init__(parent, item)
		self.ancestor = item.Ancestor
		self.reference = item.Reference
		self.done = item.Done
		self.replies = item.Replies
		self.text = ""
		self.range = item.range
		if item.Range.Text:
			self.text = item.range.text
		self.author = ""
		if item.author:
			self.author = item.Author
		self.date = ""
		if item.date:
			self.date = item.date
		self.scope = item.scope
		r = item.scope
		self.start = r.Start
		self.associatedText = ""
		if r.text:
			self.associatedText = r.Text
		self.setLineAndPageNumber()

	def modifyText(self, text):
		self.obj.Range.Text = text
		self.text = text

	def modifyAuthor(self):
		author = askForAuthor(self.author)
		if author and author != self.author:
			self.obj.Author = author
			self.author = author
			return True
		return False

	def getReplies(self):
		if self.replies.Count == 0:
			return []
		replies = []
		for index in range(1, self.replies.Count + 1):
			item = self.replies(index)
			if item.Ancestor:
				ancestorAuthor = item.Ancestor.author
				text = ""
				if item.Range.Text:
					text = item.range.text
				date = ""
				if item.date:
					date = item.date
				replies.append({"text": text, "author": item.Author, "date": date, "ancestorAuthor": ancestorAuthor})
		return replies

	def reportReplies(self, show=False):
		replies = self.getReplies()
		if not replies:
			# Translators: message to user to warn that there is no reply to this comment.
			wx.CallLater(40, ui.message, _("There is no reply"))
			return
		textList = []
		if len(self.associatedText) > 200:
			text = "{part1}\n...\n{part2}" .format(part1=self.associatedText[:100], part2=self.associatedText[100:])
		else:
			text = self.associatedText
		# Translators: text to indicate commented text
		textList.append(_("Commented text") + ":")
		textList.append("\n%s\n" % text + "\n")
		# Translators: text to indicate that the comment is solved.
		solvedText = _("solved")
		solvedText = " - %s" % solvedText if self.done else "" if self.done else ""
		# Translators: comment author.
		text = _("Comment by {author}") .format(solved=solvedText, author=self.author)
		if self.date == "":
			text = _("{text}{solved}:") .format(text=text, solved=solvedText)
		else:
			# Translators: comment author with date.
			text = _("{text} ({date}){solved}:") . format(text=text, date=self.date, solved=solvedText)
		textList.append(text + "\n")
		textList.append(self.text + "\n\n")
		for reply in replies:
			# Translators: text to indicate author reply.
			text = _("Reply from {author}") .format(author=reply["author"])
			if reply["date"] == "":
				# Translators: text to indicate only author reply.
				text = _("{text}:") .format(text=text)
			else:
				# Translators: text to indicate author reply with date.
				text = _("{text} ({date}):") .format(text=text, date=reply["date"])
			textList.append(text + "\n")
			textList.append(reply["text"] + "\n\n")
		text = "".join(textList)
		text = text[:-1]
		if not show:
			wx.CallLater(40, ui.message, text)
		else:
			# Translators: this is the title of informationdialog box.
			# to show comment replies
			dialogTitle = _("Comment replies")
			wx.CallLater(
				40,
				InformationDialog.run,
				None, dialogTitle, "", text)

	def addReply(self):
		text = askForText(
			CommentReplies.entryDialogStrings["insertDialogTitle"], CommentReplies.entryDialogStrings["entryBoxLabel"])
		if text is None:
			return
		newReply = self.replies.Add(self.scope, text)
		if not newReply or newReply.Range.Text != text:
			# Translators: message to user to warn an error on reply insert.
			wx.CallLater(40, ui.message, _("Error, the reply could not be correctly added."))
			return
		if gui.messageBox(
			# Translators: text to indicate that reply has been inserted and ask for author name change.
			_("""The reply was inserted with "{name}" as author. Do you want to change this name?""") .format(
				name=newReply.author),
			# Translators: title of insert reply dialog.
			_("Reply comment"),
			wx.YES | wx.NO | wx.ICON_INFORMATION
		) != wx.YES:
			return
		author = askForAuthor(newReply.Author)
		if author:
			newReply.Author = author

	def formatInfos(self):
		# Translators: document information text.
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


class CommentReplies(object):
	askForAuthor = True
	entryDialogStrings = {
		# Translators: title of comment insert dialog.
		"insertDialogTitle": _("Reply comment's insert"),
		# Translators: label of dialog entry text box.
		"entryBoxLabel": _("Enter the reply's text"),
		# Translators: title of modification dialog.
		"modifyDialogTitle": _("Reply's Modification"),
	}

	@classmethod
	def canInsert(cls, focus):
		comments = Comments(None, focus, "focus")
		if not comments.collection:
			# Translator: Not available, there  is no comment at cursor position
			msg = _("Not available, there is no comment at cursor position")
			return (False, msg)
		return (True, None)

	@classmethod
	def insert(cls, wordApp, text, focus):
		comments = Comments(None, focus, "focus")
		if len(comments.collection) != 1:
			return
		comment = comments.collection[0]
		newReply = comment.replies.Add(comment.scope, text)
		if not newReply or newReply.Range.Text != text:
			# Translators: message to user to warn an error on reply insert.
			wx.CallLater(40, ui.message, _("Error, the reply could not be correctly added."))
			return
		if gui.messageBox(
			# Translators: text to indicate that reply has been inserted and ask for author name change.
			_("""The reply was inserted with "{name}" as author name. Do you want to change this name?""") .format(
				name=newReply.author),
			# Translators: title of insert reply dialog.
			_("Reply comment"),
			wx.YES | wx.NO | wx.ICON_INFORMATION
		) != wx.YES:
			return
		author = askForAuthor(newReply.Author)
		if author:
			newReply.Author = author


class Comments(Collection):
	askForAuthor = True
	_propertyName = (("Comments", Comment),)
	# Translators: singular and plural comment element name.
	_name = (_("Comment"), _("Comments"))
	_wdGoToItem = wdGoToComment
	entryDialogStrings = {
		# Translators: title of comment insert dialog.
		"insertDialogTitle": _("Comment's insert"),
		# Translators: label of dialog entry text box.
		"entryBoxLabel": _("Enter the comment's text"),
		# Translators: title of comment modification dialog.
		"modifyDialogTitle": _("Comment's Modification"),
	}

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = CommentsDialog
		# Translators: to indicate that there is no comment.
		self.noElement = _("No comment")
		reference = getReferenceAtFocus(focus)
		if reference:
			self.reference = reference
		super(Comments, self).__init__(parent, focus)

	@classmethod
	def canInsert(cls, focus):
		comments = Comments(None, focus, "focus")
		if len(comments.collection):
			# Translator: Not available, there already a commentat cursor position
			msg = _("Not available, there is already comment at cursor position")
			return (False, msg)
		return (True, None)

	@classmethod
	def insert(cls, wordApp, text, focus):
		r = wordApp.ActiveDocument.Range(wordApp.Selection.Start, wordApp.Selection.End)
		comment = wordApp.ActiveDocument.Comments.Add(r, text)
		if not comment or comment.Range.Text != text:
			# Translators: message to user to warn an error on comment insert.
			wx.CallLater(40, ui.message, _("Error, the comment could not be correctly added."))
			return
		if gui.messageBox(
			# Translators: text to indicate that comment has been inserted and ask for author name change.
			_("""The comment was inserted with "{name}" as author name. Do you want to change this name?""") .format(
				name=comment.author),
			# Translators: title of insert comment dialog.
			_("Insert comment"),
			wx.YES | wx.NO | wx.ICON_INFORMATION
		) != wx.YES:
			return
		author = askForAuthor(comment.author)
		if author:
			comment.Author = author

	def reportRepliesToFocusedComment(self, show=False):
		if not hasattr(self, "reference"):
			# Translators: to indicate that there is no comment to cursor position.
			ui.message(_("There is no comment at cursor position"))
			return
		if len(self.collection) != 1:
			return
		self.collection[0].reportReplies(show)

	def replyToFocusedComment(self):
		if not hasattr(self, "reference"):
			# Translators: to indicate that there is no comment to cursor position.
			ui.message(_("There is no comment to cursor position"))
			return
		if len(self.collection) != 1:
			return
		self.collection[0].addReply()


class CommentsDialog(ReportDialog):
	collectionClass = Comments

	def __init__(self, parent, obj):
		super(CommentsDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		# Translators: label of comment column in comments result dialog.
		self.lcLabel = _("Comments:")
		self.lcColumns = (
			# Translators: label of comment number column in comments result dialog.
			(_("Number"), 100),
			# Translators: label of comment location column in comments result dialog.
			(_("Location"), 150),
			# Translators: label of comment author column in comments result dialog.
			(_("Author"), 300),
			# Translators: label of comment date column in comments result dialog.
			(_("Date"), 150)
		)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]
		self.lcSize = (lcWidth, self._defaultLCWidth)
		self.buttons = (
			# Translators: go to button label.
			(100, _("&Go to"), self.goTo),
			# Translators: reply button label.
			(101, _("Repl&y"), self.reply),
			# Translators: modify comment text button label.
			(102, _("&Modify comment text"), self.modifyTC1Text),
			# Translators: modify author button label.
			(103, _("Modify author &name"), self.modifyAuthor),
			# Translators: delete button label.
			(104, _("&Delete"), self.delete),
			# Translators: report replies button label.
			(105, _("Report &replies"), self.reportReplies),
			# Translators: delete all button label.
			(106, _("Delete &all"), self.deleteAll),
		)

		self.tc1 = {
			# Translators: label of comment text column in comment result dialog.
			"label": _("Comment's text"),
			"size": (800, 200)
		}
		self.tc2 = {
			# Translators: label of commented text column in comment result dialog.
			"label": _("Commented text"),
			"size": (800, 200)
		}

	def get_lcColumnsDatas(self, element):
		# Translators: text to indicate comment location.
		location = (_("Page {page}, line {line}")).format(
			page=element.page, line=element.line)
		index = self.collection.index(element) + 1
		datas = (index, location, element.author, element.date)
		return datas

	def get_tc1Datas(self, element):
		return element.text

	def get_tc2Datas(self, element):
		return element.associatedText

	def modifyAuthor(self):
		index = self.entriesList.GetFocusedItem()
		comment = self.collection[index]
		if not comment.modifyAuthor():
			return
		self.refreshList(index)
		self.entriesList.SetFocus()

	def reply(self):
		index = self.entriesList.GetFocusedItem()
		comment = self.collection[index]
		comment.addReply()
		self.onClose(None)

	def reportReplies(self):
		index = self.entriesList.GetFocusedItem()
		comment = self.collection[index]
		comment.reportReplies()
