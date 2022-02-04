# appModules\winword\ww_browsemode.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2022 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import ui
import browseMode
import textInfos
import config
try:
	# for nvda version < 2021.1
	from sayAllHandler import CURSOR_CARET
except (AttributeError, ImportError):
	from speech.sayAll import CURSOR
	CURSOR_CARET = CURSOR.CARET
import speech
import speech.commands
from NVDAObjects.window.winword import (
	WinWordCollectionQuicknavIterator, WordDocumentRevisionQuickNavItem,
	WordDocumentCollectionQuickNavItem, WordDocumentCommentQuickNavItem,
	SpellingErrorWinWordCollectionQuicknavIterator, ChartWinWordCollectionQuicknavIterator,
	GraphicWinWordCollectionQuicknavIterator, WordDocumentTreeInterceptor,
	BrowseModeWordDocumentTextInfo,
	wdRevisionInsert,
	wdRevisionDelete
)
# from NVDAObjects.window.winword import *
from versionInfo import version_year, version_major
from .ww_fields import Field
from .ww_keyboard import getBrowseModeQuickNavKey
from scriptHandler import willSayAllResume
import api
import winsound
import sys
import os
try:
	# for nvda version >= 2021.2
	from controlTypes.outputReason import OutputReason
	REASON_QUICKNAV = OutputReason.QUICKNAV
except ImportError:
	try:
		# for nvda version == 2021.1
		from controlTypes import OutputReason
		REASON_QUICKNAV = OutputReason.QUICKNAV
		REASON_FOCUS = OutputReason.FOCUS
	except ImportError:
		# for nvda version < 2021.1
		REASON_QUICKNAV = browseMode.REASON_QUICKNAV
		import controlTypes
		REASON_FOCUS = controlTypes.Reason.FOCUS

_curAddon = addonHandler.getCodeAddon()
path = os.path.join(_curAddon.path, "shared")
sys.path.append(path)
from ww_NVDAStrings import NVDAString
from ww_addonConfigManager import _addonConfigManager, AutoReadingWith_Beep
del sys.path[-1]

addonHandler.initTranslation()


class BrowseModeTreeInterceptorEx(browseMode.BrowseModeTreeInterceptor):
	__gestures = {}
	scriptCategory = _("Extended browse mode for Microsoft Word")

	def __init__(self, rootNVDAObject):
		super(BrowseModeTreeInterceptorEx, self).__init__(rootNVDAObject)
		# to be able to move sentences by sentence in navigation mode.
		# collapseOrExpandControl script is not usefull in word.
		try:
			del self._gestureMap["kb:alt+uparrow"]
		except KeyError:
			pass
		try:
			del self._gestureMap["kb:alt+downarrow"]
		except KeyError:
			pass

	@classmethod
	def addQuickNav(cls, itemType, key, nextDoc, nextError, prevDoc, prevError, readUnit=None):
		map = cls.__gestures
		"""Adds a script for the given quick nav item.
		@param itemType: The type of item, I.E. "heading" "Link" ...
		@param key: The quick navigation key to bind to the script.
		Shift is automatically added for the previous item gesture. E.G. h for heading
		@param nextDoc: The command description to bind to the script that yields the next quick nav item.
		@param nextError: The error message if there are no more quick nav items of type itemType in this direction.
		@param prevDoc: The command description to bind to the script that yields the previous quick nav item.
		@param prevError: The error message if there are no more quick nav items of type itemType in this direction.
		@param readUnit: The unit (one of the textInfos.UNIT_* constants)
		to announce when moving to this type of item.
			For example, only the line is read when moving to tables to avoid reading a potentially massive table.
			If None, the entire item will be announced.
		"""
		scriptSuffix = itemType[0].upper() + itemType[1:]
		scriptName = "next%s" % scriptSuffix
		funcName = "script_%s" % scriptName
		script = lambda self, gesture: self._quickNavScript(
			gesture, itemType, "next", nextError, readUnit)
		script.__doc__ = nextDoc
		script.__name__ = funcName
		script.resumeSayAllMode = CURSOR_CARET
		setattr(cls, funcName, script)
		if key is not None:
			map["kb:%s" % key] = scriptName
		scriptName = "previous%s" % scriptSuffix
		funcName = "script_%s" % scriptName
		script = lambda self, gesture: self._quickNavScript(
			gesture, itemType, "previous", prevError, readUnit)
		script.__doc__ = prevDoc
		script.__name__ = funcName
		script.resumeSayAllMode = CURSOR_CARET
		setattr(cls, funcName, script)
		if key is not None:
			map["kb:shift+%s" % key] = scriptName


# Add quick navigation scripts.
qn = BrowseModeTreeInterceptorEx.addQuickNav
qn(
	"grammaticalError",
	key=getBrowseModeQuickNavKey("grammaticalError"),
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next grammatical error"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next grammatical error"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous grammatical error"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous grammatical error"))

qn(
	"revision",
	key=getBrowseModeQuickNavKey("revision"),
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next revision"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next revision"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous revision"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous revision"))

qn(
	"comment",
	key=getBrowseModeQuickNavKey("comment"),
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next comment"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next comment"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous comment"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous comment"))

qn(
	"field",
	key=getBrowseModeQuickNavKey("field"),
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next field"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next field"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous field"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous field"),
	readUnit=textInfos.UNIT_LINE)

qn(
	"bookmark",
	key=getBrowseModeQuickNavKey("bookmark"),
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next bookmark"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next bookmark"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous bookmark"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous bookmark"))

qn(
	"endnote",
	key=getBrowseModeQuickNavKey("endnote"),
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next endnote"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next endnote"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous endnote"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous endnote"))

qn(
	"footnote",
	key=getBrowseModeQuickNavKey("footnote"),
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next footnote"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next footnote"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous footnote"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous footnote"))

qn(
	"section",
	key=getBrowseModeQuickNavKey("section"),
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next section"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next section"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous section"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous section"))

del qn


class BrowseModeDocumentTreeInterceptorEx(
	BrowseModeTreeInterceptorEx, browseMode.BrowseModeDocumentTreeInterceptor):
	pass

	def _quickNavScript(
		self, gesture, itemType, direction, errorMessage, readUnit):
		if itemType == "notLinkBlock":
			iterFactory = self._iterNotLinkBlock
		else:
			iterFactory = lambda direction, info: self._iterNodesByType(
				itemType, direction, info)
		info = self.selection
		try:
			item = next(iterFactory(direction, info))
		except NotImplementedError:
			# Translators: a message when a particular quick nav command
			# is not supported in the current document.
			ui.message(NVDAString("Not supported in this document"))
			return
		except StopIteration:
			if not _addonConfigManager.toggleLoopInNavigationModeOption(False):
				ui.message(errorMessage)
				return
			# return to the top or bottom of page and continue search
			if direction == "previous":
				info = api.getReviewPosition().obj.makeTextInfo(textInfos.POSITION_LAST)
				self._set_selection(info, reason="quickNav")
				# Translators: message to the user which indicates the return
				# to the bottom of the page.
				msg = _("Return to bottom of page")
			else:
				info = None
				# Translators: message to user which indicates the return
				# to the top of the page.
				msg = _("Return to top of page")
			try:
				item = next(iterFactory(direction, info))
			except Exception:
				ui.message(errorMessage)
				return
			ui.message(msg)
			winsound.PlaySound("default", 1)
		# #8831: Report before moving because moving might change the focus, which
		# might mutate the document, potentially invalidating info if it is
		# offset-based.
		if not gesture or not willSayAllResume(gesture):
			item.report(readUnit=readUnit)
		item.moveTo()


NVDAVersion = [version_year, version_major]
if NVDAVersion < [2019, 3]:
	# automatic reading not available
	BrowseModeWordDocumentTextInfoEx = BrowseModeWordDocumentTextInfo
else:
	from .ww_automaticReading import AutomaticReadingWordTextInfo

	class BrowseModeWordDocumentTextInfoEx(AutomaticReadingWordTextInfo, BrowseModeWordDocumentTextInfo):
		pass


class WordDocumentTreeInterceptorEx(BrowseModeDocumentTreeInterceptorEx, WordDocumentTreeInterceptor):
	disableAutoPassThrough = True
	TextInfo = BrowseModeWordDocumentTextInfoEx

	def __init__(self, obj):
		super(WordDocumentTreeInterceptorEx, self).__init__(obj)
		self.passThrough = True
		browseMode.reportPassThrough.last = True
		# to keep WordDocumentEx scripts available in browseMode on
		gestures = [
			"kb:alt+upArrow",
			"kb: alt+downArrow",
			"kb:control+upArrow",
			"kb:control+downArrow"]
		for gest in gestures:
			try:
				self.removeGestureBinding(gest)
			except Exception:
				pass

	def _get_ElementsListDialog(self):
		from .ww_elementsListDialog import ElementsListDialog
		return ElementsListDialog

	def _iterNodesByType(self, nodeType, direction="next", pos=None):
		if pos:
			rangeObj = pos.innerTextInfo._rangeObj
		else:
			rangeObj = self.rootNVDAObject.WinwordDocumentObject.range(0, 0)
		includeCurrent = False if pos else True
		if nodeType == "bookmark":
			return BookmarkWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
		elif nodeType == "comment":
			return CommentWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
		elif nodeType == "revision":
			return RevisionWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
		elif nodeType == "endnote":
			return EndnoteWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
		elif nodeType == "field":
			fields = FieldWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
			formfields = FormFieldWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
			return browseMode.mergeQuickNavItemIterators([fields, formfields], direction)
		elif nodeType == "formfield":
			return FormFieldWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
		elif nodeType == "footnote":
			return FootnoteWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
		elif nodeType == "graphic":
			graphics = GraphicWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
			charts = ChartWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
			shapes = ShapeWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
			return browseMode.mergeQuickNavItemIterators(
				[graphics, charts, shapes], direction)
		elif nodeType == "grammaticalError":
			return GrammaticalErrorWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
		elif nodeType == "spellingError":
			return SpellingErrorWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
		elif nodeType == "section":
			return SectionWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
		elif nodeType == "annotation":
			comments = CommentWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
			revisions = RevisionWinWordCollectionQuicknavIterator(
				nodeType, self, direction, rangeObj, includeCurrent).iterate()
			return browseMode.mergeQuickNavItemIterators([comments, revisions], direction)
		return super(WordDocumentTreeInterceptorEx, self)._iterNodesByType(nodeType, direction, pos)


class WordDocumentCommentQuickNavItemEx(WordDocumentCommentQuickNavItem):
	@property
	def label(self):
		author = self.collectionItem.author
		date = self.collectionItem.date
		text = self.collectionItem.range.text
		if text is None:
			text = ""
		msg = NVDAString("comment: {text} by {author} on {date}")
		return msg.format(author=author, text=text, date=date)


class CommentWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentCommentQuickNavItemEx

	def collectionFromRange(self, rangeObj):
		return rangeObj.comments


class WordDocumentBookmarkQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		item = self.collectionItem
		text = ""
		if not item.empty:
			text = item.range.text
		name = item.Name if item.Name else _("Bookmark")
		msg = _("{name} text: {text}")
		return msg.format(name=name, text=text)

	def report(self, readUnit=None):
		ui.message(_("Bookmark %s" % self.collectionItem.index))


class BookmarkWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentBookmarkQuickNavItem

	def collectionFromRange(self, rangeObj):
		return rangeObj.Bookmarks


class WordDocumentEndnoteQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		msg = _("Endnote{index}")
		return msg.format(index=self.collectionItem.index)

	def rangeFromCollectionItem(self, item):
		return item.Reference

	def report(self, readUnit=None):
		textList = []
		textList.append(self.label)
		try:
			# only for nvda version >= 2019.3
			from .ww_endnotes import Endnote
			index = int(self.collectionItem.Index)
			doc = self.collectionItem.Application.ActiveDocument
			endnoteObj = doc.EndNotes[index]
			endnote = Endnote(self.collectionItem, endnoteObj)
			from .ww_automaticReading import formatAutoSpeechSequence
			if _addonConfigManager.toggleAutomaticReadingOption(False) and (
				_addonConfigManager.toggleAutoEndnoteReadingOption(False)):
				textList.extend(formatAutoSpeechSequence([endnote.text]))
		except Exception:
			pass
		speech.speak(textList)


class EndnoteWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentEndnoteQuickNavItem

	def collectionFromRange(self, rangeObj):
		return rangeObj.Endnotes


class WordDocumentFieldQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		typeText = Field._getTypeText(self.collectionItem.type)
		# Translators: field index and type.
		msg = _("Field {index}, type: {type}")
		return msg.format(index=self.collectionItem.index, type=typeText)

	def rangeFromCollectionItem(self, item):
		return item.Code


class FieldWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentFieldQuickNavItem

	def collectionFromRange(self, rangeObj):
		return rangeObj.fields


class WordDocumentFormFieldQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		msg = "{name}"
		return msg.format(name=self.collectionItem.Name)


class FormFieldWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentFormFieldQuickNavItem

	def collectionFromRange(self, rangeObj):
		return rangeObj.FormFields


class WordDocumentFootnoteQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		# Translators: label for a footnote.
		msg = _("Footnote {index}")
		return msg.format(index=self.collectionItem.index)

	def rangeFromCollectionItem(self, item):
		return item.Reference

	def report(self, readUnit=None):
		textList = []
		textList.append(self.label)
		try:
			# only for nvda version >= 2019.3
			from .ww_footnotes import Footnote
			index = int(self.collectionItem.Index)
			doc = self.collectionItem.Application.ActiveDocument
			footnoteObj = doc.FootNotes[index]
			footnote = Footnote(self.collectionItem, footnoteObj)
			from .ww_automaticReading import formatAutoSpeechSequence
			if _addonConfigManager.toggleAutomaticReadingOption(False) and (
				_addonConfigManager.toggleAutoFootnoteReadingOption(False)):
				textList.extend(formatAutoSpeechSequence([footnote.text]))
		except Exception:
			pass
		speech.speak(textList)

	def moveTo(self):
		info = self.textInfo.copy()
		info.collapse()
		self.document._set_selection(info, reason=REASON_QUICKNAV)


class FootnoteWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentFootnoteQuickNavItem

	def collectionFromRange(self, rangeObj):
		return rangeObj.Footnotes


class WordDocumentShapeQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		return "%s" % self.collectionItem.name

	def rangeFromCollectionItem(self, item):
		return item.Anchor

	def report(self, readUnit=None):
		ui.message(_("%s" % self.collectionItem.Name))


class ShapeWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentShapeQuickNavItem

	def collectionFromRange(self, rangeObj):
		try:
			return rangeObj.ShapeRange
		except Exception:
			return None


class WordDocumentGrammaticalErrorQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		text = self.collectionItem.text
		if len(text) > 100:
			text = "%s ..." % text[:100]
		msg = "{text}"
		return msg.format(text=text)

	def rangeFromCollectionItem(self, item):
		return item


class GrammaticalErrorWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentGrammaticalErrorQuickNavItem

	def collectionFromRange(self, rangeObj):
		return rangeObj.GrammaticalErrors


class WordDocumentSectionQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):

		msg = _("Section {index}")
		return msg.format(index=self.collectionItem.index)

	def report(self, readUnit=None):
		ui.message(_("Section %s" % self.collectionItem.index))


class SectionWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentSectionQuickNavItem

	def collectionFromRange(self, rangeObj):
		return rangeObj.Sections


class WordDocumentRevisionQuickNavItemEx(WordDocumentRevisionQuickNavItem):
	def report(self, readUnit=None):
		from .ww_revisions import revisionTypeText
		revisionType = self.collectionItem.type
		revisionTypeText = revisionTypeText.get(revisionType)
		info = self.textInfo
		if readUnit:
			fieldInfo = info.copy()
			info.collapse()
			info.move(readUnit, 1, endPoint="end")
			if info.compareEndPoints(fieldInfo, "endToEnd") > 0:
				# We've expanded past the end of the field, so limit to the end of the field.
				info.setEndPoint(fieldInfo, "endToEnd")
		if revisionType == wdRevisionDelete:
			info.expand(textInfos.UNIT_CHARACTER)
			speech.speakTextInfo(info, useCache=False, reason=REASON_FOCUS)
			return
		autoReadingWith = _addonConfigManager.getAutoReadingWithOption()
		autoReadingWithBeep = _addonConfigManager.toggleAutomaticReadingOption(
			False) and (autoReadingWith == AutoReadingWith_Beep)
		autoReading = autoReadingWithBeep and (
			(revisionType == wdRevisionInsert) and (_addonConfigManager.toggleAutoInsertedTextReadingOption(False))
			or _addonConfigManager.toggleAutoRevisedTextReadingOption(False))
		if autoReading:
			# we don't ear automatic reading beep in navigation mode
			formatConfig = config.conf["documentFormatting"]
			formatConfig = formatConfig.copy()
			formatConfig["reportRevisions"] = False
			seq = []
			from .ww_revisions import Revision
			rev = Revision(None, self.collectionItem)
			seq.append(rev.FormatRevisionTypeAndAuthorText())
			seq.append(speech.commands.EndUtteranceCommand())
			speech.speak(seq)
			speech.speakTextInfo(info, useCache=False, reason=REASON_FOCUS, formatConfig=formatConfig)
			return
		speech.speakTextInfo(info, useCache=False, reason=REASON_FOCUS)


class RevisionWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass = WordDocumentRevisionQuickNavItemEx

	def collectionFromRange(self, rangeObj):
		return rangeObj.revisions
