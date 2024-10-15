# appModules\winword\ww_browsemode.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2024 paulber19, Abdel
# This file is covered by the GNU General Public License.


import addonHandler
import NVDAObjects
import UIAHandler
import textInfos
import speech
import queueHandler
import scriptHandler
import api
import controlTypes
from . import ww_browseMode
from .automaticReading import AutomaticReadingWordTextInfo
from UIAHandler.browseMode import UIATextAttributeQuicknavIterator, TextAttribUIATextInfoQuickNavItem
from NVDAObjects.UIA.wordDocument import CommentUIATextInfoQuickNavItem, RevisionUIATextInfoQuickNavItem
from .import ww_elementsListDialog

addonHandler.initTranslation()


class FootNoteUIATextInfoQuickNavItem(TextAttribUIATextInfoQuickNavItem):
	attribID = UIAHandler.UIA_AnnotationTypesAttributeId
	wantedAttribValues = {UIAHandler.AnnotationType_Footnote, }

	@property
	def label(self):
		return self.textInfo.UIAElementAtStart.CurrentName

	def report(self, readUnit=None):
		readUnit = readUnit if readUnit else textInfos.UNIT_WORD
		queueHandler.queueFunction(
			queueHandler.eventQueue,
			speech.speakTextInfo,
			self.textInfo,
			unit=readUnit,
			reason=controlTypes.OutputReason.CHANGE
		)


class EndNoteUIATextInfoQuickNavItem(TextAttribUIATextInfoQuickNavItem):
	attribID = UIAHandler.UIA_AnnotationTypesAttributeId
	wantedAttribValues = {UIAHandler.AnnotationType_Endnote, }

	@property
	def label(self):
		return self.textInfo.UIAElementAtStart.CurrentName

	def report(self, readUnit=None):
		readUnit = readUnit if readUnit else textInfos.UNIT_WORD
		queueHandler.queueFunction(
			queueHandler.eventQueue,
			speech.speakTextInfo,
			self.textInfo,
			unit=readUnit,
			reason=controlTypes.OutputReason.CHANGE
		)


class UIAWordBrowseModeDocumentTextInfo(
	AutomaticReadingWordTextInfo,
	UIAHandler.browseMode.UIABrowseModeDocumentTextInfo
):

	pass


class UIAWordBrowseModeDocument(
	ww_browseMode.BrowseModeDocumentTreeInterceptorEx,
	NVDAObjects.UIA.wordDocument.WordBrowseModeDocument
):
	ElementsListDialog = ww_elementsListDialog.UIAElementsListDialog
	TextInfo = UIAWordBrowseModeDocumentTextInfo

	def __init__(self, rootNVDAObject):
		super(UIAWordBrowseModeDocument, self).__init__(rootNVDAObject)
		self.passThrough = True

	def script_moveByCharacter_back(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByCharacter_back(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_CHARACTER)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_CHARACTER, reason=controlTypes.OutputReason.CARET)

	def script_moveByCharacter_forward(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByCharacter_forward(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_CHARACTER)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_CHARACTER, reason=controlTypes.OutputReason.CARET)

	def script_moveByWord_back(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByWord_back(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_WORD)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_WORD, reason=controlTypes.OutputReason.CARET)

	def script_moveByWord_forward(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByWord_forward(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_WORD)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_LINE, reason=controlTypes.OutputReason.CARET)

	def script_moveByLine_back(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByLine_back(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_LINE)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_LINE, reason=controlTypes.OutputReason.CARET)

	def script_moveByLine_forward(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByLine_forward(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_LINE)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_LINE, reason=controlTypes.OutputReason.CARET)

	def script_moveByParagraph_back(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByParagraph_back(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		self.selection = info
		info.expand(textInfos.UNIT_PARAGRAPH)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_PARAGRAPH, reason=controlTypes.OutputReason.CARET)

	def script_moveByParagraph_forward(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByParagraph_forward(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		self.selection = info
		info.expand(textInfos.UNIT_PARAGRAPH)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_PARAGRAPH, reason=controlTypes.OutputReason.CARET)

	def script_moveByPage_back(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByPage_back(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		self.selection = info
		info.expand(textInfos.UNIT_LINE)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_LINE, reason=controlTypes.OutputReason.CARET)

	def script_moveByPage_forward(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_moveByPage_forward(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		self.selection = info
		info.expand(textInfos.UNIT_LINE)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_LINE, reason=controlTypes.OutputReason.CARET)

	def script_topOfDocument(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_topOfDocument(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		self.selection = info
		info.expand(textInfos.UNIT_LINE)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_LINE, reason=controlTypes.OutputReason.CARET)

	def script_bottomOfDocument(self, gesture):
		savedSpeechMode = speech.getState().speechMode
		speech.setSpeechMode(speech.SpeechMode.off)
		super(UIAWordBrowseModeDocument, self).script_bottomOfDocument(gesture)
		api.processPendingEvents()
		speech.setSpeechMode(savedSpeechMode)
		info = self.currentFocusableNVDAObject.makeTextInfo(textInfos.POSITION_CARET)
		self.selection = info
		info.expand(textInfos.UNIT_LINE)
		if not scriptHandler.willSayAllResume(gesture):
			speech.speakTextInfo(info, unit=textInfos.UNIT_LINE, reason=controlTypes.OutputReason.CARET)

	def _iterNodesByType(self, nodeType, direction="next", pos=None):
		if nodeType == "comment":
			return UIATextAttributeQuicknavIterator(
				CommentUIATextInfoQuickNavItem, nodeType, self, pos, direction=direction)
		elif nodeType == "revision":
			return UIATextAttributeQuicknavIterator(
				RevisionUIATextInfoQuickNavItem, nodeType, self, pos, direction=direction)
		elif nodeType == "footnote":
			return UIATextAttributeQuicknavIterator(
				FootNoteUIATextInfoQuickNavItem, nodeType, self, pos, direction=direction)
		elif nodeType == "endnote":
			return UIATextAttributeQuicknavIterator(
				EndNoteUIATextInfoQuickNavItem, nodeType, self, pos, direction=direction)
		return super(UIAWordBrowseModeDocument, self)._iterNodesByType(
			nodeType, direction=direction, pos=pos)
