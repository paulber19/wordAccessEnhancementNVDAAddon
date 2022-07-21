# appModules\winword\__init__.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2022 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import time
import scriptHandler
from scriptHandler import script
import core
import api
import config
import os
import wx
import ui
import queueHandler
import speech
import ui
import NVDAObjects.window.winword
import NVDAObjects.UIA.wordDocument
import NVDAObjects.IAccessible.winword
from .ww_scriptTimer import stopScriptTimer, delayScriptTask
import sys
try:
	# for nvda >= 2021.2
	from controlTypes.role import Role
	ROLE_BUTTON = Role.BUTTON
	ROLE_PANE = Role.PANE
	from controlTypes.outputReason import OutputReason
	from controlTypes.role import _roleLabels as roleLabels
	REASON_FOCUS = OutputReason.FOCUS
except ImportError:
	import controlTypes
	ROLE_PANE = controlTypes.ROLE_PANE
	ROLE_BUTTON = controlTypes.ROLE_BUTTON
	from controlTypes import roleLabels
	try:
		# for nvda version == 2021.1
		from controlTypes import OutputReason
		REASON_FOCUS = OutputReason.FOCUS
	except AttributeError:
		# fornvda version <  2020.1
		REASON_FOCUS = controlTypes.REASON_FOCUS


_curAddon = addonHandler.getCodeAddon()
debugToolsPath = os.path.join(_curAddon.path, "debugTools")
sys.path.append(debugToolsPath)
try:
	from appModuleDebug import AppModuleDebug as AppModule
	from appModuleDebug import printDebug, toggleDebugFlag
except ImportError:
	from appModuleHandler import AppModule as AppModule

	def printDebug(msg):
		return

	def toggleDebugFlag():
		return
del sys.path[-1]
sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_addonConfigManager import _addonConfigManager
from ww_utils import (
	maximizeWindow,
	getSpeechMode, setSpeechMode, setSpeechMode_off)
from ww_messageBox import myMessageBox
del sys.path[-1]

addonHandler.initTranslation()
_addonSummary = _curAddon.manifest['summary']
_scriptCategory = _addonSummary


class AppModule(AppModule):
	scriptCategory = _scriptCategory
	layerMode = None
	reportAllCellsFlag = False

	def __init__(self, *args, **kwargs):
		printDebug("word appmodule init")
		super(AppModule, self).__init__(*args, **kwargs)
		# toggleDebugFlag()
		# configuration load
		self.hasFocus = False

		self.checkUseUIAForWord = True

	def terminate(self):
		if hasattr(self, "checkObjectsTimer") and self.checkObjectsTimer is not None:
			self.checkObjectsTimer.Stop()
			self.checkObjectsTimer = None
		super(AppModule, self).terminate()

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if NVDAObjects.IAccessible.winword.WordDocument in clsList:
			from . import ww_IAccessibleWordDocument
			clsList.insert(0, ww_IAccessibleWordDocument.IAccessibleWordDocument)
			clsList.remove(NVDAObjects.IAccessible.winword.WordDocument)
			from nvdaBuiltin.appModules.winword import WinwordWordDocument
			clsList.insert(0, WinwordWordDocument)
		elif NVDAObjects.UIA.wordDocument.WordDocument in clsList:
			from . import ww_UIAWordDocument
			clsList.insert(0, ww_UIAWordDocument.UIAWordDocument)
			clsList.remove(NVDAObjects.UIA.wordDocument.WordDocument)
			from nvdaBuiltin.appModules.winword import WinwordWordDocument
			clsList.insert(0, WinwordWordDocument)

	def event_appModule_gainFocus(self):
		printDebug("Word: event_appModuleGainFocus")
		self.hasFocus = True

		from . import ww_automaticReading
		ww_automaticReading.initialize()

	def event_appModule_loseFocus(self):
		printDebug("Word: event_appModuleLoseFocus")
		self.hasFocus = False
		from . import ww_automaticReading
		ww_automaticReading.terminate()

	def event_foreground(self, obj, nextHandler):
		printDebug("word: event_foreground")
		if obj.windowClassName == "OpusApp":
			self.opusApp = obj
		if obj.windowClassName in ["_WwG"]:
			maximizeWindow(obj.windowHandle)
		nextHandler()

	def event_gainFocus(self, obj, nextHandler):
		printDebug("Word: event_gainFocus: %s, %s" % (
			roleLabels.get(obj.role), obj.windowClassName))
		if not hasattr(self, "WinwordWindowObject"):
			try:
				self.WinwordWindowObject = obj.WinwordWindowObject
				self.WinwordVersion = obj.WinwordVersion
			except Exception:
				pass
		if not self.hasFocus:
			nextHandler()
			return
		if obj.windowClassName == "OpusApp":
			# to suppress double announce of document window title
			return
		# for spelling and grammar check ending window
		if obj.role == ROLE_BUTTON and obj.name.lower() == "ok":
			foreground = api.getForegroundObject()
			if foreground.windowClassName == "#32770"\
				and foreground.name == "Microsoft Word":
				lastChild = foreground.getChild(foreground.childCount - 1)
				if lastChild.windowClassName == "MSOUNISTAT":
					speech.speakMessage(foreground.description)
		nextHandler()

	def isSupportedVersion(self, obj=None):
		if obj is None:
			obj = self
		if hasattr(obj, "WinwordVersion") and obj.WinwordVersion in [15.0, 16.0]:
			return True
		return False

	def event_typedCharacter(self, obj, nextHandler, ch):
		printDebug("event_typedCharacter: %s, ch= %s" % (roleLabels.get(
			obj.role), ch))
		nextHandler()
		if not self.isSupportedVersion():
			return
		from .ww_spellingChecker import SpellingChecker
		sc = SpellingChecker(obj, self.WinwordVersion)
		if sc.isInSpellingChecker():
			# to report next error to be corrected
			# after a presse button or character shortcut hit.
			if ch > "a" and ch < "z"\
				or ord(ch) in [wx.WXK_SPACE, wx.WXK_RETURN]\
				and obj.role == ROLE_BUTTON:
				wx.CallLater(100, sc.sayErrorAndSuggestion, False, True)

	@script(
		# Translators: Input help mode message for toggle Skip Empty Paragraphs Option command.
		description=_("Toggle on or off the option to skip empty paragraphs"),
		gesture="kb:windows+alt+f4"
	)
	def script_toggleSkipEmptyParagraphsOption(self, gesture):
		stopScriptTimer()
		if _addonConfigManager.toggleSkipEmptyParagraphsOption():
			# Translators: message to user
			# to report skipping of empty paragraph when moving by paragraph.
			ui.message(_("Skip empty paragraphs"))
		else:
			# Translators: message to user
			# to report no skipping empty paragraph when moving by paragraph.
			ui.message(_("Don't skip empty paragraphs"))

	@script(
		gesture="kb:f7"
	)
	def script_f7KeyStroke(self, gesture):
		def verify(oldSpeechMode):
			api.processPendingEvents()
			speech.cancelSpeech()
			setSpeechMode(oldSpeechMode)
			focus = api.getFocusObject()
			from .ww_spellingChecker import SpellingChecker
			sc = SpellingChecker(focus, self.WinwordVersion)
			if not sc.isInSpellingChecker():
				return
			if focus.role == ROLE_PANE:
				# focus on the pane not not on an object of the pane
				queueHandler.queueFunction(
					queueHandler.eventQueue,
					ui.message,
					# Translators: message to ask user to hit tab key.
					_("Hit tab to move focus in the spelling checker pane"))
			else:
				sc.sayErrorAndSuggestion(title=True, spell=False, focusOnSuggestion=True)
				queueHandler.queueFunction(
					queueHandler.eventQueue,
					speech.speakObject, focus, REASON_FOCUS)
		stopScriptTimer()
		if not self.isSupportedVersion():
			gesture.send()
			return
		focus = api.getFocusObject()
		from .ww_spellingChecker import SpellingChecker
		sc = SpellingChecker(focus, self.WinwordVersion)
		if not sc.isInSpellingChecker():
			# moving to spelling checker
			oldSpeechMode = getSpeechMode()
			setSpeechMode_off()
			gesture.send()
			core.callLater(500, verify, oldSpeechMode)
			return
		# moving to document selection
		winwordWindowObject = self.WinwordWindowObject
		selection = winwordWindowObject.Selection
		selection.Characters(1).GoTo()
		time.sleep(0.1)
		api.processPendingEvents()
		speech.cancelSpeech()
		speech.speakObject(api.getFocusObject(), REASON_FOCUS)

	@script(
		# Translators: Input help mode message for spellingChecker Helper command.
		description=_(
			"Report error and suggestion displayed by the spelling checker."
			" Twice:spell them. Third: say help text"),
		gesture="kb:nvda+shift+f7"
	)
	def script_spellingCheckerHelper(self, gesture):
		stopScriptTimer()
		if not self.isSupportedVersion():
			# Translators: message to the user when word version is not supported.
			ui.message(_("Not available for this Word version"))
			return
		focus = api.getFocusObject()
		from .ww_spellingChecker import SpellingChecker
		sc = SpellingChecker(focus, self.WinwordVersion)
		if not sc.isInSpellingChecker():
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				ui.message,
				# Translators: message to indicate the focus is not in spellAndGrammar checker.
				_("You are Not in the spelling checker"))
			return
		if focus.role == ROLE_PANE:
			# focus on the pane not not on an object of the pane
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				ui.message,
				# Translators: message to ask user to hit tab key.
				_("Hit tab to move focus in the spelling checker pane"))
			return
		count = scriptHandler.getLastScriptRepeatCount()
		count = scriptHandler.getLastScriptRepeatCount()
		if count == 0:
			delayScriptTask(
				sc.sayErrorAndSuggestion,
				spell=False,
				focusOnSuggestion=False)
		elif count == 1:
			delayScriptTask(
				sc.sayErrorAndSuggestion,
				spell=True,
				focusOnSuggestion=False)
		else:
			wx.CallAfter(sc.sayHelpText)

	def sayCurrentSentence(self):
		selection = self.WinwordWindowObject.Application.Selection
		queueHandler.queueFunction(
			queueHandler.eventQueue, ui.message, selection.Sentences(1).Text)

	@script(
		# Translators: Input help mode message for report Current Sentence command.
		description=_("Report current sentence "),
		gesture="kb:nvda+control+f7"
	)
	def script_reportCurrentSentence(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self.sayCurrentSentence)

	@script(
		# Translators: Input help mode message for set automatic reading voice command.
		description=_("Record automatic reading voice's settings"),
		gesture="kb:windows+alt+f12"
	)
	def script_setAutomaticReadingVoice(self, gesture):
		stopScriptTimer()
		from .import ww_automaticReading
		ww_automaticReading.saveCurrentSpeechSettings()

	def isThereHiddenText(self, selection):
		wdSelectionNormal = 2
		if selection.Type == wdSelectionNormal:
			# print ("font.hidden: %s"%Selection.Font.Hidden)
			wdUndefined = 9999999
			if selection.Font.Hidden == wdUndefined or selection.Font.Hidden:
				print("text hidden")
				return True
		return False

	def isHidden(self, selection):
		rngTemp = selection.Range
		rngTemp.TextRetrievalMode.IncludeHiddenText = False
		if rngTemp.text == "":
			
			print("selection is hidden")


	def findHiddenText(self):
		focus = api.getFocusObject()
		winwordDocumentObject = focus.WinwordDocumentObject
		winwordWindowObject = self.WinwordWindowObject
		selection = winwordWindowObject.Selection
		sView = winwordDocumentObject .ActiveWindow.View.ShowHiddenText
		winwordDocumentObject .ActiveWindow.View.ShowHiddenText = True
		print ("selection.font.hidden: %s"%selection.font.hidden)
		selection.Find.ClearFormatting
		selection.Find.Replacement.ClearFormatting
		if True:
			f = selection.Find
			f.Text = ""
			f.Font.Hidden = True
			f.Replacement.Text = ""
			f.Forward = True
			wdFindContinue = 1
			wdFindStop = 0
			f.Wrap = wdFindStop
			f.Format = True
			f.MatchCase = False
			f.MatchWholeWord = False
			f.MatchWildcards = False
			f.MatchSoundsLike = False
			f.MatchAllWordForms = False
		import textInfos
		oldBookmark=focus.makeTextInfo(textInfos.POSITION_CARET).bookmark
		found = f.Execute()
		print ("found: %s"%found)
		if focus._hasCaretMoved(oldBookmark)[0]:
			info=focus.makeTextInfo(textInfos.POSITION_SELECTION)
			print ("info: %s"%info.text)



		if found:
			print ("text: %s"%selection.range.text)
			ui.message(selection.range.text)
			selection.Collapse(1)
		

		winwordDocumentObject .ActiveWindow.View.ShowHiddenText = sView
		

	def script_test(self, gesture):
		print("test word")
		ui.message("test word")
		self.findHiddenText()
		return
		focus = api.getFocusObject()
		winwordDocumentObject = focus.WinwordDocumentObject
		from .ww_wdConst import wdMainTextStory
		r = winwordDocumentObject .StoryRanges(wdMainTextStory)
		print("r: %s" % r.Font.Hidden)
		return
		winwordWindowObject = self.WinwordWindowObject
		selection = winwordWindowObject.Selection
		self.isThereHiddenText(selection)
		return
		# ActiveDocument.ActiveWindow.View.ShowHiddenText = True
		# winwordWindowObject.View.ShowHiddenText = False
		winwordWindowObject.View.ShowHiddenText = True
		return
		wdSelectionNormal = 2
		if selection.Type == wdSelectionNormal:
			selection.Font.Hidden = True

	__gestures = {
		"kb:control+windows+alt+f12": "test",
	}
