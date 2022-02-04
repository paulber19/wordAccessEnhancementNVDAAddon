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
		elif NVDAObjects.UIA.wordDocument.WordDocument in clsList:
			from . import ww_UIAWordDocument
			clsList.insert(0, ww_UIAWordDocument.UIAWordDocument)
			clsList.remove(NVDAObjects.UIA.wordDocument.WordDocument)

	def event_appModule_gainFocus(self):
		printDebug("Word: event_appModuleGainFocus")
		self.hasFocus = True

		from . import ww_automaticReading
		ww_automaticReading.initialize()
		if self.checkUseUIAForWord and config.conf["UIA"]["useInMSWordWhenAvailable"]:
			wx.CallAfter(
				myMessageBox,
				# Translators: message to user
				_(
					"""You have checked """
					"""the "Use UI Automation to access Microsoft Word document controls when available" option. """
					"""For the add-on to work properly, it is recommended not to check this option."""),
				# Translators: dialog title.
				_("Warning"),
				wx.OK | wx.ICON_ERROR)
			self.checkUseUIAForWord = False
			return
		if not config.conf["UIA"]["useInMSWordWhenAvailable"]:
			self.checkUseUIAForWord = True

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
			speech.speakMessage(_("Skip empty paragraphs"))
		else:
			# Translators: message to user
			# to report no skipping empty paragraph when moving by paragraph.
			speech.speakMessage(_("Don't skip empty paragraphs"))

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
					speech.speakMessage,
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
			speech.speakMessage(_("Not available for this Word version"))
			return
		focus = api.getFocusObject()
		from .ww_spellingChecker import SpellingChecker
		sc = SpellingChecker(focus, self.WinwordVersion)
		if not sc.isInSpellingChecker():
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				speech.speakMessage,
				# Translators: message to indicate the focus is not in spellAndGrammar checker.
				_("You are Not in the spelling checker"))
			return
		if focus.role == ROLE_PANE:
			# focus on the pane not not on an object of the pane
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				speech.speakMessage,
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

	def script_test(self, gesture):
		print("test word")
		ui.message("test word")

	__gestures = {
		"kb:control+windows+alt+f12": "test",
	}
