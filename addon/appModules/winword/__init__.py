# appModules\winword\__init__.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

import addonHandler
from versionInfo import version_year, version_major
import time
import scriptHandler
import core
import api
import config
import os
import wx
import ui
import queueHandler
import controlTypes
import speech
import NVDAObjects.window.winword
import NVDAObjects.UIA.wordDocument
import NVDAObjects.IAccessible.winword
from .ww_scriptTimer import GB_scriptTimer, stopScriptTimer, _delay
import sys
_curAddon = addonHandler.getCodeAddon()
debugToolsPath = os.path.join(_curAddon.path, "debugTools")
sys.path.append(debugToolsPath)
try:
	from appModuleDebug import AppModuleDebug as AppModule
	from appModuleDebug import printDebug, toggleDebugFlag
except ImportError:
	from appModuleHandler import AppModule as AppModule
	def printDebug(msg): return
	def toggleDebugFlag(): return
del sys.path[-1]

path = os.path.join(_curAddon.path, "shared")
sys.path.append(path)
from ww_utils import myMessageBox, maximizeWindow  # noqa:E402
from ww_py3Compatibility import _unicode  # noqa:E402
from ww_addonConfigManager import _addonConfigManager  # noqa:E402
del sys.path[-1]

addonHandler.initTranslation()
_addonSummary = _unicode(_curAddon.manifest['summary'])
_scriptCategory = _addonSummary


class AppModule(AppModule):
	layerMode = None
	reportAllCellsFlag = False

	def __init__(self, *args, **kwargs):
		printDebug("word appmodule init")
		super(AppModule, self).__init__(*args, **kwargs)
		# toggleDebugFlag()
		# configuration load
		self.hasFocus = False
		# install input gesture for recording autoreading voice profil
		NVDAVersion = [version_year, version_major]
		if NVDAVersion >= [2019, 3]:
			self.bindGesture("kb:windows+alt+f12", "setAutomaticReadingVoice")
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
		NVDAVersion = [version_year, version_major]
		if NVDAVersion >= [2019, 3]:
			from . import ww_automaticReading
			ww_automaticReading.initialize()
		if self.checkUseUIAForWord and config.conf["UIA"]["useInMSWordWhenAvailable"]:
			wx.CallAfter(
				myMessageBox,
				# Translators: message to user
				_("""You have checked the "Use UI Automation to access Microsoft &Word document controls when available")" option. For the add-on to work properly, it is recommended not to check this option."""),
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
		NVDAVersion = [version_year, version_major]
		if NVDAVersion >= [2019, 3]:
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
			controlTypes.roleLabels.get(obj.role), obj.windowClassName))
		if not hasattr(self, "WinwordWindowObject"):
			try:
				self.WinwordWindowObject = obj.WinwordWindowObject
				self.WinwordVersion = obj.WinwordVersion
			except:  # noqa:E722
				pass
		if not self.hasFocus:
			nextHandler()
			return
		if obj.windowClassName == "OpusApp":
			# to suppress double announce of document window title
			return
		# for spelling and grammar check ending window
		if obj.role == controlTypes.ROLE_BUTTON and obj.name.lower() == "ok":
			foreground = api.getForegroundObject()
			if foreground.windowClassName == "#32770"\
				and foreground.name == "Microsoft Word":
				lastChild = foreground.getChild(foreground.childCount-1)
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
		printDebug("event_typedCharacter: %s, ch= %s" % (controlTypes.roleLabels.get(obj.role), ch))  # noqa:E501
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
				and obj.role == controlTypes.ROLE_BUTTON:
				wx.CallLater(100, sc.sayErrorAndSuggestion, False, True)

	def script_toggleSkipEmptyParagraphsOption(self, gesture):
		if _addonConfigManager.toggleSkipEmptyParagraphsOption():
			# Translators: message to user
			# to report skipping of empty paragraph when moving by paragraph.
			speech.speakMessage(_("Skip empty paragraphs"))
		else:
			# Translators: message to user
			# to report no skipping empty paragraph when moving by paragraph.
			speech.speakMessage(_("Don't skip empty paragraphs"))
	# Translators: a description for a script.
	script_toggleSkipEmptyParagraphsOption.__doc__ = _("Toggle on or off the option to skip empty paragraphs")  # noqa:E501
	script_toggleSkipEmptyParagraphsOption.category = _scriptCategory

	def script_f7KeyStroke(self, gesture):
		def verify(oldSpeechMode):
			api.processPendingEvents()
			speech.cancelSpeech()
			speech.speechMode = oldSpeechMode
			focus = api.getFocusObject()
			from .ww_spellingChecker import SpellingChecker
			sc = SpellingChecker(focus, self.WinwordVersion)
			if not sc.isInSpellingChecker():
				return
			if focus.role == controlTypes.ROLE_PANE:
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
					speech.speakObject, focus, controlTypes.REASON_FOCUS)
		stopScriptTimer()
		if not self.isSupportedVersion():
			gesture.send()
			return
		focus = api.getFocusObject()
		from .ww_spellingChecker import SpellingChecker
		sc = SpellingChecker(focus, self.WinwordVersion)
		if not sc.isInSpellingChecker():
			# moving to spelling checker
			oldSpeechMode = speech.speechMode
			speech.speechMode = speech.speechMode_off
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
		speech.speakObject(api.getFocusObject(), controlTypes.REASON_FOCUS)

	def script_spellingCheckerHelper(self, gesture):
		global GB_scriptTimer
		if not self.isSupportedVersion():
			# Translators: message to the user when word version is not supported.
			speech.speakMessage(_("Not available for this Word version"))
			return
		stopScriptTimer()
		focus = api.getFocusObject()
		from .ww_spellingChecker import SpellingChecker
		sc = SpellingChecker(focus, self.WinwordVersion)
		if not sc.isInSpellingChecker():
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				speech.speakMessage,
				# Translators: message to indicate the focus is not in spellAndGrammar checker. # noqa:E501
				_("You are Not in the spelling checker"))
			return
		if focus.role == controlTypes.ROLE_PANE:
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
			GB_scriptTimer = core.callLater(
				_delay, sc.sayErrorAndSuggestion, spell=False, focusOnSuggestion=False)
		elif count == 1:
			GB_scriptTimer = core.callLater(
				_delay, sc.sayErrorAndSuggestion, spell=True, focusOnSuggestion=False)
		else:
			wx.CallAfter(sc.sayHelpText)
	# Translators: a description for a script.
	script_spellingCheckerHelper.__doc__ = _("Report error and suggestion displayed by the spelling checker.Twice:spell them. Third: say help text")  # noqa:E501
	script_spellingCheckerHelper.category = _scriptCategory

	def sayCurrentSentence(self):
		selection = self.WinwordWindowObject.Application.Selection
		queueHandler.queueFunction(
			queueHandler.eventQueue, ui.message, selection.Sentences(1).Text)

	def script_reportCurrentSentence(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self.sayCurrentSentence)
	# Translators: a description for a script.
	script_reportCurrentSentence.__doc__ = _("Report current sentence ")
	script_reportCurrentSentence.category = _scriptCategory

	def script_setAutomaticReadingVoice(self, gesture):
		from .import ww_automaticReading
		ww_automaticReading.saveCurrentSpeechSettings()

	script_setAutomaticReadingVoice .__doc__ = _("Record automatic reading voice's settings")  # noqa:E501
	script_setAutomaticReadingVoice .category = _scriptCategory

	def script_test(self, gesture):
		print("test word")
		ui.message("test word")

	__gestures = {
		"kb:control+windows+alt+f12": "test",

		# for spelling checker
		"kb:f7": "f7KeyStroke",
		"kb:nvda+shift+f7": "spellingCheckerHelper",
		"kb:nvda+control+f7": "reportCurrentSentence",
		# for empty paragraph
		"kb:windows+alt+f4": "toggleSkipEmptyParagraphsOption",
	}
