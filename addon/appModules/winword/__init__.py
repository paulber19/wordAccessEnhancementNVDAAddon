# appModules\winword\__init__.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2024 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import time
import scriptHandler
from scriptHandler import script
import config
import core
import api
import os
import wx
import queueHandler
import speech
import gui
import ui
import NVDAObjects.window.winword
import NVDAObjects.UIA.wordDocument
import NVDAObjects.IAccessible.winword
from .ww_scriptTimer import stopScriptTimer, delayScriptTask
import sys
from controlTypes.role import Role
from controlTypes.outputReason import OutputReason
from controlTypes.role import _roleLabels as roleLabels
from .NVDAAppModuleWinword import WinwordWordDocument as NVDAAppModuleWinwordDocument

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
from ww_NVDAStrings import NVDAString, NVDAString_pgettext
from ww_utils import (
	maximizeWindow, makeAddonWindowTitle, isOpened,
	getSpeechMode, setSpeechMode, setSpeechMode_off,
	messageWithSpeakOnDemand,
)
del sys.path[-1]
del sys.modules["ww_NVDAStrings"]
del sys.modules["ww_utils"]

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
			# from nvdaBuiltin.appModules.winword import WinwordWordDocument
			clsList.insert(0, NVDAAppModuleWinwordDocument)
		elif NVDAObjects.UIA.wordDocument.WordDocument in clsList:
			from . import ww_UIAWordDocument
			clsList.insert(0, ww_UIAWordDocument.UIAWordDocument)
			clsList.remove(NVDAObjects.UIA.wordDocument.WordDocument)
			# from nvdaBuiltin.appModules.winword import WinwordWordDocument
			clsList.insert(0, NVDAAppModuleWinwordDocument)

	def event_appModule_gainFocus(self):
		printDebug("Word: event_appModuleGainFocus")
		self.hasFocus = True

		from . import automaticReading
		automaticReading.initialize()

	def event_appModule_loseFocus(self):
		printDebug("Word: event_appModuleLoseFocus")
		self.hasFocus = False
		from . import automaticReading
		automaticReading.terminate()

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
		if obj.role == Role.BUTTON and obj.name.lower() == "ok":
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
				and obj.role == Role.BUTTON:
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
			if focus.role == Role.PANE:
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
					speech.speakObject, focus, OutputReason.FOCUS)
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
		speech.speakObject(api.getFocusObject(), OutputReason.FOCUS)

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
			messageWithSpeakOnDemand(_("Not available for this Word version"))
			return
		focus = api.getFocusObject()
		from .ww_spellingChecker import SpellingChecker
		sc = SpellingChecker(focus, self.WinwordVersion)
		if not sc.isInSpellingChecker():
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				messageWithSpeakOnDemand,
				# Translators: message to indicate the focus is not in spellAndGrammar checker.
				_("You are Not in the spelling checker"))
			return
		if focus.role == Role.PANE:
			# focus on the pane not not on an object of the pane
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				messageWithSpeakOnDemand,
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
			queueHandler.eventQueue,
			messageWithSpeakOnDemand,
			selection.Sentences(1).Text
		)

	@script(
		# Translators: Input help mode message for report Current Sentence command.
		description=_("Report current sentence"),
		gesture="kb:nvda+control+f7"
	)
	def script_reportCurrentSentence(self, gesture):
		stopScriptTimer()
		wx.CallAfter(self.sayCurrentSentence)

	@script(
		# Translators: Input help mode message for set automatic reading voice command.
		description=_(
			"Record settings for automatic reading voice. "
			"Twice: Switch between the automatic reading  voice and the current voice for Word"
		),
		gesture="kb:windows+alt+f12"
	)
	def script_setAutomaticReadingVoice(self, gesture):
		stopScriptTimer()
		count = scriptHandler.getLastScriptRepeatCount()
		from .import automaticReading
		if count == 0:
			delayScriptTask(
				automaticReading.saveCurrentSpeechSettings)
		else:
			automaticReading.switchToAutomaticReadingSynth()

	def isThereHiddenText(self, selection):
		wdSelectionNormal = 2
		if selection.Type == wdSelectionNormal:
			# print ("font.hidden: %s"%Selection.Font.Hidden)
			wdUndefined = 9999999
			if selection.Font.Hidden == wdUndefined or selection.Font.Hidden:
				return True
		return False

	def isHidden(self, selection):
		rngTemp = selection.Range
		rngTemp.TextRetrievalMode.IncludeHiddenText = False

	def findHiddenText(self):
		focus = api.getFocusObject()
		winwordDocumentObject = focus.WinwordDocumentObject
		winwordWindowObject = self.WinwordWindowObject
		selection = winwordWindowObject.Selection
		sView = winwordDocumentObject .ActiveWindow.View.ShowHiddenText
		winwordDocumentObject .ActiveWindow.View.ShowHiddenText = True
		selection.Find.ClearFormatting
		selection.Find.Replacement.ClearFormatting
		if True:
			f = selection.Find
			f.Text = ""
			f.Font.Hidden = True
			f.Replacement.Text = ""
			f.Forward = True
			# wdFindContinue = 1
			wdFindStop = 0
			f.Wrap = wdFindStop
			f.Format = True
			f.MatchCase = False
			f.MatchWholeWord = False
			f.MatchWildcards = False
			f.MatchSoundsLike = False
			f.MatchAllWordForms = False
		found = f.Execute()
		if found:
			ui.message(selection.range.text)
			selection.Collapse(1)
		winwordDocumentObject .ActiveWindow.View.ShowHiddenText = sView

	@script(
		# Translators: Input help mode message for display Microsoft UIA dialog command.
		description=_(
			"Display Microsoft UIA dialog to define the Use of UI Automation "
			"to access Microsoft Word document controls"
		),
	)
	def script_displayMicrosoftUIADialog(self, gesture):
		wx.CallAfter(MicrosoftUIADialog.run)

	def script_test(self, gesture):
		print("test word")
		ui.message("test word")

	__gestures = {
		# "kb:control+windows+alt+f12": "test",
	}


class MicrosoftUIADialog(wx.Dialog):
	shouldSuspendConfigProfileTriggers = True
	_instance = None
	title = None

	def __new__(cls, *args, **kwargs):
		if MicrosoftUIADialog._instance is not None:
			return MicrosoftUIADialog._instance
		return wx.Dialog.__new__(cls)

	def __init__(self, parent):
		if MicrosoftUIADialog._instance is not None:
			return
		MicrosoftUIADialog._instance = self
		dialogTitle = NVDAString("Microsoft UI Automation")
		title = MicrosoftUIADialog.title = makeAddonWindowTitle(dialogTitle)
		super().__init__(parent, -1, title)
		self.doGui()

	def doGui(self):
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		sHelper = gui.guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		labelText = NVDAString_pgettext(
			"advanced.uiaWithMSWord",
			# Translators: Label for the Use UIA with MS Word combobox, in the MicrosoftUIADialog.
			"Use UI Automation to access Microsoft &Word document controls",
		)
		wordChoices = (
			NVDAString_pgettext("advanced.uiaWithMSWord", "Default (Where suitable)"),
			NVDAString_pgettext("advanced.uiaWithMSWord", "Only when necessary"),
			NVDAString_pgettext("advanced.uiaWithMSWord", "Where suitable"),
			NVDAString_pgettext("advanced.uiaWithMSWord", "Always"),
		)
		self.UIAInMSWordListBox = sHelper.addLabeledControl(labelText, wx.ListBox, choices=wordChoices)
		self.UIAInMSWordListBox.SetSelection(config.conf["UIA"]["allowInMSWord"])
		bHelper = sHelper.addDialogDismissButtons(
			gui.guiHelper.ButtonHelper(wx.HORIZONTAL))
		OKButton = bHelper.addButton(self, id=wx.ID_OK)

		closeButton = bHelper.addButton(self, id=wx.ID_CLOSE)
		mainSizer.Add(
			sHelper.sizer,
			border=gui.guiHelper.BORDER_FOR_DIALOGS,
			flag=wx.ALL)
		mainSizer.Fit(self)
		self.SetSizer(mainSizer)
		# events
		OKButton.Bind(wx.EVT_BUTTON, self.onOKButton)
		closeButton.Bind(wx.EVT_BUTTON, lambda evt: self.Destroy())
		self.SetEscapeId(wx.ID_CLOSE)
		OKButton.SetDefault()
		self.UIAInMSWordListBox.SetFocus()

	def Destroy(self):
		MicrosoftUIADialog._instance = None
		super().Destroy()

	def onOKButton(self, evt):
		config.conf["UIA"]["allowInMSWord"] = self.UIAInMSWordListBox.GetSelection()
		self.Close()
		evt.Skip()

	@classmethod
	def run(cls):
		if isOpened(cls):
			return
		gui.mainFrame.prePopup()
		d = cls(gui.mainFrame)
		d.CentreOnScreen()
		d.Show()
		gui.mainFrame.postPopup()
