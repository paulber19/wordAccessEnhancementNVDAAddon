# globalPlugins\wordAccessEnhancement\ww_config.py
# A part of WordAccessEnhancement add-on
# Copyright (C) 2019-2025 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import os
import sys
import gui
from gui import nvdaControls
from gui.settingsDialogs import MultiCategorySettingsDialog, SettingsPanel
import wx
import characterProcessing
import speech
from versionInfo import version_year, version_major
NVDAVersion = [version_year, version_major]
_curAddon = addonHandler.getCodeAddon()
path = os.path.join(_curAddon.path, "shared")
sys.path.append(path)
from ww_informationDialog import InformationDialog
from ww_NVDAStrings import NVDAString
from ww_addonConfigManager import _addonConfigManager
del sys.path[-1]
del sys.modules["ww_informationDialog"]
del sys.modules["ww_NVDAStrings"]

addonHandler.initTranslation()


NVDASpeechSettings = [
	"autoLanguageSwitching",
	"autoDialectSwitching",
	"symbolLevel",
	"trustVoiceLanguage",
	"includeCLDR"]
NVDASpeechManySettings = [
	"capPitchChange",
	"sayCapForCapitals",
	"beepForCapitals",
	"useSpellingFunctionality"]
SCT_Speech = "speech"
SCT_Many = "__many__"


class WordOptionsPanel(SettingsPanel):
	# Translators: This is the label for the Options panel.
	title = _("Options")

	def makeSettings(self, settingsSizer):

		sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		# Translators: This is the label for a group of editing options
		# in the settings panel.
		groupText = _("Paragraph")
		group = gui.guiHelper.BoxSizerHelper(
			self,
			sizer=wx.StaticBoxSizer(wx.StaticBox(self, label=groupText), wx.VERTICAL))
		sHelper.addItem(group)
		# Translators: This is the label for a checkbox in the settings panel.
		labelText = _("&Skip empty paragraphs")
		self.skipEmptyParagraphBox = group.addItem(
			wx.CheckBox(self, wx.ID_ANY, label=labelText))
		self.skipEmptyParagraphBox.SetValue(
			_addonConfigManager.toggleSkipEmptyParagraphsOption(False))
		# Translators: This is the label for a checkbox in the settings panel.
		labelText = _("&Play sound when paragraph is skipped")
		self.playSoundOnSkippedParagraphBox = group.addItem(
			wx.CheckBox(self, wx.ID_ANY, label=labelText))
		self.playSoundOnSkippedParagraphBox.SetValue(
			_addonConfigManager.togglePlaySoundOnSkippedParagraphOption(False))
		# Translators: This is the label for a checkbox in the options settings panel.
		labelText = _("Na&vigate in loop")
		self.loopInNavigationModeOptionBox = sHelper.addItem(wx.CheckBox(self, wx.ID_ANY, label=labelText))
		self.loopInNavigationModeOptionBox.SetValue(_addonConfigManager.toggleLoopInNavigationModeOption(False))
		choice = [x for x in range(5, 125, 5)]
		choice = list(reversed(choice))
		# translators: label for a list box in Options settings panel.
		labelText = _("Maximum time of elements's search (in seconds:)")
		self.elementsSearchMaxTimeListBox = sHelper.addLabeledControl(
			labelText, wx.Choice, choices=[str(x) for x in choice])
		self.elementsSearchMaxTimeListBox.SetSelection(choice.index(_addonConfigManager.getElementsSearchMaxTime()))

	def postInit(self):
		self.skipEmptyParagraphBox .SetFocus()

	def saveSettingChanges(self):
		if self.skipEmptyParagraphBox.IsChecked() != _addonConfigManager.toggleSkipEmptyParagraphsOption(False):
			_addonConfigManager.toggleSkipEmptyParagraphsOption()
		option = _addonConfigManager.togglePlaySoundOnSkippedParagraphOption(False)
		if self.playSoundOnSkippedParagraphBox.IsChecked() != option:
			_addonConfigManager.togglePlaySoundOnSkippedParagraphOption()
		if self.loopInNavigationModeOptionBox.IsChecked() != (
			_addonConfigManager.toggleLoopInNavigationModeOption(False)):
			_addonConfigManager.toggleLoopInNavigationModeOption()
		elementsSearchMaxTime = int(
			self.elementsSearchMaxTimeListBox.GetString(self.elementsSearchMaxTimeListBox.GetSelection()))
		_addonConfigManager.setElementsSearchMaxTime(elementsSearchMaxTime)

	def onSave(self):
		self.saveSettingChanges()


class AutomaticReadingPanel(SettingsPanel):
	# Translators: This is the label for the Automatic reading panel.
	title = _("Automatic reading")
	objectsToRead = [
		"comment",
		"footnote",
		"endnote",
		"insertedText",
		"deletedText",
		"revisedText",
	]
	objectsToReadLabels = {
		"comment": _("Comment"),
		"footnote": _("Footnote"),
		"endnote": _("Endnote"),
		"insertedText": _("Inserted text"),
		"deletedText": _("Deleted text"),
		"revisedText": _("Revised text"),
	}

	def makeSettings(self, settingsSizer):
		sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		# Translators: This is the label for a checkbox in the automatic reading settings panel.
		labelText = _("&Activate automatic reading")
		self.automaticReadingCheckBox = sHelper.addItem(
			wx.CheckBox(self, wx.ID_ANY, label=labelText))
		self.automaticReadingCheckBox.SetValue(
			_addonConfigManager.toggleAutomaticReadingOption(False))
		# Translators: This is the label for a list of checkboxes
		# controlling which  object are automatically reading.
		labelText = _("Concerned elements:")
		choice = [self.objectsToReadLabels[x]for x in self.objectsToRead]
		self.objectsToReadCheckListBox = sHelper.addLabeledControl(
			labelText, nvdaControls.CustomCheckListBox, choices=choice)
		checkedItems = []
		if _addonConfigManager.toggleAutoCommentReadingOption(False):
			checkedItems.append(self.objectsToRead.index("comment"))
		if _addonConfigManager.toggleAutoFootnoteReadingOption(False):
			checkedItems.append(self.objectsToRead.index("footnote"))
		if _addonConfigManager.toggleAutoEndnoteReadingOption(False):
			checkedItems.append(self.objectsToRead.index("endnote"))
		if _addonConfigManager.toggleAutoInsertedTextReadingOption(False):
			checkedItems.append(self.objectsToRead.index("insertedText"))
		if _addonConfigManager.toggleAutoDeletedTextReadingOption(False):
			checkedItems.append(self.objectsToRead.index("deletedText"))
		if _addonConfigManager.toggleAutoRevisedTextReadingOption(False):
			checkedItems.append(self.objectsToRead.index("revisedText"))
		self.objectsToReadCheckListBox.CheckedItems = checkedItems
		self.objectsToReadCheckListBox.Select(0)
		# Translators: This is the label for a checkbox in the settings panel.
		labelText = _("Read &with:")
		choice = [
			# Translators: a choice labels for automatic reading.
			_("Current voice"),
			# Translators: a choice labels for automatic reading.
			_("Current voice and  beep at start and end"),
			# Translators: a choice labels for automatic reading.
			_("another voice")
		]
		self.autoReadingWithChoiceBox = sHelper.addLabeledControl(
			labelText, wx.Choice, choices=choice)
		self.autoReadingWithChoiceBox.SetSelection(
			_addonConfigManager.getAutoReadingWithOption())
		# Translators:  this is a label for a checkbox
		#  in automatic reading settings panel
		labelText = _("&Report revision's type with author")
		self.reportRevisionTypeWithAuthorCheckBox = sHelper.addItem(
			wx.CheckBox(self, wx.ID_ANY, label=labelText))
		self.reportRevisionTypeWithAuthorCheckBox.SetValue(
			_addonConfigManager.toggleReportRevisionTypeWithAuthorOption(False))
		# translators: label for a button in automatic reading settings panel.
		labelText = _("&Display voice's recorded setting")
		voiceInformationsButton = wx.Button(self, label=labelText)
		sHelper.addItem(voiceInformationsButton)
		voiceInformationsButton.Bind(wx.EVT_BUTTON, self.onVoiceInformationButton)

	def getSynthInformations(self, autoReadingSynth ):
		def get_NVDASpeechSettings():
			if NVDAVersion < [2024, 3]:
				return [
					"autoLanguageSwitching",
					"autoDialectSwitching",
					"symbolLevel",
					"trustVoiceLanguage",
					"delayedCharacterDescriptions"]
			else:
				return [
					"autoLanguageSwitching",
					"autoDialectSwitching",
					"symbolLevel",
					"trustVoiceLanguage",
					"unicodeNormalization",
					"reportNormalizedForCharacterNavigation",
					"delayedCharacterDescriptions"]

		def getNVDASpeechSettingsInfos():
			if NVDAVersion < [2024, 3]:
				return [
					("Automatic language switching (when supported)", boolToText),
					("Automatic dialect switching (when supported)", boolToText),
					("Punctuation/symbol level", punctuationLevelToText),
					("Trust voice's language when processing characters and symbols", boolToText),
					("Delayed descriptions for characters on cursor movement", boolToText),
				]
			else:
				return [
					("Automatic language switching (when supported)", boolToText),
					("Automatic dialect switching (when supported)", boolToText),
					(_("Punctuation/symbol level"), punctuationLevelToText),
					("Trust voice's language when processing characters and symbols", boolToText),
					("Unicode normali&zation", featureFlagToText),
					("Report '&Normalized' when navigating by character", boolToText),
					("&Delayed descriptions for characters on cursor movement", boolToText),
				]


		def boolToText(value):
			if type(value) is str and value == "True":
				return _("yes")
			if type(value) is str and value == "False":
				return _("no")
			if type(value) is bool:
				text = _("yes") if value else _("no")
				return text
			log.error("boolToText: value is out of type range:  bool, str")
			return ""

		def punctuationLevelToText(level):
			return characterProcessing.SPEECH_SYMBOL_LEVEL_LABELS[int(level)]

		def featureFlagToText(value):
			return value.calculated().displayString

		NVDASpeechManySettingsInfos = [
			("Capital pitch change percentage", None),
			("Say &cap before capitals", boolToText),
			("&Beep for capitals", boolToText),
			("Use &spelling functionality if supported", boolToText),
		]

		autoReadingSynthName = autoReadingSynth.get("synthName")
		textList = []
		# Translators: text to report synthesizer name.
		textList.append("%s = %s" %(_("Synthetizer"), autoReadingSynthName))
		synthSpeechSettings = autoReadingSynth[SCT_Speech]
		synthDisplayInfos = autoReadingSynth["SynthDisplayInfos"]
		if NVDAVersion < [2025, 1]:
			# Translators:  label to report synthesizer output device .
			textList.append("%s = %s" % (_("Audio output device"), synthSpeechSettings["outputDevice"]))
		print("synthDisplayInfos: %s" %synthDisplayInfos)
		for i in synthDisplayInfos:
			print("i: %s" %i)
			label, val = synthDisplayInfos[i]
			textList.append("%s = %s" % (label,val))

		for setting in get_NVDASpeechSettings():
			val = synthSpeechSettings.get(setting, None)
			if val is None:
				continue
			index = get_NVDASpeechSettings().index(setting)
			(name, f) = getNVDASpeechSettingsInfos()[index]
			if f is not None:
				val = f(val)
			name = NVDAString(name).replace("&", "")
			textList.append("%s = %s" % (name, val))
		for setting in NVDASpeechManySettings:
			val = synthSpeechSettings[SCT_Many][setting]
			if setting in synthSpeechSettings:
				val = synthSpeechSettings[setting]
			else:
				val = synthSpeechSettings[SCT_Many][setting]
			index = NVDASpeechManySettings.index(setting)
			(name, f) = NVDASpeechManySettingsInfos[index]
			if f is not None:
				val = f(val)
			name = NVDAString(name).replace("&", "")
			textList.append("%s = %s" % (name, val))
		return textList
		

	def onVoiceInformationButton(self, evt):
		autoReadingSynth = _addonConfigManager.getAutoReadingSynthSettings()
		if autoReadingSynth is None:
			# Translators: message to user to report no automatic reading voice.
			speech.speakMessage(_("No voice recorded for automatic reading"))
			return
		textList = self.getSynthInformations(autoReadingSynth )
		# Translators: this is the title of informationdialog box
# to show voice profile informations.
		dialogTitle = _("Voice settings for automatic reading")
		infos = "\r\n".join(textList)
		InformationDialog.run(None, dialogTitle, "", infos, False)

	def postInit(self):
		self.automaticReadingCheckBox .SetFocus()

	def saveSettingChanges(self):
		if self.automaticReadingCheckBox.IsChecked() != _addonConfigManager.toggleAutomaticReadingOption(False):
			_addonConfigManager.toggleAutomaticReadingOption()
		option = _addonConfigManager.toggleAutoCommentReadingOption(False)
		if self.objectsToReadCheckListBox.IsChecked(self.objectsToRead.index("comment")) != option:
			_addonConfigManager.toggleAutoCommentReadingOption()
		option = _addonConfigManager.toggleAutoFootnoteReadingOption(False)
		if self.objectsToReadCheckListBox.IsChecked(self.objectsToRead.index("footnote")) != option:
			_addonConfigManager.toggleAutoFootnoteReadingOption()
		option = _addonConfigManager.toggleAutoFootnoteReadingOption(False)
		if self.objectsToReadCheckListBox.IsChecked(self.objectsToRead.index("endnote")) != option:
			_addonConfigManager.toggleAutoEndnoteReadingOption()
		option = _addonConfigManager.toggleAutoInsertedTextReadingOption(False)
		if self.objectsToReadCheckListBox.IsChecked(self.objectsToRead.index("insertedText")) != option:
			_addonConfigManager.toggleAutoInsertedTextReadingOption()
		option = _addonConfigManager.toggleAutoDeletedTextReadingOption(False)
		if self.objectsToReadCheckListBox.IsChecked(self.objectsToRead.index("deletedText")) != option:
			_addonConfigManager.toggleAutoDeletedTextReadingOption()
		option = _addonConfigManager.toggleAutoRevisedTextReadingOption(False)
		if self.objectsToReadCheckListBox.IsChecked(self.objectsToRead.index("revisedText")) != option:
			_addonConfigManager.toggleAutoRevisedTextReadingOption()
		_addonConfigManager.setAutoReadingWithOption(
			self.autoReadingWithChoiceBox .GetSelection())
		option = _addonConfigManager.toggleReportRevisionTypeWithAuthorOption(False)
		if self.reportRevisionTypeWithAuthorCheckBox.IsChecked() != option:
			_addonConfigManager.toggleReportRevisionTypeWithAuthorOption()

	def onSave(self):
		self.saveSettingChanges()


class WordUpdatePanel(SettingsPanel):
	# Translators: This is the label for the Update panel.
	title = _("Update")

	def makeSettings(self, settingsSizer):
		sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		# Translators: This is the label for a checkbox in the update settings panel.
		labelText = _("Automatically check for &updates ")
		self.autoCheckForUpdatesCheckBox = sHelper.addItem(
			wx.CheckBox(self, wx.ID_ANY, label=labelText))
		self.autoCheckForUpdatesCheckBox.SetValue(
			_addonConfigManager.toggleAutoUpdateCheck(False))
		# Translators: This is the label for a checkbox in the update settings panel.
		labelText = _("Update also release versions to &developpement versions")
		self.updateReleaseVersionsToDevVersionsCheckBox = sHelper.addItem(
			wx.CheckBox(self, wx.ID_ANY, label=labelText))
		self.updateReleaseVersionsToDevVersionsCheckBox.SetValue(
			_addonConfigManager.toggleUpdateReleaseVersionsToDevVersions(False))
		# translators: label for a button in update settings panel.
		labelText = _("&Check for update")
		checkForUpdateButton = wx.Button(self, label=labelText)
		sHelper.addItem(checkForUpdateButton)
		checkForUpdateButton.Bind(wx.EVT_BUTTON, self.onCheckForUpdate)
		# translators: this is a label for a button in update settings panel.
		labelText = _("View &history")
		seeHistoryButton = wx.Button(self, label=labelText)
		sHelper.addItem(seeHistoryButton)
		seeHistoryButton.Bind(wx.EVT_BUTTON, self.onSeeHistory)

	def onCheckForUpdate(self, evt):
		from .updateHandler import addonUpdateCheck
		self.saveSettingChanges()
		releaseToDevVersion = self.updateReleaseVersionsToDevVersionsCheckBox.IsChecked()
		wx.CallAfter(addonUpdateCheck, auto=False, releaseToDev=releaseToDevVersion)
		self.Close()

	def onSeeHistory(self, evt):
		addon = addonHandler.getCodeAddon()
		from languageHandler import getLanguage
		theFile = os.path.join(addon.path, "doc", getLanguage(), "changes.html")
		if not os.path.exists(theFile):
			lang = getLanguage().split("_")[0]
			theFile = os.path.join(addon.path, "doc", lang, "changes.html")
			if not os.path.exists(theFile):
				lang = "en"
				theFile = os.path.join(addon.path, "doc", lang, "changes.html")
		os.startfile(theFile)

	def saveSettingChanges(self):
		if self.autoCheckForUpdatesCheckBox.IsChecked() != _addonConfigManager .toggleAutoUpdateCheck(False):
			_addonConfigManager .toggleAutoUpdateCheck(True)
			from . updateHandler.update_check import setCheckForUpdate
			setCheckForUpdate(_addonConfigManager.toggleAutoUpdateCheck(False))
		option = _addonConfigManager .toggleUpdateReleaseVersionsToDevVersions(False)
		if self.updateReleaseVersionsToDevVersionsCheckBox.IsChecked() != option:
			_addonConfigManager .toggleUpdateReleaseVersionsToDevVersions(True)

	def postSave(self):
		pass

	def onSave(self):
		self.saveSettingChanges()


class AddonSettingsDialog(MultiCategorySettingsDialog):
	title = "%s -%s" % (_curAddon.manifest["summary"], _("Settings"))
	INITIAL_SIZE = (1000, 480)
	# Min height required to show the OK, Cancel, Apply buttons
	MIN_SIZE = (470, 240)

	categoryClasses = [
		WordOptionsPanel,
		AutomaticReadingPanel,
		WordUpdatePanel
	]

	def __init__(self, parent, initialCategory=None):
		curAddon = addonHandler.getCodeAddon()
		# Translators: title of add-on parameters dialog.
		dialogTitle = _("Settings")
		self.title = "%s - %s" % (curAddon.manifest["summary"], dialogTitle)
		super(AddonSettingsDialog, self).__init__(parent, initialCategory)
