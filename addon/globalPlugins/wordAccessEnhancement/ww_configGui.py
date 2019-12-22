# globalPlugins\wordAccessEnhancement\ww_config.py
# A part of WordAccessEnhancement add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
import appModuleHandler
from logHandler import log
import os
import globalVars
import sys
import gui
import wx
from configobj import ConfigObj, ConfigObjError
# ConfigObj 5.1.0 and later integrates validate module.
try:
	from configobj.validate import Validator, VdtTypeError
except ImportError:
	from validate import Validator, VdtTypeError, ConfigObjError

_curAddon = addonHandler.getCodeAddon()
path = os.path.join(_curAddon.path, "shared")
sys.path.append(path)
from ww_py3Compatibility import importStringIO, _unicode
from ww_addonConfigManager import _addonConfigManager
del sys.path[-1]

StringIO = importStringIO()


class WordAddonSettingsDialog(gui.SettingsDialog):
	
	# Translators: This is the title for the Word addon settings dialog
	title = _("%s - settings")%_unicode(_curAddon.manifest['summary'])
	

	def makeSettings(self, settingsSizer):

		sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		# Translators: This is the label for a checkbox in the settings panel.
		self.skipEmptyParagraphBox = sHelper.addItem(wx.CheckBox(self,wx.ID_ANY,label=_("&Skip empty paragraph")))
		self.skipEmptyParagraphBox.SetValue(_addonConfigManager.toggleSkipEmptyParagraphsOption(False))
		# Translators: This is the label for a checkbox in the settings panel.
		self.playSoundOnSkippedParagraphBox = sHelper.addItem(wx.CheckBox(self,wx.ID_ANY,label=_("&Play sound when paragraph is skipped")))
		self.playSoundOnSkippedParagraphBox.SetValue(_addonConfigManager.togglePlaySoundOnSkippedParagraphOption(False))
		# Translators: This is the label for a group of editing options in the settings panel.
		groupText = _("Update")
		group = gui.guiHelper.BoxSizerHelper(self, sizer=wx.StaticBoxSizer(wx.StaticBox(self, label=groupText), wx.VERTICAL))
		sHelper.addItem(group)
		# Translators: This is the label for a checkbox in the settings panel.
		labelText = _("Automatically check for &updates ")
		self.autoCheckForUpdatesCheckBox=group.addItem (wx.CheckBox(self,wx.ID_ANY, label= labelText))
		self.autoCheckForUpdatesCheckBox.SetValue(_addonConfigManager.toggleAutoUpdateCheck(False))
		# Translators: This is the label for a checkbox in the settings panel.
		labelText = _("Update also release versions to &developpement versions")
		self.updateReleaseVersionsToDevVersionsCheckBox=group.addItem (wx.CheckBox(self,wx.ID_ANY, label= labelText))
		self.updateReleaseVersionsToDevVersionsCheckBox.SetValue(_addonConfigManager.toggleUpdateReleaseVersionsToDevVersions     (False))
		# Translators: This is the label for a button in the settings panel.
		labelText = _("&Check for update")
		checkForUpdateButton= wx.Button(self, label=labelText)
		group.addItem (checkForUpdateButton)
		checkForUpdateButton.Bind(wx.EVT_BUTTON,self.onCheckForUpdate)	
	
	def onCheckForUpdate(self, evt):
		from .updateHandler import addonUpdateCheck
		wx.CallAfter(addonUpdateCheck, auto = False, releaseToDev = _addonConfigManager.toggleUpdateReleaseVersionsToDevVersions(False))
		self.Close()

	
	def postInit(self):
		self.skipEmptyParagraphBox.SetFocus()
	def saveSettingChanges (self):
		if self.skipEmptyParagraphBox.IsChecked() != _addonConfigManager.toggleSkipEmptyParagraphsOption(False):
			_addonConfigManager.toggleSkipEmptyParagraphsOption()
		if self.playSoundOnSkippedParagraphBox.IsChecked() != _addonConfigManager.togglePlaySoundOnSkippedParagraphOption(False):
			_addonConfigManager.togglePlaySoundOnSkippedParagraphOption()	
		if self.autoCheckForUpdatesCheckBox.IsChecked() != _addonConfigManager .toggleAutoUpdateCheck(False):
			_addonConfigManager .toggleAutoUpdateCheck(True)
		if self.updateReleaseVersionsToDevVersionsCheckBox.IsChecked() != _addonConfigManager .toggleUpdateReleaseVersionsToDevVersions     (False):
			_addonConfigManager .toggleUpdateReleaseVersionsToDevVersions     (True)			
	
	def onOk (self, evt):
		self.saveSettingChanges()
		super(WordAddonSettingsDialog, self).onOk(evt)

