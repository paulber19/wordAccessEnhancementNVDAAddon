# globalPlugins\winword\ww_config.py
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



from .__init__ import _addonSettingsWindowTitle 

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
		self.skipEmptyParagraphBox = sHelper.addItem(wx.CheckBox(self,wx.ID_ANY,label=_("&Skip empty paragraph")))
		self.skipEmptyParagraphBox.SetValue(_addonConfigManager.toggleSkipEmptyParagraphsOption(False))
		self.playSoundOnSkippedParagraphBox = sHelper.addItem(wx.CheckBox(self,wx.ID_ANY,label=_("&Play sound when paragraph is skipped")))
		self.playSoundOnSkippedParagraphBox.SetValue(_addonConfigManager.togglePlaySoundOnSkippedParagraphOption(False))
	
	def postInit(self):
		self.skipEmptyParagraphBox.SetFocus()
	
	def onOk (self, evt):
		if self.skipEmptyParagraphBox.IsChecked() != _addonConfigManager.toggleSkipEmptyParagraphsOption(False):
			_addonConfigManager.toggleSkipEmptyParagraphsOption()
		if self.playSoundOnSkippedParagraphBox.IsChecked() != _addonConfigManager.togglePlaySoundOnSkippedParagraphOption(False):
			_addonConfigManager.togglePlaySoundOnSkippedParagraphOption()
		super(WordAddonSettingsDialog, self).onOk(evt)

