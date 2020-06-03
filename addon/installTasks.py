# -*- coding: UTF-8 -*-
# install.py
# a part of wordAccessEnhancement add-on
# Copyright 2018-2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
import os
from logHandler import log
previousNameAndAuthor = ("word", "Paulber19")
previousConfigFileName = "WordAddon.ini"
saveConfigFileName = "addonConfig_old.ini"

def saveFile(theFile, path):
	import shutil
	if not os.path.exists(theFile): return
	try:
		shutil.copy(theFile, path)
		os.remove(theFile)
		log.warning("%s file copied in %s and deleted"%(path, theFile))
	except:
		log.warning("Error: %s file cannot be move to %s"%(theFile, path))

def onInstall():
	import globalVars
	import wx
	import gui
	import sys
	if sys.version.startswith("3"):
		curPath = os.path.dirname(__file__)
	else:
		curPath = os.path.dirname(__file__).decode("mbcs")
	from addonHandler import _availableAddons 
	addon = _availableAddons [curPath]
	addonName = addon.manifest["name"]
	addonSummary = addon.manifest["summary"]
	# save old configuration
	userConfigPath = globalVars.appArgs.configPath
	curConfigFileName = "%sAddon.ini"%addonName
	f= os.path.join(userConfigPath, curConfigFileName)
	if not os.path.exists(f): return
	if gui.messageBox(
		# Translators: the label of a message box dialog  to ask the user if he wants keep current configuration settings.
		_("Do you want to keep current add-on configuration settings ?"),
		# Translators: the title of a message box dialog.
		_("%s - installation"%addonSummary),
		wx.YES|wx.NO|wx.ICON_WARNING)==wx.NO:
		return
	path = os.path.join(curPath, curConfigFileName )
	saveFile(f, path)
	curAutoReadingSynthFileName = "%s_autoReadingSynth.pickle"%addonName
	f= os.path.join(userConfigPath, curAutoReadingSynthFileName )
	path = os.path.join(curPath, curAutoReadingSynthFileName )
	saveFile(f, path)

def deleteFile(theFile):
	if not os.path.exists(theFile): return
	os.remove(theFile)
	if os.path.exists(theFile):
		log.warning("Error on deletion of%s  file"%theFile)
	else:
		log.warning("%s file deleted"%theFile)

def deleteAddonConfig():
	import globalVars
	import sys
	if sys.version.startswith("3"):
		curPath = os.path.dirname(__file__)
	else:
		curPath = os.path.dirname(__file__).decode("mbcs")
	sys.path.append(curPath)
	import buildVars
	addonName = buildVars.addon_info["addon_name"]
	del sys.path[-1]
	configFile = os.path.join(globalVars.appArgs.configPath, "%sAddon.ini"%addonName)
	deleteFile(configFile )
	autoReadingSynthFile = os.path.join(globalVars.appArgs.configPath, "%s_autoReadingSynth.pickle"%addonName)
	deleteFile(autoReadingSynthFile)

def onUninstall():
	deleteAddonConfig(  )

