# appModules\winword\ww_keyboard.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2020-2021 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
from logHandler import log
import os.path
from configobj import ConfigObj
from configobj.validate import Validator
from io import StringIO

SCT_BrowseModeQuickNavKeys = "BrowseModeQuickNavKeys"
# default key assignement (based on qwerty keyboard)
# can be modified with the "browseModeQuickNavKeys" section
# in keyboard.ini file
_browseModeQuickNavConfigSpec = """
[{browseModeQuickNavKeys}]
	"grammaticalError" = string(default="]")
	"revision" = string(default="\\")
	"comment" = string(default="j")
	"field" = string(default="z")
	"bookmark" = string(default="-")
	"endnote" = string(default=".")
	"footnote" = string(default="/")
	"section" = string(default="[")
	""".format(browseModeQuickNavKeys=SCT_BrowseModeQuickNavKeys)


def getKeyboardKeysIniFilePath():
	from languageHandler import getLanguage
	curLang = getLanguage()
	langs = [curLang, ]
	addonFolderPath = addonHandler.getCodeAddon().path
	if '_' in curLang:
		langs.append(curLang.split("_")[0])
	langs.append("en")
	for lang in langs:
		langDir = os.path.join(addonFolderPath, "locale", lang)
		if os.path.exists(langDir):
			file = os.path.join(langDir, "keyboard.ini")
			if os.path.isfile(file):
				log.debugWarning("keyboard.ini file loaded from locale\\%s folder" % lang)
				return file
	log.error("keyboard.ini file not found")
	return ""


def _getBrowseModeQuickNavKeys():
	return _getKeyboardIniConfig()[SCT_BrowseModeQuickNavKeys]


def getBrowseModeQuickNavKey(script):
	return _getBrowseModeQuickNavKeys().get(script)


def _getKeyboardIniConfig():
	global _conf
	if _conf is not None:
		return _conf
	path = getKeyboardKeysIniFilePath()
	conf = ConfigObj(
		path,
		configspec=StringIO(_browseModeQuickNavConfigSpec),
		encoding="utf-8",
		list_values=False)
	conf.newlines = "\r\n"
	val = Validator()
	res = conf.validate(val, preserve_errors=True, copy=True)
	if not res:
		log.warning("KeyboardKeys configuration file is invalid: %s" % res)
	_conf = conf
	return conf


# singleton
_conf = None
