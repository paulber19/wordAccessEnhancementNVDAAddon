# -*- coding: UTF-8 -*-
import os.path

# Build customizations
# Change this file instead of sconstruct or manifest files, whenever possible.


# Full getext (please don't change)
def _(arg):
	return arg


# Add-on information variables
addon_info = {
	# for previously unpublished addons,
	# please follow the community guidelines at:
	# https://bitbucket.org/nvdaaddonteam/todo/raw/master/guideLines.txt
	# add-on Name, internal for nvda
	"addon_name": "wordAccessEnhancement",
	# Add-on summary, usually the user visible name of the addon.
	# Translators: Summary for this add-on to be shown
	# on installation and add-on information.
	"addon_summary": _("Microsoft Word text editor:   accessibility enhancement"),
	# Add-on description
	# Translators: Long description to be shown for this add-on
	# on add-on information from add-ons manager
	"addon_description": _(
		"""This add-on adds extra functionality when working with Microsoft Word:
* a script ("windows+alt+f5") to display a dialog box to choose between most of objects 's type """
		"""to be listed """
		"""(like comments, revisions, bookmarks, fields, endnotes, footnotes, spelling errors, grammar errors,...),
* a script ("Alt+delete") to announce line, column  and page of  cursor position, """
		"""or start  and end of selection, or current table's cell,
* a script ("windows+alt+f2") to insert a comment,
* a script ("windows+alt+m") to report revision at cursor's  position,
* a script ("windows+alt+n") to report endNote  or footNote at cursor's position,
* modify the NVDA scripts "control+downArrow" and "Control+Uparrow" """
		"""(which moves the carret paragraph by paragraph) to skip the empty paragraph (optionnal),
* some scripts to move in table and read table 's elements (row, column, cell),
* adds specific Word browse mode command keys,
* possibility to move sentence by sentence ("alt+ downArrow" and "alt+upArrow"),
* a script to report document 's informations(windows+alt+f1"),
* accessibility enhancement for spelling checker (Word 2013 and 2016):
	* a script (NVDA+shift+f7") to report spelling or grammatical error
		and suggested correction by the spelling checker,
	* a script (NVDA+control+f7") to report current sentence under focus.
* automatic reading of some elements as comments ,  footnotes, or endnotes.


This add-on has been tested with Microsoft Word 2019, 2016 and 2013 (perhaps works also with Word 365).
"""),

	# version
	"addon_version": "3.6",
	# Author(s)
	"addon_author": "paulber19",
	# URL for the add-on documentation support
	"addon_url": "paulber19@laposte.net",
	# Documentation file name
	"addon_docFileName": "addonUserManual.html",
	# Minimum NVDA version supported (e.g. "2018.3")
	"addon_minimumNVDAVersion": "2023.1",
	# Last NVDA version supported/tested
	# (e.g. "2018.4", ideally more recent than minimum version)
	"addon_lastTestedNVDAVersion": "2024.4",
	# Add-on update channel (default is stable or None)
	"addon_updateChannel": None,
}


# Define the python files that are the sources of your add-on.
# You can use glob expressions here, they will be expanded.
mainPath = os.path.join("addon", "appModules", "winword")
pythonSources = [
	os.path.join("addon", "*.py"),
	os.path.join("addon", "appModules", "winword", "*.py"),
	os.path.join("addon", "appModules", "winword","automaticReading",  "*.py"),
	os.path.join("addon", "globalPlugins", "wordAccessEnhancement", "*.py"),
	os.path.join(
		"addon", "globalPlugins", "wordAccessEnhancement", "updateHandler", "*.py"),
	os.path.join("addon", "shared", "*.py"),
]

# Files that contain strings for translation. Usually your python sources
i18nSources = pythonSources

# Files that will be ignored when building the nvda-addon file
# Paths are relative to the addon directory,
# not to the root directory of your addon sources.
excludedFiles = []

# Base language for the NVDA add-on
# If your add-on is written in a language other than english, modify this variable.
# For example, set baseLanguage to "es" if your add-on is primarily written in spanish.
baseLanguage = "en"
