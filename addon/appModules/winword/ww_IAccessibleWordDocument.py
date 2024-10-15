# appModules\winword\__init__.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2022 paulber19
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

import addonHandler
# from logHandler import log
import os
import NVDAObjects.window.winword
import NVDAObjects.IAccessible.winword
from . import ww_wordDocumentBase
from NVDAObjects.IAccessible.winword import WordDocument
from . import ww_browseMode
from .automaticReading import AutomaticReadingWordTextInfo
import sys
_curAddon = addonHandler.getCodeAddon()
debugToolsPath = os.path.join(_curAddon.path, "debugTools")
sys.path.append(debugToolsPath)
try:
	from appModuleDebug import printDebug, toggleDebugFlag
except ImportError:

	def printDebug(msg):
		return

	def toggleDebugFlag():
		return
del sys.path[-1]

addonHandler.initTranslation()


class WindowWinwordWordDocumentTextInfo(
	AutomaticReadingWordTextInfo,
	NVDAObjects.window.winword.WordDocumentTextInfo):
	pass


class IAccessibleWordDocument(ww_wordDocumentBase.WordDocument, WordDocument):
	treeInterceptorClass = ww_browseMode.WordDocumentTreeInterceptorEx
	shouldCreateTreeInterceptor = True
	disableAutoPassThrough = True
	TextInfo = WindowWinwordWordDocumentTextInfo
