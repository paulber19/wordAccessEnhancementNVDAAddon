# appModules\winword\__init__.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2022 paulber19
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

import addonHandler
from versionInfo import version_year, version_major
# from logHandler import log
import os
import NVDAObjects.window.winword
import NVDAObjects.IAccessible.winword
from . import ww_wordDocumentBase
from . import ww_browseMode
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


_NVDAVersion = [version_year, version_major]
if _NVDAVersion < [2019, 3]:
	# automatic reading not available
	WindowWinwordWordDocumentTextInfo = NVDAObjects.window.winword.WordDocumentTextInfo
else:
	from .ww_automaticReading import AutomaticReadingWordTextInfo

	class WindowWinwordWordDocumentTextInfo(
		AutomaticReadingWordTextInfo,
		NVDAObjects.window.winword.WordDocumentTextInfo):
		pass
try:
	# new module winword.py in nvda 2020.4
	from appModules.winword import WordDocument
except ImportError:
	from NVDAObjects.IAccessible.winword import WordDocument


class IAccessibleWordDocument(ww_wordDocumentBase.WordDocument, WordDocument):
	treeInterceptorClass = ww_browseMode.WordDocumentTreeInterceptorEx
	shouldCreateTreeInterceptor = True
	disableAutoPassThrough = True
	TextInfo = WindowWinwordWordDocumentTextInfo
