# appModules\winword\ww_browsemode.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2024 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import NVDAObjects
from . import ww_browseMode
from UIAHandler.browseMode import UIATextAttributeQuicknavIterator
from NVDAObjects.UIA.wordDocument import CommentUIATextInfoQuickNavItem, RevisionUIATextInfoQuickNavItem
from .import ww_elementsListDialog

addonHandler.initTranslation()


class UIAWordBrowseModeDocument(
	ww_browseMode.BrowseModeDocumentTreeInterceptorEx,
	NVDAObjects.UIA.wordDocument.WordBrowseModeDocument):
	ElementsListDialog = ww_elementsListDialog.UIAElementsListDialog

	def __init__(self, rootNVDAObject):
		super(UIAWordBrowseModeDocument, self).__init__(rootNVDAObject)
		self.passThrough = True

	def _iterNodesByType(self, nodeType, direction="next", pos=None):
		if nodeType == "comment":
			return UIATextAttributeQuicknavIterator(
				CommentUIATextInfoQuickNavItem, nodeType, self, pos, direction=direction)
		elif nodeType == "revision":
			return UIATextAttributeQuicknavIterator(
				RevisionUIATextInfoQuickNavItem, nodeType, self, pos, direction=direction)
		return super(UIAWordBrowseModeDocument, self)._iterNodesByType(
			nodeType, direction=direction, pos=pos)
