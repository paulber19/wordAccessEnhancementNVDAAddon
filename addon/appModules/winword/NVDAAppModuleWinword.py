# appModules\winword\NVDAAppModuleWinword.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2024 paulber19
# This file is covered by the GNU General Public License.
#
# this a part of builtin nvda winword appModule.
# to bring to nvda versions prior to version 2025.1, its new features


import addonHandler
from scriptHandler import script
import ui
from NVDAObjects.window.winword import WordDocument
from utils.displayString import DisplayStringIntEnum

addonHandler.initTranslation()


class ViewType(DisplayStringIntEnum):
	"""Enumeration containing the possible view types in Word documents:.
	https://learn.microsoft.com/en-us/office/vba/api/word.wdviewtype
	"""

	DRAFT = 1
	OUTLINE = 2
	PRINT = 3
	WEB = 6
	READ = 7

	@property
	def _displayStringLabels(self):
		return {
			# Translators: One of the view types in Word documents.
			ViewType.DRAFT: _("DRAFT"),
			# Translators: One of the view types in Word documents.
			ViewType.OUTLINE: _("Outline"),
			# Translators: One of the view types in Word documents.
			ViewType.PRINT: _("Print layout"),
			# Translators: One of the view types in Word documents.
			ViewType.WEB: _("Web layout"),
			# Translators: One of the view types in Word documents.
			ViewType.READ: _("Read mode"),
		}


class WinwordWordDocument(WordDocument):
	def _get_description(self) -> str:
		description = super()._get_description()
		try:
			curView = self.WinwordWindowObject.view.Type
			if description:
				return f"{ViewType(curView).displayString} {description}"
			else:
				return f"{ViewType(curView).displayString}"
		except AttributeError:
			return super()._get_description()

	@script(gesture="kb:control+shift+e")
	def script_toggleChangeTracking(self, gesture):
		if not self.WinwordDocumentObject:
			# We cannot fetch the Word object model, so we therefore cannot report the status change.
			# The object model may be unavailable because it's within Windows Defender Application Guard.
			# In this case, just let the gesture through and don't report anything.
			return gesture.send()
		val = self._WaitForValueChangeForAction(
			lambda: gesture.send(),
			lambda: self.WinwordDocumentObject.TrackRevisions,
		)
		if val:
			# Translators: a message when toggling change tracking in Microsoft word
			ui.message(_("Change tracking on"))
		else:
			# Translators: a message when toggling change tracking in Microsoft word
			ui.message(_("Change tracking off"))

	@script(gesture="kb:control+shift+k")
	def script_toggleCaps(self, gesture):
		if not self.WinwordSelectionObject:
			# We cannot fetch the Word object model, so we therefore cannot report the format change.
			# The object model may be unavailable because this is a pure UIA implementation such as Windows 10 Mail,
			# or its within Windows Defender Application Guard.
			# Eventually UIA will have its own way of detecting format changes at the cursor.
			# For now, just let the gesture through and don't report anything.
			return gesture.send()
		val = self._WaitForValueChangeForAction(
			lambda: gesture.send(),
			lambda: (self.WinwordSelectionObject.font.allcaps, self.WinwordSelectionObject.font.smallcaps),
		)
		if val[0]:
			# Translators: a message when toggling formatting to 'all capital' in Microsoft word
			ui.message(_("All caps on"))
		elif val[1]:
			# Translators: a message when toggling formatting to 'small capital' in Microsoft word
			ui.message(_("Small caps on"))
		else:
			# Translators: a message when toggling formatting to 'No capital' in Microsoft word
			ui.message(_("Caps off"))

	def script_viewChange(self, gesture):
		if not self.WinwordDocumentObject:
			return gesture.send()
		val = self._WaitForValueChangeForAction(
			lambda: gesture.send(),
			lambda: self.WinwordDocumentObject.ActiveWindow.View.Type,
		)
		if val:
			view = ViewType(val).displayString
			ui.message(view)
		else:
			pass

	__gestures = {
		"kb:control+shift+b": "toggleBold",
		"kb:control+shift+w": "toggleUnderline",
		"kb:control+shift+a": "toggleCaps",
		"kb:control+alt+p": "viewChange",
		"kb:control+alt+o": "viewChange",
	}
