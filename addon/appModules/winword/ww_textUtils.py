# appModule\\winword\ww_textUtils.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2024 paulber19
# This file is covered by the GNU General Public License.
import addonHandler
from gui import mainFrame
import wx

addonHandler.initTranslation()


def askForText(dialogTitle, entryBoxLabel, defaultText=None, multiLine=True):
	with wx.TextEntryDialog(
		mainFrame,
		entryBoxLabel,
		dialogTitle,
		defaultText if defaultText else "",
		style=wx.TextEntryDialogStyle | wx.TE_MULTILINE if multiLine else wx.TextEntryDialogStyle
	) as entryDialog:
		if entryDialog.ShowModal() != wx.ID_OK:
			return None
	text = entryDialog.Value
	if len(text) == 0:
		return None
	return text


def askForAuthor(defaultAuthor=None):
	return askForText(
		# Translators: text control dialog title to confirm the author.
		_("Modification of author name"),
		# Translators: label for an entry text box.
		_("Enter Author name:"),
		defaultAuthor,
		False
	)
