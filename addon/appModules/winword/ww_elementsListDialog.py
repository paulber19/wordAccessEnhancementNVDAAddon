# appModules\winword\ww_elementsListDialog.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2021- 2025 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
from logHandler import log
import wx
import collections
import speech
import gui
import time
import core
import queueHandler
import itertools
from gui.dpiScalingHelper import DpiScalingHelperMixinWithoutInit
import sys
import os
import api
_curAddon = addonHandler.getCodeAddon()
sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_NVDAStrings import NVDAString
from ww_addonConfigManager import _addonConfigManager
from ww_utils import (
	getSpeechMode, setSpeechMode, setSpeechMode_off)
del sys.path[-1]
del sys.modules["ww_NVDAStrings"]
del sys.modules["ww_utils"]

addonHandler.initTranslation()


class ElementsListDialog(DpiScalingHelperMixinWithoutInit, wx.Dialog):
	ELEMENT_TYPES = (
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("heading", NVDAString("heading").capitalize()),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("link", NVDAString("link").capitalize()),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("graphic", NVDAString("graphic").capitalize()),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("bookmark", _("Bookmark")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("comment", _("Comment")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("endnote", _("Endnote")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("field", _("Field")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("formfield", _("FormField")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("footnote", _("Footnote")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("revision", _("Revision")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("section", _("Section")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("spellingError", _("Spelling error")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("grammaticalError", _("Grammatical error")),
	)
	Element = collections.namedtuple("Element", ("item", "parent"))
	lastSelectedElementType = 0

	def __init__(self, document):
		self.dialogTitle = NVDAString("Elements List")
		super(ElementsListDialog, self).__init__(
			parent=gui.mainFrame,
			title=self.dialogTitle)
		self.document = document
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		contentsSizer = wx.BoxSizer(wx.VERTICAL)
		childSizer = wx.BoxSizer(wx.VERTICAL)
		# Translators: The label of a list of items to select the type of element
		# in the browse mode Elements List dialog.
		childLabel = wx.StaticText(
			parent=self,
			label=_("&Type:"),
			style=wx.ALIGN_CENTRE
		)
		childSizer.Add(childLabel, )
		self.childListBox = wx.ListBox(
			self,
			wx.ID_ANY,
			name="TypeName",
			choices=tuple(et[1] for et in self.ELEMENT_TYPES),
			style=wx.LB_SINGLE,
			size=(596, 130))
		if self.childListBox.GetCount():
			self.childListBox.SetSelection(self.lastSelectedElementType)
		self.childListBox.Bind(wx.EVT_LISTBOX, self.onElementTypeChange)
		self.childListBox.Bind(wx.EVT_SET_FOCUS, self.onChildBoxFocus)
		childSizer.Add(self.childListBox)
		contentsSizer.Add(childSizer, flag=wx.EXPAND)
		contentsSizer.AddSpacer(gui.guiHelper.SPACE_BETWEEN_VERTICAL_DIALOG_ITEMS)

		self.tree = wx.TreeCtrl(
			self,
			size=self.scaleSize((500, 300)),  # height is chosen to ensure the dialog will fit on an 800x600 screen
			style=wx.TR_HAS_BUTTONS | wx.TR_HIDE_ROOT | wx.TR_LINES_AT_ROOT | wx.TR_SINGLE | wx.TR_EDIT_LABELS
		)
		self.tree.Bind(wx.EVT_SET_FOCUS, self.onTreeSetFocus)
		self.tree.Bind(wx.EVT_CHAR, self.onTreeChar)
		self.treeRoot = self.tree.AddRoot("root")
		contentsSizer.Add(self.tree, flag=wx.EXPAND)
		contentsSizer.AddSpacer(gui.guiHelper.SPACE_BETWEEN_VERTICAL_DIALOG_ITEMS)

		# Translators: The label of an editable text field to filter the elements
		# in the browse mode Elements List dialog.
		filterText = NVDAString("Filter b&y:")
		labeledCtrl = gui.guiHelper.LabeledControlHelper(self, filterText, wx.TextCtrl)
		self.filterEdit = labeledCtrl.control
		self.filterEdit.Bind(wx.EVT_TEXT, self.onFilterEditTextChange)
		contentsSizer.Add(labeledCtrl.sizer)
		contentsSizer.AddSpacer(gui.guiHelper.SPACE_BETWEEN_VERTICAL_DIALOG_ITEMS)

		bHelper = gui.guiHelper.ButtonHelper(wx.HORIZONTAL)
		# Translators: The label of a button to activate an element
		# in the browse mode Elements List dialog.
		self.activateButton = bHelper.addButton(self, label=_("&Activate"))
		self.activateButton.Bind(wx.EVT_BUTTON, lambda evt: self.onAction(True))

		# Translators: The label of a button to move to an element
		# in the browse mode Elements List dialog.
		self.moveButton = bHelper.addButton(self, label=NVDAString("&Move to"))
		self.moveButton.Bind(wx.EVT_BUTTON, lambda evt: self.onAction(False))
		bHelper.addButton(self, id=wx.ID_CANCEL)

		contentsSizer.Add(bHelper.sizer, flag=wx.ALIGN_RIGHT)
		mainSizer.Add(contentsSizer, border=gui.guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		mainSizer.Fit(self)
		self.SetSizer(mainSizer)

		self.tree.SetFocus()
		self.initElementTypeIsRunning = False
		self.initElementType(self.ELEMENT_TYPES[self.lastSelectedElementType][0])
		self.CentreOnScreen()
		self.refreshDialogTitle()

	def refreshDialogTitle(self):
		if self.childListBox.GetCount():
			type = self.childListBox.GetStringSelection()
			title = "%s - %s" % (self.dialogTitle, type)
			self.SetTitle(title)

	def onElementTypeChange(self, evt):
		if self.initElementTypeIsRunning:
			self.abort = True
			wx.CallLater(100, self.onElementTypeChange, evt)
			return
		elementType = self.childListBox.GetSelection()
		# We need to make sure this gets executed after the focus event.
		# Otherwise, NVDA doesn't seem to get the event.
		# modified because of error when type spelling error type
		wx.CallLater(50, self.initElementType, self.ELEMENT_TYPES[elementType][0])
		self.lastSelectedElementType = elementType
		self.refreshDialogTitle()

	def initElementType(self, elType):
		from .ww_tones import RepeatBeep
		rb = RepeatBeep()
		rb.start()
		self.initElementTypeIsRunning = True
		self.startTime = int(time.time())
		if elType in ("link", "button", "radioButton", "checkBox"):
			# Links, buttons, radio button, check box can be activated.
			self.activateButton.Enable()
			self.SetAffirmativeId(self.activateButton.GetId())
		else:
			# No other element type can be activated.
			self.activateButton.Disable()
			self.SetAffirmativeId(self.moveButton.GetId())

		# Gather the elements of this type.
		self.abort = False
		self._elements = []
		self._initialElement = None
		parentElements = []
		isAfterSelection = False
		try:
			limited = False
			for item in self.document._iterNodesByType(elType):
				wx.Yield()
				if self.abort:
					rb.stop()
					self.initElementTypeIsRunning = False
					return
				# Find the parent element, if any.
				for parent in reversed(parentElements):
					if item.isChild(parent.item):
						break
					else:
						# We're not a child of this parent, so this parent has no more children
						# and can be removed from the stack.
						parentElements.pop()
				else:
					# No parent found, so we're at the root.
					# Note that parentElements will be empty at this point, as all parents are no longer relevant
					# and have thus been removed from the stack.
					parent = None

				element = self.Element(item, parent)
				self._elements.append(element)
				elapsedTime = int(time.time()) - self.startTime
				if elapsedTime > _addonConfigManager.getElementsSearchMaxTime():
					limited = True
					break

				if not isAfterSelection:
					isAfterSelection = item.isAfterSelection
					if not isAfterSelection:
						# The element immediately preceding or overlapping the caret should be the initially selected element.
						# Since we have not yet passed the selection, use this as the initial element.
						try:
							self._initialElement = self._elements[-1]
						except IndexError:
							# No previous element.
							pass

				# This could be the parent of a subsequent element, so add it to the parents stack.
				parentElements.append(element)

		except Exception:
			log.error("initElementType: Can not find all elements")
		rb.stop()
		# Start with no filtering.
		self.filterEdit.ChangeValue("")
		self.filter("", newElementType=True)
		self.sayNumberOfElements(limited)
		self.initElementTypeIsRunning = False

	def sayNumberOfElements(self, limited=False):
		count = self.tree.Count
		if not self.childListBox.HasFocus():
			return
		if count:
			if limited:
				# Translators: message to indicate the number of elements is limited.
				msg = _("more than %s elements") % str(count) if count > 1 else _("One element")
			else:
				# Translators: message to the user to report number of elements.
				msg = _("%s elements") % str(count) if count > 1 else _("One element")
		else:
			# Translators: message to the user when there is no element.
			msg = _("no element")
		wx.CallLater(100, queueHandler.queueFunction, queueHandler.eventQueue, speech.speakMessage, msg)

	def onChildBoxFocus(self, evt):
		if not hasattr(self, "sayNumberOfElementsEnabled"):
			#  don't speak number of elements of first focus
			self.sayNumberOfElementsEnabled = True
			return
		self.sayNumberOfElements()

	def filter(self, filterText, newElementType=False):
		# If this is a new element type, use the element nearest the cursor.
		# Otherwise, use the currently selected element.
		try:
			defaultElement = self._initialElement if newElementType else self.tree.GetItemData(
				self.tree.GetSelection())
		except Exception:
			defaultElement = self._initialElement
		# Clear the tree.
		self.tree.DeleteChildren(self.treeRoot)

		# Populate the tree with elements matching the filter text.
		elementsToTreeItems = {}
		defaultItem = None
		matched = False
		# Do case-insensitive matching by lowering both filterText
		# and each element's text.
		filterText = filterText.lower()
		for element in self._elements:
			label = element.item.label
			if filterText and filterText not in label.lower():
				continue
			matched = True
			parent = element.parent
			if parent:
				parent = elementsToTreeItems.get(parent)
			item = self.tree.AppendItem(parent or self.treeRoot, label)
			self.tree.SetItemData(item, element)
			elementsToTreeItems[element] = item
			if element == defaultElement:
				defaultItem = item

		self.tree.ExpandAll()

		if not matched:
			# No items, so disable the buttons.
			self.activateButton.Disable()
			self.moveButton.Disable()
			return

		# If there's no default item, use the first item in the tree.
		self.tree.SelectItem(defaultItem or self.tree.GetFirstChild(self.treeRoot)[0])
		# Enable the button(s).
		# If the activate button isn't the default button, it is disabled for this element type
		# and shouldn't be enabled here.
		if self.AffirmativeId == self.activateButton.Id:
			self.activateButton.Enable()
		self.moveButton.Enable()

	def onTreeSetFocus(self, evt):
		# Start with no search.
		self._searchText = ""
		self._searchCallLater = None
		evt.Skip()

	def onTreeChar(self, evt):
		key = evt.KeyCode

		if key == wx.WXK_RETURN:
			# The enter key should be propagated to the dialog and thus activate the default button,
			# but this is broken (wx ticket #3725).
			# Therefore, we must catch the enter key here.
			# Activate the current default button.
			evt = wx.CommandEvent(wx.wxEVT_COMMAND_BUTTON_CLICKED, wx.ID_ANY)
			button = self.FindWindowById(self.AffirmativeId)
			if button.Enabled:
				button.ProcessEvent(evt)
			else:
				wx.Bell()

		elif key == wx.WXK_F2:
			item = self.tree.GetSelection()
			if item:
				self.tree.EditLabel(item)
				evt.Skip()

		elif key >= wx.WXK_START or key == wx.WXK_BACK:
			# Non-printable character.
			self._searchText = ""
			evt.Skip()

		else:
			# Search the list.
			# We have to implement this ourselves, as tree views don't accept space as a search character.
			char = chr(evt.UnicodeKey).lower()
			# IF the same character is typed twice, do the same search.
			if self._searchText != char:
				self._searchText += char
			if self._searchCallLater:
				self._searchCallLater.Restart()
			else:
				self._searchCallLater = core.callLater(1000, self._clearSearchText)
			self.search(self._searchText)

	def onTreeLabelEditBegin(self, evt):
		item = self.tree.GetSelection()
		selectedItemType = self.tree.GetItemPyData(item).item
		if not selectedItemType.isRenameAllowed:
			evt.Veto()

	def onTreeLabelEditEnd(self, evt):
		selectedItemNewName = evt.GetLabel()
		item = self.tree.GetSelection()
		selectedItemType = self.tree.GetItemData(item).item
		selectedItemType.rename(selectedItemNewName)

	def _clearSearchText(self):
		self._searchText = ""

	def search(self, searchText):
		item = self.tree.GetSelection()
		if not item:
			# No items.
			return

		# First try searching from the current item.
		# Failing that, search from the first item.
		items = itertools.chain(self._iterReachableTreeItemsFromItem(
			item), self._iterReachableTreeItemsFromItem(self.tree.GetFirstChild(self.treeRoot)[0]))
		if len(searchText) == 1:
			# If only a single character has been entered, skip (search after) the current item.
			next(items)

		for item in items:
			if self.tree.GetItemText(item).lower().startswith(searchText):
				self.tree.SelectItem(item)
				return

		# Not found.
		wx.Bell()

	def _iterReachableTreeItemsFromItem(self, item):
		while item:
			yield item

			childItem = self.tree.GetFirstChild(item)[0]
			if childItem and self.tree.IsExpanded(item):
				# Has children and is reachable, so recurse.
				for childItem in self._iterReachableTreeItemsFromItem(childItem):
					yield childItem

			item = self.tree.GetNextSibling(item)

	def onFilterEditTextChange(self, evt):
		self.filter(self.filterEdit.GetValue())
		evt.Skip()

	def onAction(self, activate):
		self.__class__.lastSelectedElementType = self.lastSelectedElementType
		item = self.tree.GetSelection()
		item = self.tree.GetItemData(item).item
		if activate:
			self.Close()
			item.activate()
		else:
			def move():
				oldSpeechMode = getSpeechMode()
				setSpeechMode_off()
				self.Close()
				api.processPendingEvents()
				item.moveTo()

				def report():
					speech.cancelSpeech()
					setSpeechMode(oldSpeechMode)
					item.report()
				core.callLater(1, report)

			# We must use core.callLater rather than wx.CallLater to ensure that the callback runs
			# within NVDA's core pump.
			# If it didn't, and it directly or indirectly called wx.Yield, it could start executing NVDA's core
			# pump from within the yield, causing recursion.
			core.callLater(20, move)


class UIAElementsListDialog(ElementsListDialog):
	ELEMENT_TYPES = (
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("heading", _("Heading")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("link", _("Link")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("graphic", _("Graphic")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		# ("bookmark", _("Bookmark")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("comment", _("Comment")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("endnote", _("Endnote")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		# ("field", _("Field")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		# ("formfield", _("FormField")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("footnote", _("Footnote")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("revision", _("Revision")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		# ("section", _("Section")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		# ("spellingError", _("Spelling error")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		# ("grammaticalError", _("Grammatical error")),
		# Translators: The label of a list item to select the type of element
		# in the browse mode Elements List dialog.
		("error", _("Error")),
	)
	lastSelectedElementType = 0
