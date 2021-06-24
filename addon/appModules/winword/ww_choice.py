# appModules\winword\ww.choice.py
# A part of WordAccessEnhancement add-on
# Copyright (C) 2019-2021 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import wx
import speech
from gui import mainFrame, guiHelper
from .ww_spellingErrors import SpellingErrors
from .ww_grammaticalErrors import GrammaticalErrors
from .ww_comments import Comments
from .ww_revisions import Revisions
from .ww_bookmarks import Bookmarks
from .ww_footnotes import Footnotes
from .ww_endnotes import Endnotes
from .ww_shapes import Shapes, InLineShapes
from .ww_hyperlinks import Hyperlinks
from .ww_fields import Fields
from .ww_formfields import FormFields
from .ww_contentControl import ContentControls
from .ww_tables import Tables
from .ww_headings import Headings
from .ww_frames import Frames
from .ww_sections import Sections
import sys
import os
_curAddon = addonHandler.getCodeAddon()
debugToolsPath = os.path.join(_curAddon.path, "debugTools")
sys.path.append(debugToolsPath)
try:
	from appModuleDebug import printDebug, toggleDebugFlag
except ImportError:
	def printDebug(msg): return
	def toggleDebugFlag(): return
del sys.path[-1]

sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_utils import putWindowOnForeground, makeAddonWindowTitle  # noqa:E402
del sys.path[-1]

addonHandler.initTranslation()

_collectionClassDic = {
	"bookmark": Bookmarks,
	"comment": Comments,
	"contentControl": ContentControls,
	"endnote": Endnotes,
	"field": Fields,
	"formfield": FormFields,
	"footnote": Footnotes,
	"frame": Frames,
	"grammaticalError": GrammaticalErrors,
	"heading": Headings,
	"hyperlink": Hyperlinks,
	"drawingLayerObject": Shapes,
	"textLayerObject": InLineShapes,
	"revision": Revisions,
	"section": Sections,
	"spellingError": SpellingErrors,
	"table": Tables
	}
_wordObjectList = (
			(SpellingErrors, "document"),
			(GrammaticalErrors, "document"),
			(Comments, "document"),
			(Revisions, "document"),
			(Bookmarks, "document"),
			(Footnotes, "document"),
			(Endnotes, "document"),
			(Shapes, "document"),
			(InLineShapes, "document"),
			(Hyperlinks, "document"),
			(Fields, "document"),
			(FormFields, "document"),
			(ContentControls, "document"),
			(Tables, "document"),
			(Headings, "document"),
			(Frames, "document"),
			(Sections, "document")
			)


class PopulateAndReportCollection(wx.Dialog):
	def __init__(
			self, parent, collectionClass, rangeType, focus, copyToClipboard=False):
		super(PopulateAndReportCollection, self).__init__(
			parent=parent,
			# Translators: title of dialog.
			title=_("Informations 's collection"))
		self.parent = parent
		self.collectionClass = collectionClass
		self.rangeType = rangeType
		self.copyToClipboard = copyToClipboard
		# Translators: message to user for waiting result
		self.waitMessage = _("Please wait")
		self.doGui()
		self.canceled = False
		self.interrupted = False
		self.isActive = True  # forced, because don't True when created
		wx.CallLater(1000, self.populateAndReportCollection, focus)
		self.CentreOnScreen()
		self.Show()

	def populateAndReportCollection(self, focus):
		self.collection = self.collectionClass(self, focus, self.rangeType)
		if self.canceled:
			self.Close()
			return
		wx.CallAfter(
			self.collection.reportCollection, self.copyToClipboard, self.parent)
		self.Close()

	def doGui(self):
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		sHelper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		# the text control
		sHelper.addItem(wx.StaticText(self, label=self.waitMessage))
		# the buttons
		bHelper = sHelper.addDialogDismissButtons(
			guiHelper.ButtonHelper(wx.HORIZONTAL))
		interruptButton = bHelper.addButton(
			self, id=wx.ID_ANY,
			# Translators: label of interrupt to clipboard button.
			label=_("&interrupt"))
		cancelButton = bHelper.addButton(
			self, id=wx.ID_CANCEL, label=_("&Cancel"))
		mainSizer.Add(
			sHelper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		mainSizer.Fit(self)
		self.SetSizer(mainSizer)
		cancelButton.SetFocus()
		# the events
		cancelButton.Bind(wx.EVT_BUTTON, self.onCancelButton)
		interruptButton.Bind(wx.EVT_BUTTON, self.onInterruptButton)
		self.Bind(wx.EVT_CLOSE, self.onClose)
		self.SetEscapeId(wx.ID_CANCEL)
		self.Bind(wx.EVT_ACTIVATE, self.onActivate)

	def onActivate(self, evt):
		printDebug("collectInformation active: %s" % evt.GetActive())
		if evt.GetActive():
			self.isActive = True
		else:
			self.isActive = False
		evt.Skip()

	def onClose(self, evt):
		self.Destroy()
		evt.Skip()

	def onInterruptButton(self, evt):
		self.interrupted = True
		speech.speakMessage(_("Interrupted"))
		speech.speakMessage(self.waitMessage)

	def onCancelButton(self, evt):
		self.canceled = True
		self.Destroy()


class ChoiceDialog(wx.Dialog):
	_instance = None
	title = None

	def __init__(self, parent=None, focus=None):
		ChoiceDialog._instance = self
		# Translators: title of Choice dialog
		dialogTitle = _("Elements's search")
		title = ChoiceDialog.title = makeAddonWindowTitle(dialogTitle)
		super(ChoiceDialog, self).__init__(
			None, -1, title=title, style=wx.CAPTION | wx.CLOSE_BOX | wx.TAB_TRAVERSAL)
		self.focus = focus
		self.wordApp = focus._get_WinwordApplicationObject()
		selection = self.wordApp.Selection
		# keep if there is a selection
		self.isASelection = False
		if selection.Start != selection.End:
			self.isASelection = True
		self.isAlive = True
		self.initializeGUI()
		self.doGUI()
		putWindowOnForeground(self.GetHandle())

	def __del__(self):
		ChoiceDialog._instance = None
		if getattr(self, 'dialog', None):
			self.dialog.Destroy()

	def Destroy(self):
		ChoiceDialog._instance = None
		super(ChoiceDialog, self).Destroy()

	def initializeGUI(self):
		# Translators: label of the element's type listbox.
		self.lcLabel = _("Element's type:")
		self.lcSize = (300, 250)
		self.lcID = 300
		self.wordObjectList = _wordObjectList[:]
		self.choice = []
		for (cls, r) in self.wordObjectList:
			name = cls._name[0]
			self.choice.append(name)
		self.choice.sort()
		self.rangeTypes = (
			"focus", "selection",
			"document", "page",
			"startToFocus", "focusToEnd")
		self.lc2Label = _("Part of document:")
		self.lc2Size = (300, 250)
		self.lc2ID = 301
		self.rangeTitles = {
			"focus": _("Cursor's positionFocus"),
			"document": _("Document"),
			"page": _("Page"),
			"selection": _("Selection"),
			"focusToEnd": _("Focus to end"),
			"startToFocus": _("Start to focus")
		}
		self.parts = []
		for item in self.rangeTypes:
			if item == "selection" and not self.isASelection:
				continue

			self.parts.append(self.rangeTitles[item])

	def doGUI(self):
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		settingsSizer = wx.BoxSizer(wx.HORIZONTAL)
		entriesSizer = wx.BoxSizer(wx.VERTICAL)
		entriesLabel = wx.StaticText(self, -1, label=self.lcLabel)
		entriesSizer.Add(entriesLabel)
		self.entriesList = wx.ListBox(
			self,
			self.lcID,
			name="Select",
			choices=self.choice,
			style=wx.ALIGN_CENTRE | wx.WANTS_CHARS,
			size=self.lcSize)
		self.entriesList.SetSelection(0)
		self.entriesList.Bind(wx.EVT_LISTBOX, self.onListItemFocused)
		self.entriesList.Bind(wx.EVT_KEY_DOWN, self.onListKeyDown)
		# for index in range(0,self.entriesList.GetCount()):
	# self.entriesList.SetItemBackgroundColour(index, "white")
		entriesSizer.Add(self.entriesList, proportion=3)
		settingsSizer.Add(entriesSizer)
		# part of document
		lc2Sizer = wx.BoxSizer(wx.VERTICAL)
		lc2Label = wx.StaticText(self, -1, label=self.lc2Label)
		lc2Sizer.Add(lc2Label)
		self.lc2List = wx.ListBox(
			self,
			self.lc2ID,
			name="Part",
			choices=self.parts,
			style=wx.ALIGN_CENTRE | wx.WANTS_CHARS,
			size=self.lc2Size)
		self.lc2List.SetSelection(0)
		self.lc2List.Bind(wx.EVT_KEY_DOWN, self.onListKeyDown)
		# for index in range(0,self.lc2List.GetCount()):
		# self.lc2List.SetItemBackgroundColour(index, "white")
		lc2Sizer.Add(self.lc2List, proportion=3)
		settingsSizer.Add(lc2Sizer)

		# buttons
		buttonsSizer = wx.BoxSizer(wx.VERTICAL)
		# Translators: label of a button
		searchButton = wx.Button(self, wx.ID_ANY, label=_("Search an&d display"))
		searchButton.Bind(wx.EVT_BUTTON, self.onSearchButton)
		searchButton.Bind(wx.EVT_KEY_DOWN, self.onButtonKeyDown)
		buttonsSizer.Add(
			searchButton,
			border=20,
			flag=wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.CENTER)
		copyToClipButton = wx.Button(
			self, wx.ID_ANY, label=_("Search and C&opy to clipboard"))
		copyToClipButton.Bind(wx.EVT_BUTTON, self.onCopyToClipButton)
		copyToClipButton.Bind(wx.EVT_KEY_DOWN, self.onButtonKeyDown)
		buttonsSizer.Add(
			copyToClipButton,
			border=20,
			flag=wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.CENTER)
		closeButton = wx.Button(self, wx.ID_CLOSE, label=_("&Close"))
		closeButton.Bind(wx.EVT_BUTTON, self.onCloseButton)
		closeButton.Bind(wx.EVT_KEY_DOWN, self.onButtonKeyDown)
		buttonsSizer.Add(
			closeButton,
			border=20,
			flag=wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.CENTER)
		settingsSizer.Add(buttonsSizer)
		mainSizer.Add(
			settingsSizer,
			border=20,
			flag=wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.TOP)

		self.Bind(wx.EVT_CLOSE, self.onClose)
		self.Bind(wx.EVT_ACTIVATE, self.onActivate)
		mainSizer.Fit(self)
		self.SetSizer(mainSizer)
		self.entriesList.SetFocus()

	def getCollectionClassAndRangeType(self):
		index = self.entriesList.GetSelection()
		for (cls, r) in self.wordObjectList:
			name = cls._name[0]
			if name == self.choice[index]:
				collectionClass = cls
				break

		indexPart = self.lc2List.GetSelection()
		titlePart = self.lc2List.GetString(indexPart)
		for key in self.rangeTitles:
			if self.rangeTitles[key] == titlePart:
				rangeType = key

		return (collectionClass, rangeType)

	def onSearchButton(self, evt):
		(collectionClass, rangeType) = self.getCollectionClassAndRangeType()
		PopulateAndReportCollection(
			self, collectionClass, rangeType, self.focus, False)

	def onCopyToClipButton(self, evt):
		(collectionClass, rangeType) = self.getCollectionClassAndRangeType()
		PopulateAndReportCollection(self, collectionClass, rangeType, self.focus, True)

	def onCloseButton(self, evt):
		self.Close()
		evt.Skip()

	def onClose(self, evt):
		self.alive = False
		self.Destroy()

	def onButtonKeyDown(self, evt):
		key = evt.GetKeyCode()

		if (key == wx.WXK_ESCAPE):
			self.Close()
			return
		evt.Skip()

	def onListKeyDown(self, evt):
		key = evt.GetKeyCode()
		if (key == wx.WXK_ESCAPE):
			self.Close()
			return

		if key == wx.WXK_RETURN:
			self.onSearchButton(None)
			return

		if key == wx.WXK_TAB:
			id = evt.GetId()
			lc = self.entriesList
			if id == 301:
				lc = self.lc2List

			shiftDown = evt.ShiftDown()
			if shiftDown:
				wx.Window.Navigate(lc, wx.NavigationKeyEvent.IsBackward)
			else:
				wx.Window.Navigate(lc, wx.NavigationKeyEvent.IsForward)
			return
		evt.Skip()

	def onActivate(self, evt):
		printDebug("makeChoice active: %s" % evt.GetActive())
		evt.Skip()

	def onListItemFocused(self, evt):
		if self.isASelection:
			index = self.rangeTypes.index("selection")
		else:
			typeIndex = self.entriesList.GetSelection()
			for (cls, rangeType) in self.wordObjectList:
				name = cls._name[0]
				if name == self.choice[typeIndex]:

					break
			title = self.rangeTitles[rangeType]
			index = self.parts.index(title)
		self.lc2List.SetSelection(index)

	@staticmethod
	def run(focus):
		if ChoiceDialog._instance is not None:
			putWindowOnForeground(ChoiceDialog._instance.GetHandle())
			return
		d = ChoiceDialog(mainFrame, focus)
		d.CentreOnScreen()
		d.ShowModal()
