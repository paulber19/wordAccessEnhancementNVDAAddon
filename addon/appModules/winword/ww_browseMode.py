# appModules\winword\ww_browsemode.py
#A part of wordAccessEnhancement add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
import api
import wx
import ui
import textInfos
import sayAllHandler
import speech
import time
import gui
import core
import queueHandler
from NVDAObjects.window.winword import *
from .ww_fields import Field


import sys
import os
_curAddon = addonHandler.getCodeAddon()
path = os.path.join(_curAddon.path, "shared")
sys.path.append(path)
from ww_utils import printDebug
from ww_py3Compatibility import _unicode, uniCHR
from ww_NVDAStrings  import NVDAString
del sys.path[-1]

class BrowseModeTreeInterceptorEx(browseMode.BrowseModeTreeInterceptor):
	__gestures = {}
	scriptCategory = _("Extended browse mode for Microsoft Word")
	
	@classmethod
	def addQuickNav(cls, itemType, key, nextDoc, nextError, prevDoc, prevError, readUnit=None):
		map=  cls.__gestures

		"""Adds a script for the given quick nav item.
		@param itemType: The type of item, I.E. "heading" "Link" ...
		@param key: The quick navigation key to bind to the script. Shift is automatically added for the previous item gesture. E.G. h for heading
		@param nextDoc: The command description to bind to the script that yields the next quick nav item.
		@param nextError: The error message if there are no more quick nav items of type itemType in this direction.
		@param prevDoc: The command description to bind to the script that yields the previous quick nav item.
		@param prevError: The error message if there are no more quick nav items of type itemType in this direction.
		@param readUnit: The unit (one of the textInfos.UNIT_* constants) to announce when moving to this type of item. 
			For example, only the line is read when moving to tables to avoid reading a potentially massive table. 
			If None, the entire item will be announced.
		"""
		scriptSuffix = itemType[0].upper() + itemType[1:]
		scriptName = "next%s" % scriptSuffix
		funcName = "script_%s" % scriptName
		script = lambda self,gesture: self._quickNavScript(gesture, itemType, "next", nextError, readUnit)
		script.__doc__ = nextDoc
		script.__name__ = funcName
		script.resumeSayAllMode=sayAllHandler.CURSOR_CARET
		setattr(cls, funcName, script)
		map["kb:%s" % key] = scriptName
		scriptName = "previous%s" % scriptSuffix
		funcName = "script_%s" % scriptName
		script = lambda self,gesture: self._quickNavScript(gesture, itemType, "previous", prevError, readUnit)
		script.__doc__ = prevDoc
		script.__name__ = funcName
		script.resumeSayAllMode=sayAllHandler.CURSOR_CARET
		setattr(cls, funcName, script)
		map["kb:shift+%s" % key] = scriptName




# Add quick navigation scripts.
BrowseModeTreeInterceptorEx._BrowseModeTreeInterceptorEx__gestures= browseMode.BrowseModeTreeInterceptor._BrowseModeTreeInterceptor__gestures.copy()
qn = BrowseModeTreeInterceptorEx.addQuickNav
qn("grammaticalError", key="$",
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next grammatical error"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next grammatical error"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous grammatical error"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous grammatical error"))
		
	
qn("revision", key="<",
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next revision"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next revision"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous revision"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous revision"))

qn("comment", key="j",
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next comment"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next comment"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous comment"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous comment"))

qn("field", key="y",
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next field"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next field"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous field"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous field"),
	readUnit=textInfos.UNIT_LINE)

qn("bookmark", key=";",
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next bookmark"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next bookmark"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous bookmark"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous bookmark"))

qn("endnote", key="!",
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next endnote"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next endnote"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous endnote"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous endnote"))
qn("footnote", key=":",
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next footnote"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next footnote"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous footnote"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous footnote"))

qn("section", key=")",
	# Translators: Input help message for a quick navigation command in browse mode.
	nextDoc=_("moves to the next section"),
	# Translators: Message presented when the browse mode element is not found.
	nextError=_("no next section"),
	# Translators: Input help message for a quick navigation command in browse mode.
	prevDoc=_("moves to the previous section"),
	# Translators: Message presented when the browse mode element is not found.
	prevError=_("no previous section"))

del qn

class BrowseModeDocumentTreeInterceptorEx(BrowseModeTreeInterceptorEx, browseMode.BrowseModeDocumentTreeInterceptor):
	pass



class WordDocumentTreeInterceptorEx(BrowseModeDocumentTreeInterceptorEx, WordDocumentTreeInterceptor):
	TextInfo=BrowseModeWordDocumentTextInfo
	def __init__(self,obj):
		super(WordDocumentTreeInterceptorEx,self).__init__(obj)
		# to keep WordDocumentEx scripts available in browseMode on
		gestures = ["kb:alt+upArrow", "kb: alt+downArrow","kb:control+upArrow", "kb:control+downArrow"]
		for gest in gestures:
			try:
				self.removeGestureBinding(gest)
			except:
				pass
	
	def _get_ElementsListDialog(self):
		return ElementsListDialog
	
	def _iterNodesByType(self,nodeType,direction="next",pos=None):
		if pos:
			rangeObj=pos.innerTextInfo._rangeObj 
		else:
			rangeObj=self.rootNVDAObject.WinwordDocumentObject.range(0,0)
		includeCurrent=False if pos else True
		if nodeType=="bookmark":
			return BookmarkWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
		elif nodeType=="comment":
			return CommentWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
		elif nodeType=="revision":
			return RevisionWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
		elif nodeType=="endnote":
			return EndnoteWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
		elif nodeType=="field":
			fields = FieldWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
			formfields=FormFieldWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
			return browseMode.mergeQuickNavItemIterators([fields,formfields],direction)
		elif nodeType=="formfield":
			return FormFieldWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
		elif nodeType=="footnote":
			return FootnoteWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
		elif nodeType=="graphic":
			graphics = GraphicWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
			charts = ChartWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
			shapes = ShapeWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
			return browseMode.mergeQuickNavItemIterators([graphics, charts,shapes],direction)

		elif nodeType=="grammaticalError":
			return GrammaticalErrorWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
		elif nodeType=="spellingError":
			return SpellingErrorWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
		elif nodeType=="section":
			return SectionWinWordCollectionQuicknavIterator(nodeType,self,direction,rangeObj,includeCurrent).iterate()
		return super(WordDocumentTreeInterceptorEx, self)._iterNodesByType(nodeType,direction,pos)

class WordDocumentCommentQuickNavItemEx(WordDocumentCommentQuickNavItem):
	@property
	def label(self):
		author=self.collectionItem.author
		date=self.collectionItem.date
		text=self.collectionItem.range.text
		if text == None:
			text = ""
		return NVDAString(_unicode("comment: {text} by {author} on {date}")).format(author=author,text=text,date=date)


class CommentWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass=WordDocumentCommentQuickNavItemEx
	def collectionFromRange(self,rangeObj):
		return rangeObj.comments

class WordDocumentBookmarkQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		item = self.collectionItem
		text = ""
		if not item.empty:
			text = item.range.text
		name = item.Name  if item.Name else _("Bookmark")
		return _("{name} text: {text}") .format(name = name, text= text)
	def report(self,readUnit=None):
		ui.message(_("Bookmark %s" %self.collectionItem.index))



class BookmarkWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass=WordDocumentBookmarkQuickNavItem
	def collectionFromRange(self,rangeObj):
		return rangeObj.Bookmarks

class WordDocumentEndnoteQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		return _("Endnote{index}").format(index =self.collectionItem.index) 
		
	def rangeFromCollectionItem(self,item):
		return item.Reference
	def report(self,readUnit=None):
		ui.message(_("Endnote {index}") .format(index = self.collectionItem.index))



class EndnoteWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass=WordDocumentEndnoteQuickNavItem
	def collectionFromRange(self,rangeObj):
		return rangeObj.Endnotes
class WordDocumentFieldQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		typeText = Field._getTypeText(self.collectionItem.type)
		return _(_unicode("Field {index}, type: {type}")).format(index =self.collectionItem.index, type = typeText) 
	def rangeFromCollectionItem(self,item):
		return item.Code


class FieldWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass=WordDocumentFieldQuickNavItem
	def collectionFromRange(self,rangeObj):
		return rangeObj.fields

class WordDocumentFormFieldQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
	
		return _unicode("{name}").format(name =self.collectionItem.Name) 

	
class FormFieldWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass=WordDocumentFormFieldQuickNavItem
	def collectionFromRange(self,rangeObj):
		return rangeObj.FormFields



	

class WordDocumentFootnoteQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		return _("Footnote {index}").format(index =self.collectionItem.index) 
	def rangeFromCollectionItem(self,item):
		return item.Reference
			
	def report(self,readUnit=None):
		ui.message(_("Footnote {index}") .format(index = self.collectionItem.index))


	def moveTo(self):
		info=self.textInfo.copy()
		info.collapse()
		self.document._set_selection(info,reason=browseMode.REASON_QUICKNAV)


class FootnoteWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass=WordDocumentFootnoteQuickNavItem
	def collectionFromRange(self,rangeObj):
		return rangeObj.Footnotes



class WordDocumentShapeQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		return _unicode("%s") %self.collectionItem.name
	def rangeFromCollectionItem(self,item):
		return item.Anchor
	def report(self,readUnit=None):
		ui.message(_("%s" %self.collectionItem.Name))

class ShapeWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass=WordDocumentShapeQuickNavItem
	def collectionFromRange(self,rangeObj):
		try:
			return rangeObj.ShapeRange
		except:
			return None

	
class WordDocumentGrammaticalErrorQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):
		text = self.collectionItem.text
		if len(text) > 100:
			text = "%s ..." %text[:100]
		return _(_unicode("{text}")).format(text =text) 
		
	def rangeFromCollectionItem(self,item):
		return item

		

class GrammaticalErrorWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass=WordDocumentGrammaticalErrorQuickNavItem
	def collectionFromRange(self,rangeObj):
		return rangeObj.GrammaticalErrors
		
	
			
class WordDocumentSectionQuickNavItem(WordDocumentCollectionQuickNavItem):
	@property
	def label(self):

		return _(_unicode("Section {index}")).format(index =self.collectionItem.index) 
	def report(self,readUnit=None):
		ui.message(_("Section %s" %self.collectionItem.index))




class SectionWinWordCollectionQuicknavIterator(WinWordCollectionQuicknavIterator):
	quickNavItemClass=WordDocumentSectionQuickNavItem
	def collectionFromRange(self,rangeObj):
		return rangeObj.Sections

class ElementsListDialog(wx.Dialog):

	ELEMENT_TYPES=(
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
	lastSelectedElementType=0
	_timer = None
	
	def __init__(self, document):
		self.document = document
		# Translators: The title of the browse mode Elements List dialog.
		super(ElementsListDialog, self).__init__(gui.mainFrame, wx.ID_ANY, NVDAString("Elements List"))
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		contentsSizer = wx.BoxSizer(wx.VERTICAL)
		childSizer=wx.BoxSizer(wx.VERTICAL)
		# Translators: The label of a list of items to select the type of element
		# in the browse mode Elements List dialog.
		childLabel=wx.StaticText(self,wx.NewId(),label= NVDAString("&Type:") , style =wx.ALIGN_CENTRE )
		childSizer.Add(childLabel, )
		self.childListBox =wx.ListBox(self,wx.ID_ANY,name= "TypeName" ,choices=tuple(et[1] for et in self.ELEMENT_TYPES),  style = wx.LB_SINGLE ,size= (596,130))
		if self.childListBox.GetCount():
			self.childListBox.SetSelection(self.lastSelectedElementType)
		self.childListBox.Bind(wx.EVT_LISTBOX, self.onElementTypeChange)
		self.childListBox.Bind(wx.EVT_SET_FOCUS, self.onChildBoxFocus)
		childSizer.Add(self.childListBox)
		contentsSizer.Add(childSizer, flag=wx.EXPAND)
		contentsSizer.AddSpacer(gui.guiHelper.SPACE_BETWEEN_VERTICAL_DIALOG_ITEMS)

		self.tree = wx.TreeCtrl(self, size=wx.Size(500, 600), style=wx.TR_HAS_BUTTONS | wx.TR_HIDE_ROOT | wx.TR_LINES_AT_ROOT | wx.TR_SINGLE | wx.TR_EDIT_LABELS)
		self.tree.Bind(wx.EVT_SET_FOCUS, self.onTreeSetFocus)
		self.tree.Bind(wx.EVT_CHAR, self.onTreeChar)
		self.treeRoot = self.tree.AddRoot("root")
		contentsSizer.Add(self.tree,flag=wx.EXPAND)
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
		self.activateButton = bHelper.addButton(self, label= NVDAString("&Activate"))
		self.activateButton.Bind(wx.EVT_BUTTON, lambda evt: self.onAction(True))
		
		# Translators: The label of a button to move to an element
		# in the browse mode Elements List dialog.
		self.moveButton = bHelper.addButton(self, label= NVDAString("&Move to"))
		self.moveButton.Bind(wx.EVT_BUTTON, lambda evt: self.onAction(False))
		bHelper.addButton(self, id=wx.ID_CANCEL)

		contentsSizer.Add(bHelper.sizer, flag=wx.ALIGN_RIGHT)
		mainSizer.Add(contentsSizer, border=gui.guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		mainSizer.Fit(self)
		self.SetSizer(mainSizer)

		self.tree.SetFocus()
		self.initElementType(self.ELEMENT_TYPES[self.lastSelectedElementType][0])
		self.CentreOnScreen()

	def onElementTypeChange(self, evt):
		if self._timer:
			self._timer.Stop()
		elementType = self.childListBox.GetSelection()
		# We need to make sure this gets executed after the focus event.
		# Otherwise, NVDA doesn't seem to get the event.
		# modified because of error  when type spelling error type
		#queueHandler.queueFunction(queueHandler.eventQueue, self.initElementType, self.ELEMENT_TYPES[elementType][0])
		wx.CallLater(100, self.initElementType, self.ELEMENT_TYPES[elementType][0])
		self.lastSelectedElementType=elementType


	
	def initElementType(self, elType):
		from .ww_tones import RepeatBeep
		rb = RepeatBeep()
		rb.start()
		maxElement = None
		if elType in ["grammaticalError","spellingError"]:
			maxElement = 80
		if elType in ("link","button", "radioButton", "checkBox"):
			# Links, buttons  , radio button, check box can be activated.
			self.activateButton.Enable()
			self.SetAffirmativeId(self.activateButton.GetId())
		else:
			# No other element type can be activated.
			self.activateButton.Disable()
			self.SetAffirmativeId(self.moveButton.GetId())

		# Gather the elements of this type.
		self._elements = []
		self._initialElement = None
		parentElements = []
		isAfterSelection=False
		try:
			for item in self.document._iterNodesByType(elType):
				# Find the parent element, if any.
				for parent in reversed(parentElements):
					if item.isChild(parent.item):
						break
					else:
						# We're not a child of this parent, so this parent has no more children and can be removed from the stack.
						parentElements.pop()
				else:
					# No parent found, so we're at the root.
					# Note that parentElements will be empty at this point, as all parents are no longer relevant and have thus been removed from the stack.
					parent = None
	
				element=self.Element(item,parent)
				self._elements.append(element)
				if maxElement is not None:
					maxElement -=1
					if maxElement == 0:
						break
	
				if not isAfterSelection:
					isAfterSelection=item.isAfterSelection
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
			
		except:
			log.error("initElementType: Can not find all elements")
		rb.stop()
		# Start with no filtering.
		self.filterEdit.ChangeValue("")
		self.filter("", newElementType=True)
		self.sayNumberOfElements(maxElement is not None and maxElement == 0)
		
	def sayNumberOfElements(self,limited = False ):
		count = self.tree.Count
		if not self.childListBox.HasFocus(): return
		if count:
			if limited:
				# Translators: message to indicate the number of elements is limited.
				msg = _("limited to  %s elements") %str(count) if count > 1 else _("One element")
			else:
				# Translators: message to the user to report number of elements
				msg = _("%s elements") %str(count) if count > 1 else _("One element")
		else:
			# Translators: message to the user when there is no element
			msg =_("no element")
		queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, msg)


	def onChildBoxFocus(self, evt):
		self.sayNumberOfElements()		

	def filter(self, filterText, newElementType=False):
		# If this is a new element type, use the element nearest the cursor.
		# Otherwise, use the currently selected element.
		if wx.version().startswith("4"):
			defaultElement = self._initialElement if newElementType else self.tree.GetItemData(self.tree.GetSelection())
		else:
			defaultElement = self._initialElement if newElementType else self.tree.GetItemPyData(self.tree.GetSelection())
		# Clear the tree.
		self.tree.DeleteChildren(self.treeRoot)

		# Populate the tree with elements matching the filter text.
		elementsToTreeItems = {}
		defaultItem = None
		matched = False
		#Do case-insensitive matching by lowering both filterText and each element's text.
		filterText=filterText.lower()
		for element in self._elements:
			label=element.item.label
			if filterText and filterText not in label.lower():
				continue
			matched = True
			parent = element.parent
			if parent:
				parent = elementsToTreeItems.get(parent)
			item = self.tree.AppendItem(parent or self.treeRoot, label)
			if wx.version().startswith("4"):
				self.tree.SetItemData(item, element)
			else:
				self.tree.SetItemPyData(item, element)
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
		# If the activate button isn't the default button, it is disabled for this element type and shouldn't be enabled here.
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
			item=self.tree.GetSelection()
			if item:
				if wx.version().startswith("4"):
					selectedItemType=self.tree.GetItemData(item).item
				else:
					selectedItemType=self.tree.GetItemPyData(item).item
				self.tree.EditLabel(item)
				evt.Skip()

		elif key >= wx.WXK_START or key == wx.WXK_BACK:
			# Non-printable character.
			self._searchText = ""
			evt.Skip()

		else:
			# Search the list.
			# We have to implement this ourselves, as tree views don't accept space as a search character.
			char = uniCHR(evt.UnicodeKey).lower()
			# IF the same character is typed twice, do the same search.
			if self._searchText != char:
				self._searchText += char
			if self._searchCallLater:
				self._searchCallLater.Restart()
			else:
				self._searchCallLater = core.callLater(1000, self._clearSearchText)
			self.search(self._searchText)

	def onTreeLabelEditBegin(self,evt):
		item=self.tree.GetSelection()
		if wx.version().startswith("4"):
			selectedItemType = self.tree.GetItemPyData(item).item
		else:
			selectedItemType = self.tree.GetItemPyData(item).item
		if not selectedItemType.isRenameAllowed:
			evt.Veto()

	def onTreeLabelEditEnd(self,evt):
			selectedItemNewName=evt.GetLabel()
			item=self.tree.GetSelection()
			if wx.version().startswith("4"):
				selectedItemType = self.tree.GetItemData(item).item
			else:
				selectedItemType = self.tree.GetItemPyData(item).item
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
		items = itertools.chain(self._iterReachableTreeItemsFromItem(item), self._iterReachableTreeItemsFromItem(self.tree.GetFirstChild(self.treeRoot)[0]))
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
		self.Close()
		# Save off the last selected element type on to the class so its used in initialization next time.
		self.__class__.lastSelectedElementType=self.lastSelectedElementType
		item = self.tree.GetSelection()
		if wx.version().startswith("4"):
			item = self.tree.GetItemData(item).item
		else:
			item = self.tree.GetItemPyData(item).item
		if activate:
			item.activate()
		else:
			def move():
				speech.cancelSpeech()
				item.moveTo()
				item.report()
			core.callLater(100, move)

