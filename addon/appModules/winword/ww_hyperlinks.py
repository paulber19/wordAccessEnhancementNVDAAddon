# appModules\winword\ww_hyperlinks.py
#A part of wordAccessEnhancement add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
import api
import wx
import gui
import ui
from NVDAObjects.window.winword import WordDocument
import time
from eventHandler import queueEvent
from .ww_wdConst import wdActiveEndPageNumber , wdFirstCharacterLineNumber , wdGoToHyperlink
from .ww_collection import Collection, CollectionElement,ReportDialog

class Hyperlink(CollectionElement):
	def __init__(self, parent, item):
		super(Hyperlink, self).__init__(parent, item)
		self.type = item.Type
		try:
			self.address = item.Address
		except:
			self.address = ""
		try:
			self.subAddress = item.SubAddress
		except:
			self.subAddress = ""
		try:
			self.target = item.Target
		except:
			self.target = ""
		try:
			self.textToDisplay = item.TextToDisplay
		except:
			self.textToDisplay = ""


		try:
			self.emailSubject = item.EmailSubject
		except:
			self.emailSubject = ""

		try:
			self.name = item.Name
		except:
			self.name = ""
			
		try:
					self.screenTyp = item.Screentyp
		except:
			self.screenTyp = ""
		try:
			self.shape = item.Shape
		except:
			self.shape = None
		self.start = item.Range.Start
		r= self.parent.doc.range (self.start, self.start)
		self.line = r.information(wdFirstCharacterLineNumber )
		self.page = r.Information(wdActiveEndPageNumber )
		
	def getTypeText(self):
		msoHyperLinkTypes = {
			0: _("Hyperlink to range"),# msoHyperlinkRange 
			1 : _("Hyperlink to shape"), # msoHyperlinkShape 
			2:  _("Hyperlink to inLine Shape") # msoHyperlinkInlineShape 
			}
		return msoHyperLinkTypes[self.type]


	def formatInfos(self):

		sInfo = _("""Page {page}, line {line}
Name: {name}
type: {type}
Address: {address}
Sub-address: {subAddress}
""")
		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(page =self.page, line= self.line, name = self.name, type = self.getTypeText(), address = self.address, subAddress = self.subAddress)
		
		
class Hyperlinks(Collection):
	_propertyName = (("Hyperlinks",Hyperlink),)
	_name = (_("Hyperlink"), _("Hyperlinks"))
	_wdGoToItem = None

	def __init__(self, parent, obj, rangeType):
		self.rangeType = rangeType
		self.dialogClass = HyperlinksDialog
		self.noElement = _("No hyperlink")
		super(Hyperlinks, self).__init__( parent, obj)

class HyperlinksDialog(ReportDialog):

	def __init__(self, parent,obj ):
		super(HyperlinksDialog, self).__init__(parent,obj)
		
	def initializeGUI (self):
		self.lcLabel = _("Hyperlinks")
		self.lcColumns = (
			(_("Number"), 100),
			(_("Location"),150),
			(_("Name"),300),
			(_("Type"), 150)
			)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]
		
		self.lcSize = (lcWidth, self._defaultLCWidth)
		
		self.buttons = (
			(100, _("&Go to"),self.goTo),
			(101, _("&Delete"), self.delete)
			)


		self.tc1 = {
		"label":   _("Text to display"),
		"size": (800,200)
		}
		self.tc2 = None



	def get_lcColumnsDatas(self, element):
		location = (_("Page {page}, line {line}")).format(page = element.page, line= element.line)
		index = self.collection.index(element)+1
		datas = (index,location, element.name, element.getTypeText())
		return datas

	def get_tc1Datas(self, element):
		return element.textToDisplay 
