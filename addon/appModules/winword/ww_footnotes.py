# appModules\winword\ww_footnotes.py
#A part of wordAccessEnhancement add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
import api
import wx
import gui
import ui
import time
from eventHandler import queueEvent
import textInfos
from .ww_wdConst import wdActiveEndPageNumber , wdFirstCharacterLineNumber , wdGoToFootnote
from .ww_collection import Collection, CollectionElement,ReportDialog





		
class Footnote(CollectionElement):
	def __init__(self, parent, item):
		super(Footnote, self).__init__(parent, item)
		self.start = item.reference.Start
		self.end = item.reference.End
		self.text = " "
		if item.Range.Text:
			self.text = item.Range.Text

		r = self.parent.doc.range (self.start, self.end)
		r.Collapse()
		self.line = r.information(wdFirstCharacterLineNumber )
		self.page = r.Information(wdActiveEndPageNumber )
		
			
	def formatInfos(self):
		sInfo = _("""Page {page}, line {line}
Note's Ttext:
{text}
""")

		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(page =self.page, line= self.line, text = self.text)


class Footnotes(Collection):
	_propertyName = (("Footnotes",Footnote),)
	_name = (_("Footnote"), _("Footnotes"))
	_wdGoToItem = wdGoToFootnote
	
	def __init__(self, parent, obj, rangeType):
		self.rangeType = rangeType
		self.dialogClass = FootnotesDialog
		self.noElement = _("No footnote")
		super(Footnotes, self).__init__( parent, obj)
		self.__class__._elementUnit = textInfos.UNIT_CHARACTER

	

class FootnotesDialog(ReportDialog):

	def __init__(self, parent,obj ):
		super(FootnotesDialog, self).__init__(parent,obj)
		
	def initializeGUI (self):
		self.lcLabel = _("Notes:")
		self.lcColumns = (
			(_("Number"), 100),
			(_("Location"),150),
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
		"label": _("Note's text"),
		"size": (800,200)
		}
		
		self.tc2 = None


	def get_lcColumnsDatas(self, element):
		location = _("Page {page}, line {line}") .format(page = element.page, line = element.line)
		index = self.collection.index(element)+1
		datas = (index, location)
		return datas
	def get_tc1Datas(self, element):
		return element.text

