# appModules\winword\ww_grammaticalErrors.py
#A part of wordAccesseEnhancement add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
import api
import wx
import ui
import time
from eventHandler import queueEvent
from .ww_wdConst import wdActiveEndPageNumber , wdFirstCharacterLineNumber , wdSentence, wdGoToGrammaticalError
from .ww_collection import Collection, CollectionElement,ReportDialog



class GrammaticalError(CollectionElement):


	def __init__(self, parent, item ):
		super(GrammaticalError, self).__init__(parent, item)
		self.start = item.Start
		self.end = item.End
		self.text = ""
		if item.Text:
			self.text = item.Text
		self.setLineAndPageNumber()

		
			
	def formatInfos(self):
		sInfo = _("""Page {page}, line {line}
Ttext involved:
{text}
""")


		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(page =self.page, line= self.line, text = self.text)
	def goTo(self):
		self.obj.Select()
class GrammaticalErrors(Collection):
	_propertyName = (("GrammaticalErrors",GrammaticalError),)
	_name = (_("Grammatical error"), _("Grammatical errors"))
	_wdGoToItem = None #wdGoToGrammaticalError don't work
	def __init__(self, parent,obj, rangeType):
		self.rangeType = rangeType
		self.dialogClass = GrammaticalErrorsDialog
		self.noElement = _("No gramatical error")
		wordApp = obj._WinwordWindowObject.Application
		if not wordApp.Options.CheckGrammarAsYouType:
			self.collection = None
			ui.message(_("Not available, Check grammar as you type is not activated")) 
			time.sleep(3.0)
			api.processPendingEvents()
			time.sleep(0.2)
			obj  = api.getFocusObject()
			queueEvent("gainFocus",obj)
			return
		super(GrammaticalErrors, self).__init__( parent,obj)
		
	

class GrammaticalErrorsDialog(ReportDialog):
	def __init__(self, parent,obj ):
		super(GrammaticalErrorsDialog, self).__init__(parent,obj)
		
	def initializeGUI (self):
		self.lcLabel = _("Errors:")
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
			)


		self.tc1 = {
		"label": _("Involved sentence"),
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
