# appModules\winword\ww_comments.py
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
import textInfos
from eventHandler import queueEvent
from .ww_wdConst import wdActiveEndPageNumber , wdFirstCharacterLineNumber , wdGoToComment, wdGoToAbsolute
from .ww_collection import Collection, CollectionElement, ReportDialog



def insertComment(wordApp):

	def callback(result):
		if result == wx.ID_OK:
			text =  d.GetValue()
			comments = doc.Comments
			r = doc.Range(selection.Start, selection.End)
			comments.Add(r, text)
		else:
			pass


	# ask the user to type the comment's text
	d = wx.TextEntryDialog(gui.mainFrame, _("Enter the comment's text"), _("Insert comment"), style = wx.TE_MULTILINE|wx.OK|wx.CANCEL)
	doc = wordApp.ActiveDocument
	selection = wordApp.Selection
	gui.runScriptModalDialog(d, callback)
	




class Comment(CollectionElement):
	def __init__(self, parent, item):
		super(Comment, self).__init__(parent, item)
		self.reference = item.Reference

		self.text = ""
		if item.Range.Text:
			self.text = item.range.text

		self.author = ""
		if item.author:
			self.author = item.Author
		self.date = ""
		if item.date:
			self.date = item.date
		r = self.associatedTextRange = item.scope
		self.start = r.Start
		self.associatedText = ""
		if r.text:
			self.associatedText = r.text
		self.setLineAndPageNumber()
	
	def formatInfos(self):
		sInfo = _("""Page {page}, line {line}
Author: {author}
Date: {date}
Comment's text:
{text}
Commented text:
{commentedText}
""")

		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(page =self.page, line= self.line, text = self.text, author = self.author, date = self.date, commentedText= self.associatedText)
	
		
class Comments(Collection):
	_propertyName = (("Comments",Comment),)
	_name = (_("Comment"), _("Comments"))
	_wdGoToItem =  wdGoToComment
	
	def __init__(self, parent, obj, rangeType):
		self.obj = obj
		self.rangeType = rangeType
		self.dialogClass = CommentsDialog
		self.noElement = _("No comment")
		reference = self.getReferenceAtFocus()
		if  reference:
			self.reference = reference
		super(Comments, self).__init__(parent, obj)


	
	def getReferenceAtFocus(self):
		info=self.obj.makeTextInfo(textInfos.POSITION_CARET)
		info.expand(textInfos.UNIT_CHARACTER)
		fields=info.getTextWithFields(formatConfig={'reportComments':True})
		for field in reversed(fields):
			if isinstance(field,textInfos.FieldCommand) and isinstance(field.field,textInfos.FormatField): 
				reference=field.field.get('comment')
				if reference:
					return int(reference)
		return None
					

class CommentsDialog(ReportDialog):

	def __init__(self, parent,obj ):
		super(CommentsDialog, self).__init__(parent,obj)
	
	def initializeGUI (self):
		self.lcLabel = _("Comments:")
		self.lcColumns = (
			(_("Number"), 100),
			(_("Location"),150),
			(_("Author"), 300),
			(_("Date"), 150)
			)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]
			
		self.lcSize = (lcWidth, self._defaultLCWidth)
		
		self.buttons = (
			(100, _("&Go to"),self.goTo),
			(101, _("&Delete"), self.delete),
			(102, _("Delete &all"), self.deleteAll)
			)

		self.tc1 = {
		"label": _("Comment's text"),
		"size": (800,200)
		}
		self.tc2 = {
		"label":  _("Commented text"),
		"size": (800,200)
		}





	def get_lcColumnsDatas(self, element):

		location = (_("Page {page}, line {line}")).format(page = element.page, line= element.line)
		index = self.collection.index(element)+1
		datas = (index, location, element.author, element.date)
		return datas
	def get_tc1Datas(self, element):
		return element.text
	def get_tc2Datas(self, element):
		return element.associatedText 
	



