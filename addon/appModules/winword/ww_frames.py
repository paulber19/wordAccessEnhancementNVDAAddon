# appModules\winword\ww_frames.py
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
from .ww_wdConst import wdActiveEndPageNumber , wdFirstCharacterLineNumber , wdGoToFrame
from .ww_collection import Collection, CollectionElement,ReportDialog, convertPixelToUnit


wdFrameSizeRules = {
0 : "Auto", # wdFrameAuto
1 : "At least",#wdFrameAtLeast
2 :"Exact" # wdFrameExact
}

class Frame(CollectionElement):
	def __init__(self, parent, item):
		super(Frame, self).__init__(parent, item)
		self.borders = item.borders
		self.range = item.Range
		self.height = item.Height # in points
		self.heightRule = item.HeightRule
		self.width = item.Width # in points
		self.widthRule = item.WidthRule
		self.lockAnchor = item.LockAnchor
		self.horizontalPosition = item.HorizontalPosition
		self.relativeHorizontalPosition = item.RelativeHorizontalPosition
		self.verticalPosition = item.VerticalPosition
		self.relativeVerticalPosition = item.RrelativeVerticalPosition
		self.textWrap = item.TextWrap
		self.orizontalDistanceFromText    = item.HorizontalDistanceFromText    # in points
		self.verticalDistanceFromText    = item.VerticalDistanceFromText   # in points
		r = self.parent.doc.range (self.range.Start, self.range.Start)
		self.line = r.information(wdFirstCharacterLineNumber )
		self.page = r.Information(wdActiveEndPageNumber )




	
	def formatInfos(self):
		height = "%.2f" %convertPixelToUnit(self.height)
		width = "%.2f" %convertPixelToUnit(self.width)
		sInfo = _("""Page {page}, line {line}
Height: {height}
Width:
{width}
""")
		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(page =self.page, line= self.line, height = height, width = width)
		

class Frames(Collection):
	_propertyName = (("Frames",Frame),)
	_name = (_("Frame"), _("Frames"))
	_wdGoToItem = None 
	

	def __init__(self, parent, obj, rangeType):
		self.rangeType = rangeType
		self.dialogClass = FramesDialog
		self.noElement = _("No frame")
		super(Frames, self).__init__( parent, obj)
				
class FramesDialog(ReportDialog):

	def __init__(self, parent,obj ):
		super(FramesDialog, self).__init__(parent,obj)
		
	def initializeGUI (self):
		self.lcLabel = _("Frames")
		self.lcColumns = (
			(_("Frame"), 100),
			(_("Location"),150),
			(_("Width"),1500),
			(_("Height"), 150)
			)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]
		
		self.lcSize = (lcWidth, self._defaultLCWidth)
		
		self.buttons = (
			(100, _("&Go to"),self.goTo),
			(101, _("&Delete"), self.delete),
			)


		self.tc1 = None
		self.tc2 = None



	def get_lcColumnsDatas(self, element):
		location = (_("Page {page}, line {line}")).format(page = element.page, line= element.line)
		height = "%.2f" %convertPixelToUnit(element.height)
		width = "%.2f" %convertPixelToUnit(element.width)
		index = self.collection.index(element)+1
		datas = (index,location, width, height)
		return datas

