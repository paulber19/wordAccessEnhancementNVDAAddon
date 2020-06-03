# appModules\winword\ww_shapes.py
#A part of wordAccessEnhancement add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
import api
import wx
import ui
from .ww_wdConst import wdActiveEndPageNumber , wdFirstCharacterLineNumber , wdGoToObject
from .ww_collection import Collection, CollectionElement, ReportDialog, convertPixelToUnit


class Shape(CollectionElement):
	def __init__(self, parent, item):
		super(Shape, self).__init__(parent, item)
		self.width = item.width
		self.height = item.height
		self.type = item.Type
		self.name = ""
		if item.Name:
			self.name = item.name

			self.range =item.anchor
		self.start = item.Anchor.Start
		self.end = item.Anchor.End
		self.text = item.Anchor.text
		self.alternativeText = ""
		if item.AlternativeText:
			self.alternativeText = item.AlternativeText
		try:
			self.title = item.Title
		except:
			self.title = ""
		self.setLineAndPageNumber()		
	
	def getTypeText(self):
		typeTexts = {
			-2 :  _("Mixed shape type"),#  msoShapeTypeMixed 
			1: _("AutoShape"), # msoAutoShape 
			2: _(" Callout"), # msoCallout 
			3 : _("Chart"),  #msoChart 
			4 : _("Comment"), #  msoComment 
			5 : _("Freeform"), # msoFreeform 
			6 : _("Group"),  # msoGroup 
			7 : _("Embedded OLE object"), # msoEmbeddedOLEObject 
			8 : _("Form control"), #  msoFormControl 
			9 : _("Line"), #  msoLine 
			10 : _("Linked OLE object"), #  msoLinkedOLEObject 
			11 : _("Linked picture"), #  msoLinkedPicture 
			12 : _("OLE control object"), #  msoOLEControlObject 
			13: _("Picture"), #  msoPicture 
			14 : _("Placeholder"), #  msoPlaceholder 
			15 : _("Text effect"), # msoTextEffect 
			16 : _("Media. "), # msoMedia 
			17 : _("Text box") , # msoTextBox 
			18 : "Script anchor", # msoScriptAnchor 
			19 : _("Table"), # msoTable 
			20 : _("Canvas"), #  msoCanvas 
			21 : _("Diagram"), # msoDiagram 
			22 : _("Ink"), #  msoInk 
			23 :  _("Ink comment"), # msoInkComment 
			24 : _("Smart art"), # msoSmartArt  
			25 : _("Slicer"), # msoSlicer  
			26 : _("Web video") #  msoWebVideo 
			}
		return typeTexts[self.type]
			
	def formatInfos(self):
		height = "%.2f" %convertPixelToUnit(self.height)
		width = "%.2f" %convertPixelToUnit(self.width)
		sInfo = _("""Page {page}, line {line}
Name: {name}
type: {typeText}
Width: {width}
Height: {height}
Alternative text:
{alternativeText}
""")

		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(page =self.page, line= self.line, typeText = self.getTypeText(), name = self.name, title = self.title,
		width = width, height = height, alternativeText = self.alternativeText)


class InLineShape(CollectionElement):
	def __init__(self, parent, item):
		super(InLineShape, self).__init__(parent, item)
		self.width = item.width
		self.height = item.height
		self.name = ""
		self.type = item.Type
		self.start = item.Range.Start
		self.end = item.Range.End
		self.text = item.Range.text
		r = self.parent.doc.range(self.start, self.start)
		self.line = r.information(wdFirstCharacterLineNumber )
		self.page = r.Information(wdActiveEndPageNumber )
		self.alternativeText = ""
		if item.AlternativeText:
			self.alternativeText = item.AlternativeText
		self.title = ""
		#if item.Title:
			#self.title = item.Title
	

	def getTypeText(self):
		typeTexts = {
			1: _("Embedded OLE object"), # wdInlineShapeEmbeddedOLEObject 
			2 :_("Linked OLE object"), #  wdInlineShapeLinkedOLEObject 
			3: _("Picture"),  #  wdInlineShapePicture 
			4 : _("Linked picture"), #  wdInlineShapeLinkedPicture 
			5: _("OLE control object"), #  wdInlineShapeOLEControlObject 
			6 : _("Horizontal line"),  #  wdInlineShapeHorizontalLine 
			7 :_("Picture with horizontal line"), #  wdInlineShapePictureHorizontalLine 
			8 : _("Linked picture with horizontal line"), #  wdInlineShapeLinkedPictureHorizontalLine 
			9 : _("Picture used as a bullet"),  #  wdInlineShapePictureBullet 
			10 : "Script anchor.Anchor for script", #  wdInlineShapeScriptAnchor 
			11 : _("OWS anchor"), #  wdInlineShapeOWSAnchor 
			12 : "Inline chart",  #  wdInlineShapeChart 
			13 : _("Inline diagram"), #  wdInlineShapeDiagram 
			14 : "Locked inline shape canvas", #  wdInlineShapeLockedCanvas 
			15 : _("Smart art") #  wdInlineShapeSmartArt 
			}
			
		return typeTexts[self.type]
			
	def formatInfos(self):
		height = "%.2f" %convertPixelToUnit(self.height)
		width = "%.2f" %convertPixelToUnit(self.width)
		sInfo = _("""Page {page}, line {line}
type: {typeText}
Width: {width}
Height: {height}
Alternative ttext:
{alternativeText}
""")

		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(page =self.page, line= self.line, typeText = self.getTypeText(),  title = self.title,
		width = width, height = height, alternativeText = self.alternativeText)


class Shapes(Collection):
	_propertyName = (("ShapeRange", Shape),)
	_name = (_("Drawing layer object"), _("Drawing layer objects"))
	_wdGoToItem = None #wdGoToObject
	
	def __init__(self, parent, obj, rangeType):
		self.rangeType = rangeType
		self.dialogClass = ShapesDialog
		self.noElement = _("No object")
		super(Shapes, self).__init__( parent, obj)

class InLineShapes(Collection):
	_propertyName = (("InLineShapes", InLineShape),)
	_name = (_("Text layer object"), _("Text layer objects"))
	_wdGoToItem =  None # wdGoToObject
	
	def __init__(self, parent, obj, rangeType):
		self.rangeType = rangeType
		self.dialogClass = ShapesDialog
		self.noElement = _("No object")
		super(InLineShapes, self).__init__( parent, obj)
		
class ShapesDialog(ReportDialog):

	def __init__(self, parent,obj ):
		super(ShapesDialog, self).__init__(parent,obj)
		
	def initializeGUI (self):
		self.lcLabel = _("Objects:")
		self.lcColumns = (
			(_("Number"),50),
			(_("Location"),150),
			(_("Name"),150),
			("Type",300),
			(_("Height(in centimeter)"),100),
			(_("Width (in centimeter)"),100)
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
			"label" : _("Alternative text"),
			"size": (800,200)
			}
		self.tc2 = None

	def get_lcColumnsDatas(self, element):
		location = (_("Page {page}, line {line}")).format(page = element.page, line= element.line)
		height = "%.2f" %convertPixelToUnit(element.height)
		width = "%.2f" %convertPixelToUnit(element.width)
		typeText = element.getTypeText()
		index = self.collection.index(element)+1
		datas = (index,location, element.name,typeText, width, height)
		return datas
	


	def get_tc1Datas(self, element):
		return element.alternativeText