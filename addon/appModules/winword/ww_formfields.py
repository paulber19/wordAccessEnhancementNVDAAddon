# appModules\winword\ww_formfields.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
from .ww_wdConst import wdGoToField
from .ww_collection import Collection, CollectionElement, ReportDialog

addonHandler.initTranslation()


class FormField(CollectionElement):
	def __init__(self, parent, item):
		super(FormField, self).__init__(parent, item)
		self.type = item.Type
		self.enabled = item.Enabled
		self.range = item.Range
		self.checkbox = item.CheckBox
		self.resultText = ""
		try:
			if item.Result.Text:
				self.resultText = item.Result.Text
		except Exception:
			pass
		self.statusText = item.StatusText
		self.start = item.range.Start
		self.name = ""
		try:
			if item.name:
				self.name = item.Name
		except Exception:
			pass
		self.setLineAndPageNumber()

	def goTo(self):
		self.obj.Select()

	def formatInfos(self):
		sInfo = _("""Page {page}, line {line}
""")
		sInfo = sInfo.replace("\n", "\r\n")
		return sInfo.format(
			page=self.page, line=self.line)

	@classmethod
	def _getTypeText(cls, type):
		try:
			text = wdFieldTypes[type]
		except Exception:
			text = str(type)
		return text

	def getTypeText(self):
		return self.__class__._getTypeText(self.type)


class FormFields(Collection):
	_propertyName = (("FormFields", FormField),)
	_name = (_("Formfield"), _("Formfields"))
	_wdGoToItem = wdGoToField

	def __init__(self, parent, focus, rangeType):
		self.rangeType = rangeType
		self.dialogClass = FormFieldsDialog
		self.noElement = _("No formfield")
		super(FormFields, self).__init__(parent, focus)


class FormFieldsDialog(ReportDialog):
	def __init__(self, parent, obj):
		super(FormFieldsDialog, self).__init__(parent, obj)

	def initializeGUI(self):
		self.lcLabel = _("FormFields")
		self.lcColumns = (
			(_("Number"), 100),
			(_("Location"), 150),
			(_("Type"), 300)
		)
		lcWidth = 0
		for column in self.lcColumns:
			lcWidth = lcWidth + column[1]
		self.lcSize = (lcWidth, self._defaultLCWidth)
		self.buttons = (
			(100, _("&Go to"), self.goTo),
			(101, _("&Delete"), self.delete),
			(102, _("Delete &all"), self.deleteAll)
		)
		self.tc1 = {
			"label": _("Result"),
			"size": (800, 200)
		}

		self.tc2 = {
			"label": _("Status"),
			"size": (800, 200)
		}

	def get_lcColumnsDatas(self, element):
		typeText = element.getTypeText()
		location = (_("Page {page}, line {line}")).format(
			page=element.page, line=element.line)
		index = self.collection.index(element) + 1
		datas = (index, location, typeText)
		return datas

	def get_tc1Datas(self, element):
		return element.resultText

	def get_tc2Datas(self, element):
		return element.statusText


wdFieldTypes = {
	-1: _("Empty"),  # wdFieldEmpty
	3: _("Ref"),  # wdFieldRef
	4: "IndexEntry",  # wdFieldIndexEntry
	5: _("FootnoteRef"),  # wdFieldFootnoteRef
	6: "Set",  # wdFieldSet
	7: "If",  # wdFieldIf
	8: "Index",  # wdFieldIndex
	9: "TOCEntry",  # wdFieldTOCEntry
	10: _("StyleRef"),  # wdFieldStyleRef
	11: "RefDoc",  # wdFieldRefDoc
	12: "Sequence",  # wdFieldSequence
	13: _("TOC"),  # wdFieldTOC
	14: "Info",  # wdFieldInfo
	15: _("Title"),  # wdFieldTitle
	16: _("Subject"),  # wdFieldSubject
	17: _("Author"),  # wdFieldAuthor
	18: _("Keyword"),  # wdFieldKeyWord
	19: _("Comment"),  # wdFieldComments
	20: "LastSavedBy",  # wdFieldLastSavedBy
	21: _("CreateDate"),  # wdFieldCreateDate
	22: "SaveDate",  # wdFieldSaveDate
	23: "FindDate",  # wdFieldPrintDate
	24: "RevisionNum",  # wdFieldRevisionNum
	25: "EditTime",  # wdFieldEditTime
	26: "NumPage",  # wdFieldNumPages
	27: "NumWord",  # wdFieldNumWords
	28: "NumChar",  # wdFieldNumChars
	29: "Fielname",  # wdFieldFileName
	30: "Template",  # wdFieldTemplate
	31: _("Date"),  # wdFieldDate
	32: "Time",  # wdFieldTime
	33: _("Page"),  # wdFieldPage
	34: _("Expression"),  # wdFieldExpression
	35: _("Quote"),  # wdFieldQuote
	36: _("Include"),  # wdFieldInclude
	37: "PageRef",  # wdFieldPageRef
	38: "Ask",  # wdFieldAsk
	39: "FillIn",  # wdFieldFillIn
	40: "Data",  # wdFieldData
	41: "Next",  # wdFieldNext
	42: "NextIf",  # wdFieldNextIf
	43: "SkipIf",  # wdFieldSkipIf
	44: "MergeRec",  # wdFieldMergeRec
	45: "DDE",  # wdFieldDDE
	46: "DDEAuto",  # wdFieldDDEAuto
	47: "Glossary",  # wdFieldGlossary
	48: "Print",  # wdFieldPrint
	49: "Formula",  # wdFieldFormula
	50: "GoToButton",  # wdFieldGoToButton
	51: "MacroButton",  # wdFieldMacroButton
	52: "AutoNumOutline",  # wdFieldAutoNumOutline
	53: "AutoNumLegal",  # wdFieldAutoNumLegal
	54: "AutoNum",  # wdFieldAutoNum
	55: "import",  # wdFieldImport
	56: _("Link"),  # wdFieldLink
	57: _("Symbol"),  # wdFieldSymbol
	58: "Embedded",  # wdFieldEmbed
	59: "MergeField",  # wdFieldMergeField
	60: "UserName",  # wdFieldUserName
	61: "UserInitials",  # wdFieldUserInitials
	62: "UserAddress",  # wdFieldUserAddress
	63: "BarCode",  # wdFieldBarCode
	64: "DocVariable",  #: wdFieldDocVariable
	65: "Section",  # wdFieldSection
	66: "SectionPage",  # wdFieldSectionPages
	67: "IncludePicture",  # wdFieldIncludePicture
	68: "IncludeText",  # wdFieldIncludeText
	69: "Size",  # wdFieldFileSize
	70: "FormTextInput",  # wdFieldFormTextInput
	71: "FormCheckBox",  # wdFieldFormCheckBox
	72: "NoteRef",  # wdFieldNoteRef
	73: "TOA",  # wdFieldTOA
	74: "TOAEntry",  # wdFieldTOAEntry
	75: "MergeSeq",  # wdFieldMergeSeq
	77: "Private",  # wdFieldPrivate
	78: "Database",  # wdFieldDatabase
	79: "AutoText",  # wdFieldAutoText
	80: "Compare",  # wdFieldCompare
	81: "Addin",  # wdFieldAddin
	82: "Subscriber",  # wdFieldSubscriber
	83: "FormDropDown",  # wdFieldFormDropDown
	84: "Advance",  # wdFieldAdvance
	85: "DocProperty",  # wdFieldDocProperty
	87: "OCX",  # wdFieldOCX
	88: "Hyperlink",  # wdFieldHyperlink
	89: "AutoTextList",  # wdFieldAutoTextList
	90: "ListNum",  # wdFieldListNum
	91: "HTMLActiveX",  # wdFieldHTMLActiveX
	92: "BidiOutline ",  # wdFieldBidiOutline
	93: "AddressBlock",  # wdFieldAddressBlock
	94: "GreetingLine ",  # wdFieldGreetingLine
	95: "Shape"  # wdFieldShape
}
