# appModules\winword\ww_automaticReading.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
from versionInfo import version_year, version_major
import config
import textInfos
from textInfos import SpeechSequence
from typing import List, Optional, Dict
import synthDriverHandler
import speech.commands
from speech import *  # noqa:F403
import sys
import os
try:
	# minimum nvda 2020.1
	from controlTypes import OutputReason as OutputReason
except ImportError:
	# nvda 2019.3
	OutputReason = str

_curAddon = addonHandler.getCodeAddon()
sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_NVDAStrings import NVDAString  # noqa:E402
from ww_addonConfigManager import _addonConfigManager, AutoReadingWith_Beep, AutoReadingWith_Voice  # noqa:E402
from ww_py3Compatibility import baseString  # noqa:E402
del sys.path[-1]

addonHandler.initTranslation()

SCT_Speech = "speech"
SCT_Many = "__many__"


def formatVoiceAutoSpeechSequence(seq):
	curSynthName = synthDriverHandler.getSynth().name
	autoReadingSynth = _addonConfigManager.getAutoReadingSynthSettings()
	if not autoReadingSynth:
		return seq
	currentSpeechSettings = getCurrentSpeechSettings()
	(autoReadingSynthName, autoReadingSynthSpeechSettings) = (autoReadingSynth["synthName"], autoReadingSynth)

	textList = []
	textList.append(speech.commands.EndUtteranceCommand())
	# start use of automatic reading synth
	textList.append(SetSynthCommand(autoReadingSynthName, autoReadingSynthSpeechSettings, temporary=True))
	textList.append(speech.commands.EndUtteranceCommand())
	textList.extend(seq)
	textList.append(speech.commands.EndUtteranceCommand())
	# restore main synth
	textList.append(SetSynthCommand(curSynthName, currentSpeechSettings, temporary=False))
	textList.append(speech.commands.EndUtteranceCommand())
	return textList


def formatBeepAutoSpeechSequence(seq):
	textList = []
	textList.append(speech.commands.EndUtteranceCommand())
	textList.append(speech.commands.EndUtteranceCommand())
	textList.append(speech.commands.BeepCommand(600, 40))
	textList.append(speech.commands.EndUtteranceCommand())
	textList.extend(seq)
	textList.append(speech.commands.EndUtteranceCommand())
	textList.append(speech.commands.BeepCommand(400, 40))
	textList.append(speech.commands.EndUtteranceCommand())
	return textList


def formatAutoSpeechSequence(seq):
	autoReadingWithOption = _addonConfigManager.getAutoReadingWithOption()
	if autoReadingWithOption == AutoReadingWith_Beep:
		return formatBeepAutoSpeechSequence(seq)
	if autoReadingWithOption == AutoReadingWith_Voice:
		return formatVoiceAutoSpeechSequence(seq)
	return seq


def getNotePropertiesSpeech(reason, value, note):
	seq = getPropertiesSpeech(reason=reason, value=value)
	seq.extend(formatAutoSpeechSequence([note]))
	return seq




class AutomaticReadingWordTextInfo(textInfos.TextInfo):
	def getCommentFormatFieldSpeech(self, commentReference):
		# with UIA enabled, commentReference is a boolean (True)
		if not(_addonConfigManager.toggleAutomaticReadingOption(False) and _addonConfigManager.toggleAutoCommentReadingOption(False)) or\
			type(commentReference) != baseString:
			# Translators: Reported when text contains a comment.
			text = NVDAString("has comment")
			return [text, ]
		from .ww_comments import Comment
		offset = int(commentReference)
		if hasattr(self.obj, "rootNVDAObject"):
			doc = self.obj.rootNVDAObject.WinwordApplicationObject.ActiveDocument
		else:
			doc = self.obj.WinwordApplicationObject.ActiveDocument
		textRange = doc.range(offset, offset + 1)
		commentObj = textRange.Comments[1]
		comment = Comment(self.obj, commentObj)
		commentText = comment.text
		commentAuthor = comment.author
		if commentText != "":
			seq = []
			# Translators: Reported when text contains a comment.
			text = _("comment of %s:") % commentAuthor if commentAuthor != "" else _("comment:") % commentText
			seq.append(text)
			seq.extend(formatAutoSpeechSequence([commentText]))
			return seq

	def getFootnote(self, footnoteReference):
		from .ww_footnotes import Footnote
		offset = int(footnoteReference)
		if hasattr(self.obj, "rootNVDAObject"):
			doc = self.obj.rootNVDAObject.WinwordApplicationObject.ActiveDocument
		else:
			doc = self.obj.WinwordApplicationObject.ActiveDocument
		footnoteObj = doc.FootNotes[offset]
		footnote = Footnote(self.obj, footnoteObj)
		return footnote.text
	
	def getEndNote(self, endnoteReference):
		from .ww_endnotes import Endnote
		offset = int(endnoteReference)
		if hasattr(self.obj, "rootNVDAObject"):
			doc = self.obj.rootNVDAObject.WinwordApplicationObject.ActiveDocument
		else:
			doc = self.obj.WinwordApplicationObject.ActiveDocument
		endnoteObj = doc.EndNotes[offset]
		endnote = Endnote(self.obj, endnoteObj)
		return endnote.text


	def getFormatFieldSpeech(
			self,
			attrs: textInfos.Field,
			attrsCache: Optional[textInfos.Field] = None,
			formatConfig: Optional[Dict[str, bool]] = None,
			reason: Optional[OutputReason] = None,
			unit: Optional[str] = None,
			extraDetail: bool = False,
			initialFormat: bool = False,
	) -> SpeechSequence:
		import speech
		# from speech import getTableInfoSpeech
		if not formatConfig:
			formatConfig = config.conf["documentFormatting"]
		textList = []
		if formatConfig["reportTables"]:
			tableInfo = attrs.get("table-info")
			oldTableInfo = attrsCache.get("table-info") if attrsCache is not None else None
			tableSequence = speech.getTableInfoSpeech(
				tableInfo, oldTableInfo, extraDetail=extraDetail
			)
			if tableSequence:
				textList.extend(tableSequence)

		if formatConfig["reportPage"]:
			pageNumber = attrs.get("page-number")
			oldPageNumber = attrsCache.get("page-number") if attrsCache is not None else None
			if pageNumber and pageNumber != oldPageNumber:
				# Translators: Indicates the page number in a document.
				# %s will be replaced with the page number.
				text = NVDAString("page %s") % pageNumber
				textList.append(text)
			sectionNumber = attrs.get("section-number")
			oldSectionNumber = attrsCache.get("section-number") if attrsCache is not None else None
			if sectionNumber and sectionNumber != oldSectionNumber:
				# Translators: Indicates the section number in a document.
				# %s will be replaced with the section number.
				text = NVDAString("section %s") % sectionNumber
				textList.append(text)

			textColumnCount = attrs.get("text-column-count")
			oldTextColumnCount = attrsCache.get("text-column-count") if attrsCache is not None else None
			textColumnNumber = attrs.get("text-column-number")
			oldTextColumnNumber = attrsCache.get("text-column-number") if attrsCache is not None else None

			# Because we do not want to report the number of columns when a document is just opened and there is only
			# one column. This would be verbose, in the standard case.
			# column number has changed, or the columnCount has changed
			# but not if the columnCount is 1 or less and there is no old columnCount.
			if (
				(
					(textColumnNumber and textColumnNumber != oldTextColumnNumber)
					or (textColumnCount and textColumnCount != oldTextColumnCount))
				and not (
					textColumnCount and int(textColumnCount) <= 1
					and oldTextColumnCount is None)
				):
				if textColumnNumber and textColumnCount:
					# Translators: Indicates the text column number in a document.
					# {0} will be replaced with the text column number.
					# {1} will be replaced with the number of text columns.
					text = NVDAString("column {0} of {1}").format(textColumnNumber, textColumnCount)
					textList.append(text)
				elif textColumnCount:
					# Translators: Indicates the text column number in a document.
					# %s will be replaced with the number of text columns.
					text = NVDAString("%s columns") % (textColumnCount)
					textList.append(text)

		sectionBreakType = attrs.get("section-break")
		if sectionBreakType:
			if sectionBreakType == "0":  # Continuous section break.
				text = NVDAString("continuous section break")
			elif sectionBreakType == "1":  # New column section break.
				text = NVDAString("new column section break")
			elif sectionBreakType == "2":  # New page section break.
				text = NVDAString("new page section break")
			elif sectionBreakType == "3":  # Even pages section break.
				text = NVDAString("even pages section break")
			elif sectionBreakType == "4":  # Odd pages section break.
				text = NVDAString("odd pages section break")
			else:
				text = ""
			textList.append(text)
		columnBreakType = attrs.get("column-break")
		if columnBreakType:
			textList.append(NVDAString("column break"))
		if formatConfig["reportHeadings"]:
			headingLevel = attrs.get("heading-level")
			oldHeadingLevel = attrsCache.get("heading-level") if attrsCache is not None else None
			# headings should be spoken not only if they change, but also when beginning to speak lines or paragraphs
			# Ensuring a similar experience to if a heading was a controlField
			if headingLevel and (initialFormat and (reason == controlTypes.REASON_FOCUS or unit in (textInfos.UNIT_LINE, textInfos.UNIT_PARAGRAPH)) or headingLevel != oldHeadingLevel):
				# Translators: Speaks the heading level (example output: heading level 2).
				text = NVDAString("heading level %d") % headingLevel
				textList.append(text)

		if formatConfig["reportStyle"]:
			style = attrs.get("style")
			oldStyle = attrsCache.get("style") if attrsCache is not None else None
			if style != oldStyle:
				if style:
					# Translators: Indicates the style of text.
					# A style is a collection of formatting settings and depends on the application.
					# %s will be replaced with the name of the style.
					text = NVDAString("style %s") % style
				else:
					# Translators: Indicates that text has reverted to the default style.
					# A style is a collection of formatting settings and depends on the application.
					text = NVDAString("default style")
				textList.append(text)
		if formatConfig["reportBorderStyle"]:
			borderStyle = attrs.get("border-style")
			oldBorderStyle = attrsCache.get("border-style") if attrsCache is not None else None
			if borderStyle != oldBorderStyle:
				if borderStyle:
					text = borderStyle
				else:
					# Translators: Indicates that cell does not have border lines.
					text = NVDAString("no border lines")
				textList.append(text)
		if formatConfig["reportFontName"]:
			fontFamily = attrs.get("font-family")
			oldFontFamily = attrsCache.get("font-family") if attrsCache is not None else None
			if fontFamily and fontFamily != oldFontFamily:
				textList.append(fontFamily)
			fontName = attrs.get("font-name")
			oldFontName = attrsCache.get("font-name") if attrsCache is not None else None
			if fontName and fontName != oldFontName:
				textList.append(fontName)
		if formatConfig["reportFontSize"]:
			fontSize = attrs.get("font-size")
			oldFontSize = attrsCache.get("font-size") if attrsCache is not None else None
			if fontSize and fontSize != oldFontSize:
				textList.append(fontSize)
		if formatConfig["reportColor"]:
			color = attrs.get("color")
			oldColor = attrsCache.get("color") if attrsCache is not None else None
			backgroundColor = attrs.get("background-color")
			oldBackgroundColor = attrsCache.get("background-color") if attrsCache is not None else None
			backgroundColor2 = attrs.get("background-color2")
			oldBackgroundColor2 = attrsCache.get("background-color2") if attrsCache is not None else None
			bgColorChanged = backgroundColor != oldBackgroundColor or backgroundColor2 != oldBackgroundColor2
			bgColorText = backgroundColor.name if isinstance(backgroundColor, colors.RGB) else backgroundColor
			if backgroundColor2:
				bg2Name = backgroundColor2.name if isinstance(backgroundColor2, colors.RGB) else backgroundColor2
				# Translators: Reported when there are two background colors.
				# This occurs when, for example, a gradient pattern is applied to a spreadsheet cell.
				# {color1} will be replaced with the first background color.
				# {color2} will be replaced with the second background color.
				bgColorText = NVDAString("{color1} to {color2}").format(color1=bgColorText, color2=bg2Name)
			if color and backgroundColor and color != oldColor and bgColorChanged:
				# Translators: Reported when both the text and background colors change.
				# {color} will be replaced with the text color.
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{color} on {backgroundColor}").format(
					color=color.name if isinstance(color, colors.RGB) else color,
					backgroundColor=bgColorText))
			elif color and color != oldColor:
				# Translators: Reported when the text color changes (but not the background color).
				# {color} will be replaced with the text color.
				textList.append(NVDAString("{color}").format(
					color=color.name if isinstance(color, colors.RGB) else color))
			elif backgroundColor and bgColorChanged:
				# Translators: Reported when the background color changes (but not the text color).
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{backgroundColor} background").format(backgroundColor=bgColorText))
			backgroundPattern = attrs.get("background-pattern")
			oldBackgroundPattern = attrsCache.get("background-pattern") if attrsCache is not None else None
			if backgroundPattern and backgroundPattern != oldBackgroundPattern:
				textList.append(NVDAString("background pattern {pattern}").format(pattern=backgroundPattern))

		if formatConfig["reportLineNumber"]:
			lineNumber = attrs.get("line-number")
			oldLineNumber = attrsCache.get("line-number") if attrsCache is not None else None
			if lineNumber is not None and lineNumber != oldLineNumber:
				# Translators: Indicates the line number of the text.
				# %s will be replaced with the line number.
				text = NVDAString("line %s") % lineNumber
				textList.append(text)
		if formatConfig["reportRevisions"]:
			# Insertion
			revision = attrs.get("revision-insertion")
			oldRevision = attrsCache.get("revision-insertion") if attrsCache is not None else None
			if (revision or oldRevision is not None) and revision != oldRevision:
				# Translators: Reported when text is marked as having been inserted
				text = (
					NVDAString("inserted") if revision
					# Translators: Reported when text is no longer marked as having been inserted.
					else NVDAString("not inserted"))
				textList.append(text)
			revision = attrs.get("revision-deletion")
			oldRevision = attrsCache.get("revision-deletion") if attrsCache is not None else None
			if (revision or oldRevision is not None) and revision != oldRevision:
				# Translators: Reported when text is marked as having been deleted
				text = (
					NVDAString("deleted") if revision
					# Translators: Reported when text is no longer marked as having been deleted.
					else NVDAString("not deleted"))
				textList.append(text)
			revision = attrs.get("revision")
			oldRevision = attrsCache.get("revision") if attrsCache is not None else None
			if (revision or oldRevision is not None) and revision != oldRevision:
				if revision:
					# Translators: Reported when text is revised.
					text = NVDAString("revised %s") % revision
				else:
					# Translators: Reported when text is not revised.
					text = NVDAString("no revised %s") % oldRevision
				textList.append(text)
		if formatConfig["reportEmphasis"]:
			# marked text
			marked = attrs.get("marked")
			oldMarked = attrsCache.get("marked") if attrsCache is not None else None
			if (marked or oldMarked is not None) and marked != oldMarked:
				# Translators: Reported when text is marked
				text = (
					NVDAString("marked") if marked
					# Translators: Reported when text is no longer marked
					else NVDAString("not marked"))
				textList.append(text)
			# strong text
			strong = attrs.get("strong")
			oldStrong = attrsCache.get("strong") if attrsCache is not None else None
			if (strong or oldStrong is not None) and strong != oldStrong:
				# Translators: Reported when text is marked as strong (e.g. bold)
				text = (
					NVDAString("strong") if strong
					# Translators: Reported when text is no longer marked as strong (e.g. bold)
					else NVDAString("not strong"))
				textList.append(text)
			# emphasised text
			emphasised = attrs.get("emphasised")
			oldEmphasised = attrsCache.get("emphasised") if attrsCache is not None else None
			if (emphasised or oldEmphasised is not None) and emphasised != oldEmphasised:
				# Translators: Reported when text is marked as emphasised
				text = (
					NVDAString("emphasised") if emphasised
					# Translators: Reported when text is no longer marked as emphasised
					else NVDAString("not emphasised"))
				textList.append(text)

		if formatConfig["reportFontAttributes"]:
			bold = attrs.get("bold")
			oldBold = attrsCache.get("bold") if attrsCache is not None else None
			if (bold or oldBold is not None) and bold != oldBold:
				# Translators: Reported when text is bolded.
				text = (
					NVDAString("bold") if bold
					# Translators: Reported when text is not bolded.
					else NVDAString("no bold"))
				textList.append(text)
			italic = attrs.get("italic")
			oldItalic = attrsCache.get("italic") if attrsCache is not None else None
			if (italic or oldItalic is not None) and italic != oldItalic:
				# Translators: Reported when text is italicized.
				text = (
					NVDAString("italic") if italic
					# Translators: Reported when text is not italicized.
					else NVDAString("no italic"))
				textList.append(text)
			strikethrough = attrs.get("strikethrough")
			oldStrikethrough = attrsCache.get("strikethrough") if attrsCache is not None else None
			if (strikethrough or oldStrikethrough is not None) and strikethrough != oldStrikethrough:
				if strikethrough:
					# Translators: Reported when text is formatted with double strikethrough.
					# See http://en.wikipedia.org/wiki/Strikethrough
					text = (
						NVDAString("double strikethrough") if strikethrough == "double"
						# Translators: Reported when text is formatted with strikethrough.
						# See http://en.wikipedia.org/wiki/Strikethrough
						else NVDAString("strikethrough"))
				else:
					# Translators: Reported when text is formatted without strikethrough.
					# See http://en.wikipedia.org/wiki/Strikethrough
					text = NVDAString("no strikethrough")
				textList.append(text)
			underline = attrs.get("underline")
			oldUnderline = attrsCache.get("underline") if attrsCache is not None else None
			if (underline or oldUnderline is not None) and underline != oldUnderline:
				# Translators: Reported when text is underlined.
				text = (
					NVDAString("underlined") if underline
					# Translators: Reported when text is not underlined.
					else NVDAString("not underlined"))
				textList.append(text)
			hidden = attrs.get("hidden")
			oldHidden = attrsCache.get("hidden") if attrsCache is not None else None
			if (hidden or oldHidden is not None) and hidden != oldHidden:
				text = (
					# Translators: Reported when text is hidden.
					NVDAString("hidden")if hidden
					# Translators: Reported when text is not hidden.
					else NVDAString("not hidden")
				)
				textList.append(text)
		try:
			# for NVDA version > 2019.3
			reportSuperscriptsAndSubscripts = formatConfig["reportSuperscriptsAndSubscripts"]
		except:  # noqa:E722
			reportSuperscriptsAndSubscripts = True
		# if formatConfig["reportSuperscriptsAndSubscripts"]:
		if reportSuperscriptsAndSubscripts:
			textPosition = attrs.get("text-position")
			oldTextPosition = attrsCache.get("text-position") if attrsCache is not None else None
			if (textPosition or oldTextPosition is not None) and textPosition != oldTextPosition:
				textPosition = textPosition.lower() if textPosition else textPosition
				if textPosition == "super":
					# Translators: Reported for superscript text.
					text = NVDAString("superscript")
				elif textPosition == "sub":
					# Translators: Reported for subscript text.
					text = NVDAString("subscript")
				else:
					# Translators: Reported for text which is at the baseline position;
					# i.e. not superscript or subscript.
					text = NVDAString("baseline")
				textList.append(text)
		if formatConfig["reportAlignment"]:
			textAlign = attrs.get("text-align")
			oldTextAlign = attrsCache.get("text-align") if attrsCache is not None else None
			if (textAlign or oldTextAlign is not None) and textAlign != oldTextAlign:
				textAlign = textAlign.lower() if textAlign else textAlign
				if textAlign == "left":
					# Translators: Reported when text is left-aligned.
					text = NVDAString("align left")
				elif textAlign == "center":
					# Translators: Reported when text is centered.
					text = NVDAString("align center")
				elif textAlign == "right":
					# Translators: Reported when text is right-aligned.
					text = NVDAString("align right")
				elif textAlign == "justify":
					# Translators: Reported when text is justified.
					# See http://en.wikipedia.org/wiki/Typographic_alignment#Justified
					text = NVDAString("align justify")
				elif textAlign == "distribute":
					# Translators: Reported when text is justified with character spacing (Japanese etc)
					# See http://kohei.us/2010/01/21/distributed-text-justification/
					text = NVDAString("align distributed")
				else:
					# Translators: Reported when text has reverted to default alignment.
					text = NVDAString("align default")
				textList.append(text)
			verticalAlign = attrs.get("vertical-align")
			oldverticalAlign = attrsCache.get("vertical-align") if attrsCache is not None else None
			if (verticalAlign or oldverticalAlign is not None) and verticalAlign != oldverticalAlign:
				verticalAlign = verticalAlign.lower() if verticalAlign else verticalAlign
				if verticalAlign == "top":
					# Translators: Reported when text is vertically top-aligned.
					text = NVDAString("vertical align top")
				elif verticalAlign in ("center", "middle"):
					# Translators: Reported when text is vertically middle aligned.
					text = NVDAString("vertical align middle")
				elif verticalAlign == "bottom":
					# Translators: Reported when text is vertically bottom-aligned.
					text = NVDAString("vertical align bottom")
				elif verticalAlign == "baseline":
					# Translators: Reported when text is vertically aligned on the baseline.
					text = NVDAString("vertical align baseline")
				elif verticalAlign == "justify":
					# Translators: Reported when text is vertically justified.
					text = NVDAString("vertical align justified")
				elif verticalAlign == "distributed":
					# Translators: Reported when text is vertically justified but with character spacing (For some Asian content).
					text = NVDAString("vertical align distributed")
				else:
					# Translators: Reported when text has reverted to default vertical alignment.
					text = NVDAString("vertical align default")
				textList.append(text)
		if formatConfig["reportParagraphIndentation"]:
			indentLabels = {
				'left-indent': (
					# Translators: the label for paragraph format left indent
					NVDAString("left indent"),
					# Translators: the message when there is no paragraph format left indent
					NVDAString("no left indent"),
				),
				'right-indent': (
					# Translators: the label for paragraph format right indent
					NVDAString("right indent"),
					# Translators: the message when there is no paragraph format right indent
					NVDAString("no right indent"),
				),
				'hanging-indent': (
					# Translators: the label for paragraph format hanging indent
					NVDAString("hanging indent"),
					# Translators: the message when there is no paragraph format hanging indent
					NVDAString("no hanging indent"),
				),
				'first-line-indent': (
					# Translators: the label for paragraph format first line indent
					NVDAString("first line indent"),
					# Translators: the message when there is no paragraph format first line indent
					NVDAString("no first line indent"),
				),
			}
			for attr, (label, noVal) in indentLabels.items():
				newVal = attrs.get(attr)
				oldVal = attrsCache.get(attr) if attrsCache else None
				if (newVal or oldVal is not None) and newVal != oldVal:
					if newVal:
						textList.append(u"%s %s" % (label, newVal))
					else:
						textList.append(noVal)
		if formatConfig["reportLineSpacing"]:
			lineSpacing = attrs.get("line-spacing")
			oldLineSpacing = attrsCache.get("line-spacing") if attrsCache is not None else None
			if (lineSpacing or oldLineSpacing is not None) and lineSpacing != oldLineSpacing:
				# Translators: a type of line spacing (E.g. single line spacing)
				textList.append(NVDAString("line spacing %s") % lineSpacing)
		if formatConfig["reportLinks"]:
			link = attrs.get("link")
			oldLink = attrsCache.get("link") if attrsCache is not None else None
			if (link or oldLink is not None) and link != oldLink:
				text = NVDAString("link") if link else NVDAString("out of %s") % NVDAString("link")
				textList.append(text)
		if formatConfig["reportComments"]:
			comment = attrs.get("comment")
			oldComment = attrsCache.get("comment") if attrsCache is not None else None
			if (comment or oldComment is not None) and comment != oldComment:
				if comment:
					textList.extend(self.getCommentFormatFieldSpeech(comment))
				elif extraDetail:
					# Translators: Reported when text no longer contains a comment.
					text = NVDAString("out of comment")
					textList.append(text)

		if formatConfig["reportSpellingErrors"]:
			invalidSpelling = attrs.get("invalid-spelling")
			oldInvalidSpelling = attrsCache.get("invalid-spelling") if attrsCache is not None else None
			if (invalidSpelling or oldInvalidSpelling is not None) and invalidSpelling != oldInvalidSpelling:
				if invalidSpelling:
					# Translators: Reported when text contains a spelling error.
					text = NVDAString("spelling error")
				elif extraDetail:
					# Translators: Reported when moving out of text containing a spelling error.
					text = NVDAString("out of spelling error")
				else:
					text = ""
				if text:
					textList.append(text)
			invalidGrammar = attrs.get("invalid-grammar")
			oldInvalidGrammar = attrsCache.get("invalid-grammar") if attrsCache is not None else None
			if (invalidGrammar or oldInvalidGrammar is not None) and invalidGrammar != oldInvalidGrammar:
				if invalidGrammar:
					# Translators: Reported when text contains a grammar error.
					text = NVDAString("grammar error")
				elif extraDetail:
					# Translators: Reported when moving out of text containing a grammar error.
					text = NVDAString("out of grammar error")
				else:
					text = ""
				if text:
					textList.append(text)
		# The line-prefix formatField attribute contains the text for a bullet or number for a list item, when the bullet or number does not appear in the actual text content.
		# Normally this attribute could be repeated across formatFields within a list item and therefore is not safe to speak when the unit is word or character.
		# However, some implementations (such as MS Word with UIA) do limit its useage to the very first formatField of the list item.
		# Therefore, they also expose a line-prefix_speakAlways attribute to allow its usage for any unit.
		linePrefix_speakAlways = attrs.get('line-prefix_speakAlways', False)
		if linePrefix_speakAlways or unit in (textInfos.UNIT_LINE, textInfos.UNIT_SENTENCE, textInfos.UNIT_PARAGRAPH, textInfos.UNIT_READINGCHUNK):
			linePrefix = attrs.get("line-prefix")
			if linePrefix:
				textList.append(linePrefix)
		if attrsCache is not None:
			attrsCache.clear()
			attrsCache.update(attrs)
		# types.logBadSequenceTypes(textList)

		return textList

	def getControlFieldSpeech(
			self,
			attrs: textInfos.ControlField,
			ancestorAttrs: List[textInfos.Field],
			fieldType: str,
			formatConfig: Optional[Dict[str, bool]] = None,
			extraDetail: bool = False,
			reason: Optional[OutputReason] = None
	) -> SpeechSequence:

		if attrs.get('isHidden'):
			return []
		if not formatConfig:
			formatConfig = config.conf["documentFormatting"]
		presCat = attrs.getPresentationCategory(ancestorAttrs, formatConfig, reason=reason)
		childControlCount = int(attrs.get('_childcontrolcount', "0"))
		role = attrs.get('role', controlTypes.ROLE_UNKNOWN)
		if (
			reason == controlTypes.REASON_FOCUS
			or attrs.get('alwaysReportName', False)
		):
			name = attrs.get('name', "")
		else:
			name = ""
		states = attrs.get('states', set())
		keyboardShortcut = attrs.get('keyboardShortcut', "")
		ariaCurrent = attrs.get('current', None)
		placeholderValue = attrs.get('placeholder', None)
		value = attrs.get('value', "")
		if reason == controlTypes.REASON_FOCUS or attrs.get('alwaysReportDescription', False):
			description = attrs.get('description', "")
		else:
			description = ""
		level = attrs.get('level', None)

		if presCat != attrs.PRESCAT_LAYOUT:
			tableID = attrs.get("table-id")
		else:
			tableID = None

		roleText = attrs.get('roleText')
		landmark = attrs.get("landmark")
		if roleText:
			roleTextSequence = [roleText, ]
		elif role == controlTypes.ROLE_LANDMARK and landmark:
			roleTextSequence = [
				f"{aria.landmarkRoles[landmark]} {controlTypes.roleLabels[controlTypes.ROLE_LANDMARK]}",
			]
		else:
			roleTextSequence = getPropertiesSpeech(reason=reason, role=role)
		stateTextSequence = getPropertiesSpeech(reason=reason, states=states, _role=role)
		keyboardShortcutSequence = []
		if config.conf["presentation"]["reportKeyboardShortcuts"]:
			keyboardShortcutSequence = getPropertiesSpeech(
				reason=reason, keyboardShortcut=keyboardShortcut
			)
		ariaCurrentSequence = getPropertiesSpeech(reason=reason, current=ariaCurrent)
		placeholderSequence = getPropertiesSpeech(reason=reason, placeholder=placeholderValue)
		nameSequence = getPropertiesSpeech(reason=reason, name=name)
		valueSequence = getPropertiesSpeech(reason=reason, value=value)
		if role == controlTypes.ROLE_FOOTNOTE and _addonConfigManager.toggleAutomaticReadingOption(False) and _addonConfigManager.toggleAutoFootnoteReadingOption(False):
			footnote = self.getFootnote(value)
			if footnote != "":
				valueSequence = getNotePropertiesSpeech(reason, value, footnote)   # "%s (%s)" % (value, footnote))
		if role == controlTypes.ROLE_ENDNOTE and _addonConfigManager.toggleAutomaticReadingOption(False) and _addonConfigManager.toggleAutoEndnoteReadingOption(False):
			endnote = self.getEndNote(value)
			if endnote != "":
				valueSequence = getNotePropertiesSpeech(reason, value, endnote)
		descriptionSequence = []
		if config.conf["presentation"]["reportObjectDescriptions"]:
			descriptionSequence = getPropertiesSpeech(
				reason=reason, description=description
			)
		levelSequence = getPropertiesSpeech(reason=reason, positionInfo_level=level)

		# Determine under what circumstances this node should be spoken.
		# speakEntry: Speak when the user enters the control.
		# speakWithinForLine: When moving by line, speak when the user is already within the control.
		# speakExitForLine: When moving by line, speak when the user exits the control.
		# speakExitForOther: When moving by word or character, speak when the user exits the control.
		speakEntry = speakWithinForLine = speakExitForLine = speakExitForOther = False
		if presCat == attrs.PRESCAT_SINGLELINE:
			speakEntry = True
			speakWithinForLine = True
			speakExitForOther = True
		elif presCat in (attrs.PRESCAT_MARKER, attrs.PRESCAT_CELL):
			speakEntry = True
		elif presCat == attrs.PRESCAT_CONTAINER:
			speakEntry = True
			speakExitForLine = bool(
				attrs.get('roleText')
				or role != controlTypes.ROLE_LANDMARK
			)
			speakExitForOther = True

		# Determine the order of speech.
		# speakContentFirst: Speak the content before the control field info.
		speakContentFirst = (
			reason == controlTypes.REASON_FOCUS
			and presCat != attrs.PRESCAT_CONTAINER
			and role not in (
				controlTypes.ROLE_EDITABLETEXT,
				controlTypes.ROLE_COMBOBOX,
				controlTypes.ROLE_TREEVIEW,
				controlTypes.ROLE_LIST,
				controlTypes.ROLE_LANDMARK,
				controlTypes.ROLE_REGION,
			)
			and not tableID
			and controlTypes.STATE_EDITABLE not in states
		)
		# speakStatesFirst: Speak the states before the role.
		speakStatesFirst = role == controlTypes.ROLE_LINK

		containerContainsText = ""  #: used for item counts for lists

		# Determine what text to speak.
		# Special cases
		if(
			childControlCount
			and fieldType == "start_addedToControlFieldStack"
			and role == controlTypes.ROLE_LIST
			and controlTypes.STATE_READONLY in states
		):
			# List.
			# #7652: containerContainsText variable is set here, but the actual generation of all other output is
			# handled further down in the general cases section.
			# This ensures that properties such as name, states and level etc still get reported appropriately.
			# Translators: Number of items in a list (example output: list with 5 items).
			containerContainsText = NVDAString("with %s items") % childControlCount
		elif fieldType == "start_addedToControlFieldStack" and role == controlTypes.ROLE_TABLE and tableID:
			# Table.
			rowCount = (attrs.get("table-rowcount-presentational") or attrs.get("table-rowcount"))
			columnCount = (attrs.get("table-columncount-presentational") or attrs.get("table-columncount"))
			tableSeq = nameSequence[:]
			tableSeq.extend(roleTextSequence)
			tableSeq.extend(stateTextSequence)
			tableSeq.extend(
				getPropertiesSpeech(
					_tableID=tableID,
					rowCount=rowCount,
					columnCount=columnCount))
			tableSeq.extend(levelSequence)
			types.logBadSequenceTypes(tableSeq)
			return tableSeq
		elif (
			nameSequence
			and reason == controlTypes.REASON_FOCUS
			and fieldType == "start_addedToControlFieldStack"
			and role in (controlTypes.ROLE_GROUPING, controlTypes.ROLE_PROPERTYPAGE)
		):
			# #10095, #3321, #709: Report the name and description of groupings (such as fieldsets) and tab pages
			nameAndRole = nameSequence[:]
			nameAndRole.extend(roleTextSequence)
			types.logBadSequenceTypes(nameAndRole)
			return nameAndRole
		elif (
			fieldType in ("start_addedToControlFieldStack", "start_relative")
			and role in (
				controlTypes.ROLE_TABLECELL,
				controlTypes.ROLE_TABLECOLUMNHEADER,
				controlTypes.ROLE_TABLEROWHEADER
			)
			and tableID
		):
			# Table cell.
			reportTableHeaders = formatConfig["reportTableHeaders"]
			reportTableCellCoords = formatConfig["reportTableCellCoords"]
			getProps = {
				'rowNumber': (attrs.get("table-rownumber-presentational") or attrs.get("table-rownumber")),
				'columnNumber': (attrs.get("table-columnnumber-presentational") or attrs.get("table-columnnumber")),
				'rowSpan': attrs.get("table-rowsspanned"),
				'columnSpan': attrs.get("table-columnsspanned"),
				'includeTableCellCoords': reportTableCellCoords
			}
			if reportTableHeaders:
				getProps['rowHeaderText'] = attrs.get("table-rowheadertext")
				getProps['columnHeaderText'] = attrs.get("table-columnheadertext")
			tableCellSequence = getPropertiesSpeech(_tableID=tableID, **getProps)
			tableCellSequence.extend(stateTextSequence)
			tableCellSequence.extend(ariaCurrentSequence)
			types.logBadSequenceTypes(tableCellSequence)
			return tableCellSequence

		# General cases.

		if ((
			speakEntry and ((
				speakContentFirst
				and fieldType in ("end_relative", "end_inControlFieldStack")
			)
			or (
				not speakContentFirst
				and fieldType in ("start_addedToControlFieldStack", "start_relative")
			))
			)
			or (
				speakWithinForLine
				and not speakContentFirst
				and not extraDetail
				and fieldType == "start_inControlFieldStack"
			)):
			out = []
			content = attrs.get("content")
			if content and speakContentFirst:
				out.append(content)
			if placeholderValue:
				if valueSequence:
					log.error(
						f"valueSequence exists when expected none: "
						f"valueSequence: {valueSequence!r} placeholderSequence: {placeholderSequence!r}"
					)
				valueSequence = placeholderSequence

			# Avoid speaking name twice. Which may happen if this controlfield is labelled by
			# one of it's internal fields. We determine this by checking for 'labelledByContent'.
			# An example of this situation is a checkbox element that has aria-labelledby pointing to a child
			# element.
			if (
				# Don't speak name when labelledByContent. It will be spoken by the subsequent controlFields instead.
				attrs.get("IAccessible2::attribute_explicit-name", False)
				and attrs.get("labelledByContent", False)
			):
				log.debug("Skipping name sequence: control field is labelled by content")
			else:
				out.extend(nameSequence)

			out.extend(stateTextSequence if speakStatesFirst else roleTextSequence)
			out.extend(roleTextSequence if speakStatesFirst else stateTextSequence)
			out.append(containerContainsText)
			out.extend(ariaCurrentSequence)
			out.extend(valueSequence)
			out.extend(descriptionSequence)
			out.extend(levelSequence)
			out.extend(keyboardShortcutSequence)
			if content and not speakContentFirst:
				out.append(content)

			types.logBadSequenceTypes(out)
			return out
		elif (
			fieldType in (
				"end_removedFromControlFieldStack",
				"end_relative",
			)
			and roleTextSequence
			and (
				(not extraDetail and speakExitForLine)
				or (extraDetail and speakExitForOther))):
			if all(isinstance(item, str) for item in roleTextSequence):
				joinedRoleText = " ".join(roleTextSequence)
				out = [
					# Translators: Indicates end of something (example output: at the end of a list, speaks out of list).
					NVDAString("out of %s") % joinedRoleText,
				]
			else:
				out = roleTextSequence

			types.logBadSequenceTypes(out)
			return out

		# Special cases
		elif not speakEntry\
			and fieldType in ("start_addedToControlFieldStack", "start_relative"):
			out = []
			if ariaCurrent:
				out.extend(ariaCurrentSequence)
				# types.logBadSequenceTypes(out)
			return out
		else:
			return []


class SetSynthCommand(speech.commands.CallbackCommand):
	def __init__(self, synthName, speechSettings, temporary):
		self.synthName = synthName
		self.speechSettings = speechSettings
		self.temporary = temporary

	def run(self):
		if self.synthName is None:
			return
		_setSynth(self.synthName, self.speechSettings, self.temporary)

	def __repr__(self):
		return "SetSynthCommand (%s)" % self.synthName


def getSynthDisplayInfos(synth, synthConf):
	conf = synthConf
	import driverHandler
	id = "id"
	numericSynthSetting = [driverHandler.NumericDriverSetting, ]
	booleanSynthSetting = [driverHandler.BooleanDriverSetting, ]
	infos = []
	for setting in synth.supportedSettings:
		settingID = getattr(setting, id)
		if type(setting) in numericSynthSetting:
			info = str(conf[settingID])
			infos.append((setting.displayName, info))
		elif type(setting) in booleanSynthSetting:
			info = _("yes") if conf[settingID] else _("no")
			infos.append((setting.displayName, info))
		else:
			if hasattr(synth, "available%ss" % settingID.capitalize()):
				tempList = list(getattr(synth, "available%ss" % settingID.capitalize()).values())
				cur = conf[settingID]
				try:
					# for nvda =>2019.3
					i = [x.id for x in tempList].index(cur)
					v = tempList[i].displayName
				except:  # noqa:E722
					i = [x.ID for x in tempList].index(cur)
					v = tempList[i].name
				info = v
				infos.append((setting.displayName, info))
	d = {}
	for i in range(0, len(infos)):
		item = infos[i]
		d[str(i+1)] = [item[0], item[1]]
	return d


def getCurrentSpeechSettings():
	currentSynth = synthDriverHandler.getSynth()
	d = {}
	d[SCT_Speech] = config.conf[SCT_Speech].dict()
	for key in config.conf[SCT_Speech].copy():
		val = config.conf[SCT_Speech][key]
		if type(val) == config.AggregatedSection and key not in [SCT_Many, currentSynth.name]:
			del d[SCT_Speech][key]
	d["SynthDisplayInfos"] = getSynthDisplayInfos(currentSynth, d[SCT_Speech][currentSynth.name])
	return d


def saveCurrentSpeechSettings():
	autoReadingSynth = {}
	autoReadingSynth["synthName"] = synthDriverHandler.getSynth().name
	autoReadingSynth.update(getCurrentSpeechSettings())
	_addonConfigManager.saveAutoReadingSynthSettings(autoReadingSynth)
	# Translators: message to user to report automatic reading voice record.
	speech.speakMessage(_("Current voice settings have beenset for automatic reading"))


# to memorize the suspended synth settings when automacic reading synth is running
_suspendedSynth = (None, None)


def _setSynth(synthName, speechSettings, temporary=False):
	global _suspendedSynth
	if temporary:
		curSynth = synthDriverHandler.getSynth()
		currentSynthName = curSynth.name
		currentSpeechSettings = getCurrentSpeechSettings()
		_suspendedSynth = (currentSynthName, currentSpeechSettings)
	else:
		_suspendedSynth = (None, None)
	# when changing sapi5 to sapi5, or oneCore to oneCore,
	# if we don't change with another synthetizer, nothing appends.
	# for NVDA version >= 2020.3,  we switch to embedded synthetizer.
	NVDAVersion = [version_year, version_major]
	if NVDAVersion >= [2020, 3]:
		synthDriverHandler.setSynth("espeak")
	else:
		synthDriverHandler.setSynth(None)
	if speechSettings is not None:
		synthSpeechConfig = config.conf[SCT_Speech] .dict()
		synthSpeechConfig.update(speechSettings[SCT_Speech])
		config.conf[SCT_Speech] = synthSpeechConfig.copy()
	synthDriverHandler.setSynth(synthName)


# used for saving current nvda cancelSpeech
# we must hake nvda cancelSpeech to restore previous synth when cancel speech occurs during automatic reading
_NVDACancelSpeech = None


def myCancelSpeech():
	from speech.priorities import SpeechPriority
	global _suspendedSynth
	_NVDACancelSpeech()
	# if we are in automatic reading, we must restore previous synth.
	if _suspendedSynth[0] is None:
		return
	# restore main synth
	_setSynth(_suspendedSynth[0], _suspendedSynth[1], temporary=False)

def initialize():
	global _NVDACancelSpeech
	if _NVDACancelSpeech is not None:
		log.warning("cancelSpeech already patched")
		return
	_NVDACancelSpeech = speech.cancelSpeech
	speech.cancelSpeech = myCancelSpeech


def terminate():
	global _NVDACancelSpeech
	if _NVDACancelSpeech is not None:
		speech.cancelSpeech = _NVDACancelSpeech
		_NVDACancelSpeech = None


_autoReadingSynth = None
