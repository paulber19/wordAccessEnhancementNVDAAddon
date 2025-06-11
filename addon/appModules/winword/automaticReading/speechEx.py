# appModules\winword\automaticReading\__init__.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2025 paulber19, Abdel
# This file is covered by the GNU General Public License.

import addonHandler
# from logHandler import log
from versionInfo import version_year, version_major
import config
import textInfos
from textInfos import SpeechSequence

import colors
from typing import Optional, Dict

from speech import getTableInfoSpeech
from speech.types import logBadSequenceTypes
from config.configFlags import ReportCellBorders
import sys
import os
from controlTypes import OutputReason, TextPosition

_curAddon = addonHandler.getCodeAddon()
sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_NVDAStrings import NVDAString, NVDAString_ngettext
del sys.path[-1]
del sys.modules["ww_NVDAStrings"]

addonHandler.initTranslation()

SCT_Speech = "speech"
SCT_Many = "__many__"


class FormatFieldSpeech(textInfos.TextInfo):
	def getFormatFieldSpeech(
		self,
		*args, **kwargs
	) -> SpeechSequence:
		# same as nvda speech method, but modified to speak comment
		NVDAVersion = [version_year, version_major]
		if NVDAVersion >= [2024, 4]:
			funct = self.getFormatFieldSpeech_2024_4
		elif NVDAVersion >= [2024, 1]:
			funct = self.getFormatFieldSpeech_2024_1
		elif NVDAVersion >= [2023, 3]:
			funct = self.getFormatFieldSpeech_2023_3
		elif NVDAVersion >= [2023, 2]:
			funct = self.getFormatFieldSpeech_2023_2
		else:
			funct = self.getFormatFieldSpeech_2023_1
		seq = funct(
			*args, **kwargs
		)
		return seq

	def getFormatFieldSpeech_2024_4(  # noqa: C901
		self,
		attrs: textInfos.Field,
		attrsCache: Optional[textInfos.Field] = None,
		formatConfig: Optional[Dict[str, bool]] = None,
		reason: Optional[OutputReason] = None,
		unit: Optional[str] = None,
		extraDetail: bool = False,
		initialFormat: bool = False,
	) -> SpeechSequence:
		if not formatConfig:
			formatConfig = config.conf["documentFormatting"]
		from config.configFlags import (
			ReportCellBorders,
			OutputMode,
		)

		textList = []
		if formatConfig["reportTables"]:
			tableInfo = attrs.get("table-info")
			oldTableInfo = attrsCache.get("table-info") if attrsCache is not None else None
			tableSequence = getTableInfoSpeech(
				tableInfo,
				oldTableInfo,
				extraDetail=extraDetail,
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
				(textColumnNumber and textColumnNumber != oldTextColumnNumber)
				or (textColumnCount and textColumnCount != oldTextColumnCount)
			) and not (textColumnCount and int(textColumnCount) <= 1 and oldTextColumnCount == None):  # noqa: E711
				if textColumnNumber and textColumnCount:
					# Translators: Indicates the text column number in a document.
					# {0} will be replaced with the text column number.
					# {1} will be replaced with the number of text columns.
					text = NVDAString("column {0} of {1}") .format(textColumnNumber, textColumnCount)
					textList.append(text)
				elif textColumnCount:
					# Translators: Indicates the text column number in a document.
					# %s will be replaced with the number of text columns.
					text = NVDAString_ngettext("%s column", "%s columns", textColumnCount) % textColumnCount
					textList.append(text)
				elif textColumnNumber:
					# Translators: Indicates the text column number in a document.
					text = NVDAString("column {columnNumber}") .format(columnNumber=textColumnNumber)
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
			if headingLevel and (
				initialFormat
				and (
					reason in [OutputReason.FOCUS, OutputReason.QUICKNAV]
					or unit in (textInfos.UNIT_LINE, textInfos.UNIT_PARAGRAPH)
				)
				or headingLevel != oldHeadingLevel
			):
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
		if formatConfig["reportCellBorders"] != ReportCellBorders.OFF:
			borderStyle = attrs.get("border-style")
			oldBorderStyle = attrsCache.get("border-style") if attrsCache is not None else None
			if (borderStyle or oldBorderStyle is not None) and borderStyle != oldBorderStyle:
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
				bgColorText = NVDAString("{color1} to {color2}") .format(color1=bgColorText, color2=bg2Name)
			if color and backgroundColor and color != oldColor and bgColorChanged:
				textList.append(
					# Translators: Reported when both the text and background colors change.
					# {color} will be replaced with the text color.
					# {backgroundColor} will be replaced with the background color.
					NVDAString("{color} on {backgroundColor}") .format(
						color=color.name if isinstance(color, colors.RGB) else color,
						backgroundColor=bgColorText,
					),
				)
			elif color and color != oldColor:
				# Translators: Reported when the text color changes (but not the background color).
				# {color} will be replaced with the text color.
				textList.append(
					NVDAString("{color}") .format(color=color.name if isinstance(color, colors.RGB) else color))
			elif backgroundColor and bgColorChanged:
				# Translators: Reported when the background color changes (but not the text color).
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{backgroundColor} background") .format(backgroundColor=bgColorText))
			backgroundPattern = attrs.get("background-pattern")
			oldBackgroundPattern = attrsCache.get("background-pattern") if attrsCache is not None else None
			if (
				backgroundPattern or oldBackgroundPattern is not None
			) and backgroundPattern != oldBackgroundPattern:
				if not backgroundPattern:
					# Translators: A type of background pattern in Microsoft Excel.
					# No pattern
					backgroundPattern = NVDAString("none")
				textList.append(NVDAString("background pattern {pattern}") .format(pattern=backgroundPattern))
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
				text = (
					# Translators: Reported when text is marked as having been inserted
					NVDAString("inserted")
					if revision
					# Translators: Reported when text is no longer marked as having been inserted.
					else NVDAString("not inserted")
				)
				textList.append(text)
			revision = attrs.get("revision-deletion")
			oldRevision = attrsCache.get("revision-deletion") if attrsCache is not None else None
			if (revision or oldRevision is not None) and revision != oldRevision:
				text = (
					# Translators: Reported when text is marked as having been deleted
					NVDAString("deleted")
					if revision
					# Translators: Reported when text is no longer marked as having been  deleted.
					else NVDAString("not deleted")
				)
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
		if formatConfig["reportHighlight"]:
			# marked text
			marked = attrs.get("marked")
			oldMarked = attrsCache.get("marked") if attrsCache is not None else None
			if (marked or oldMarked is not None) and marked != oldMarked:
				text = (
					# Translators: Reported when text is marked
					NVDAString("marked")
					if marked
					# Translators: Reported when text is no longer marked
					else NVDAString("not marked")
				)
				textList.append(text)
			# color-highlighted text in Word
			hlColor = attrs.get("highlight-color")
			oldHlColor = attrsCache.get("highlight-color") if attrsCache is not None else None
			if (hlColor or oldHlColor is not None) and hlColor != oldHlColor:
				colorName = hlColor.name if isinstance(hlColor, colors.RGB) else hlColor
				text = (
					# Translators: Reported when text is color-highlighted
					NVDAString("highlighted in {color}") .format(color=colorName)
					if hlColor
					# Translators: Reported when text is no longer marked
					else NVDAString("not highlighted")
				)
				textList.append(text)
		if formatConfig["reportEmphasis"]:
			# strong text
			strong = attrs.get("strong")
			oldStrong = attrsCache.get("strong") if attrsCache is not None else None
			if (strong or oldStrong is not None) and strong != oldStrong:
				text = (
					# Translators: Reported when text is marked as strong (e.g. bold)
					NVDAString("strong")
					if strong
					# Translators: Reported when text is no longer marked as strong (e.g. bold)
					else NVDAString("not strong")
				)
				textList.append(text)
			# emphasised text
			emphasised = attrs.get("emphasised")
			oldEmphasised = attrsCache.get("emphasised") if attrsCache is not None else None
			if (emphasised or oldEmphasised is not None) and emphasised != oldEmphasised:
				text = (
					# Translators: Reported when text is marked as emphasised
					NVDAString("emphasised")
					if emphasised
					# Translators: Reported when text is no longer marked as emphasised
					else NVDAString("not emphasised")
				)
				textList.append(text)
		if formatConfig["fontAttributeReporting"] & OutputMode.SPEECH:
			bold = attrs.get("bold")
			oldBold = attrsCache.get("bold") if attrsCache is not None else None
			if (bold or oldBold is not None) and bold != oldBold:
				text = (
					# Translators: Reported when text is bolded.
					NVDAString("bold")
					if bold
					# Translators: Reported when text is not bolded.
					else NVDAString("no bold")
				)
				textList.append(text)
			italic = attrs.get("italic")
			oldItalic = attrsCache.get("italic") if attrsCache is not None else None
			if (italic or oldItalic is not None) and italic != oldItalic:
				# Translators: Reported when text is italicized.
				text = (
					NVDAString("italic")
					if italic
					# Translators: Reported when text is not italicized.
					else NVDAString("no italic")
				)
				textList.append(text)
			strikethrough = attrs.get("strikethrough")
			oldStrikethrough = attrsCache.get("strikethrough") if attrsCache is not None else None
			if (strikethrough or oldStrikethrough is not None) and strikethrough != oldStrikethrough:
				if strikethrough:
					text = (
						# Translators: Reported when text is formatted with double strikethrough.
						# See http://en.wikipedia.org/wiki/Strikethrough
						NVDAString("double strikethrough")
						if strikethrough == "double"
						# Translators: Reported when text is formatted with strikethrough.
						# See http://en.wikipedia.org/wiki/Strikethrough
						else NVDAString("strikethrough")
					)
				else:
					# Translators: Reported when text is formatted without strikethrough.
					# See http://en.wikipedia.org/wiki/Strikethrough
					text = NVDAString("no strikethrough")
				textList.append(text)
			underline = attrs.get("underline")
			oldUnderline = attrsCache.get("underline") if attrsCache is not None else None
			if (underline or oldUnderline is not None) and underline != oldUnderline:
				text = (
					# Translators: Reported when text is underlined.
					NVDAString("underlined")
					if underline
					# Translators: Reported when text is not underlined.
					else NVDAString("not underlined")
				)
				textList.append(text)
			hidden = attrs.get("hidden")
			oldHidden = attrsCache.get("hidden") if attrsCache is not None else None
			if (hidden or oldHidden is not None) and hidden != oldHidden:
				text = (
					# Translators: Reported when text is hidden.
					NVDAString("hidden")
					if hidden
					# Translators: Reported when text is not hidden.
					else NVDAString("not hidden")
				)
				textList.append(text)
		if formatConfig["reportSuperscriptsAndSubscripts"]:
			textPosition = attrs.get("text-position", TextPosition.UNDEFINED)
			attrs["text-position"] = textPosition
			oldTextPosition = attrsCache.get("text-position") if attrsCache is not None else None
			if textPosition != oldTextPosition and (
				textPosition in [TextPosition.SUPERSCRIPT, TextPosition.SUBSCRIPT]
				or (
					textPosition == TextPosition.BASELINE
					and (oldTextPosition is not None and oldTextPosition != TextPosition.UNDEFINED)
				)
			):
				textList.append(textPosition.displayString)
		if formatConfig["reportAlignment"]:
			textAlign = attrs.get("text-align")
			oldTextAlign = attrsCache.get("text-align") if attrsCache is not None else None
			if textAlign and textAlign != oldTextAlign:
				textList.append(textAlign.displayString)
			verticalAlign = attrs.get("vertical-align")
			oldVerticalAlign = attrsCache.get("vertical-align") if attrsCache is not None else None
			if verticalAlign and verticalAlign != oldVerticalAlign:
				textList.append(verticalAlign.displayString)
		if formatConfig["reportParagraphIndentation"]:
			indentLabels = {
				"left-indent": (
					# Translators: the label for paragraph format left indent
					NVDAString("left indent"),
					# Translators: the message when there is no paragraph format left indent
					NVDAString("no left indent"),
				),
				"right-indent": (
					# Translators: the label for paragraph format right indent
					NVDAString("right indent"),
					# Translators: the message when there is no paragraph format right indent
					NVDAString("no right indent"),
				),
				"hanging-indent": (
					# Translators: the label for paragraph format hanging indent
					NVDAString("hanging indent"),
					# Translators: the message when there is no paragraph format hanging indent
					NVDAString("no hanging indent"),
				),
				"first-line-indent": (
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
						textList.append("%s %s" % (label, newVal))
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
					# my modification
					textList.extend(self.getCommentFormatFieldSpeech(comment))
					"""
					if comment is textInfos.CommentType.DRAFT:
						# Translators: Reported when text contains a draft comment.
						text = NVDAString("has draft comment")
					elif comment is textInfos.CommentType.RESOLVED:
						# Translators: Reported when text contains a resolved comment.
						text = NVDAString("has resolved comment")
					else:  # generic
						# Translators: Reported when text contains a generic comment.
						text = NVDAString("has comment")
					textList.append(text)
					"""
				elif extraDetail:
					# Translators: Reported when text no longer contains a comment.
					text = NVDAString("out of comment")
					textList.append(text)
		if formatConfig["reportBookmarks"]:
			bookmark = attrs.get("bookmark")
			oldBookmark = attrsCache.get("bookmark") if attrsCache is not None else None
			if (bookmark or oldBookmark is not None) and bookmark != oldBookmark:
				if bookmark:
					# Translators: Reported when text contains a bookmark
					text = NVDAString("bookmark")
					textList.append(text)
				elif extraDetail:
					# Translators: Reported when text no longer contains a bookmark
					text = NVDAString("out of bookmark")
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
		# The line-prefix formatField attribute contains the text for a bullet or number for a list item,
					# when the bullet or number does not appear in the actual text content.
		# Normally this attribute could be repeated across formatFields within a list item
					# and therefore is not safe to speak when the unit is word or character.
		# However, some implementations (such as MS Word with UIA)
					# do limit its useage to the very first formatField of the list item.
		# Therefore, they also expose a line-prefix_speakAlways attribute
					# to allow its usage for any unit.
		linePrefix_speakAlways = attrs.get("line-prefix_speakAlways", False)
		if linePrefix_speakAlways or unit in (
			textInfos.UNIT_LINE,
			textInfos.UNIT_SENTENCE,
			textInfos.UNIT_PARAGRAPH,
			textInfos.UNIT_READINGCHUNK,
		):
			linePrefix = attrs.get("line-prefix")
			if linePrefix:
				textList.append(linePrefix)
		if attrsCache is not None:
			attrsCache.clear()
			attrsCache.update(attrs)
		logBadSequenceTypes(textList)
		return textList

	def getFormatFieldSpeech_2024_1(  # noqa: C901
		self,
		attrs: textInfos.Field,
		attrsCache: Optional[textInfos.Field] = None,
		formatConfig: Optional[Dict[str, bool]] = None,
		reason: Optional[OutputReason] = None,
		unit: Optional[str] = None,
		extraDetail: bool = False,
		initialFormat: bool = False,
	) -> SpeechSequence:
		if not formatConfig:
			formatConfig = config.conf["documentFormatting"]
		textList = []
		if formatConfig["reportTables"]:
			tableInfo = attrs.get("table-info")
			oldTableInfo = attrsCache.get("table-info") if attrsCache is not None else None
			tableSequence = getTableInfoSpeech(
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
			#  Because we do not want to report the number of columns when a document is just opened and there is only
			#  one column. This would be verbose, in the standard case.
			#  column number has changed, or the columnCount has changed
			#  but not if the columnCount is 1 or less and there is no old columnCount.
			if (((
				textColumnNumber and textColumnNumber != oldTextColumnNumber)
				or (textColumnCount and textColumnCount != oldTextColumnCount))
				and not (textColumnCount and int(textColumnCount) <= 1 and oldTextColumnCount is None)):
				if textColumnNumber and textColumnCount:
					#  Translators: Indicates the text column number in a document.
					#  {0} will be replaced with the text column number.
					#  {1} will be replaced with the number of text columns.
					text = NVDAString("column {0} of {1}") .format(textColumnNumber, textColumnCount)
					textList.append(text)
				elif textColumnCount:
					# Translators: Indicates the text column number in a document.
					# %s will be replaced with the number of text columns.
					text = NVDAString_ngettext("%s column", "%s columns", textColumnCount) % textColumnCount
					textList.append(text)
				elif textColumnNumber:
					# Translators: Indicates the text column number in a document.
					text = NVDAString("column {columnNumber}") .format(columnNumber=textColumnNumber)
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
			if (
				headingLevel
				and (
					initialFormat
					and (
						reason in [OutputReason.FOCUS, OutputReason.QUICKNAV]
						or unit in (textInfos.UNIT_LINE, textInfos.UNIT_PARAGRAPH)
					)
					or headingLevel != oldHeadingLevel
				)
			):
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
		if formatConfig["reportCellBorders"] != ReportCellBorders.OFF:
			borderStyle = attrs.get("border-style")
			oldBorderStyle = attrsCache.get("border-style") if attrsCache is not None else None
			if (borderStyle or oldBorderStyle is not None) and borderStyle != oldBorderStyle:
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
				bgColorText = NVDAString("{color1} to {color2}") .format(color1=bgColorText, color2=bg2Name)
			if color and backgroundColor and color != oldColor and bgColorChanged:
				# Translators: Reported when both the text and background colors change.
				# {color} will be replaced with the text color.
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{color} on {backgroundColor}") .format(
					color=color.name if isinstance(color, colors.RGB) else color,
					backgroundColor=bgColorText))
			elif color and color != oldColor:
				# Translators: Reported when the text color changes (but not the background color).
				# {color} will be replaced with the text color.
				textList.append(NVDAString(
					"{color}") .format(color=color.name if isinstance(color, colors.RGB) else color))
			elif backgroundColor and bgColorChanged:
				# Translators: Reported when the background color changes (but not the text color).
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{backgroundColor} background") .format(backgroundColor=bgColorText))
			backgroundPattern = attrs.get("background-pattern")
			oldBackgroundPattern = attrsCache.get("background-pattern") if attrsCache is not None else None
			if (backgroundPattern or oldBackgroundPattern is not None) and backgroundPattern != oldBackgroundPattern:
				if not backgroundPattern:
					# Translators: A type of background pattern in Microsoft Excel.
					# No pattern
					backgroundPattern = NVDAString("none")
				textList.append(NVDAString("background pattern {pattern}") .format(pattern=backgroundPattern))
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
					# Translators: Reported when text is no longer marked as having been  deleted.
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
		if formatConfig["reportHighlight"]:
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
			# color-highlighted text in Word
			hlColor = attrs.get("highlight-color")
			oldHlColor = attrsCache.get("highlight-color") if attrsCache is not None else None
			if (hlColor or oldHlColor is not None) and hlColor != oldHlColor:
				colorName = hlColor.name if isinstance(hlColor, colors.RGB) else hlColor
				text = (
					# Translators: Reported when text is color-highlighted
					NVDAString("highlighted in {color}") .format(color=colorName) if hlColor
					# Translators: Reported when text is no longer marked
					else NVDAString("not highlighted"))
				textList.append(text)
		if formatConfig["reportEmphasis"]:
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
		if formatConfig["reportSuperscriptsAndSubscripts"]:
			textPosition = attrs.get("text-position", TextPosition.UNDEFINED)
			attrs["text-position"] = textPosition
			oldTextPosition = attrsCache.get("text-position") if attrsCache is not None else None
			if (
				textPosition != oldTextPosition
				and (
					textPosition in [TextPosition.SUPERSCRIPT, TextPosition.SUBSCRIPT]
					or (
						textPosition == TextPosition.BASELINE
						and (oldTextPosition is not None and oldTextPosition != TextPosition.UNDEFINED)
					)
				)
			):
				textList.append(textPosition.displayString)
		if formatConfig["reportAlignment"]:
			textAlign = attrs.get("text-align")
			oldTextAlign = attrsCache.get("text-align") if attrsCache is not None else None
			if textAlign and textAlign != oldTextAlign:
				textList.append(textAlign.displayString)
			verticalAlign = attrs.get("vertical-align")
			oldVerticalAlign = attrsCache.get("vertical-align") if attrsCache is not None else None
			if verticalAlign and verticalAlign != oldVerticalAlign:
				textList.append(verticalAlign.displayString)
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
					# my modification
					textList.extend(self.getCommentFormatFieldSpeech(comment))
					"""
					if comment is textInfos.CommentType.DRAFT:
						# Translators: Reported when text contains a draft comment.
						text = NVDAString("has draft comment")
					elif comment is textInfos.CommentType.RESOLVED:
						# Translators: Reported when text contains a resolved comment.
						text = NVDAString("has resolved comment")
					else:  # generic
						# Translators: Reported when text contains a generic comment.
						text = NVDAString("has comment")
					textList.append(text)
					"""
					# end modification
				elif extraDetail:
					# Translators: Reported when text no longer contains a comment.
					text = NVDAString("out of comment")
					textList.append(text)
		if formatConfig["reportBookmarks"]:
			bookmark = attrs.get("bookmark")
			oldBookmark = attrsCache.get("bookmark") if attrsCache is not None else None
			if (bookmark or oldBookmark is not None) and bookmark != oldBookmark:
				if bookmark:
					# Translators: Reported when text contains a bookmark
					text = NVDAString("bookmark")
					textList.append(text)
				elif extraDetail:
					# Translators: Reported when text no longer contains a bookmark
					text = NVDAString("out of bookmark")
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
		# The line-prefix formatField attribute contains the text for a bullet or number for a list item,
					# when the bullet or number does not appear in the actual text content.
		# Normally this attribute could be repeated across formatFields within a list item
					# and therefore is not safe to speak when the unit is word or character.
		# However, some implementations (such as MS Word with UIA)
					# do limit its useage to the very first formatField of the list item.
		# Therefore, they also expose a line-prefix_speakAlways attribute
					# to allow its usage for any unit.
		linePrefix_speakAlways = attrs.get('line-prefix_speakAlways', False)
		if linePrefix_speakAlways or unit in (
			textInfos.UNIT_LINE, textInfos.UNIT_SENTENCE, textInfos.UNIT_PARAGRAPH, textInfos.UNIT_READINGCHUNK):
			linePrefix = attrs.get("line-prefix")
			if linePrefix:
				textList.append(linePrefix)
		if attrsCache is not None:
			attrsCache.clear()
			attrsCache.update(attrs)
		logBadSequenceTypes(textList)
		return textList

	def getFormatFieldSpeech_2023_3(  # noqa: C901
		self,
		attrs: textInfos.Field,
		attrsCache: Optional[textInfos.Field] = None,
		formatConfig: Optional[Dict[str, bool]] = None,
		reason: Optional[OutputReason] = None,
		unit: Optional[str] = None,
		extraDetail: bool = False,
		initialFormat: bool = False,
	) -> SpeechSequence:

		if not formatConfig:
			formatConfig = config.conf["documentFormatting"]
		textList = []
		if formatConfig["reportTables"]:
			tableInfo = attrs.get("table-info")
			oldTableInfo = attrsCache.get("table-info") if attrsCache is not None else None
			tableSequence = getTableInfoSpeech(
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
			if (((
				textColumnNumber and textColumnNumber != oldTextColumnNumber)
				or (textColumnCount and textColumnCount != oldTextColumnCount)) and not
				(textColumnCount and int(textColumnCount) <= 1 and oldTextColumnCount is None)):
				if textColumnNumber and textColumnCount:
					# Translators: Indicates the text column number in a document.
					# {0} will be replaced with the text column number.
					# {1} will be replaced with the number of text columns.
					text = NVDAString("column {0} of {1}") .format(textColumnNumber, textColumnCount)
					textList.append(text)
				elif textColumnCount:
					# Translators: Indicates the text column number in a document.
					# %s will be replaced with the number of text columns.
					text = NVDAString("%s columns") % (textColumnCount)
					textList.append(text)
				elif textColumnNumber:
					# Translators: Indicates the text column number in a document.
					text = NVDAString("column {columnNumber}") .format(columnNumber=textColumnNumber)
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
			if (
				headingLevel
				and (
					initialFormat
					and (
						reason in [OutputReason.FOCUS, OutputReason.QUICKNAV]
						or unit in (textInfos.UNIT_LINE, textInfos.UNIT_PARAGRAPH)
					)
					or headingLevel != oldHeadingLevel
				)
			):
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
		if formatConfig["reportCellBorders"] != ReportCellBorders.OFF:
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
				bgColorText = NVDAString("{color1} to {color2}") .format(color1=bgColorText, color2=bg2Name)
			if color and backgroundColor and color != oldColor and bgColorChanged:
				# Translators: Reported when both the text and background colors change.
				# {color} will be replaced with the text color.
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{color} on {backgroundColor}") .format(
					color=color.name if isinstance(color, colors.RGB) else color,
					backgroundColor=bgColorText))
			elif color and color != oldColor:
				# Translators: Reported when the text color changes (but not the background color).
				# {color} will be replaced with the text color.
				textList.append(
					NVDAString("{color}") .format(color=color.name if isinstance(color, colors.RGB) else color))
			elif backgroundColor and bgColorChanged:
				# Translators: Reported when the background color changes (but not the text color).
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{backgroundColor} background") .format(backgroundColor=bgColorText))
			backgroundPattern = attrs.get("background-pattern")
			oldBackgroundPattern = attrsCache.get("background-pattern") if attrsCache is not None else None
			if backgroundPattern and backgroundPattern != oldBackgroundPattern:
				textList.append(NVDAString("background pattern {pattern}") .format(pattern=backgroundPattern))
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
					# Translators: Reported when text is no longer marked as having been  deleted.
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
		if formatConfig["reportHighlight"]:
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
			# color-highlighted text in Word
			hlColor = attrs.get("highlight-color")
			oldHlColor = attrsCache.get("highlight-color") if attrsCache is not None else None
			if (hlColor or oldHlColor is not None) and hlColor != oldHlColor:
				colorName = hlColor.name if isinstance(hlColor, colors.RGB) else hlColor
				text = (
					# Translators: Reported when text is color-highlighted
					NVDAString("highlighted in {color}") .format(color=colorName) if hlColor
					# Translators: Reported when text is no longer marked
					else NVDAString("not highlighted"))
				textList.append(text)
		if formatConfig["reportEmphasis"]:
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
		if formatConfig["reportSuperscriptsAndSubscripts"]:
			textPosition = attrs.get("text-position", TextPosition.UNDEFINED)
			attrs["text-position"] = textPosition
			oldTextPosition = attrsCache.get("text-position") if attrsCache is not None else None
			if (
				textPosition != oldTextPosition
				and (
					textPosition in [TextPosition.SUPERSCRIPT, TextPosition.SUBSCRIPT]
					or (
						textPosition == TextPosition.BASELINE
						and (oldTextPosition is not None and oldTextPosition != TextPosition.UNDEFINED)
					)
				)
			):
				textList.append(textPosition.displayString)
		if formatConfig["reportAlignment"]:
			textAlign = attrs.get("text-align")
			oldTextAlign = attrsCache.get("text-align") if attrsCache is not None else None
			if textAlign and textAlign != oldTextAlign:
				textList.append(textAlign.displayString)
			verticalAlign = attrs.get("vertical-align")
			oldVerticalAlign = attrsCache.get("vertical-align") if attrsCache is not None else None
			if verticalAlign and verticalAlign != oldVerticalAlign:
				textList.append(verticalAlign.displayString)
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
					# my modification
					textList.extend(self.getCommentFormatFieldSpeech(comment))
					"""
					if comment is textInfos.CommentType.DRAFT:
						# Translators: Reported when text contains a draft comment.
						text = NVDAString("has draft comment")
					elif comment is textInfos.CommentType.RESOLVED:
						# Translators: Reported when text contains a resolved comment.
						text = NVDAString("has resolved comment")
					else:  # generic
						# Translators: Reported when text contains a generic comment.
						text = NVDAString("has comment")
					textList.append(text)
					"""
					# end modification
				elif extraDetail:
					# Translators: Reported when text no longer contains a comment.
					text = NVDAString("out of comment")
					textList.append(text)
		if formatConfig["reportBookmarks"]:
			bookmark = attrs.get("bookmark")
			oldBookmark = attrsCache.get("bookmark") if attrsCache is not None else None
			if (bookmark or oldBookmark is not None) and bookmark != oldBookmark:
				if bookmark:
					# Translators: Reported when text contains a bookmark
					text = NVDAString("bookmark")
					textList.append(text)
				elif extraDetail:
					# Translators: Reported when text no longer contains a bookmark
					text = NVDAString("out of bookmark")
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
		# The line-prefix formatField attribute contains the text for a bullet or number for a list item,
					# when the bullet or number does not appear in the actual text content.
		# Normally this attribute could be repeated across formatFields within a list item
		# and therefore is not safe to speak when the unit is word or character.
		# However, some implementations (such as MS Word with UIA)
		# do limit its useage to the very first formatField of the list item.
		# Therefore, they also expose a line-prefix_speakAlways attribute to allow its usage for any unit.
		linePrefix_speakAlways = attrs.get('line-prefix_speakAlways', False)
		if linePrefix_speakAlways or unit in (
			textInfos.UNIT_LINE, textInfos.UNIT_SENTENCE, textInfos.UNIT_PARAGRAPH, textInfos.UNIT_READINGCHUNK):
			linePrefix = attrs.get("line-prefix")
			if linePrefix:
				textList.append(linePrefix)
		if attrsCache is not None:
			attrsCache.clear()
			attrsCache.update(attrs)
		logBadSequenceTypes(textList)
		return textList

	def getFormatFieldSpeech_2023_2(  # noqa: C901
		self,
		attrs: textInfos.Field,
		attrsCache: Optional[textInfos.Field] = None,
		formatConfig: Optional[Dict[str, bool]] = None,
		reason: Optional[OutputReason] = None,
		unit: Optional[str] = None,
		extraDetail: bool = False,
		initialFormat: bool = False,
	) -> SpeechSequence:
		if not formatConfig:
			formatConfig = config.conf["documentFormatting"]
		textList = []
		if formatConfig["reportTables"]:
			tableInfo = attrs.get("table-info")
			oldTableInfo = attrsCache.get("table-info") if attrsCache is not None else None
			tableSequence = getTableInfoSpeech(
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
			if (((
				textColumnNumber and textColumnNumber != oldTextColumnNumber)
				or (textColumnCount and textColumnCount != oldTextColumnCount)) and not
				(textColumnCount and int(textColumnCount) <= 1 and oldTextColumnCount is None)):
				if textColumnNumber and textColumnCount:
					# Translators: Indicates the text column number in a document.
					# {0} will be replaced with the text column number.
					# {1} will be replaced with the number of text columns.
					text = NVDAString("column {0} of {1}") .format(textColumnNumber, textColumnCount)
					textList.append(text)
				elif textColumnCount:
					# Translators: Indicates the text column number in a document.
					# %s will be replaced with the number of text columns.
					text = NVDAString("%s columns") % (textColumnCount)
					textList.append(text)
				elif textColumnNumber:
					# Translators: Indicates the text column number in a document.
					text = NVDAString("column {columnNumber}") .format(columnNumber=textColumnNumber)
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
			if (
				headingLevel
				and (
					initialFormat
					and (
						reason in [OutputReason.FOCUS, OutputReason.QUICKNAV]
						or unit in (textInfos.UNIT_LINE, textInfos.UNIT_PARAGRAPH)
					)
					or headingLevel != oldHeadingLevel
				)
			):
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
		if formatConfig["reportCellBorders"] != ReportCellBorders.OFF:
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
				bgColorText = NVDAString("{color1} to {color2}") .format(color1=bgColorText, color2=bg2Name)
			if color and backgroundColor and color != oldColor and bgColorChanged:
				# Translators: Reported when both the text and background colors change.
				# {color} will be replaced with the text color.
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{color} on {backgroundColor}") .format(
					color=color.name if isinstance(color, colors.RGB) else color,
					backgroundColor=bgColorText))
			elif color and color != oldColor:
				# Translators: Reported when the text color changes (but not the background color).
				# {color} will be replaced with the text color.
				textList.append(
					NVDAString("{color}") .format(color=color.name if isinstance(color, colors.RGB) else color))
			elif backgroundColor and bgColorChanged:
				# Translators: Reported when the background color changes (but not the text color).
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{backgroundColor} background") .format(backgroundColor=bgColorText))
			backgroundPattern = attrs.get("background-pattern")
			oldBackgroundPattern = attrsCache.get("background-pattern") if attrsCache is not None else None
			if backgroundPattern and backgroundPattern != oldBackgroundPattern:
				textList.append(NVDAString("background pattern {pattern}") .format(pattern=backgroundPattern))
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
				text = (NVDAString(
					"inserted") if revision
					# Translators: Reported when text is no longer marked as having been inserted.
					else NVDAString("not inserted"))
				textList.append(text)
			revision = attrs.get("revision-deletion")
			oldRevision = attrsCache.get("revision-deletion") if attrsCache is not None else None
			if (revision or oldRevision is not None) and revision != oldRevision:
				# Translators: Reported when text is marked as having been deleted
				text = (NVDAString(
					"deleted") if revision
					# Translators: Reported when text is no longer marked as having been  deleted.
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
		if formatConfig["reportHighlight"]:
			# marked text
			marked = attrs.get("marked")
			oldMarked = attrsCache.get("marked") if attrsCache is not None else None
			if (marked or oldMarked is not None) and marked != oldMarked:
				# Translators: Reported when text is marked
				text = (NVDAString(
					"marked") if marked
					# Translators: Reported when text is no longer marked
					else NVDAString("not marked"))
				textList.append(text)
			# color-highlighted text in Word
			hlColor = attrs.get("highlight-color")
			oldHlColor = attrsCache.get("highlight-color") if attrsCache is not None else None
			if (hlColor or oldHlColor is not None) and hlColor != oldHlColor:
				colorName = hlColor.name if isinstance(hlColor, colors.RGB) else hlColor
				text = (
					# Translators: Reported when text is color-highlighted
					NVDAString("highlighted in {color}") .format(color=colorName) if hlColor
					# Translators: Reported when text is no longer marked
					else NVDAString("not highlighted"))
				textList.append(text)
		if formatConfig["reportEmphasis"]:
			# strong text
			strong = attrs.get("strong")
			oldStrong = attrsCache.get("strong") if attrsCache is not None else None
			if (strong or oldStrong is not None) and strong != oldStrong:
				# Translators: Reported when text is marked as strong (e.g. bold)
				text = (NVDAString(
					"strong") if strong
					# Translators: Reported when text is no longer marked as strong (e.g. bold)
					else NVDAString("not strong"))
				textList.append(text)
			# emphasised text
			emphasised = attrs.get("emphasised")
			oldEmphasised = attrsCache.get("emphasised") if attrsCache is not None else None
			if (emphasised or oldEmphasised is not None) and emphasised != oldEmphasised:
				# Translators: Reported when text is marked as emphasised
				text = (NVDAString(
					"emphasised") if emphasised
					# Translators: Reported when text is no longer marked as emphasised
					else NVDAString("not emphasised"))
				textList.append(text)
		if formatConfig["reportFontAttributes"]:
			bold = attrs.get("bold")
			oldBold = attrsCache.get("bold") if attrsCache is not None else None
			if (bold or oldBold is not None) and bold != oldBold:
				# Translators: Reported when text is bolded.
				text = (NVDAString(
					"bold") if bold
					# Translators: Reported when text is not bolded.
					else NVDAString("no bold"))
				textList.append(text)
			italic = attrs.get("italic")
			oldItalic = attrsCache.get("italic") if attrsCache is not None else None
			if (italic or oldItalic is not None) and italic != oldItalic:
				# Translators: Reported when text is italicized.
				text = (NVDAString(
					"italic") if italic
					# Translators: Reported when text is not italicized.
					else NVDAString("no italic"))
				textList.append(text)
			strikethrough = attrs.get("strikethrough")
			oldStrikethrough = attrsCache.get("strikethrough") if attrsCache is not None else None
			if (strikethrough or oldStrikethrough is not None) and strikethrough != oldStrikethrough:
				if strikethrough:
					# Translators: Reported when text is formatted with double strikethrough.
					# See http://en.wikipedia.org/wiki/Strikethrough
					text = (NVDAString(
						"double strikethrough") if strikethrough == "double"
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
		if formatConfig["reportSuperscriptsAndSubscripts"]:
			textPosition = attrs.get("text-position", TextPosition.UNDEFINED)
			attrs["text-position"] = textPosition
			oldTextPosition = attrsCache.get("text-position") if attrsCache is not None else None
			if (
				textPosition != oldTextPosition
				and (
					textPosition in [TextPosition.SUPERSCRIPT, TextPosition.SUBSCRIPT]
					or (
						textPosition == TextPosition.BASELINE
						and (oldTextPosition is not None and oldTextPosition != TextPosition.UNDEFINED)
					)
				)
			):
				textList.append(textPosition.displayString)
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
					# Translators: Reported when text is vertically justified
					# but with character spacing (For some Asian content).
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
					# my modification
					textList.extend(self.getCommentFormatFieldSpeech(comment))
					"""
					if comment is textInfos.CommentType.DRAFT:
						# Translators: Reported when text contains a draft comment.
						text = NVDAString("has draft comment")
					elif comment is textInfos.CommentType.RESOLVED:
						# Translators: Reported when text contains a resolved comment.
						text = NVDAString("has resolved comment")
					else:  # generic
						# Translators: Reported when text contains a generic comment.
						text = NVDAString("has comment")
					textList.append(text)
					"""
					# end modification
				elif extraDetail:
					# Translators: Reported when text no longer contains a comment.
					text = NVDAString("out of comment")
					textList.append(text)
		if formatConfig["reportBookmarks"]:
			bookmark = attrs.get("bookmark")
			oldBookmark = attrsCache.get("bookmark") if attrsCache is not None else None
			if (bookmark or oldBookmark is not None) and bookmark != oldBookmark:
				if bookmark:
					# Translators: Reported when text contains a bookmark
					text = NVDAString("bookmark")
					textList.append(text)
				elif extraDetail:
					# Translators: Reported when text no longer contains a bookmark
					text = NVDAString("out of bookmark")
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
		# The line-prefix formatField attribute contains the text for a bullet
					# or number for a list item, when the bullet or number does not appear in the actual text content.
		# Normally this attribute could be repeated across formatFields within a list item
					#  and therefore is not safe to speak when the unit is word or character.
		# However, some implementations (such as MS Word with UIA)
					#  do limit its useage to the very first formatField of the list item.
		# Therefore, they also expose a line-prefix_speakAlways attribute to allow its usage for any unit.
		linePrefix_speakAlways = attrs.get('line-prefix_speakAlways', False)
		if linePrefix_speakAlways or unit in (
			textInfos.UNIT_LINE, textInfos.UNIT_SENTENCE, textInfos.UNIT_PARAGRAPH, textInfos.UNIT_READINGCHUNK):
			linePrefix = attrs.get("line-prefix")
			if linePrefix:
				textList.append(linePrefix)
		if attrsCache is not None:
			attrsCache.clear()
			attrsCache.update(attrs)
		logBadSequenceTypes(textList)
		return textList

	def getFormatFieldSpeech_2023_1(
		self,
		attrs: textInfos.Field,
		attrsCache: Optional[textInfos.Field] = None,
		formatConfig: Optional[Dict[str, bool]] = None,
		reason: Optional[OutputReason] = None,
		unit: Optional[str] = None,
		extraDetail: bool = False,
		initialFormat: bool = False,
	) -> SpeechSequence:

		if not formatConfig:
			formatConfig = config.conf["documentFormatting"]
		textList = []
		if formatConfig["reportTables"]:
			tableInfo = attrs.get("table-info")
			oldTableInfo = attrsCache.get("table-info") if attrsCache is not None else None
			tableSequence = getTableInfoSpeech(
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
			if (((
				textColumnNumber and textColumnNumber != oldTextColumnNumber)
				or (textColumnCount and textColumnCount != oldTextColumnCount)) and not
				(textColumnCount and int(textColumnCount) <= 1 and oldTextColumnCount is None)):
				if textColumnNumber and textColumnCount:
					# Translators: Indicates the text column number in a document.
					# {0} will be replaced with the text column number.
					# {1} will be replaced with the number of text columns.
					text = NVDAString("column {0} of {1}") .format(textColumnNumber, textColumnCount)
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
			if (
				headingLevel
				and (
					initialFormat
					and (
						reason in [OutputReason.FOCUS, OutputReason.QUICKNAV]
						or unit in (textInfos.UNIT_LINE, textInfos.UNIT_PARAGRAPH)
					)
					or headingLevel != oldHeadingLevel
				)
			):
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
		if formatConfig["reportCellBorders"] != ReportCellBorders.OFF:
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
				bgColorText = NVDAString("{color1} to {color2}") .format(color1=bgColorText, color2=bg2Name)
			if color and backgroundColor and color != oldColor and bgColorChanged:
				# Translators: Reported when both the text and background colors change.
				# {color} will be replaced with the text color.
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{color} on {backgroundColor}") .format(
					color=color.name if isinstance(color, colors.RGB) else color,
					backgroundColor=bgColorText))
			elif color and color != oldColor:
				# Translators: Reported when the text color changes (but not the background color).
				# {color} will be replaced with the text color.
				textList.append(
					NVDAString("{color}") .format(color=color.name if isinstance(color, colors.RGB) else color))
			elif backgroundColor and bgColorChanged:
				# Translators: Reported when the background color changes (but not the text color).
				# {backgroundColor} will be replaced with the background color.
				textList.append(NVDAString("{backgroundColor} background") .format(backgroundColor=bgColorText))
			backgroundPattern = attrs.get("background-pattern")
			oldBackgroundPattern = attrsCache.get("background-pattern") if attrsCache is not None else None
			if backgroundPattern and backgroundPattern != oldBackgroundPattern:
				textList.append(NVDAString("background pattern {pattern}") .format(pattern=backgroundPattern))
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
				text = (
					# Translators: Reported when text is marked as having been inserted
					NVDAString("inserted") if revision else
					# Translators: Reported when text is no longer marked as having been inserted.
					NVDAString("not inserted"))
				textList.append(text)
			revision = attrs.get("revision-deletion")
			oldRevision = attrsCache.get("revision-deletion") if attrsCache is not None else None
			if (revision or oldRevision is not None) and revision != oldRevision:
				text = (
					# Translators: Reported when text is marked as having been deleted
					NVDAString("deleted") if revision else
					# Translators: Reported when text is no longer marked as having been  deleted.
					NVDAString("not deleted"))
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
		if formatConfig["reportHighlight"]:
			# marked text
			marked = attrs.get("marked")
			oldMarked = attrsCache.get("marked") if attrsCache is not None else None
			if (marked or oldMarked is not None) and marked != oldMarked:
				text = (
					# Translators: Reported when text is marked
					NVDAString("marked") if marked else
					# Translators: Reported when text is no longer marked
					NVDAString("not marked"))
				textList.append(text)
		if formatConfig["reportEmphasis"]:
			# strong text
			strong = attrs.get("strong")
			oldStrong = attrsCache.get("strong") if attrsCache is not None else None
			if (strong or oldStrong is not None) and strong != oldStrong:
				text = (
					# Translators: Reported when text is marked as strong (e.g. bold)
					NVDAString("strong") if strong else
					# Translators: Reported when text is no longer marked as strong (e.g. bold)
					NVDAString("not strong"))
				textList.append(text)
			# emphasised text
			emphasised = attrs.get("emphasised")
			oldEmphasised = attrsCache.get("emphasised") if attrsCache is not None else None
			if (emphasised or oldEmphasised is not None) and emphasised != oldEmphasised:
				text = (
					# Translators: Reported when text is marked as emphasised
					NVDAString("emphasised") if emphasised else
					# Translators: Reported when text is no longer marked as emphasised
					NVDAString("not emphasised"))
				textList.append(text)
		if formatConfig["reportFontAttributes"]:
			bold = attrs.get("bold")
			oldBold = attrsCache.get("bold") if attrsCache is not None else None
			if (bold or oldBold is not None) and bold != oldBold:
				text = (
					# Translators: Reported when text is bolded.
					NVDAString("bold") if bold else
					# Translators: Reported when text is not bolded.
					NVDAString("no bold"))
				textList.append(text)
			italic = attrs.get("italic")
			oldItalic = attrsCache.get("italic") if attrsCache is not None else None
			if (italic or oldItalic is not None) and italic != oldItalic:
				text = (
					# Translators: Reported when text is italicized.
					NVDAString("italic") if italic else
					# Translators: Reported when text is not italicized.
					NVDAString("no italic"))
				textList.append(text)
			strikethrough = attrs.get("strikethrough")
			oldStrikethrough = attrsCache.get("strikethrough") if attrsCache is not None else None
			if (strikethrough or oldStrikethrough is not None) and strikethrough != oldStrikethrough:
				if strikethrough:
					text = (
						# Translators: Reported when text is formatted with double strikethrough.
						# See http://en.wikipedia.org/wiki/Strikethrough
						NVDAString("double strikethrough") if strikethrough == "double" else
						# Translators: Reported when text is formatted with strikethrough.
						# See http://en.wikipedia.org/wiki/Strikethrough
						NVDAString("strikethrough"))
				else:
					# Translators: Reported when text is formatted without strikethrough.
					# See http://en.wikipedia.org/wiki/Strikethrough
					text = NVDAString("no strikethrough")
				textList.append(text)
			underline = attrs.get("underline")
			oldUnderline = attrsCache.get("underline") if attrsCache is not None else None
			if (underline or oldUnderline is not None) and underline != oldUnderline:
				text = (
					# Translators: Reported when text is underlined.
					NVDAString("underlined") if underline else
					# Translators: Reported when text is not underlined.
					NVDAString("not underlined"))
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
		if formatConfig["reportSuperscriptsAndSubscripts"]:
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
					# Translators: Reported when text is vertically justified
					# but with character spacing (For some Asian content).
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
					# my modification
					textList.extend(self.getCommentFormatFieldSpeech(comment))
					# end modification
				elif extraDetail:
					# Translators: Reported when text no longer contains a comment.
					text = NVDAString("out of comment")
					textList.append(text)
		if formatConfig["reportBookmarks"]:
			bookmark = attrs.get("bookmark")
			oldBookmark = attrsCache.get("bookmark") if attrsCache is not None else None
			if (bookmark or oldBookmark is not None) and bookmark != oldBookmark:
				if bookmark:
					# Translators: Reported when text contains a bookmark
					text = NVDAString("bookmark")
					textList.append(text)
				elif extraDetail:
					# Translators: Reported when text no longer contains a bookmark
					text = NVDAString("out of bookmark")
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
		# The line-prefix formatField attribute contains the text for a bullet
					# or number for a list item, when the bullet or number does not appear in the actual text content.
		# Normally this attribute could be repeated across formatFields within a list item
		# and therefore is not safe to speak when the unit is word or character.
		# However, some implementations (such as MS Word with UIA)
		# do limit its useage to the very first formatField of the list item.
		# Therefore, they also expose a line-prefix_speakAlways attribute to allow its usage for any unit.
		linePrefix_speakAlways = attrs.get('line-prefix_speakAlways', False)
		if linePrefix_speakAlways or unit in (
			textInfos.UNIT_LINE, textInfos.UNIT_SENTENCE, textInfos.UNIT_PARAGRAPH, textInfos.UNIT_READINGCHUNK):
			linePrefix = attrs.get("line-prefix")
			if linePrefix:
				textList.append(linePrefix)
		if attrsCache is not None:
			attrsCache.clear()
			attrsCache.update(attrs)
		logBadSequenceTypes(textList)
		return textList
