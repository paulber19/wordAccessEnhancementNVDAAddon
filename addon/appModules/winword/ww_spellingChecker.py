# -*- coding: UTF-8 -*-
# appModules\word\appModules/winword/ww_spellingChecker.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import api
import ui
import eventHandler
import queueHandler
import controlTypes
import speech

addonHandler.initTranslation()

ID_InSpellingChecker_No = 0
ID_InSpellingChecker_2016AndLess = 1
ID_InSpellingChecker_2019AndMore = 2


class BaseSpellingChecker(object):
	def __init__(self, focus):
		self.focus = focus
		(self._SpellingCheckerID, self.pane) = self._getPane(focus)

	def isInSpellingChecker(self):
		return self._SpellingCheckerID > ID_InSpellingChecker_No

	def _getPane(self, obj):
		o = obj
		oParents = []
		pane = None
		id = ID_InSpellingChecker_No
		i = 50
		while i and o and o.windowClassName != "#32769":
			i = i-1
			oParents.append(o)
			if o.windowClassName == "MsoCommandBarDock" and o.name == "MsoDockRight":
				pane = oParents[-8] if len(oParents) >= 8 else None
				if pane is not None and pane.childCount > 1:
					# word 2016 and less
					id = ID_InSpellingChecker_2016AndLess
					break
				pane = oParents[-9] if len(oParents) >= 9 else None
				if pane is not None and pane.childCount > 1:
					# word 2019 and more
					id = ID_InSpellingChecker_2019AndMore
					break
			o = o.parent

		return (id, pane)

	def sayErrorAndSuggestion(
		self, title=False, spell=False, focusOnSuggestion=False):
		return


class SpellingChecker_2016(object):
	def __init__(self, bsc):
		self.bsc = bsc
		self.pane = bsc.pane

	def isInSpellingChecker(self):
		return True

	def getErrorInformations(self):
		infos = {"title": None, "type": None, "error": None, "suggestion": None}
		pane = self.pane
		o = pane.parent.parent.parent.parent.parent
		infos["title"] = o.name
		infos["type"] = pane.name
		infos["error"] = pane.firstChild.name if pane.firstChild is not None else ""
		oList = None
		for o in pane.children:
			if o.role == controlTypes.ROLE_LIST:
				oList = o
				break
		if oList:
			if oList.firstChild.role == controlTypes.ROLE_PANE:
				childs = oList.firstChild.firstChild.children
			else:
				childs = oList.children
			for o in childs:
				if o.name and controlTypes.STATE_SELECTED in o.states:
					# Translators: message to indicate the suggested correction.
					infos["suggestion"] = o.name
					break
		return infos

	def sayErrorAndSuggestion(
		self, title=False, spell=False, focusOnSuggestion=False):
		infos = self.getErrorInformations()
		if title and infos["title"] is not None:
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				ui.message,
				infos["title"])
		if not spell:
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				ui.message,
				infos["type"])
		queueHandler.queueFunction(
			queueHandler.eventQueue,
			ui.message,
			# Translators: message to indicate the spell or grammar error.
			_("Error: %s") % infos["error"])
		if spell:
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				speech.speakSpelling,
				infos["error"])
		if infos["suggestion"] is not None:
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				ui.message,
				# Translators: message to indicate the suggested correction.
				_("Suggestion: %s") % infos["suggestion"])
			if spell:
				queueHandler.queueFunction(
					queueHandler.eventQueue,
					speech.speakSpelling,
					infos["suggestion"])

	def sayHelpText(self):

		def getExplanationText():
			textList = []
			for o in self.pane.children[1:]:
				if o.role in [controlTypes.ROLE_STATICTEXT, controlTypes.ROLE_LINK]\
					and controlTypes.STATE_INVISIBLE not in o.states:
					textList.append(o.name)
			if len(textList):
				return "\n".join(textList)
			else:
				# Translators: message to user that there is no help.
				return _("No explanation help")

		text = getExplanationText()
		queueHandler.queueFunction(queueHandler.eventQueue, ui.message, text)


class SpellingChecker_2019(object):
	def __init__(self, bsc):
		self.bsc = bsc
		self.pane = bsc.pane

	def isInSpellingChecker(self):
		return True

	def getErrorInformations(self):

		def getSpellSuggestionText(buttonText):
			if buttonText is None:
				return (None, None)
			buttonTextList = buttonText.split("\xa0:")
			s = buttonTextList[0].split(", ")
			suggestion = s[0]
			s1 = buttonTextList[0].split(", ")
			if len(s1) > 1:
				s1 = s1[1:]
			s1.extend(buttonTextList[1:])
			helpText = "\xa0:".join(s1)
			return (suggestion, helpText)

		def getSelectedSuggestion(pane):
			oSuggestions = pane.getChild(1)
			if oSuggestions.role == controlTypes.ROLE_GROUPING:
				suggestionCount = oSuggestions .childCount-1
				if suggestionCount:
					sg = None
					firstSuggestion = None
					for o in oSuggestions.children:
						if o.role != controlTypes.ROLE_SPLITBUTTON:
							continue
						if firstSuggestion is None:
							firstSuggestion = o
						if controlTypes.STATE_FOCUSED in o.states:
							sg = o
							break
					if sg:
						(suggestion, helpText) = getSpellSuggestionText(sg.name)
						return (suggestion, helpText, suggestionCount)
					if firstSuggestion:
						# return first suggestion to put focus on
						(suggestion, helpText) = getSpellSuggestionText(firstSuggestion.name)
						return (suggestion, helpText, suggestionCount)
			return (None, None, suggestionCount)

		infos = {"title": None, "type": None, "error": None, "suggestion": None}
		pane = self.pane
		o = pane.parent.parent.parent.parent.parent
		infos["title"] = o.name
		errorObj = pane.firstChild
		# grammar or spelling
		o = errorObj.getChild(1)
		spelling = True
		if o.role == controlTypes.ROLE_TOGGLEBUTTON:
			# grammar error
			spelling = False
		error = ""
		type = ""
		if errorObj.name != "":
			tempList = errorObj.name.split(chr(160)+":")
			if spelling:
				error = tempList[0].strip()
				type = tempList[1].split(", ")[-1]
				type = type + " : " + pane.firstChild.getChild(1).name
			else:
				# grammar
				e = tempList[0].split(", ")
				error = e[0].strip()
				type = e[-1].strip()
				type = type + " : " + tempList[-1]
		infos["error"] = error
		infos["type"] = type
		# suggestion
		infos["suggestion"] = getSelectedSuggestion(pane)
		return infos

	def sayErrorAndSuggestion(
		self, title=False, spell=False, focusOnSuggestion=False):

		infos = self.getErrorInformations()
		if title and infos["title"] is not None:
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				ui.message,
				infos["title"])
		if not spell:
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				ui.message,
				infos["type"])
		queueHandler.queueFunction(
			queueHandler.eventQueue,
			ui.message,
			# Translators: message to indicate the spell or grammar error.
			_("Error: %s") % infos["error"])
		if spell:
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				speech.speakSpelling,
				infos["error"])
		(suggestion, suggestionHelpText, suggestionCount) = infos["suggestion"]
		if suggestionCount == 0:
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				ui.message,
				# Translators: message to user to report no suggestion.
				_("No suggestion"))
			return

		if suggestion:
			queueHandler.queueFunction(
				queueHandler.eventQueue,
				ui.message,
				# Translators: message to indicate the suggested correction.
				_("Suggestion: %s") % suggestion)
			if spell:
				queueHandler.queueFunction(
					queueHandler.eventQueue,
					speech.speakSpelling,
					suggestion)
			if suggestionCount > 1:
				queueHandler.queueFunction(
					queueHandler.eventQueue,
					ui.message,
					# Translators: message to user to report most of one suggestions.
					_("%s suggestions") % suggestionCount)
			else:
				queueHandler.queueFunction(
					queueHandler.eventQueue,
					ui.message,
					# Translators: message to user to report one suggestion.
					_("One suggestions"))
		if spell:
			return
		if focusOnSuggestion:
			# focus on first suggestion
			firstSuggestion = self.pane.getChild(1).getChild(1)
			if api.getFocusObject() != firstSuggestion:
				firstSuggestion.setFocus()
			else:
				eventHandler.queueEvent("focusEntered", firstSuggestion.parent)
				eventHandler.queueEvent("gainFocus", firstSuggestion)

	def sayHelpText(self):
		def getDisplayExplanationButton():
			button = None
			for o in self.pane.firstChild.children:
				if o.role == controlTypes.ROLE_TOGGLEBUTTON\
					and controlTypes.STATE_INVISIBLE not in o.states:
					button = o
					break

			return button

		def getExplanationText():
			focus = api.getFocusObject()
			button = getDisplayExplanationButton()
			if button is None:
				return None
			oldSpeechMode = speech.speechMode
			speech.speechMode = speech.speechMode_off
			stateHasChanged = False
			if controlTypes.STATE_PRESSED not in button.states:
				button.doAction()
				stateHasChanged = True
			oText = button.next
			textList = [oText.name]
			textList.append(oText.value)
			if stateHasChanged:
				button.doAction()
			focus.setFocus()
			api.processPendingEvents()
			speech.speechMode = oldSpeechMode
			speech.cancelSpeech()
			return "\r\n".join(textList)

		def getDefaultHelpText():
			infos = self.getErrorInformations()
			(suggestion, suggestionHelpText, suggestionCount) = infos["suggestion"]
			return suggestionHelpText

		text = getExplanationText()
		if text is None:
			text = getDefaultHelpText()
			if text is None:
				text = _("No explanation help")
		queueHandler.queueFunction(queueHandler.eventQueue, ui.message, text)


def SpellingChecker(focus, winwordVersion):
	bsc = BaseSpellingChecker(focus)
	if bsc._SpellingCheckerID == ID_InSpellingChecker_2016AndLess:
		return SpellingChecker_2016(bsc)
	elif bsc._SpellingCheckerID == ID_InSpellingChecker_2019AndMore:
		return SpellingChecker_2019(bsc)
	return bsc
