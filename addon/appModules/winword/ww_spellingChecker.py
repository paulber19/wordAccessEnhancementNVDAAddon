# appModules\word\appModules/winword/ww_spellingChecker.py
#A part of wordAccessEnhancement add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
from logHandler import log
import time
import api
import os
import wx
import ui
import eventHandler
import queueHandler
import keyboardHandler
import textInfos
import controlTypes
import speech
import NVDAObjects.IAccessible.winword
import inputCore

class SpellingChecker(object):
	def __init__(self, focus, winwordVersion):
		self.winwordVersion = winwordVersion
		(self._inSpellingChecker, self.pane) = self._getPane(focus)
	
	def inSpellingChecker(self):
		return self._inSpellingChecker
	
	def _getPane(self, obj):
		o = obj
		oParents = []
		pane = None
		inSpellingChecker = False
		i = 50
		while i and o and o.windowClassName != "#32769":
			i = i-1
			oParents.append(o)
			if o.windowClassName == "MsoCommandBarDock" and o.name == "MsoDockRight":
				pane= oParents[-8] if len(oParents) >= 8 else None
				inSpellingChecker = True
				break
			o = o.parent
		
		return (inSpellingChecker, pane)
	def getErrorInformations(self):
		if not self.inSpellingChecker: return None
		infos = {"title": None, "type": None, "error": None, "suggestion": None}
		pane = self.pane
		o = pane.parent.parent.parent.parent.parent
		infos["title"]  = o.name
		infos["type"] = pane.name
		infos["error"] = pane.firstChild.name if pane.firstChild is not None else ""
		oList = None
		for o in pane.children:
			if o.role == controlTypes.ROLE_LIST:
				oList = o
				break
		if oList:
			for o in oList.children:
				if controlTypes.STATE_SELECTED in o.states:
					# Translators: message to indicate the suggested correction.
					infos["suggestion"] = o.name
					break
		return infos

	def sayErrorAndSuggestion(self, title = False, spell = False):
		if not self.inSpellingChecker: return
		infos = self.getErrorInformations()
		if title and infos["title"] is not None:
			queueHandler.queueFunction(queueHandler.eventQueue, ui.message, infos["title"])
		if not spell:
			queueHandler.queueFunction(queueHandler.eventQueue, ui.message, infos["type"])
		# Translators: message to indicate the spell or grammar error.
		queueHandler.queueFunction(queueHandler.eventQueue,ui.message, _("Error: %s")%infos["error"])
		if spell:
			queueHandler.queueFunction(queueHandler.eventQueue, speech.speakSpelling,infos["error"])
		if infos["suggestion"] is not None:
			# Translators: message to indicate the suggested correction.
			queueHandler.queueFunction(queueHandler.eventQueue, ui.message,_("Suggestion: %s") %infos["suggestion"])
			if spell:
				queueHandler.queueFunction(queueHandler.eventQueue, speech.speakSpelling, infos["suggestion"])


	def sayHelpText(self):
		textList = []
		for o in self.pane.children[1:]:
			if o.role in [controlTypes.ROLE_STATICTEXT, controlTypes.ROLE_LINK]:
				textList.append(o.name)
		if len(textList):
			queueHandler.queueFunction(queueHandler.eventQueue, ui.message, "\n".join(textList))
	