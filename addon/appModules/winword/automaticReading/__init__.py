# appModules\winword\automaticReading\__init__.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2024-2025 paulber19, Abdel
# This file is covered by the GNU General Public License.

import addonHandler
from logHandler import log
from versionInfo import version_year, version_major
import config
import gui
import wx
import textInfos
from textInfos import SpeechSequence
from typing import List, Optional, Dict
import synthDriverHandler
import ui
import speech.commands
from speech import getPropertiesSpeech
import sys
import os
import controlTypes
from controlTypes import OutputReason
from .speechEx import FormatFieldSpeech

_curAddon = addonHandler.getCodeAddon()
sharedPath = os.path.join(_curAddon.path, "shared")
sys.path.append(sharedPath)
from ww_NVDAStrings import NVDAString
from ww_addonConfigManager import (
	_addonConfigManager, AutoReadingWith_CurrentVoice, AutoReadingWith_Beep, AutoReadingWith_Voice
)
from messages import confirm_YesNo, ReturnCode
del sys.path[-1]
del sys.modules["ww_NVDAStrings"]
del sys.modules["ww_addonConfigManager"]
del sys.modules["messages"]

addonHandler.initTranslation()

SCT_Speech = "speech"
SCT_Audio = "audio"
SCT_Many = "__many__"

NVDAVersion = [version_year, version_major]


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


_curSynthName = None
_currentSpeechSettings = None


def formatVoiceAutoReadingStartSequence(start=True):
	global _curSynthName, _currentSpeechSettings
	autoReadingSynth = _addonConfigManager.getAutoReadingSynthSettings()
	if not autoReadingSynth:
		return []
	seq = []
	seq.append(speech.commands.EndUtteranceCommand())
	if start:
		# start use of automatic reading synth
		_curSynthName = synthDriverHandler.getSynth().name
		_currentSpeechSettings = getCurrentSpeechSettings()
		(autoReadingSynthName, autoReadingSynthSpeechSettings) = (autoReadingSynth["synthName"], autoReadingSynth)
		seq.append(SetSynthCommand(autoReadingSynthName, autoReadingSynthSpeechSettings, temporary=True))
	else:
		# restore main synth
		seq.append(SetSynthCommand(_curSynthName, _currentSpeechSettings, temporary=False))
		_curSynthName = None
		_currentSpeechSettings = None
	seq.append(speech.commands.EndUtteranceCommand())
	return seq


def startOrStopAutomaticReadingVoice(start=True):
	seq = []
	autoReadingWithOption = _addonConfigManager.getAutoReadingWithOption()
	if autoReadingWithOption == AutoReadingWith_Beep:
		seq.append(speech.commands.EndUtteranceCommand())
		if start:
			seq.append(speech.commands.BeepCommand(600, 40))
		else:
			seq.append(speech.commands.BeepCommand(400, 40))
		seq.append(speech.commands.EndUtteranceCommand())
	elif autoReadingWithOption == AutoReadingWith_Voice:
		seq = formatVoiceAutoReadingStartSequence(start)
	return seq


def formatRevisionAutoReadingSequence(rangeObj, revision, text):
	seq = []
	autoReadingWithOption = _addonConfigManager.getAutoReadingWithOption()
	autoReading = autoReadingWithOption != AutoReadingWith_CurrentVoice
	if revision:
		# speak text and change voice if autoReading
		if not autoReading:
			seq.append(text)
			return seq
		from .ww_revisions import Revision
		rev = Revision(None, rangeObj.Revisions[1])
		seq.append(rev.FormatRevisionTypeAndAuthorText())
		seq .extend(startOrStopAutomaticReadingVoice())
	else:
		# restore current voice if autoReading and speak text
		if autoReading:
			seq.extend(startOrStopAutomaticReadingVoice(False))
		else:
			seq.append(text)
	return seq


def formatDeletedRevisionAutoReadingSequence(rangeObj, revision, text):
	seq = []
	if not revision or rangeObj.Revisions.Count == 0:
		seq.append(text)
		return seq
	from .ww_revisions import Revision
	rev = Revision(None, rangeObj.Revisions[1])
	seq.append(rev.FormatRevisionTypeAndAuthorText())
	seq.append(speech.commands.EndUtteranceCommand())
	autoReadingWithOption = _addonConfigManager.getAutoReadingWithOption()
	if autoReadingWithOption == AutoReadingWith_Voice:
		seq.extend(startOrStopAutomaticReadingVoice())
		seq.append(speech.commands.EndUtteranceCommand())
		seq.append(rev.text)
		seq.append(speech.commands.EndUtteranceCommand())
		seq.extend(startOrStopAutomaticReadingVoice(False))
	elif autoReadingWithOption == AutoReadingWith_Beep:
		seq.append(speech.commands.BeepCommand(600, 40))
		seq.append(speech.commands.EndUtteranceCommand())
		seq.append(rev.text)
		seq.append(speech.commands.EndUtteranceCommand())
		seq.append(speech.commands.BeepCommand(400, 40))
	else:
		seq.append(rev.text)
	seq.append(speech.commands.EndUtteranceCommand())
	return seq


class AutomaticReadingWordTextInfo(FormatFieldSpeech, textInfos.TextInfo):
	def getCommentFormatFieldSpeech(self, commentReference):
		def getHasCommentTextSequence(comment):
			if comment is textInfos.CommentType.DRAFT:
				# Translators: Reported when text contains a draft comment.
				text = NVDAString("has draft comment")
			elif comment is textInfos.CommentType.RESOLVED:
				# Translators: Reported when text contains a resolved comment.
				text = NVDAString("has resolved comment")
			else:  # generic
				# Translators: Reported when text contains a generic comment.
				text = NVDAString("has comment")
			return [text]
		if not (
			_addonConfigManager.toggleAutomaticReadingOption(False)
			and _addonConfigManager.toggleAutoCommentReadingOption(False)
		):
			return getHasCommentTextSequence(commentReference)
		# with UIA enabled, commentReference is a boolean (True)
		# with IAccessible, it's a string
		if type(commentReference) is not str:
			from NVDAObjects.UIA.wordDocument import getCommentInfoFromPosition
			commentInfo = getCommentInfoFromPosition(self)
			if not commentInfo:
				return getHasCommentTextSequence(commentReference)
			commentText = commentInfo["comment"]
			commentAuthor = commentInfo["author"]
		else:
			from ..ww_comments import Comment
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
			if not commentText:
				return getHasCommentTextSequence(commentReference)

		seq = []
		if commentReference is textInfos.CommentType.DRAFT:
			# Translators: Reported when text contains a draft comment.
			text = _("draft comment of %s:") % commentAuthor if commentAuthor != "" else _("comment:") % commentText
		elif commentReference is textInfos.CommentType.RESOLVED:
			# Translators: Reported when text contains a resolved comment.
			text = _("resolved comment of %s:") % (
				commentAuthor if commentAuthor != "" else _("comment:") % commentText
			)
		else:
			# Translators: Reported when text contains a comment.
			text = _("comment of %s:") % commentAuthor if commentAuthor != "" else _("comment:") % commentText
		seq.append(text)
		seq.extend(formatAutoSpeechSequence([commentText]))
		return seq

	def getFootnote(self, footnoteReference):
		from ..ww_footnotes import Footnote
		try:
			offset = int(footnoteReference)
		except ValueError:
			return ""
		if hasattr(self.obj, "rootNVDAObject"):
			doc = self.obj.rootNVDAObject.WinwordApplicationObject.ActiveDocument
		else:
			doc = self.obj.WinwordApplicationObject.ActiveDocument
		try:
			footnoteObj = doc.FootNotes[offset]
			footnote = Footnote(self.obj, footnoteObj)
			return footnote.text
		except Exception:
			return ""

	def getEndNote(self, endnoteReference):
		from ..ww_endnotes import Endnote
		try:
			offset = int(endnoteReference)
		except ValueError:
			return ""
		if hasattr(self.obj, "rootNVDAObject"):
			doc = self.obj.rootNVDAObject.WinwordApplicationObject.ActiveDocument
		else:
			doc = self.obj.WinwordApplicationObject.ActiveDocument
		try:
			endnoteObj = doc.EndNotes[offset]
			endnote = Endnote(self.obj, endnoteObj)
			return endnote.text
		except Exception:
			return ""

	def getAutomaticReadingSequence(self, attrs):
		seq = []
		role = attrs.get('role', controlTypes.Role.UNKNOWN)
		value = attrs.get('value', "")
		extendedRole = attrs.get("extendedRole")
		if controlTypes.Role.FOOTNOTE in (role, extendedRole) and (
			_addonConfigManager.toggleAutomaticReadingOption(False)
			and _addonConfigManager.toggleAutoFootnoteReadingOption(False)):
			footnote = self.getFootnote(value)
			if footnote != "":
				seq.extend(formatAutoSpeechSequence([footnote]))
		elif controlTypes.Role.ENDNOTE in (role, extendedRole) and (
			_addonConfigManager.toggleAutomaticReadingOption(False)
			and _addonConfigManager.toggleAutoEndnoteReadingOption(False)):
			endnote = self.getEndNote(value)
			if endnote != "":
				seq.extend(formatAutoSpeechSequence([endnote]))
		return seq

	def getControlFieldSpeech(
		self,
		attrs: textInfos.ControlField,
		ancestorAttrs: List[textInfos.Field],
		fieldType: str,
		formatConfig: Optional[Dict[str, bool]] = None,
		extraDetail: bool = False,
		reason: Optional[OutputReason] = None
	) -> SpeechSequence:
		out = super().getControlFieldSpeech(
			attrs, ancestorAttrs, fieldType, formatConfig, extraDetail, reason)
		outofString = NVDAString("out of %s") % ""
		if len(out) and outofString not in out[0]:
			out.extend(self.getAutomaticReadingSequence(attrs))
		return out


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


def getSynthDisplayInfos(synth, synthConf, outputDeviceName):
	conf = synthConf
	id = "id"
	import autoSettingsUtils.driverSetting
	numericSynthSetting = [autoSettingsUtils.driverSetting.NumericDriverSetting, ]
	booleanSynthSetting = [autoSettingsUtils.driverSetting.BooleanDriverSetting, ]
	textList = []
	for setting in synth.supportedSettings:
		settingID = getattr(setting, id)
		if type(setting) in numericSynthSetting:
			info = str(conf[settingID])
			textList.append((setting.displayName, info))
		elif type(setting) in booleanSynthSetting:
			info = _("yes") if conf[settingID] else _("no")
			textList.append((setting.displayName, info))
		else:
			if hasattr(synth, "available%ss" % settingID.capitalize()):
				tempList = list(getattr(synth, "available%ss" % settingID.capitalize()).values())
				cur = conf[settingID]
				i = [x.id for x in tempList].index(cur)
				v = tempList[i].displayName
				info = v
				textList.append((setting.displayName, info))
	d = {}
	i=1
	if NVDAVersion >= [2025, 1]:
		# Translators:  label to report synthesizer output device .
		d[str(1)] = [_("Audio output device"), outputDeviceName]
		i+= 1
	for label, val in textList:
		d[str(i )] = (label, val)
		i += 1
	return d


def getCurrentSpeechSettings():
	currentSynth = synthDriverHandler.getSynth()
	d = {}
	d[SCT_Speech] = config.conf[SCT_Speech].dict()
	for key in config.conf[SCT_Speech].copy():
		val = config.conf[SCT_Speech][key]
		if type(val) is config.AggregatedSection and key not in [SCT_Many, currentSynth.name]:
			del d[SCT_Speech][key]
	outputDeviceName = ""
	# for nvda >= 2025.1, outputDevice is stored in "audio" section instead of "speech" section
	# and it is stored by its id instead its name
	if "outputDevice" in config.conf[SCT_Audio]:
		outputDevice = config.conf[SCT_Audio]["outputDevice"]
		d[SCT_Speech]["outputDevice"] = outputDevice
		from utils import mmdevice
		deviceIds, deviceNames = zip(*mmdevice._getOutputDevices(includeDefault=True))
		try:
			outputDeviceName = deviceNames[deviceIds.index(outputDevice)]
		except ValueError:
			pass
	
	d["SynthDisplayInfos"] = getSynthDisplayInfos(currentSynth, d[SCT_Speech][currentSynth.name], outputDeviceName)
	# nvda 2024.4: include CLDR check box is replaced by symbolDictionnaries list
	# so we exclude from setting to keep and restore
	if "includeCLDR" in d[SCT_Speech]:
		del d[SCT_Speech]["includeCLDR"]
	if "symbolDictionaries" in d[SCT_Speech]:
		del d[SCT_Speech]["symbolDictionaries"]
	return d


def saveCurrentSpeechSettings():
	autoReadingSynth = _addonConfigManager.getAutoReadingSynthSettings()
	curSynthName = synthDriverHandler.getSynth().name
	currentSpeechSettings = getCurrentSpeechSettings()
	if autoReadingSynth:
		if confirm_YesNo(
			_(
				"There are already voice settings saved for automatic reading.\n"
				"Do you want to replace them with the current voice settings for Word?"
			),
			# Translators: title of message box
			"{addon} - {title}" .format(addon=_curAddon.manifest["summary"], title=_("Warning")),
		) != ReturnCode.YES:
			return
	autoReadingSynth = {}
	autoReadingSynth["synthName"] = curSynthName 
	autoReadingSynth.update(currentSpeechSettings)
	_addonConfigManager.saveAutoReadingSynthSettings(autoReadingSynth)
	# Translators: message to user to report automatic reading voice record.
	wx.CallLater(50, ui.message, _("Current voice settings have beenset for automatic reading"))


# to memorize the suspended synth settings when automacic reading synth is running
_suspendedSynth = (None, None)
settingsBeforeSwitch = None
def switchToAutomaticReadingSynth():
	global settingsBeforeSwitch 
	autoReadingSynth = _addonConfigManager.getAutoReadingSynthSettings()
	if not autoReadingSynth:
		return
	curSynthName = synthDriverHandler.getSynth().name
	currentSpeechSettings = getCurrentSpeechSettings()
	(autoReadingSynthName, autoReadingSynthSpeechSettings) = (autoReadingSynth["synthName"], autoReadingSynth)
	if settingsBeforeSwitch  is None:
		settingsBeforeSwitch  = (curSynthName, currentSpeechSettings )
		_setSynth(autoReadingSynthName, autoReadingSynthSpeechSettings, temporary=False)
		msg = _("Automatic reading voice")
	else:
		(curSynthName, currentSpeechSettings ) = settingsBeforeSwitch  
		settingsBeforeSwitch   = None
		_setSynth(curSynthName, currentSpeechSettings , temporary=False) 
		msg = _("Current voice for Word")
	ui.message(msg)
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
	synthDriverHandler.setSynth(None)
	synthSpeechConfig = config.conf[SCT_Speech] .dict()
	if speechSettings is not None:
		synthSpeechConfig.update(speechSettings[SCT_Speech])
	newOutputDevice = synthSpeechConfig["outputDevice"]
	if "outputDevice" in config.conf[SCT_Audio]:
		#curOutputDevice = config.conf[SCT_Audio]["outputDevice"] 
		config.conf[SCT_Audio]["outputDevice"] = newOutputDevice
		del synthSpeechConfig["outputDevice"]
	config.conf[SCT_Speech] = synthSpeechConfig.copy()
	synthDriverHandler.setSynth(synthName)
	config.conf[SCT_Speech] = synthSpeechConfig.copy()


# used for saving current nvda cancelSpeech
# we must hake nvda cancelSpeech to restore previous synth when cancel speech occurs during automatic reading
_NVDACancelSpeech = None


def onSpeechCanceled():
	if _suspendedSynth[0] is None:
		return
	# restore main synth
	_setSynth(_suspendedSynth[0], _suspendedSynth[1], temporary=False)


def myCancelSpeech():
	global _suspendedSynth
	_NVDACancelSpeech()
	onSpeechCanceled()


def initialize():
	if NVDAVersion >= [2024, 1]:
		from speech.extensions import speechCanceled
		speechCanceled.register(onSpeechCanceled)
		return
	global _NVDACancelSpeech
	if _NVDACancelSpeech is not None:
		log.warning("cancelSpeech already patched")
		return
	_NVDACancelSpeech = speech.cancelSpeech
	speech.cancelSpeech = myCancelSpeech


def terminate():
	if NVDAVersion >= [2024, 1]:
		from speech.extensions import speechCanceled
		speechCanceled.unregister(onSpeechCanceled)
		return
	global _NVDACancelSpeech
	# to restore suspended synth
	speech.cancelSpeech()
	if _NVDACancelSpeech is not None:
		speech.cancelSpeech = _NVDACancelSpeech
		_NVDACancelSpeech = None
