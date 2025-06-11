# shared\winword\ww_configManager.py
# a part of wordAccessEnhancement add-on
# Copyright 2019-2025, paulber19
# released under GPL.


from logHandler import log
import addonHandler
import os
import config
import wx
import gui
import globalVars
from configobj import ConfigObj
from configobj.validate import Validator, ValidateError
from io import StringIO
import _pickle as cPickle
from versionInfo import version_year, version_major


NVDAVersion = [version_year, version_major]
addonHandler.initTranslation()

# config section
SCT_General = "General"
SCT_Document = "Document"
SCT_AutoReport = "AutomaticReport"
SCT_AutomaticReadingSynthSettings = "AutomaticReadingSynthSettings"

# general section items
ID_ConfigVersion = "ConfigVersion"
ID_AutoUpdateCheck = "AutoUpdateCheck"
ID_UpdateReleaseVersionsToDevVersions = "UpdateReleaseVersionsToDevVersions"
# move in Document section
# ID_SkipEmptyParagraphs= "SkipEmptyParagraphs"
# ID_PlaySoundOnSkippedParagraph = "PlaySoundOnSkippedParagraph"
# ID_UseQuickNavigationMode = "UseQuickNavigationMode"
ID_ElementsSearchMaxTime = "ElementsSearchMaxTime"
# Document section items
ID_SkipEmptyParagraphs = "SkipEmptyParagraphs"
ID_PlaySoundOnSkippedParagraph = "PlaySoundOnSkippedParagraph"
ID_UseQuickNavigationMode = "UseQuickNavigationMode"
ID_LoopInNavigationMode = "LoopInNavigationMode"
# Automatic report section items
ID_AutomaticReading = "AutomaticReading"
ID_CommentReport = "CommentReport"
ID_FootnoteReport = "FootnoteReport"
ID_EndnoteReport = "EndnoteReport"
ID_InsertedTextReport = "InsertedTextReport"
ID_DeletedTextReport = "DeletedTextReport"
ID_RevisedTextReport = "RevisedTextReport"
ID_ReportRevisionTypeWithAuthor = "ReportRevisionTypeWithAuthor"

ID_AutoReadingWith = "AutoReadingWith"
ID_AutoReadingSynthName = "AutoReadingSynthName"
ID_AutoReadingSynthSpeechSettings = "AutoReadingSynthSpeechSettings"
# values for AutoReadingWith option
AutoReadingWith_CurrentVoice = 0
AutoReadingWith_Beep = 1
AutoReadingWith_Voice = 2

_curAddon = addonHandler.getCodeAddon()
_addonName = _curAddon.manifest["name"]


def renameFile(file, dest):
	try:
		if os.path.exists(dest):
			os.remove(dest)
		os.rename(file, dest)
		log.debug("current configuration file: %s renamed to : %s" % (file, dest))
	except Exception:
		log.error("current configuration file: %s  cannot be renamed to: %s" % (file, dest))


class BaseAddonConfiguration(ConfigObj):
	_version = ""
	""" Add-on configuration file. It contains metadata about add-on . """
	_GeneralConfSpec = """[{section}]
	{idConfigVersion} = string(default = " ")

	""".format(section=SCT_General, idConfigVersion=ID_ConfigVersion)
	configspec = ConfigObj(StringIO("""# addon Configuration File
	{0}""".format(_GeneralConfSpec)
	), list_values=False, encoding="UTF-8")

	def __init__(self, input):
		""" Constructs an L{AddonConfiguration} instance from manifest string data
		@param input: data to read the addon configuration information
		@type input: a fie-like object.
		"""
		super(BaseAddonConfiguration, self).__init__(
			input,
			configspec=self.configspec,
			encoding='utf-8',
			default_encoding='utf-8')
		self.newlines = "\r\n"
		self._errors = []
		val = Validator()
		result = self.validate(val, copy=True, preserve_errors=True)
		if type(result) is dict:
			self._errors = self.getValidateErrorsText(result)
		else:
			self._errors = None

	def getValidateErrorsText(self, result):
		textList = []
		for name, section in result.items():
			if section is True:
				continue
			textList.append("section [%s]" % name)
			for key, value in section.items():
				if isinstance(value, ValidateError):
					textList.append(
						'key "{}": {}'.format(
							key, value))
		return "\n".join(textList)

	@property
	def errors(self):
		return self._errors


class AddonConfiguration10(BaseAddonConfiguration):
	_version = "1.0"
	_GeneralConfSpec = """[{section}]
	{configVersion} = string(default = "1.0")
	{autoUpdateCheck} = boolean(default=True)
	{updateReleaseVersionsToDevVersions} = boolean(default=False)
	{skipEmptyParagraphs} = boolean(default=True)
	{playSoundOnSkippedParagraph} =boolean(default=True)
	{useQuickNavigationMode} = boolean(default=False)
	""".format(
		section=SCT_General,
		configVersion=ID_ConfigVersion,
		autoUpdateCheck=ID_AutoUpdateCheck,
		updateReleaseVersionsToDevVersions=ID_UpdateReleaseVersionsToDevVersions,
		skipEmptyParagraphs=ID_SkipEmptyParagraphs,
		playSoundOnSkippedParagraph=ID_PlaySoundOnSkippedParagraph,
		useQuickNavigationMode=ID_UseQuickNavigationMode)

	configspec = ConfigObj(StringIO("""# addon Configuration File
	{0}""".format(_GeneralConfSpec)
	), list_values=False, encoding="UTF-8")


class AddonConfiguration20(BaseAddonConfiguration):
	_version = "2.0"
	_GeneralConfSpec = """[{section}]
	{configVersion} = string(default = "2.0")
	{autoUpdateCheck} = boolean(default=True)
	{updateReleaseVersionsToDevVersions} = boolean(default=False)
	""".format(
		section=SCT_General,
		configVersion=ID_ConfigVersion,
		autoUpdateCheck=ID_AutoUpdateCheck,
		updateReleaseVersionsToDevVersions=ID_UpdateReleaseVersionsToDevVersions)

	_DocumentConfSpec = """[{section}]
	{skipEmptyParagraphs} = boolean(default=True)
	{playSoundOnSkippedParagraph} =boolean(default=True)
	{useQuickNavigationMode} = boolean(default=False)
	""".format(
		section=SCT_Document,
		skipEmptyParagraphs=ID_SkipEmptyParagraphs,
		playSoundOnSkippedParagraph=ID_PlaySoundOnSkippedParagraph,
		useQuickNavigationMode=ID_UseQuickNavigationMode)
	_AutoReportConfSpec = """[{section}]
	{automaticReading} = boolean(default=False)
	{commentReport} =boolean(default=True)
	{footnoteReport} = boolean(default=True)
	{autoReadingWith} = integer(default=1)
	""".format(
		section=SCT_AutoReport,
		automaticReading=ID_AutomaticReading,
		commentReport=ID_CommentReport,
		footnoteReport=ID_FootnoteReport,
		autoReadingWith=ID_AutoReadingWith)

	configspec = ConfigObj(StringIO("""# addon Configuration File
	{0}{1}{2}""".format(_GeneralConfSpec, _DocumentConfSpec, _AutoReportConfSpec)
	), list_values=False, encoding="UTF-8")

	def mergeWithPreviousConfigurationVersion(self, previousConfig):
		previousVersion = previousConfig[SCT_General][ID_ConfigVersion]
		if previousVersion != "1.0":
			log.warning(
				"%s: AddonConfigManager mergeWithPreviousConfiguration error: bad previous configuration version number"
				% _addonName)
			return
		# configuration 1 to 2
		# 3 options are moved from General section to Document section
		self[SCT_Document][ID_SkipEmptyParagraphs] = previousConfig[SCT_General][ID_SkipEmptyParagraphs]
		self[SCT_Document][ID_PlaySoundOnSkippedParagraph] = (
			previousConfig[SCT_General][ID_PlaySoundOnSkippedParagraph])
		self[SCT_Document][ID_UseQuickNavigationMode] = (
			previousConfig[SCT_General][ID_UseQuickNavigationMode])
		log.warning("%s: Merge with previous configuration version: %s" % (_addonName, previousVersion))


class AddonConfiguration21(BaseAddonConfiguration):
	_version = "2.1"
	_GeneralConfSpec = """[{section}]
	{configVersion} = string(default = "2.1")
	{autoUpdateCheck} = boolean(default=True)
	{updateReleaseVersionsToDevVersions} = boolean(default=False)
	""".format(
		section=SCT_General,
		configVersion=ID_ConfigVersion,
		autoUpdateCheck=ID_AutoUpdateCheck,
		updateReleaseVersionsToDevVersions=ID_UpdateReleaseVersionsToDevVersions)

	_DocumentConfSpec = """[{section}]
	{skipEmptyParagraphs} = boolean(default=True)
	{playSoundOnSkippedParagraph} =boolean(default=True)
	{useQuickNavigationMode} = boolean(default=False)
	{elementsSearchMaxTime} = integer(default = 20)
	{loopInNavigationMode} = boolean(default=False)
	""".format(
		section=SCT_Document,
		skipEmptyParagraphs=ID_SkipEmptyParagraphs,
		playSoundOnSkippedParagraph=ID_PlaySoundOnSkippedParagraph,
		useQuickNavigationMode=ID_UseQuickNavigationMode,
		elementsSearchMaxTime=ID_ElementsSearchMaxTime,
		loopInNavigationMode=ID_LoopInNavigationMode
	)

	_AutoReportConfSpec = """[{section}]
	{automaticReading} = boolean(default=False)
	{commentReport} =boolean(default=True)
	{footnoteReport} = boolean(default=True)
	{endnoteReport} = boolean(default=True)
	{insertedTextReport} = boolean(default=True)
	{deletedTextReport} = boolean(default=True)
		{revisedTextReport} = boolean(default=True)
				{reportRevisionTypeWithAuthor} = boolean(default=True)
	{autoReadingWith} = integer(default=1)
	""".format(
		section=SCT_AutoReport,
		automaticReading=ID_AutomaticReading,
		commentReport=ID_CommentReport,
		footnoteReport=ID_FootnoteReport,
		endnoteReport=ID_EndnoteReport,
		insertedTextReport=ID_InsertedTextReport,
		deletedTextReport=ID_DeletedTextReport,
		revisedTextReport=ID_RevisedTextReport,
		reportRevisionTypeWithAuthor=ID_ReportRevisionTypeWithAuthor,
		autoReadingWith=ID_AutoReadingWith)

	configspec = ConfigObj(StringIO("""# addon Configuration File
	{0}{1}{2}""".format(_GeneralConfSpec, _DocumentConfSpec, _AutoReportConfSpec)
	), list_values=False, encoding="UTF-8")

	def mergeWithPreviousConfigurationVersion(self, previousConfig):
		previousVersion = previousConfig[SCT_General][ID_ConfigVersion]
		if previousVersion != "1.0":
			# nothing to do
			return

		# configuration 1.0 to 2.1
		# 3 options are moved from General section to Document section
		self[SCT_Document][ID_SkipEmptyParagraphs] = (
			previousConfig[SCT_General][ID_SkipEmptyParagraphs])
		self[SCT_Document][ID_PlaySoundOnSkippedParagraph] = (
			previousConfig[SCT_General][ID_PlaySoundOnSkippedParagraph])
		self[SCT_Document][ID_UseQuickNavigationMode] = (
			previousConfig[SCT_General][ID_UseQuickNavigationMode])
		log.warning("%s: Merge with previous configuration version: %s" % (_addonName, previousVersion))


PREVIOUSCONFIGURATIONFILE_SUFFIX = ".prev"
DELETECONFIGURATIONFILE_SUFFIX = ".delete"


class AddonConfigurationManager():
	_currentConfigVersion = "2.1"
	_versionToConfiguration = {
		"1.0": AddonConfiguration10,
		"2.0": AddonConfiguration20,
		"2.1": AddonConfiguration21,
	}
	# keep the synthetizer used for automatic reading
	_autoReadingSynth = None

	def __init__(self):
		self.configFileName = "%sAddon.ini" % _addonName
		self.autoReadingSynthFileName = "%s_autoReadingSynth.pickle" % _addonName
		self.loadSettings()
		config.post_configSave.register(self.handlePostConfigSave)

	def warnConfigurationReset(self):
		from messages import alert
		wx.CallLater(
			100,
			alert,
			# Translators: A message warning configuration reset.
			_(
				"The configuration file of the add-on contains errors. "
				"The configuration has been reset to factory defaults"),
			# Translators: title of message box
			"{addon} - {title}" .format(addon=_curAddon.manifest["summary"], title=_("Warning")),
		)

	def loadSettings(self):
		self.restorePreviousAutoReadingSynth()
		self.getAutoReadingSynthSettings()
		userConfig = globalVars.appArgs.configPath
		self.addonConfigFile = addonConfigFile = os.path.join(userConfig, self.configFileName)
		self.oldConfigFile = addonConfigFile + PREVIOUSCONFIGURATIONFILE_SUFFIX
		# after add-on installation and and the user does not want to keep the configuration
		# the configuration has been renamed with .delete extension
		# if this file exists, it must be deleted
		self.deleteConfigFile = self.addonConfigFile + DELETECONFIGURATIONFILE_SUFFIX
		if os.path.exists(self.deleteConfigFile):
			os.remove(self.deleteConfigFile)
		doMerge = True
		if os.path.exists(addonConfigFile):
			# there is allready a config file
			try:
				baseConfig = BaseAddonConfiguration(addonConfigFile)
				if baseConfig.errors:
					e = Exception("Error parsing configuration file:\n%s" % baseConfig.errors)
					raise e
				if baseConfig[SCT_General][ID_ConfigVersion] != self._currentConfigVersion:
					# it's an old config, but old config file must not exist here.
					# Must be deleted
					os.remove(addonConfigFile)
					log.warning(
						"%s: Old configuration version found. Config file is removed: %s" % (_addonName, addonConfigFile))
				else:
					# it's the same version of config, so no merge
					doMerge = False
			except Exception as e:
				log.warning(e)
				# error on reading config file, so delete it
				os.remove(addonConfigFile)
				self.warnConfigurationReset()
				log.warning(
					"%s Addon configuration file error: configuration reset to factory defaults" % _addonName)

		if os.path.exists(addonConfigFile):
			self.addonConfig =\
				self._versionToConfiguration[self._currentConfigVersion](addonConfigFile)
			if self.addonConfig.errors:
				log.warning(self.addonConfig.errors)
				log.warning(
					"%s Addon configuration file error: configuration reset to factory defaults" % _addonName)
				os.remove(addonConfigFile)
				self.warnConfigurationReset()
				# reset configuration to factory defaults
				self.addonConfig =\
					self._versionToConfiguration[self._currentConfigVersion](None)
				self.addonConfig.filename = addonConfigFile
				doMerge = False
		else:
			# no add-on configuration file found
			self.addonConfig =\
				self._versionToConfiguration[self._currentConfigVersion](None)
			self.addonConfig.filename = addonConfigFile
		# merge step
		if os.path.exists(self.oldConfigFile):
			if doMerge:
				self.mergeSettings(self.oldConfigFile)
			os.remove(self.oldConfigFile)
		if not os.path.exists(addonConfigFile):
			self.saveSettings(True)

	def mergeSettings(self, previousConfigFile):
		baseConfig = BaseAddonConfiguration(previousConfigFile)
		previousVersion = baseConfig[SCT_General][ID_ConfigVersion]
		if previousVersion not in self._versionToConfiguration:
			log.warning("%s: Configuration merge error: unknown previous configuration version number" % _addonName)
			return
		previousConfig = self._versionToConfiguration[previousVersion](previousConfigFile)
		if previousVersion == self.addonConfig[SCT_General][ID_ConfigVersion]:
			# same config version, update data from previous config
			self.addonConfig.update(previousConfig)
			log.warning("%s: Configuration updated with previous configuration file" % _addonName)
			return
		# different config version, so do a merge with previous config.
		try:
			self.addonConfig.mergeWithPreviousConfigurationVersion(previousConfig)
		except Exception:
			pass

	def restorePreviousAutoReadingSynth(self):
		userConfig = globalVars.appArgs.configPath
		autoReadingSynthFile = os.path.join(userConfig, self.autoReadingSynthFileName)
		prevAutoReadingSynthFile = autoReadingSynthFile + PREVIOUSCONFIGURATIONFILE_SUFFIX
		if os.path.exists(prevAutoReadingSynthFile):
			# rename it
			renameFile(prevAutoReadingSynthFile, autoReadingSynthFile)
		# after add-on installation and and the user does not want to keep the configuration
		# the autoReadingSynth File has been renamed with .delete extension
		# if this file exists, it must be deleted
		deleteAutoReadingSynthFile = autoReadingSynthFile + DELETECONFIGURATIONFILE_SUFFIX
		if os.path.exists(deleteAutoReadingSynthFile):
			os.remove(deleteAutoReadingSynthFile)

	def handlePostConfigSave(self):
		self.saveSettings(True)

	def canConfigurationBeSaved(self, force):
		# Never save config or state if running securely or if running from the launcher.
		try:
			# for NVDA version >= 2023.2
			from NVDAState import shouldWriteToDisk
			writeToDisk = shouldWriteToDisk()
		except ImportError:
			# for NVDA version < 2023.2
			writeToDisk = not (globalVars.appArgs.secure or globalVars.appArgs.launcher)
		if not writeToDisk:
			log.debug("Not writing add-on configuration, either --secure or --launcher args present")
			return False
		# after add-on installation and and the user does not want to keep the configuration
		# the configuration has been renamed with .delete extension
		# if this file exists, configuration should not be saved
		if os.path.exists(self.deleteConfigFile):
			return False
		# after an add-on removing, configuration is deleted
			# so  don't save configuration if there is no nvda restart
		if _curAddon.isPendingRemove:
			return False
		# We don't save the configuration, in case the user
			# would not have checked the "Save configuration on exit
			# " checkbox in General settings and force is False
		if not force and not config.conf['general']['saveConfigurationOnExit']:
			return False
		return True

	def _saveAutoReadingSynthSettings(self):
		userConfig = config.getUserDefaultConfigPath()
		autoReadingSynthFile = os.path.join(userConfig, self.autoReadingSynthFileName)
		with open(autoReadingSynthFile, 'wb') as f:
			cPickle.dump(self._autoReadingSynth, f, 0)
		# if an installation took place, the configuration file was renamed.
		# so you have to do the same thing after saving
		dest = autoReadingSynthFile + ".prev"
		if os.path.exists(dest):
			renameFile(autoReadingSynthFile, dest)

	def saveSettings(self, force=False):
		if self.addonConfig is None:
			return
		if not self.canConfigurationBeSaved(force):
			return
		self._saveAutoReadingSynthSettings()
		try:
			val = Validator()
			self.addonConfig.validate(val, copy=True)
			self.addonConfig.write()
			log.warning("%s: configuration saved" % _addonName)
			# if an installation took place, the configuration file was renamed.
			# so you have to do the same thing after saving
			if os.path.exists(self.oldConfigFile):
				renameFile(self.addonConfigFile, self.oldConfigFile)
		except Exception:
			log.warning("%s: Could not save configuration - probably read only file system" % _addonName)

	def terminate(self):
		self.saveSettings()
		config.post_configSave.unregister(self.handlePostConfigSave)

	def _toggleOption(self, sct, id, toggle=True):
		conf = self.addonConfig
		if toggle:
			conf[sct][id] = not conf[sct][id]
		return conf[sct][id]

	def _toggleGeneralOption(self, id, toggle=True):
		return self._toggleOption(SCT_General, id, toggle)

	def toggleAutoUpdateCheck(self, toggle=True):
		return self._toggleGeneralOption(ID_AutoUpdateCheck, toggle)

	def toggleUpdateReleaseVersionsToDevVersions(self, toggle=True):
		return self._toggleGeneralOption(
			ID_UpdateReleaseVersionsToDevVersions, toggle)

	def _toggleDocumentOption(self, id, toggle=True):
		return self._toggleOption(SCT_Document, id, toggle)

	def toggleSkipEmptyParagraphsOption(self, toggle=True):
		return self._toggleDocumentOption(ID_SkipEmptyParagraphs, toggle)

	def togglePlaySoundOnSkippedParagraphOption(self, toggle=True):
		return self._toggleDocumentOption(ID_PlaySoundOnSkippedParagraph, toggle)

	def toggleUseQuickNavigationModeOption(self, toggle=True):
		return self._toggleDocumentOption(ID_UseQuickNavigationMode, toggle)

	def toggleLoopInNavigationModeOption(self, toggle=True):
		return self._toggleDocumentOption(ID_LoopInNavigationMode, toggle)

	def getElementsSearchMaxTime(self):
		conf = self.addonConfig[SCT_Document]
		return conf[ID_ElementsSearchMaxTime]

	def setElementsSearchMaxTime(self, maxTime):
		conf = self.addonConfig[SCT_Document]
		conf[ID_ElementsSearchMaxTime] = maxTime

	def _toggleAutoReportOption(self, id, toggle=True):
		return self._toggleOption(SCT_AutoReport, id, toggle)

	def toggleAutomaticReadingOption(self, toggle=True):
		return self._toggleAutoReportOption(ID_AutomaticReading, toggle)

	def getAutoReadingWithOption(self):
		conf = self.addonConfig[SCT_AutoReport]
		return conf[ID_AutoReadingWith]

	def setAutoReadingWithOption(self, option):
		conf = self.addonConfig[SCT_AutoReport]
		conf[ID_AutoReadingWith] = option

	def toggleAutoCommentReadingOption(self, toggle=True):
		return self._toggleAutoReportOption(ID_CommentReport, toggle)

	def toggleAutoFootnoteReadingOption(self, toggle=True):
		return self._toggleAutoReportOption(ID_FootnoteReport, toggle)

	def toggleAutoEndnoteReadingOption(self, toggle=True):
		return self._toggleAutoReportOption(ID_EndnoteReport, toggle)

	def toggleAutoInsertedTextReadingOption(self, toggle=True):
		return self._toggleAutoReportOption(ID_InsertedTextReport, toggle)

	def toggleAutoDeletedTextReadingOption(self, toggle=True):
		return self._toggleAutoReportOption(ID_DeletedTextReport, toggle)

	def toggleReportRevisionTypeWithAuthorOption(self, toggle=True):
		return self._toggleAutoReportOption(ID_ReportRevisionTypeWithAuthor, toggle)

	def toggleAutoRevisedTextReadingOption(self, toggle=True):
		return self._toggleAutoReportOption(ID_RevisedTextReport, toggle)
	def checkAutomaticReadingSynthSettingsCompatibility(self):
		from messages import alert
		# until nvda 2024.4, audio output device is stored by it name
		# since nvda 2025.1, it is stored by it id and for automatic voice  reading, itname is recorded in voice information
		if NVDAVersion != [2025, 1]or self._autoReadingSynth is None:
			return
		c = self._autoReadingSynth["SynthDisplayInfos"].copy()
		if c["1"][0] == _("Audio output device"):
			return
		self._autoReadingSynth = None
		self._saveAutoReadingSynthSettings()
		wx.CallLater(
			100,
			alert,
			# Translators: A message warning automatic reading synth settings clearedconfiguration reset.
			_(
				"To ensure compatibility with this NVDA's version, the voice settings saved for automatic reading had to be cleared.\n"
				"You need to register them again.\n"
				"Sorry for the inconvenience."),
			# Translators: title of message box
			"{addon} - {title}" .format(addon=_curAddon.manifest["summary"], title=_("Warning")),

		)


	
	def getAutoReadingSynthSettings(self):
		if self._autoReadingSynth is None:
			path = os.path.join(
				config.getUserDefaultConfigPath(),
				"wordAccessEnhancement_autoReadingSynth.pickle")
			if not os.path.exists(path):
				return None
			with open(path, 'rb') as f:
				self._autoReadingSynth = cPickle.load(f)
				log.debug("loading autoReadingSynth settings")
		if self._autoReadingSynth  is None:
			return
		# nvda 2024.4: include CLDR check box is replaced by symbolDictionnaries list
		# so we exclude from setting to keep and restore
		SCT_Speech = "speech"
		if "includeCLDR" in self._autoReadingSynth[SCT_Speech]:
			del self._autoReadingSynth[SCT_Speech]["includeCLDR"]
		if "symbolDictionaries" in self._autoReadingSynth[SCT_Speech]:
			del self._autoReadingSynth[SCT_Speech]["symbolDictionaries"]
		self.checkAutomaticReadingSynthSettingsCompatibility()
		return self._autoReadingSynth

	def saveAutoReadingSynthSettings(self, synthSettings):
		self._autoReadingSynth = synthSettings.copy()


# singleton for addon config manager
_addonConfigManager = AddonConfigurationManager()
