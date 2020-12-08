# appModules\winword\ww_funct.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
import api
import ui
import speech

addonHandler.initTranslation()

# word versions:
Word_11 = 11
Word_12 = 12
Word_14 = 14
Word_15 = 15
# Word application owner names:
WordApp_2003 = "winword"
WordApp_2007 = "wwlib"  # also true for 2010

# story types
wdMainTextStory = 1  # Main text story
wdFootnotesStory = 2  # Footnotes story
wdEndnotesStory = 3  # Endnotes story
wdCommentsStory = 4  # Comments story
wdTextFrameStory = 5  # Text frame story
wdEvenPagesHeaderStory = 6  # Even pages header story
wdPrimaryHeaderStory = 7  # Primary header story
wdEvenPagesFooterStory = 8  # Even pages footer story
wdPrimaryFooterStory = 9  # Primary footer story
wdFirstPageHeaderStory = 10  # First page header story
wdFirstPageFooterStory = 11  # First page footer story
wdFootnoteSeparatorStory = 12  # Footnote separator story
wdFootnoteContinuationSeparatorStory = 13  # Footnote continuation separator story # noqa:E501
wdFootnoteContinuationNoticeStory = 14  # Footnote continuation notice story
wdEndnoteSeparatorStory = 15  # Endnote separator story
wdEndnoteContinuationSeparatorStory = 16  # Endnote continuation separator story # noqa:E501
wdEndnoteContinuationNoticeStory = 17  # Endnote continuation notice story


def isWordActive():
	oFocus = api.getFocusObject()
	sOwner = oFocus.appModule.appName
	return (WordApp_2007 in sOwner) or (WordApp_2003 in sOwner)


def getPaneName(storyType):
	_storyTypeNames = {
		wdMainTextStory: _("Main text story"),
		wdFootnotesStory: _("Footnotes story"),
		wdEndnotesStory: _("Endnotes story"),
		wdCommentsStory: _("Comments story"),
		wdTextFrameStory: _("Text frame story"),
		wdEvenPagesHeaderStory: _("Even pages header story"),
		wdPrimaryHeaderStory: _("Primary header story"),
		wdEvenPagesFooterStory:  _("Even pages footer story"),
		wdPrimaryFooterStory: _("Primary footer story"),
		wdFirstPageHeaderStory: _("First page header story"),
		wdFirstPageFooterStory: _("First page footer story"),
		wdFootnoteSeparatorStory: _("Footnote separator story"),
		wdFootnoteContinuationSeparatorStory: _("Footnote continuation separator story"),  # noqa:E501
		wdFootnoteContinuationNoticeStory: _("Footnote continuation notice story"),
		wdEndnoteSeparatorStory: _("Endnote separator story"),
		wdEndnoteContinuationSeparatorStory: _("Endnote continuation separator story"),  # noqa:E501
		wdEndnoteContinuationNoticeStory: _("Endnote continuation notice story")
		}
	return _storyTypeNames[storyType]


def sayActivePane():
	oFocus = api.getFocusObject()
	storyType = oFocus._WinwordSelectionObject.StoryType
	sPaneName = getPaneName(storyType)
	if storyType:
		oActiveWindow = oFocus._WinwordWindowObject.Application.ActiveWindow
		oActivePane = oActiveWindow.ActivePane
		try:
			oFrameset = oActivePane.Frameset
		except:  # noqa:E722
			oFrameset = None
		if oFrameset:
			ui.message(_("Frame {name} document pane {index}") .format(
				name=oFrameset.FrameName, index=str(oActivePane.index)))
		elif sPaneName != "":
			ui.message(sPaneName)


def sayActivePaneName(speechMode):
	speech.speechMode = speechMode
	oFocus = api.getFocusObject()
	appVersion = oFocus.appModule.productVersion.split(".")[0]
	if (appVersion >= Word_12) and not isWordActive():
		wn_ApplyStyles = "appliquer les styles",
		wn_Styles = "Styles",
		sRealName = api.getForegroundObject().name
		if sRealName in [wn_ApplyStyles, wn_Styles]:
			ui.message(wn_Styles)
		return
	sayActivePane()
