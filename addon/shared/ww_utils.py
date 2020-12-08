# shared\winword\ww_utils.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2020 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
# from logHandler import log
import winUser
import keyboardHandler
import gui
import time
import api
import speech
import wx
import config
from gui import mainFrame, guiHelper
import queueHandler
from ww_NVDAStrings import NVDAString

addonHandler.initTranslation()


def putWindowOnForeground(hwnd, sleepNb=10, sleepTime=0.1):
	winUser.setForegroundWindow(hwnd)
	try:
		winUser.setForegroundWindow(hwnd)
	except:  # noqa:E722
		pass
	for i in [sleepTime]*(sleepNb-1):
		time.sleep(i)
		if winUser.getForegroundWindow() == hwnd:
			return True
	# last chance
	keyboardHandler.KeyboardInputGesture.fromName("alt+Tab").send()
	keyboardHandler.KeyboardInputGesture.fromName("alt+Tab").send()
	time.sleep(sleepTime)
	if winUser.getForegroundWindow() == hwnd:
		return True
	return False


def leftClick(x, y):
	winUser.setCursorPos(x, y)
	winUser.mouse_event(winUser.MOUSEEVENTF_LEFTDOWN, 0, 0, None, None)
	winUser.mouse_event(winUser.MOUSEEVENTF_LEFTUP, 0, 0, None, None)


# winuser.h constant
SC_MAXIMIZE = 0xF030
WS_MAXIMIZE = 0x01000000
WM_SYSCOMMAND = 0x112


def maximizeWindow(hWnd):
	winUser.sendMessage(hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)


def isOpened(dialog):
	if dialog._instance is None:
		return False
	# Translators: the label of a message box dialog.
	msg = _("%s dialog is allready open") % dialog.title
	queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, msg)
	return True


class InformationDialog(wx.Dialog):
	_instance = None
	# Translators: this is the default title of Information dialog
	title = _("%s add-on - Informations")

	def __new__(cls, *args, **kwargs):
		if InformationDialog._instance is not None:
			return InformationDialog._instance
		return super(InformationDialog, cls).__new__(cls, *args, **kwargs)

	def __init__(self, parent, informationLabel="", information=""):
		if InformationDialog._instance is not None:
			return
		InformationDialog._instance = self
		curAddon = addonHandler.getCodeAddon()
		summary = curAddon.manifest['summary']
		title = self.title % summary
		super(InformationDialog, self).__init__(parent, wx.ID_ANY, title)
		self.informationLabel = informationLabel
		self.information = information
		self.doGui()

	def doGui(self):
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		sHelper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		# the text control
		sHelper.addItem(wx.StaticText(self, label=self.informationLabel))
		self.tc = sHelper.addItem(wx.TextCtrl(
			self,
			id=wx.ID_ANY,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH,
			size=(1000, 600)))
		self.tc.AppendText(self.information)
		self.tc.SetInsertionPoint(0)
		# the buttons
		bHelper = sHelper.addDialogDismissButtons(
			guiHelper.ButtonHelper(wx.HORIZONTAL))
		# Translators: label of copy to clipboard button
		copyToClipboardButton = bHelper.addButton(
			self, id=wx.ID_ANY, label=_("Co&py to Clipboard"))
		closeButton = bHelper.addButton(
			self, id=wx.ID_CLOSE, label=NVDAString("&Close"))
		mainSizer.Add(
			sHelper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		mainSizer.Fit(self)
		self.SetSizer(mainSizer)
		# events
		copyToClipboardButton.Bind(wx.EVT_BUTTON, self.onCopyToClipboardButton)
		closeButton.Bind(wx.EVT_BUTTON, lambda evt: self.Destroy())
		self.tc.SetFocus()
		self.SetEscapeId(wx.ID_CLOSE)

	def Destroy(self):
		InformationDialog._instance = None
		super(InformationDialog, self).Destroy()

	def onCopyToClipboardButton(self, event):
		if api.copyToClip(self.information):
			# Translators: message to user when the information has been copied
			# to clipboard
			text = _("Copied")
			speech.speakMessage(text)
			time.sleep(0.8)
			self.Close()
		else:
			# Translators: message to user when the information cannot be copied
			# to clipboard
			text = _("Error, the information cannot be copied to the clipboard")
			speech.speakMessage(text)

	@classmethod
	def run(cls, parent, informationLabel, information):
		if isOpened(InformationDialog):
			return
		if not parent:
			mainFrame.prePopup()
		d = InformationDialog(parent or mainFrame, informationLabel, information)
		d.CentreOnScreen()
		d.Show()
		if not parent:
			mainFrame.postPopup()


def makeAddonWindowTitle(dialogTitle):
	curAddon = addonHandler.getCodeAddon()
	addonSummary = curAddon.manifest['summary']
	return "%s - %s" % (addonSummary, dialogTitle)


def myMessageBox(
	message,
	caption=wx.MessageBoxCaptionStr,
	style=wx.OK | wx.CENTER,
	parent=None):
	# NVDA function modified to say in all cases the window content
	global isInMessageBox
	option = config.conf["presentation"]["reportObjectDescriptions"]
	config.conf["presentation"]["reportObjectDescriptions"] = True
	res = gui.messageBox(message, caption, style, parent or gui.mainFrame)
	config.conf["presentation"]["reportObjectDescriptions"] = option
	return res
