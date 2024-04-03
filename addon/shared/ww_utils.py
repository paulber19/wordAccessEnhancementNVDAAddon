# shared\winword\ww_utils.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2019-2024 paulber19
# This file is covered by the GNU General Public License.


import addonHandler
# from logHandler import log
import winUser
import keyboardHandler
import time
import speech
import queueHandler
import ui
try:
	# NVDA >= 2024.1
	speech.speech.SpeechMode.onDemand
	speakOnDemand = {"speakOnDemand": True}
except AttributeError:
	# NVDA <= 2023.3
	speakOnDemand = {}

addonHandler.initTranslation()


def putWindowOnForeground(hwnd, sleepNb=10, sleepTime=0.1):
	winUser.setForegroundWindow(hwnd)
	try:
		winUser.setForegroundWindow(hwnd)
	except Exception:
		pass
	for i in [sleepTime] * (sleepNb - 1):
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


def makeAddonWindowTitle(dialogTitle):
	curAddon = addonHandler.getCodeAddon()
	addonSummary = curAddon.manifest['summary']
	return "%s - %s" % (addonSummary, dialogTitle)


def getSpeechMode():
	try:
		# for nvda  version >= 2021.1
		return speech.getState().speechMode
	except AttributeError:
		return speech.speechMode


def setSpeechMode(mode):
	try:
		# for nvda version >= 2021.1
		speech.setSpeechMode(mode)
	except AttributeError:
		speech.speechMode = mode


def setSpeechMode_off():
	try:
		# for nvda version >= 2021.1
		speech.setSpeechMode(speech.SpeechMode.off)
	except AttributeError:
		speech.speechMode = speech.speechMode_off


def executeWithSpeakOnDemand(func, *args, **kwargs):
	from speech.speech import _speechState, SpeechMode
	if not speakOnDemand or _speechState.speechMode != SpeechMode.onDemand:
		return func(*args, **kwargs)
	_speechState.speechMode = SpeechMode.talk
	ret = func(*args, **kwargs)
	_speechState.speechMode = SpeechMode.onDemand
	return ret


def messageWithSpeakOnDemand(msg):
	executeWithSpeakOnDemand(ui.message, msg)
