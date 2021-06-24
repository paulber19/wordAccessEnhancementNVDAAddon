# appModules\winword\__scriptTimer.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2020 paulber19
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
import wx

# global timer
_scriptTimer = None
# maximum delay for waiting new script call
_delay = 250


def delayScriptTask(func, *args, **kwargs):
	global _scriptTimer
	_scriptTimer = wx.CallLater(_delay, func, *args, **kwargs)


def stopScriptTimer():
	global _scriptTimer
	if _scriptTimer is None:
		return
	_scriptTimer.Stop()
	_scriptTimer = None


def clearScriptTimer():
	global _scriptTimer
	_scriptTimer = None
