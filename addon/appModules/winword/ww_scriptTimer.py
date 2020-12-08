# appModules\winword\__scriptTimer.py
# A part of wordAccessEnhancement add-on
# Copyright (C) 2020 paulber19
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

# global timer
GB_scriptTimer = None
# maximum delay for waiting new script call
_delay = 250


def stopScriptTimer():
	global GB_scriptTimer
	if GB_scriptTimer is not None:
		GB_scriptTimer.Stop()
		GB_scriptTimer = None
