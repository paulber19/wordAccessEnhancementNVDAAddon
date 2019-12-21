# appModules\winword\ww_tones.py
# A part of WordAccessEnhancement  add-on
#Copyright (C) 2019 paulber19
#This file is covered by the GNU General Public License.


import addonHandler
addonHandler.initTranslation()
from logHandler import log
import tones
import threading

class RepeatTask(threading.Thread): 
	_delay = None
	
	def __init__(self): 
		if self._delay is None:
			log.error("Cannot create repeatTask thread because delay is none")
			return
		self.noNextTask = True
		super(RepeatTask, self).__init__()
		self._stopevent = threading.Event( ) 
	
	def task (self):
		return
	
	def run(self): 
		if self._delay is None:
			log.error("Cannot start repeatTask thread because not delay")
			return
		while not self._stopevent.isSet() :
			self._stopevent.wait(self._delay) 
			if self.noNextTask:
				self.noNextTask = False
			else:
				self.task()
	
	def stop(self): 
		self.noNextTask = True
		self._stopevent.set() 

class RepeatBeep(RepeatTask): 
	def __init__(self, delay = 2.0, beep = (200,200) ):
		self._delay = delay
		self.beep = beep
		super(RepeatBeep,self).__init__()
	
	def task (self):
		(frequence, length) = self.beep
		tones.beep(frequence,length)
