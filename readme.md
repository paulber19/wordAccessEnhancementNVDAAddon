# Microsoft Word text editor: accessibility enhancement #

* Author: PaulBer19
* URL: paulber19@laposte.net
* Download:
	* [stable version][1]
	* [developpement' versions][2]
* Compatibility:
	* Minimum required NVDA version: 2020.4
	* Last NVDA version tested: 2022.2


This add-on adds extra functionality when working with Microsoft Word:

* a script ("windows+alt+f5") to display a dialog box to choose between most of objects 's type to be listed (like comments, revisions, bookmarks, fields, endnotes, footnotes, spelling errors, grammar errors,...),
* a script ("Alt+delete"") to announce line, column and page of cursor position, or start and end of selection, or current table's cell,
* a script (windows+alt+f2) to insert an element like a comment, a footnote, an endnote, without moving the focus,
* a script ("windows+alt+m") to report revision at cursor's position,
* a script ("windows+alt+n") to report endNote or footNote at cursor's position,
* modify the NVDA scripts "control+downArrow" and "Control+Uparrow" (which moves the carret paragraph by paragraph) to skip the empty paragraph (optionnal),
* some scripts to move in table and read table 's elements (row, column, cell),
* adds specific Word browse mode command keys,
* possibility to move sentence by sentence ("alt+ downArrow" and "alt+upArrow"),
* a script to report document 's informations(windows+alt+f1"),
* accessibility enhancement for spelling checker (Word 2013 and 2016):
	* a script (NVDA+shift+f7") to report spelling or grammatical error and suggested correction by the spelling checker,
	* a script (NVDA+control+f7") to report current sentence under focus.
* automatic reading of some elements as comments, footnotes or endnotes.


This module has been tested with Microsoft Word 2019, 2016 and 2013 (perhaps works also with Word 365).


[1]: https://github.com/paulber007/AllMyNVDAAddons/raw/master/wordAccessEnhancement/wordAccessEnhancement-3.3.2.nvda-addon
[2]: https://github.com/paulber007/AllMyNVDAAddons/tree/master/wordAccessEnhancement/dev
