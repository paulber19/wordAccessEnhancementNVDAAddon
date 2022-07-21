# Editor de texto Microsoft Word: complemento de accesibilidad #

* Autor: PaulBer19
* URL: paulber19@laposte.net
* Descargar:
	* [Versión estable][1]
	* [Versión de desarrollo][2]
* Compatibilidad:
	* Versión mínima de NVDA requerida: 2020.4
	* Última versión de NVDA probada: 2022.2

Este complemento complementa los scripts integrados en la versión básica de NVDA y proporciona:

* el script("windows+alt+F5") para obtener una lista de varios objetos Word (comentarios, cambios, marcadores, , campos, notas finales, notas al pie de página, palabras mal escritas, errores gramaticales,...),
* el script ("Alt+suprimir") para conocer según corresponda: el número de fila, columna y de página de la posición del cursor del sistema,  el principio y el final de la selección, o el número de fila y columna de la celda actual de una tabla (con indicación eventual de la posición con respecto al borde izquierdo y al borde superior),
* el script ("windows+alt+f2") para insertar un comentario en la posición del cursor,
* el script ("windows+alt+m") para leer la modificación del texto en la posición del cursor,
* el script ("windows+alt+n") para leer la nota al pie o al final en la posición del cursor,
* se añade en los scripts básicos de moverse de un párrafo a otro ("Control+Flecha abajo o Flecha arriba"), con la posibilidad de configurarlo para saltar los párrafos vacíos,
* los scripts para moverse  en una tabla o leer sus elementos (fila, columna, celda),
* completa el modo "exploración" de NVDA por  las nuevas órdenes de teclado propios a Microsoft Word,
* para pasar de una frase a otra ("alt+ flecha abajo  o arriba"),
* el script para mostrar información sobre el documento ("windows+alt+f1"),
* accesibilidad mejorada del corrector ortográfico de Word 2013 y 2016:
	* el script ("NVDA+mayúscula+f7") para anunciar el error y la sugerencia,
	* el script (NVDA+control+f7") para anunciar la frase actual bajo el foco.
* lectura automática de algunos elementos como comentarios, notas al pie de página o notas finales.


Este complemento ha sido probado con Microsoft Word 2019, 2016 y 2013 (quizá también funciona con Word 365).



[1]: https://github.com/paulber007/AllMyNVDAAddons/raw/master/wordAccessEnhancement/wordAccessEnhancement-3.3.1.nvda-addon
[2]: https://github.com/paulber007/AllMyNVDAAddons/tree/master/wordAccessEnhancement/dev
