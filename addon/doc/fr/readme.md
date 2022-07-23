# Éditeur de texte Microsoft Word: complément d'accessiblité #

* Auteur : PaulBer19
* URL : paulber19@laposte.net
* Téléchargement :
	* [version stable][1]
	* [version de développement][2]
* Compatibilité :
	* Version minimum de NVDA requise : 2020.4
	* Dernière version de NVDA testée : 2022.2


Cette extension complète les scripts intégrés dans la version de base de NVDA et apporte :

* le script("windows+alt+F5") pour obtenir la liste de divers objets Word (commentaires, modifications, signets, champs, note de fin, note de bas de page, mots mal orthographiés, erreurs de grammaire...),
* le script ("Alt+éfacement") pour connaitre suivant le cas : le numéro de ligne, colonne et de page de la position du curseur système, le début et la fin de la sélection, ou le numéro de ligne et colonne de la cellule courante d'un tableau (avec indication éventuelle de la position par rapport au bord gauche et supérieur),
* le script ("windows+alt+f2") pour insérer un un élément comme un commentaire, une note de bas de page, une note de fin, à la position du curseur,
* le script ("windows+alt+m") pour lire la modification de texte à la position du curseur,
* le script ("windows+alt+n") pour lire la note de bas de page ou de fin à la position du curseur,
* ajout dans les scripts de base de déplacement de paragraphe en paragraphe ("Control+Flèche bas ou haut"), la possiblité configurable de sauter les paragraphes vides,
* les scripts pour se déplacer dans un tableau ou lire ses éléments (ligne, colonne, cellule),
* complète le mode "navigation" de NVDA par de nouvelles commandes propre à Microsoft Word,
* le déplacement de phrase en phrase ("alt+ flèche bas ou haut"),
* le script pour afficher quelques informations sur le document ("windows+alt+f1"),
* amélioration de l'accessiblité du correcteur orthographique de Word 2013 et 2016 :
	* le script ("NVDA+majuscule+f7") pour annoncer l'erreur et la suggestion,
	* le script (NVDA+control+f7") pour annoncer la phrase courante sous le focus.
* lecture automatique de certains éléments comme les commentaires, notes de bas de page ou notes de fin.


Cette extension a été testée avec MS Word 2019, 2016 et 2013 (marche peut-être aussi avec Word 365).



[1]: https://github.com/paulber007/AllMyNVDAAddons/raw/master/wordAccessEnhancement/wordAccessEnhancement-3.3.2.nvda-addon
[2]: https://github.com/paulber007/AllMyNVDAAddons/tree/master/wordAccessEnhancement/dev
