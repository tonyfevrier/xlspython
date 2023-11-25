# Fonctions utiles pour travailler automatiquement sur des tableurs excel

## Classe File qui prend un fichier et qui possède des méthodes.

    1. (- urgent) (A préciser : quel but) Fonction sur le modèle de "cherche" ou plutôt "recherche_chaine_et_retourne_ligne" qui recherche une     donnée dans une colonne donnée et qui renvoie une autre donnée d'une autre colonne ainsi que la ligne.
    2) * Fonction qui parcourt une colonne C et qui crée (ou insère pour éviter l'écrasement de données) une nouvelle colonne à une position donnée, cette nouvelle colonne étant le résultat d'une fonction appliquée à la colonne C et passée en argument.
        FAIT : a Fonction qui parcourt une colonne qui contient plusieurs types de réponses et qui crée une nouvelle colonne à une position donnée qui contient 1 ou 0. Pourrait prendre en argument deux listes de réponses associées par le prog à 0 ou 1. A mon avis vu la fonction de la ligne précédente, il suffit de créer une fonction qui transforme un str en 0 ou 1 et de l'appliquer à la précédente fonction.
        FAIT : b Fonction style xlsparse de dataset2 qui parcourt une colonne qui contient une chaîne à séparer et qui écrit les morceaux séparés en insérant des colonnes (autant que le nb de morceaux de la chaîne) à partir d'une colonne fixée en argument.  
        c Avec la fonction globale, il ne resterait qu'à écrire une fonction spécifique décrivant une action sur la chaîne de chaque cellule de cette colonne (exe : les deux ci-dessous) voir d'autres.
        FAIT : d Fonction qui sous conditions d'une colonne colore une case.
        e Fonction qui si il y a une couleur insère une colonne et y met qqch.
    FAIT : 3) Même chose qu'en ligne 8 mais cette fois en remplaçant la même colonne (juste appeler la fonction ligne 6 et bien choisir la position de la nouvelle colonne = à l'ancienne)
    FAIT : 4) *Fonction qui parcourt une colonne C et qui supprime une ligne si la cellule contient qqch.
    5) (- urgent) *Fonction qui parcourt plusieurs colonnes d'un fichier et qui crée une nouvelle colonne contenant des choses dépendant du contenu des cellules (même style qu'en ligne 6 mais avec plusieurs colonnes au départ) : on aurait aussi une fonction générique en argument.
        (- urgent) a Fonction gén 1 : si on a ça et ça, on met un 1 dans la nouvelle colonne.
        (- urgent) b Fonction gén 2 : on fait la somme, la moyenne de colonnes chiffrées.
    FAIT : 6) *Fonction ajout_colonne_autre_fichier(file1, file2,column): qui parcourt les mails ou un élément caractérisant les participants d'un fichier et ajoute une des caractéristiques dans un second fichier si les mails ou la caractéristique est présent dans ce fichier (les mails sont dans un ordre différent du fichier de départ). Il faut passer en arg les onglets et les colonnes de travail des deux fichiers. Idem peut sûrement se baser sur celle 2 lignes au-dessus
    FAIT : 6bis) *Améliorer la fonction précédente et qui fait copie non pas une mais plusieurs colonnes (l'idée est que si on doit copier plusieurs colonnes, on ne fasse pas plusieurs fois la recherche des mails dans le fichier d'arrivée car c'est coûteux).
    FAIT : 6ter) Même chose mais cette fois en créant une colonne (et pas en copiant). Voir si c'est réellement utile
    7) (-utile) *Fonction ajout ligne_autre_fichier : qui fait comme ajout colonne.  
    8) (-utile) *Fonction qui prend tous les fichiers d'un dossier et qui fait la même action sur chacun de ces fichiers.
    9) (-utile)En combinant les deux précédentes fonctions, on peut créer un fichier de data à partir n fichiers individuels.
    FAIT : 10) Fonction qui regarde si une colonne contient ou non des choses : on pourra s'en servir afin d'éviter d'écraser des données déjà écrite.
    11) (- urgent) Fonction qui trie les lignes suivant un ou plusieurs critères avec des ordres de priorité suivant les critères. Par exe, pour le recrutement de l'institut, on veut trier les femmes en premier critère puis par handicap puis...
    12) (- urgent) Fonction qui filtre les lignes suivant un critère.
    13) (- urgent) Fonction qui copie des lignes dans un autre fichier. On pourra la combiner à 12) pour classer pour le recrutt Charpak.
    FAIT : 14) Fonction qui si qqch est écrit dans une case la colore en une couleur choisie par l'utilisateur. A décliner sur une colonne ou sur l'ensemble d'une feuille.
    FAIT : 14bis) Fonction qui si dans une case d'une ligne, il y a une str particulière (genre un tiret s'il n'y a pas de réponse), colore la ligne entière d'une certaine couleur entrée par l'utilisateur.
    15) Fonction qui prend un fichier avec une colonne de data mais un participant qui est sur plusieurs lignes et qui crée une fonction avec une seule ligne par participant.
    15bis) Fonction qui fait l'inverse : qui prend une ligne et plusieurs colonnes et le coupe en plusieurs lignes pour un même participant.
 
    
## Classe chaine :
    FAIT : Fonction qui prend une str et qui la sépare en plusieurs chaines, la sparation étant donnée par un séparateur.
    (-utile) Fonction qui enlève les guillemets ou un symbole qcq autour d'une chaine si ce symbole est là.

Pour la programmation par classe, la logique voudrait une classe File parent, une classe enfant onglet, puis une classe petit enfant colonne

## Version ++ :
     on fait une interface graphique ou web permettant d'entrer un excel et faire ces opérations sans code.


# Déroulé et prochaines étapes :

    FAIT Imaginer la strcuture par classe 
    FAIT Fabriquer un excel jouet puis un micro test pour la fonction column transform string. 
    FAIT Tester l'ouverture de l'attribut sheet (bonne page) de la classe sheet. 
    FAIT Programmer la fonction et la tester : ajouter l'intervalle des lignes où l'opération a lieu.
    FAIT Factoriser : Nettoyer le fichier des commentaires inutiles
    FAIT : Factoriser: Voir les méthodes qui doivent renvoyer l'objet complet. 
    FAIT : Factoriser si c'est possible : notamment voir la page openclassroom sur la poo : normalement sheet devrait avoir un attribut readbook puisqu'elle hérite.
    FAIT : Factoriser : voir aussi comment utiliser args, kwargs.
    FAIT : Poo : voir s'il ne vaut pas mieux créer une classe Files avec deux noms : celui du fichier à lire et celui à écrire.
    FAIT : Factoriser : Certains arguments des méthode ne seraient-ils pas mieux comme attributs de classe?
    FAIT : Créer un repository git (j'aurais dû le faire bien avant).
    FAIT : Passer à openpyxl : modifier avec les nouvelles commandes.
    FAIT : Faire et retester une fonction sécurité qui empêche d'écrire dans une colonne contenant des choses. Pour cela ajouter dans les fonctions un paramètre security = True qui mis à False permettra d'écrire dans une colonne déjà remplie.
    FAIT : Ajouter dans la classe File une méthode permettant de créer une sauvegarde du fichier de départ 
    FAIT : Ecrire la fonction test_files_identical
    FAIT : Améliorer la fonction copy afin de conserver aussi le format des cellules, les couleurs de fond et de texte.
    FAIT : Voir aussi pour obtenir un nom plus pertinent pour le fichier copié. Mettre test_2023_04_25 pour avoir un historique des copies. Il faudrait alors changer ma fonction del_extension pour supprimer aussi la date si on sauve un fichier déjà daté.
    FAIT : Modifier mes classes de sorte que les modifications se fassent sur le même fichier (en ayant bien vérifié que la sauvegarde fonctionne avant).
    FAIT : Ajouter l'heure au nom du fichier sauvegardé.
    FAIT : Modifier ma fonction 2a avec un paramètre insert = True qui choisit si on insère ou non une colonne à la position column_write. Si on n'insère pas, le paramètre security permet alors d'éviter d'écraser.
    FAIT : Tester ma fonction dans les deux cas : insert = True ou False.    
    FAIT : Débug : comprendre pourquoi dans color_special_cases_in_column il ne rentre jamais dans la condition.
    Débugger le test de color_special_cases_in_sheet : le code affiche FF alors que l'opacité est bien de 0% quand on va dans format cellule.
    FAIT : Fonction 6 : imaginer un test avec un fichier d'arrivée déjà écrit à la main (avec les colonnes séparées).
    FAIT : Programmer le test.
    FAIT : Programmer la fonction.
    FAIT : Modifier la fonction add_col_diff_sorted : pour qu'elle copie aussi les éventuelles couleurs du fichier de départ.
    FAIT : Faire la fonction 14 bis qui colore les lignes.
    FAIT : Fabriquer un test pour la fonction qui doit couper la chaîne en plusieurs
    FAIT : Me relancer dans la fonction 2b : commencer par écrire la fonction qui sépare une chaîne (voir fichier dataset)
    FAIT : Fabriquer un test pour la fonction xlsparse (préparer un fichier.)
    FAIT : Ecrire la fonction équivalente à xlsparse.
    FAIT : Ecrire un test pr la 4)
    FAIT : Programmer la 4) 
    FAIT : Récupérer un vieux test.xlsx dans les commit précédents.
    FAIT : Mettre toutes les micro fonctions utilisées dans Sheet dans un fichier à part qu'on importe pour ne garder que les grosses fonctions de la classe Sheet qui font les gros changements.
    FAIT : Fonction qui prend pleins d'onglets de structure identiques et qui copie une même colonne choisie dans chaque onglet (ou dans un sous ensemble d'onglets) dans une nouvelle feuille en mettant le numéro de l'onglet en haut.
    FAIT : Modifier les fonctions afin d'injecter non pas le numéro de la colonne mais la lettre. Traitement particulier à réserver à addsheetdifferentlysorted, revoir la façon dont je transmets l'autre fichier.
    Fonction qui prend une feuille avec une colonne où des partcpts ont pu répondre plusieurs fois et qui met dans une feuille ceux qui ont répondu plusieurs fois avec des colonnes contenant la valeur d'une cellule donnée lors des différentes réponses (exemple : le temps de réponse lors des différentes réponses).
    Ajouter des exemples avec label=False dans les docstrings.
    Fonction qui prend pleins d'onglets et qui copie les formules d'une colonne (H par exemple) créée dans un onglet (le premier) pour la reproduire sur toutes les colonnes H des autres onglets.
    FAIT:Voir comment générer automatiquement une doc à partir de mes docstrings.
    Vérifier que quand je sauvegarde un fichier avec formules avec dataonly = True, il ne garde que les valeurs. Si oui modifier la doc pour signaler ce point. Voir s'il vaut mieux garder par défaut à True ou le passer à False. Qui est majoritaire?
    Modifier les fonctions pour qu'on puisse aussi entrer les arguments par la lettre de colonne et pas le numéro peu pratique quand il y a beaucoup de colonnes.
    line_end vraiment utile dans les fonctions : max_row ne suffit pas?
    Voir pour regrouper les trois fonctions column_transform_string_in_binary et columnsetansweringroup, columnconvertinminutes la première étant un cas particulier de la seconde. 
    FAIT: Ecrire une documentation pour un utilisateur lambda.
    Prendre un temps pour réfléchir à quelle interface on pourrait utiliser pour qu'un utilisateur n'ait pas à utiliser python. Faire un freeplane.
    
    Créer une classe Path avec attribut nom de dossier
    Voir comment ça peut modifier les classes File et sheet à qui je donnais un argument path.
    Comment dans act on files traduire le fait que fonction peut avoir plusieurs arguments?
    Quel test écrire pour la fonction ci-dessous?
    Programmer une fonction qui fait une même action donnée par une fonction sur un ensemble de fichiers du dossier.
    Modifier les fonctions au cas où il y a des données sur une dernière ligne (dès fois il y a une valeur juste sur une case) : dire que s'il n'y a rien, on s'arrête pour cut in parts.
    Voir si on ne peut pas faire une seule fonction pour 2a et 2b qui utilise en argument les ss fonctions transform_string_in_binary et ...
    Quand on met 0 ou '0' ce n'est pas pareil, modifier les fonctions ou non pour que l'utilisateur ne voit pas la différence? Ou alors mettre un message d'erreur style on veut une chaîne.

# Bug intéressants : 
    -oubli de sauvegarder la feuille en fin de fonction : le prog ne fait alors rien.     
    -certaines str sortant d'excel ont des espaces insécables \xa0 différents des espaces réguliers. Python voit ainsi parfois des str qui semblent identiques différemment.
    -quand on charge un fichier, mettre data_only=True si on veut que lors d'une copie, on ait les valeurs et pas les formules.
