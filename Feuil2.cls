VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub mon_code()
'
'' Ici se trouve tout le code principal avec les appels de fonctions et les traitements à faire au préalable
'
' Déclaration des variables.
Dim x0, row1, row2, pas, n_max, n As Variant
Dim name As String
'
' Récupération de la taille de l'échantillon.
n_max = CLng(InputBox("Donnez la taille de l'échantillon. Nous vous conseillons de sélectionner un échantillon de grande taille pour obtenir le biais de sélection le plus petit possible. (de préférence une taille supérieure à 150)"))
Do While n_max <= 76 Or n_max > 9994
'
' Si la taille de l'échantillon est inférieure ou égale à 76 certains échantillons (c'est à dire certains échantillons relevés
' par unités primaires) risquent d'être trop petits (de taille 1 ou 0)
' La formule de la variance corrigée à pour dénominateur n-1 ainsi la valeur de n doit être différente de 1 et de 0.
If n_max <= 25 Then
MsgBox ("Vous devez choisir une taille adéquate pour l'échantillon. Elle ne doit ni être inférieure à 77 ni dépasser la valeur de 9994")
Else
MsgBox ("La taille choisie doit être supérieure ou égale à 77 pour éviter que certaines unités secondaires aient 1 pour taille. Car cela peut constituer un problème pour le calcul de la variance corrigée (n-1 <> 0).")
End If
'
' On redemande à l'utilisateur de donner la taille de l'échantillon tant que la taille n'est pas supérieure à 76.
n_max = CLng(InputBox("Redonnez la taille des échantillons. Nous vous conseillons de sélectionner un échantillon de grande taille pour obtenir le biais de sélection le plus petit possible. (de préférence une taille supérieure à 150)"))
Loop
'
' On attribue à B2 la valeur de n_max (taille de l'échantillon) récupérée.
Range("B2").value = n_max
'
' On initialise les comptages des lignes (donc aux premières lignes) pour les deux feuilles "Données_filtrées" et "Echantillon"
' Pour la feuille "Données_filtrées"
row1 = 1
' Pour la feuille "Echantillon"
row2 = 1
'
' On enlève les styles et les contenus des cellules de la feuille "Echantillon" (initialisation de la feuille)
With Sheets("Echantillon").Range("A1:XFD1048576")
.Style = "Normal"
.ClearContents
End With
'
' On initialise pour chaque unité primaire un aléa entre 1 et le pas calculé pour cette unité
For Each cellule In Range("E5:E19")
Range("G" & (cellule.row)).value = CLng((cellule.value + 1) * Rnd + 1)
Next cellule
'
' On crée une boucle pour afficher les données des unités secondaires et les calculs qui y sont associés dans la feuille "Echantillon"
' La boucle va faire varier le numéro de ligne contenant les données pour chaque unité primaire dans la feuille "SPD-SPI"
For i = 5 To 19
'
' On récupère le pas pour l'unité primaire à la ligne i
pas = Range("E" & (i)).value
'
' Pour le nom de l'Etat
name = Range("A" & (i)).value
'
' Pour le nombre d'unités, concernant un Etat, dans la population (n = Ni)
n = Range("F" & (i)).value
'
' Récupération de l'aléa
x0 = Range("G" & (i)).value
'
' On initialise le début de la plage dans laquelle les données filtrées, concernant un Etat, seront récupérées
' On décale de deux lignes par rapport à la ligne où on se trouve, au niveau de la feuille "Données_filtrées", pour
' pouvoir sauter lignes où sont écrit le nom de l'Etat et les noms des colonnes
begin_plage = row1 + 2
'
' On initialise la fin de la plage dans laquelle les données filtrées de l'Etat seront récupérées
' On décale de n (nombre de données à recueillir) - 1 par rapport au début de la plage
end_plage = begin_plage + n - 1
'
' On appelle la fonction get_samples qui nous permet de récupérer/afficher l'échantillon (ou les échantillons) et de faire
' les traitements appropriés
' On récupère également le numéro de la ligne où on se trouve, au niveau de la feuille "Echantillon", après affichage des données
row2 = get_samples(Sheets("Données_filtrées").Range("A" & (begin_plage) & ":A" & (end_plage)), row1, row2, x0, name, pas, n)
'
' On décale de deux lignes, au niveau des feuilles "Données_filtrées" et "Echantillon"
row1 = row1 + 2
row2 = row2 + 2
Next i
'
' Si le numéro de la ligne où on se trouve, au niveau la feuille "Données_filtrées", correspond au numéro de la ligne de
' fin (là aucune donnée ne reste à être récupérée) alors on affiche que l'échantillonnage s'est bien déroulé.
If row1 = 7942 Then
MsgBox ("L'échantillonnage s'est bien déroulé. Le résultat se trouve dans la feuille 'Echantillon'")
End If
End Sub

Function get_samples(plage As Range, row1 As Variant, row2 As Variant, x0 As Variant, name As Variant, pas As Variant, n As Variant)
'
'' Cette fonction nous permet d'effectuer tous les traitements concernant la récupération des données de l'echantillon,
'' la stylisation des cellules et le calcul des estimateurs.
'' Elle accepte en paramètres la plage contenant les données filtrées, les numéros des lignes actuelles, au niveau des feuilles
'' 'Données_filtrées' et 'Echantillon', l'aléa entre bornes, le nom de l'Etat, le nombre de pas et la taille de la population
'
' On écrit le nom de l'Etat et on attribue à la cellule correspondante le style 'Accent6'
With Sheets("Echantillon").Cells(row2, 1)
.value = name
.Style = "Accent6"
End With
'
' On décale d'une ligne vers le bas, au niveau des deux feuilles
row1 = row1 + 1
row2 = row2 + 1
'
' On attribue le style 'Accent2' pour toutes les cellules qui vont contenir les noms des colonnes
Sheets("Echantillon").Range("A" & (row2) & ":O" & (row2)).Style = "Accent2"
'
' On crée une boucle pour afficher les noms des colonnes sur une ligne
For i = 1 To 15
Sheets("Echantillon").Cells(row2, i).value = Sheets("Données_filtrées").Cells(row1, i).value
Next i
'
' On décale d'une ligne vers le bas, au niveau des deux feuilles
row1 = row1 + 1
row2 = row2 + 1
'
' On récupère les numéros des lignes où on se trouve actuellement, au niveau des deux feuilles, pour les
' traitements ultérieurs à faire
b1 = row1
b2 = row2
'
' On crée une boucle pour récupérer une à une les numéros attribués aux unités de la population
For Each cellule In plage
'
' On vérifie si le numéro de la cellule correspond à la valeur de l'aléa. Si c'est le cas, on change les styles des 15
' cellules (sur la même ligne) qui vont contenir les données propres au numéro de l'unité. Ces données sont receuillies
' au niveau des données filtrées
' puis on récupère et attribue à ces cellules les valeurs des données correspondantes au numéro
If x0 = cellule.value Then
Sheets("Echantillon").Range("A" & (row2)).Style = "60 % - Accent1"
Sheets("Echantillon").Range("B" & (row2) & ":O" & (row2)).Style = "20 % - Accent5"
For i = 1 To 15
Sheets("Echantillon").Cells(row2, i).value = Sheets("Données_filtrées").Cells(row1, i).value
Next i
'
' On décale d'une ligne vers le bas, au niveau de la feuille "Echantillon"
row2 = row2 + 1
'
' L'aléa est augmentée du nombre de pas
x0 = x0 + pas
End If
'
' A chaque itération on décale d'une ligne vers le bas, au niveau de la feuille "Données_filtrées"
row1 = row1 + 1
Next cellule
'
' On réserve 7 lignes dans la feuille "Echantillon", à la sortie de la boucle, pour pouvoir y afficher les calculs à faire
row2 = row2 + 7
'
' Les lignes de codes suivantes concernent les calculs et leurs affichages dans la feuille "Echantillon"
With Sheets("Echantillon")
'
' Attribuons des noms aux cellules qui vont permettre à l'utilisateur d'identifier les calculs
.Range("D" & (row2 - 7)).value = "Estimations ponctuelles"
.Range("E" & (row2 - 7)).value = "Valeurs estimées"
.Range("D" & (row2 - 6)).value = "Estimation_moyenne_ventes"
.Range("D" & (row2 - 5)).value = "Estimation_quantité_totale"
.Range("D" & (row2 - 4)).value = "Estimation_Home_Office_proportion"
.Range("D" & (row2 - 3)).value = "Intervalles de confiances"
.Range("E" & (row2 - 3)).value = "Bornes inférieures"
.Range("F" & (row2 - 3)).value = "Bornes supérieures"
.Range("G" & (row2 - 3)).value = "Vraies valeurs observées"
.Range("H" & (row2 - 3)).value = "Comparaisons"
.Range("D" & (row2 - 2)).value = "Intervalle_de_confiance_moyenne_ventes"
.Range("D" & (row2 - 1)).value = "Intervalle_de_confiance_quantité_totale"
.Range("D" & (row2)).value = "Intervalle_de_confiance_proportion_HO"
'
' Affichage de l'estimation de la moyenne des ventes
.Range("E" & (row2 - 6)).value = get_mean(.Range("M" & (b2) & ":M" & (row2 - 8)))
'
' Affichage de l'estimation de la totale des quantités
.Range("E" & (row2 - 5)).value = get_total(.Range("N" & (b2) & ":N" & (row2 - 8)), n)
'
' Affichage de l'estimation de la proportion de segments ayant pour nom Home Office
.Range("E" & (row2 - 4)).value = get_proportion(.Range("C" & (b2) & ":C" & (row2 - 8)))
'
' On récupère la valeur de l'estimation de la moyenne des ventes
moyenne = .Range("E" & (row2 - 6)).value
'
' On récupère la valeur de l'estimation de la totale des quantités
totale = .Range("E" & (row2 - 5)).value
'
' On récupère la valeur de l'estimation de la proportion de segments de nom Home Office
proportion = .Range("E" & (row2 - 4)).value
'
' Affichage des bornes de l'intervalle de confiance de la moyenne
.Range("E" & (row2 - 2)).value = get_borne_intervale_moyenne(.Range("M" & (b2) & ":M" & (row2 - 8)), n, moyenne, True)
.Range("F" & (row2 - 2)).value = get_borne_intervale_moyenne(.Range("M" & (b2) & ":M" & (row2 - 8)), n, moyenne, False)
'
' Affichage de la valeur de la moyenne des ventes calculée sur la population
.Range("G" & (row2 - 2)).value = get_mean(Sheets("Données_filtrées").Range("M" & (b1) & ":M" & (row1 - 1)))
'
' On vérifie si la valeur réelle de la moyenne des ventes se trouve dans l'intervalle de confiance de la moyenne des ventes
' Si c'est le cas on affiche qu'elle s'y trouve et on attribue à la cellule correspondante le style Satisfaisant
' Sinon on affiche qu'elle ne s'y trouve pas et on attribue à la cellule correspondante le style Insatisfaisant
If .Range("G" & (row2 - 2)).value >= .Range("E" & (row2 - 2)).value And .Range("G" & (row2 - 2)).value <= .Range("F" & (row2 - 2)).value Then
.Range("H" & (row2 - 2)).value = "Appartient à l'intervalle de confiance"
With .Range("H" & (row2 - 2), "J" & (row2 - 2))
.Style = "Satisfaisant"
.Font.Bold = True
End With
Else
.Range("H" & (row2 - 2)).value = "N'appartient pas à l'intervalle de confiance"
With .Range("H" & (row2 - 2), "J" & (row2 - 2))
.Style = "Insatisfaisant"
.Font.Bold = True
End With
End If
'
' Affichage des bornes de l'intervalle de confiance de la totale des quantités
.Range("E" & (row2 - 1)).value = get_borne_intervale_totale(.Range("N" & (b2) & ":N" & (row2 - 8)), n, totale, True)
.Range("F" & (row2 - 1)).value = get_borne_intervale_totale(.Range("N" & (b2) & ":N" & (row2 - 8)), n, totale, False)
.Range("G" & (row2 - 1)).value = get_somme(Sheets("Données_filtrées").Range("N" & (b1) & ":N" & (row1 - 1)))
'
' On vérifie si la valeur réelle de la totale des quantités se trouve dans l'intervalle de confiance de la totale des quantités
' Si c'est le cas on affiche qu'elle s'y trouve et on attribue à la cellule correspondante le style Satisfaisant
' Sinon on affiche qu'elle ne s'y trouve pas et on attribue à la cellule correspondante le style Insatisfaisant
If .Range("G" & (row2 - 1)).value >= .Range("E" & (row2 - 1)).value And .Range("G" & (row2 - 1)).value <= .Range("F" & (row2 - 1)).value Then
.Range("H" & (row2 - 1)).value = "Appartient à l'intervalle de confiance"
With .Range("H" & (row2 - 1), "J" & (row2 - 1))
.Style = "Satisfaisant"
.Font.Bold = True
End With
Else
.Range("H" & (row2 - 1)).value = "N'appartient pas à l'intervalle de confiance"
With .Range("H" & (row2 - 1), "J" & (row2 - 1))
.Style = "Insatisfaisant"
.Font.Bold = True
End With
End If
'
' Affichage des bornes de l'intervalle de confiance de la proportion des segments de nom Home Office
.Range("E" & (row2)).value = get_borne_intervale_proportion(.Range("C" & (b2) & ":C" & (row2 - 8)), proportion, True)
.Range("F" & (row2)).value = get_borne_intervale_proportion(.Range("C" & (b2) & ":C" & (row2 - 8)), proportion, False)
.Range("G" & (row2)).value = get_proportion(Sheets("Données_filtrées").Range("C" & (b1) & ":C" & (row1 - 1)))
'
' On vérifie si la valeur réelle de la proportion des segments de nom Home Office se trouve dans l'intervalle
' de confiance de la proportion des segments de nom Home Office
' Si c'est le cas on affiche qu'elle s'y trouve et on attribue à la cellule correspondante le style Satisfaisant
' Sinon on affiche qu'elle ne s'y trouve pas et on attribue à la cellule correspondante le style Insatisfaisant
If .Range("G" & (row2)).value >= .Range("E" & (row2)).value And .Range("G" & (row2)).value <= .Range("F" & (row2)).value Then
.Range("H" & (row2)).value = "Appartient à l'intervalle de confiance"
With .Range("H" & (row2), "J" & (row2))
.Style = "Satisfaisant"
.Font.Bold = True
End With
Else
.Range("H" & (row2)).value = "N'appartient pas à l'intervalle de confiance"
With .Range("H" & (row2), "J" & (row2))
.Style = "Insatisfaisant"
.Font.Bold = True
End With
End If
'
' Attribution de styles aux cellules réservées aux calculs
With .Range("D" & (row2 - 7) & ":D" & (row2))
.Style = "Accent1"
.Font.Bold = True
End With
With .Range("E" & (row2 - 7) & ":G" & (row2))
.Style = "Sortie"
.Style = "20 % - Accent1"
End With
.Range("D" & (row2 - 7)).Style = "60 % - Accent6"
.Range("D" & (row2 - 3)).Style = "60 % - Accent6"
.Range("E" & (row2 - 3) & ":F" & (row2 - 3)).Style = "40 % - Accent6"
.Range("G" & (row2 - 3), "J" & (row2 - 3)).Style = "20 % - Accent2"
.Range("H" & (row2 - 3)).Font.Bold = True
.Range("E" & (row2 - 7)).Style = "40 % - Accent6"
.Range("F" & (row2 - 7), "G" & (row2 - 4)).Style = "Normal"
End With
'
' On retourne le numéro de la ligne où on se trouvera, au niveau de la feuille "Echantillon", après avoir décalé d'une ligne
' vers le bas
row2 = row2 + 1
get_samples = row2
End Function

Function get_mean(plage As Range)
'
'' Cette fonction permet de calculer une moyenne d'un échantillon quelconque
'' Elle accepte en paramètre la plage comportant les données qui vont être sommer. Cette somme sera divisée par le nombre
'' de données présentes dans la plage puis le résultat de ce calcul sera retourné
Dim somme As Variant
somme = 0
For Each cellule In plage
somme = somme + cellule.value
Next cellule
get_mean = round(somme / plage.Rows.Count, 3)
End Function

Function get_total(plage As Range, n As Variant)
'
'' Cette fonction permet de calculer l'estimation de la totale des quantités
'' Elle accepte en entrées la plage où sont présentes les quantités et la taille de la population (n).
'' On calcule la moyenne à l'aide de la fonction get_mean puis on retourne le résultat de la multiplication de la moyenne par
'' n
moyenne = get_mean(plage)
get_total = round(moyenne * n)
End Function

Function get_proportion(plage As Range) As Variant
'
'' Cette fonction permet de calculer la proportion d'une certaine catégorie dans une plage donnée
'' Elle accepte ainsi en entrée la plage où sont présentes les données
'' La fonction va vérifier combien d'occurences de cette catégorie (Home Office ici) sont présentes
'' dans la plage puis va diviser cette somme par le nombre de lignes contenues dans la plage et retourner le résultat du calcul
Dim somme As Variant
somme = 0
For Each cellule In plage
If cellule.value = "Home Office" Then
somme = somme + 1
End If
Next cellule
get_proportion = round(somme / plage.Rows.Count, 3)
End Function

Function get_var_corrig(plage As Range, moyenne As Variant)
'
'' Cette fonction permet de calculer la variance corrigée d'un échantillon quelconque
'' Elle accepte en entrées la plage contenant les données et la moyenne de l'échantillon
'' La variance corrigée est calculée en sommant les carrés des différences entre les données et la moyenne
'' puis on divise cette somme par la taille de l'échantillon - 1. Ce résultat sera ensuite retourné
Dim somme As Variant
somme = 0
For Each cellule In plage
somme = somme + (cellule.value - moyenne) ^ 2
Next cellule
get_var_corrig = somme / (plage.Rows.Count - 1)
End Function

Function get_borne_intervale_moyenne(plage As Range, n As Variant, moy As Variant, inf As Boolean)
'
'' Cette fonction permet de déterminer la valeur de la borne inférieure ou supérieure d'un intervalle
'' de confiance de la moyenne
'' Elle accepte en paramètres la plage où sont présentes les données, la taille de la population, l'estimation de
'' la moyenne et une valeur booléenne indiquant si on doit déterminer la borne inférieure ou supérieure de l'intervalle de confiance
'' Pour calculer la borne (inférieure ou supérieure) on fait soit la somme, soit la différence (selon la valeur
'' de la variable inf) entre l'estimation de la moyenne et le quantile d'ordre 0.95 de la loi normale multiplié par
'' la racine carrée de l'estimation de la variance de l'estimateur de la moyenne
'' Le résultat du calcul sera ensuite retourné
n_echant = plage.Rows.Count
variance = get_var_corrig(plage, moy)
If inf Then
get_borne_intervale_moyenne = round(moy - 1.96 * Sqr((1 - n_echant / n) * variance / n_echant), 3)
Else
get_borne_intervale_moyenne = round(moy + 1.96 * Sqr((1 - n_echant / n) * variance / n_echant), 3)
End If
End Function

Function get_borne_intervale_totale(plage As Range, n As Variant, totale As Variant, inf As Boolean)
'
'' Cette fonction permet de déterminer la valeur de la borne inférieure ou supérieure d'un intervalle
'' de confiance de la totale
'' Elle accepte en paramètres la plage où sont présentes les données, la taille de la population, l'estimation de
'' la totale et une valeur booléenne indiquant si on doit déterminer la borne inférieure ou supérieure de l'intervalle de confiance
'' Pour calculer la borne (inférieure ou supérieure) on fait soit la somme, soit la différence (selon la valeur
'' de la variable inf) entre l'estimation de la totale et le quantile d'ordre 0.95 de la loi normale multiplié par
'' la racine carrée de l'estimation de la variance de l'estimateur de la totale
'' Le résultat du calcul sera ensuite retourné
moy = get_mean(plage)
variance = get_var_corrig(plage, moy)
n_echant = plage.Rows.Count
If inf Then
get_borne_intervale_totale = round(totale - 1.96 * Sqr((n ^ 2) * (1 - n_echant / n) * (variance ^ 2) / n_echant))
Else
get_borne_intervale_totale = round(totale + 1.96 * Sqr((n ^ 2) * (1 - n_echant / n) * (variance ^ 2) / n_echant))
End If
End Function

Function get_borne_intervale_proportion(plage As Range, proportion As Variant, inf As Boolean)
'
'' Cette fonction permet de déterminer la valeur de la borne inférieure ou supérieure d'un intervalle
'' de confiance de la proportion
'' Elle accepte en paramètres la plage où sont présentes les données, l'estimation de la proportion
'' et une valeur booléenne indiquant si on doit déterminer la borne inférieure ou supérieure de l'intervalle de confiance
'' Pour calculer la borne (inférieure ou supérieure) on fait soit la somme, soit la différence (selon la valeur
'' de la variable inf) entre l'estimation de la proportion et le quantile d'ordre 0.95 de la loi normale multiplié par
'' la racine carrée de l'estimation de la variance de l'estimateur de la proportion
'' Le résultat du calcul sera ensuite retourné
n_echant = plage.Rows.Count
If inf Then
get_borne_intervale_proportion = round(proportion - 1.96 * Sqr(proportion * (1 - proportion) / n_echant), 3)
Else
get_borne_intervale_proportion = round(proportion + 1.96 * Sqr(proportion * (1 - proportion) / n_echant), 3)
End If
End Function

Function get_somme(plage As Range)
'
'' Cette fonction permet de calculer la somme des valeurs contenues dans les cellules d'une plage donnée puis de retourner
'' cette somme
Dim somme As Variant
For Each cellule In plage
somme = somme + cellule.value
Next cellule
get_somme = somme
End Function

