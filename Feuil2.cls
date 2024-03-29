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
'' Ici se trouve tout le code principal avec les appels de fonctions et les traitements � faire au pr�alable
'
' D�claration des variables.
Dim x0, row1, row2, pas, n_max, n As Variant
Dim name As String
'
' R�cup�ration de la taille de l'�chantillon.
n_max = CLng(InputBox("Donnez la taille de l'�chantillon. Nous vous conseillons de s�lectionner un �chantillon de grande taille pour obtenir le biais de s�lection le plus petit possible. (de pr�f�rence une taille sup�rieure � 150)"))
Do While n_max <= 76 Or n_max > 9994
'
' Si la taille de l'�chantillon est inf�rieure ou �gale � 76 certains �chantillons (c'est � dire certains �chantillons relev�s
' par unit�s primaires) risquent d'�tre trop petits (de taille 1 ou 0)
' La formule de la variance corrig�e � pour d�nominateur n-1 ainsi la valeur de n doit �tre diff�rente de 1 et de 0.
If n_max <= 25 Then
MsgBox ("Vous devez choisir une taille ad�quate pour l'�chantillon. Elle ne doit ni �tre inf�rieure � 77 ni d�passer la valeur de 9994")
Else
MsgBox ("La taille choisie doit �tre sup�rieure ou �gale � 77 pour �viter que certaines unit�s secondaires aient 1 pour taille. Car cela peut constituer un probl�me pour le calcul de la variance corrig�e (n-1 <> 0).")
End If
'
' On redemande � l'utilisateur de donner la taille de l'�chantillon tant que la taille n'est pas sup�rieure � 76.
n_max = CLng(InputBox("Redonnez la taille des �chantillons. Nous vous conseillons de s�lectionner un �chantillon de grande taille pour obtenir le biais de s�lection le plus petit possible. (de pr�f�rence une taille sup�rieure � 150)"))
Loop
'
' On attribue � B2 la valeur de n_max (taille de l'�chantillon) r�cup�r�e.
Range("B2").value = n_max
'
' On initialise les comptages des lignes (donc aux premi�res lignes) pour les deux feuilles "Donn�es_filtr�es" et "Echantillon"
' Pour la feuille "Donn�es_filtr�es"
row1 = 1
' Pour la feuille "Echantillon"
row2 = 1
'
' On enl�ve les styles et les contenus des cellules de la feuille "Echantillon" (initialisation de la feuille)
With Sheets("Echantillon").Range("A1:XFD1048576")
.Style = "Normal"
.ClearContents
End With
'
' On initialise pour chaque unit� primaire un al�a entre 1 et le pas calcul� pour cette unit�
For Each cellule In Range("E5:E19")
Range("G" & (cellule.row)).value = CLng((cellule.value + 1) * Rnd + 1)
Next cellule
'
' On cr�e une boucle pour afficher les donn�es des unit�s secondaires et les calculs qui y sont associ�s dans la feuille "Echantillon"
' La boucle va faire varier le num�ro de ligne contenant les donn�es pour chaque unit� primaire dans la feuille "SPD-SPI"
For i = 5 To 19
'
' On r�cup�re le pas pour l'unit� primaire � la ligne i
pas = Range("E" & (i)).value
'
' Pour le nom de l'Etat
name = Range("A" & (i)).value
'
' Pour le nombre d'unit�s, concernant un Etat, dans la population (n = Ni)
n = Range("F" & (i)).value
'
' R�cup�ration de l'al�a
x0 = Range("G" & (i)).value
'
' On initialise le d�but de la plage dans laquelle les donn�es filtr�es, concernant un Etat, seront r�cup�r�es
' On d�cale de deux lignes par rapport � la ligne o� on se trouve, au niveau de la feuille "Donn�es_filtr�es", pour
' pouvoir sauter lignes o� sont �crit le nom de l'Etat et les noms des colonnes
begin_plage = row1 + 2
'
' On initialise la fin de la plage dans laquelle les donn�es filtr�es de l'Etat seront r�cup�r�es
' On d�cale de n (nombre de donn�es � recueillir) - 1 par rapport au d�but de la plage
end_plage = begin_plage + n - 1
'
' On appelle la fonction get_samples qui nous permet de r�cup�rer/afficher l'�chantillon (ou les �chantillons) et de faire
' les traitements appropri�s
' On r�cup�re �galement le num�ro de la ligne o� on se trouve, au niveau de la feuille "Echantillon", apr�s affichage des donn�es
row2 = get_samples(Sheets("Donn�es_filtr�es").Range("A" & (begin_plage) & ":A" & (end_plage)), row1, row2, x0, name, pas, n)
'
' On d�cale de deux lignes, au niveau des feuilles "Donn�es_filtr�es" et "Echantillon"
row1 = row1 + 2
row2 = row2 + 2
Next i
'
' Si le num�ro de la ligne o� on se trouve, au niveau la feuille "Donn�es_filtr�es", correspond au num�ro de la ligne de
' fin (l� aucune donn�e ne reste � �tre r�cup�r�e) alors on affiche que l'�chantillonnage s'est bien d�roul�.
If row1 = 7942 Then
MsgBox ("L'�chantillonnage s'est bien d�roul�. Le r�sultat se trouve dans la feuille 'Echantillon'")
End If
End Sub

Function get_samples(plage As Range, row1 As Variant, row2 As Variant, x0 As Variant, name As Variant, pas As Variant, n As Variant)
'
'' Cette fonction nous permet d'effectuer tous les traitements concernant la r�cup�ration des donn�es de l'echantillon,
'' la stylisation des cellules et le calcul des estimateurs.
'' Elle accepte en param�tres la plage contenant les donn�es filtr�es, les num�ros des lignes actuelles, au niveau des feuilles
'' 'Donn�es_filtr�es' et 'Echantillon', l'al�a entre bornes, le nom de l'Etat, le nombre de pas et la taille de la population
'
' On �crit le nom de l'Etat et on attribue � la cellule correspondante le style 'Accent6'
With Sheets("Echantillon").Cells(row2, 1)
.value = name
.Style = "Accent6"
End With
'
' On d�cale d'une ligne vers le bas, au niveau des deux feuilles
row1 = row1 + 1
row2 = row2 + 1
'
' On attribue le style 'Accent2' pour toutes les cellules qui vont contenir les noms des colonnes
Sheets("Echantillon").Range("A" & (row2) & ":O" & (row2)).Style = "Accent2"
'
' On cr�e une boucle pour afficher les noms des colonnes sur une ligne
For i = 1 To 15
Sheets("Echantillon").Cells(row2, i).value = Sheets("Donn�es_filtr�es").Cells(row1, i).value
Next i
'
' On d�cale d'une ligne vers le bas, au niveau des deux feuilles
row1 = row1 + 1
row2 = row2 + 1
'
' On r�cup�re les num�ros des lignes o� on se trouve actuellement, au niveau des deux feuilles, pour les
' traitements ult�rieurs � faire
b1 = row1
b2 = row2
'
' On cr�e une boucle pour r�cup�rer une � une les num�ros attribu�s aux unit�s de la population
For Each cellule In plage
'
' On v�rifie si le num�ro de la cellule correspond � la valeur de l'al�a. Si c'est le cas, on change les styles des 15
' cellules (sur la m�me ligne) qui vont contenir les donn�es propres au num�ro de l'unit�. Ces donn�es sont receuillies
' au niveau des donn�es filtr�es
' puis on r�cup�re et attribue � ces cellules les valeurs des donn�es correspondantes au num�ro
If x0 = cellule.value Then
Sheets("Echantillon").Range("A" & (row2)).Style = "60�% - Accent1"
Sheets("Echantillon").Range("B" & (row2) & ":O" & (row2)).Style = "20�% - Accent5"
For i = 1 To 15
Sheets("Echantillon").Cells(row2, i).value = Sheets("Donn�es_filtr�es").Cells(row1, i).value
Next i
'
' On d�cale d'une ligne vers le bas, au niveau de la feuille "Echantillon"
row2 = row2 + 1
'
' L'al�a est augment�e du nombre de pas
x0 = x0 + pas
End If
'
' A chaque it�ration on d�cale d'une ligne vers le bas, au niveau de la feuille "Donn�es_filtr�es"
row1 = row1 + 1
Next cellule
'
' On r�serve 7 lignes dans la feuille "Echantillon", � la sortie de la boucle, pour pouvoir y afficher les calculs � faire
row2 = row2 + 7
'
' Les lignes de codes suivantes concernent les calculs et leurs affichages dans la feuille "Echantillon"
With Sheets("Echantillon")
'
' Attribuons des noms aux cellules qui vont permettre � l'utilisateur d'identifier les calculs
.Range("D" & (row2 - 7)).value = "Estimations ponctuelles"
.Range("E" & (row2 - 7)).value = "Valeurs estim�es"
.Range("D" & (row2 - 6)).value = "Estimation_moyenne_ventes"
.Range("D" & (row2 - 5)).value = "Estimation_quantit�_totale"
.Range("D" & (row2 - 4)).value = "Estimation_Home_Office_proportion"
.Range("D" & (row2 - 3)).value = "Intervalles de confiances"
.Range("E" & (row2 - 3)).value = "Bornes inf�rieures"
.Range("F" & (row2 - 3)).value = "Bornes sup�rieures"
.Range("G" & (row2 - 3)).value = "Vraies valeurs observ�es"
.Range("H" & (row2 - 3)).value = "Comparaisons"
.Range("D" & (row2 - 2)).value = "Intervalle_de_confiance_moyenne_ventes"
.Range("D" & (row2 - 1)).value = "Intervalle_de_confiance_quantit�_totale"
.Range("D" & (row2)).value = "Intervalle_de_confiance_proportion_HO"
'
' Affichage de l'estimation de la moyenne des ventes
.Range("E" & (row2 - 6)).value = get_mean(.Range("M" & (b2) & ":M" & (row2 - 8)))
'
' Affichage de l'estimation de la totale des quantit�s
.Range("E" & (row2 - 5)).value = get_total(.Range("N" & (b2) & ":N" & (row2 - 8)), n)
'
' Affichage de l'estimation de la proportion de segments ayant pour nom Home Office
.Range("E" & (row2 - 4)).value = get_proportion(.Range("C" & (b2) & ":C" & (row2 - 8)))
'
' On r�cup�re la valeur de l'estimation de la moyenne des ventes
moyenne = .Range("E" & (row2 - 6)).value
'
' On r�cup�re la valeur de l'estimation de la totale des quantit�s
totale = .Range("E" & (row2 - 5)).value
'
' On r�cup�re la valeur de l'estimation de la proportion de segments de nom Home Office
proportion = .Range("E" & (row2 - 4)).value
'
' Affichage des bornes de l'intervalle de confiance de la moyenne
.Range("E" & (row2 - 2)).value = get_borne_intervale_moyenne(.Range("M" & (b2) & ":M" & (row2 - 8)), n, moyenne, True)
.Range("F" & (row2 - 2)).value = get_borne_intervale_moyenne(.Range("M" & (b2) & ":M" & (row2 - 8)), n, moyenne, False)
'
' Affichage de la valeur de la moyenne des ventes calcul�e sur la population
.Range("G" & (row2 - 2)).value = get_mean(Sheets("Donn�es_filtr�es").Range("M" & (b1) & ":M" & (row1 - 1)))
'
' On v�rifie si la valeur r�elle de la moyenne des ventes se trouve dans l'intervalle de confiance de la moyenne des ventes
' Si c'est le cas on affiche qu'elle s'y trouve et on attribue � la cellule correspondante le style Satisfaisant
' Sinon on affiche qu'elle ne s'y trouve pas et on attribue � la cellule correspondante le style Insatisfaisant
If .Range("G" & (row2 - 2)).value >= .Range("E" & (row2 - 2)).value And .Range("G" & (row2 - 2)).value <= .Range("F" & (row2 - 2)).value Then
.Range("H" & (row2 - 2)).value = "Appartient � l'intervalle de confiance"
With .Range("H" & (row2 - 2), "J" & (row2 - 2))
.Style = "Satisfaisant"
.Font.Bold = True
End With
Else
.Range("H" & (row2 - 2)).value = "N'appartient pas � l'intervalle de confiance"
With .Range("H" & (row2 - 2), "J" & (row2 - 2))
.Style = "Insatisfaisant"
.Font.Bold = True
End With
End If
'
' Affichage des bornes de l'intervalle de confiance de la totale des quantit�s
.Range("E" & (row2 - 1)).value = get_borne_intervale_totale(.Range("N" & (b2) & ":N" & (row2 - 8)), n, totale, True)
.Range("F" & (row2 - 1)).value = get_borne_intervale_totale(.Range("N" & (b2) & ":N" & (row2 - 8)), n, totale, False)
.Range("G" & (row2 - 1)).value = get_somme(Sheets("Donn�es_filtr�es").Range("N" & (b1) & ":N" & (row1 - 1)))
'
' On v�rifie si la valeur r�elle de la totale des quantit�s se trouve dans l'intervalle de confiance de la totale des quantit�s
' Si c'est le cas on affiche qu'elle s'y trouve et on attribue � la cellule correspondante le style Satisfaisant
' Sinon on affiche qu'elle ne s'y trouve pas et on attribue � la cellule correspondante le style Insatisfaisant
If .Range("G" & (row2 - 1)).value >= .Range("E" & (row2 - 1)).value And .Range("G" & (row2 - 1)).value <= .Range("F" & (row2 - 1)).value Then
.Range("H" & (row2 - 1)).value = "Appartient � l'intervalle de confiance"
With .Range("H" & (row2 - 1), "J" & (row2 - 1))
.Style = "Satisfaisant"
.Font.Bold = True
End With
Else
.Range("H" & (row2 - 1)).value = "N'appartient pas � l'intervalle de confiance"
With .Range("H" & (row2 - 1), "J" & (row2 - 1))
.Style = "Insatisfaisant"
.Font.Bold = True
End With
End If
'
' Affichage des bornes de l'intervalle de confiance de la proportion des segments de nom Home Office
.Range("E" & (row2)).value = get_borne_intervale_proportion(.Range("C" & (b2) & ":C" & (row2 - 8)), proportion, True)
.Range("F" & (row2)).value = get_borne_intervale_proportion(.Range("C" & (b2) & ":C" & (row2 - 8)), proportion, False)
.Range("G" & (row2)).value = get_proportion(Sheets("Donn�es_filtr�es").Range("C" & (b1) & ":C" & (row1 - 1)))
'
' On v�rifie si la valeur r�elle de la proportion des segments de nom Home Office se trouve dans l'intervalle
' de confiance de la proportion des segments de nom Home Office
' Si c'est le cas on affiche qu'elle s'y trouve et on attribue � la cellule correspondante le style Satisfaisant
' Sinon on affiche qu'elle ne s'y trouve pas et on attribue � la cellule correspondante le style Insatisfaisant
If .Range("G" & (row2)).value >= .Range("E" & (row2)).value And .Range("G" & (row2)).value <= .Range("F" & (row2)).value Then
.Range("H" & (row2)).value = "Appartient � l'intervalle de confiance"
With .Range("H" & (row2), "J" & (row2))
.Style = "Satisfaisant"
.Font.Bold = True
End With
Else
.Range("H" & (row2)).value = "N'appartient pas � l'intervalle de confiance"
With .Range("H" & (row2), "J" & (row2))
.Style = "Insatisfaisant"
.Font.Bold = True
End With
End If
'
' Attribution de styles aux cellules r�serv�es aux calculs
With .Range("D" & (row2 - 7) & ":D" & (row2))
.Style = "Accent1"
.Font.Bold = True
End With
With .Range("E" & (row2 - 7) & ":G" & (row2))
.Style = "Sortie"
.Style = "20�% - Accent1"
End With
.Range("D" & (row2 - 7)).Style = "60�% - Accent6"
.Range("D" & (row2 - 3)).Style = "60�% - Accent6"
.Range("E" & (row2 - 3) & ":F" & (row2 - 3)).Style = "40�% - Accent6"
.Range("G" & (row2 - 3), "J" & (row2 - 3)).Style = "20�% - Accent2"
.Range("H" & (row2 - 3)).Font.Bold = True
.Range("E" & (row2 - 7)).Style = "40�% - Accent6"
.Range("F" & (row2 - 7), "G" & (row2 - 4)).Style = "Normal"
End With
'
' On retourne le num�ro de la ligne o� on se trouvera, au niveau de la feuille "Echantillon", apr�s avoir d�cal� d'une ligne
' vers le bas
row2 = row2 + 1
get_samples = row2
End Function

Function get_mean(plage As Range)
'
'' Cette fonction permet de calculer une moyenne d'un �chantillon quelconque
'' Elle accepte en param�tre la plage comportant les donn�es qui vont �tre sommer. Cette somme sera divis�e par le nombre
'' de donn�es pr�sentes dans la plage puis le r�sultat de ce calcul sera retourn�
Dim somme As Variant
somme = 0
For Each cellule In plage
somme = somme + cellule.value
Next cellule
get_mean = round(somme / plage.Rows.Count, 3)
End Function

Function get_total(plage As Range, n As Variant)
'
'' Cette fonction permet de calculer l'estimation de la totale des quantit�s
'' Elle accepte en entr�es la plage o� sont pr�sentes les quantit�s et la taille de la population (n).
'' On calcule la moyenne � l'aide de la fonction get_mean puis on retourne le r�sultat de la multiplication de la moyenne par
'' n
moyenne = get_mean(plage)
get_total = round(moyenne * n)
End Function

Function get_proportion(plage As Range) As Variant
'
'' Cette fonction permet de calculer la proportion d'une certaine cat�gorie dans une plage donn�e
'' Elle accepte ainsi en entr�e la plage o� sont pr�sentes les donn�es
'' La fonction va v�rifier combien d'occurences de cette cat�gorie (Home Office ici) sont pr�sentes
'' dans la plage puis va diviser cette somme par le nombre de lignes contenues dans la plage et retourner le r�sultat du calcul
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
'' Cette fonction permet de calculer la variance corrig�e d'un �chantillon quelconque
'' Elle accepte en entr�es la plage contenant les donn�es et la moyenne de l'�chantillon
'' La variance corrig�e est calcul�e en sommant les carr�s des diff�rences entre les donn�es et la moyenne
'' puis on divise cette somme par la taille de l'�chantillon - 1. Ce r�sultat sera ensuite retourn�
Dim somme As Variant
somme = 0
For Each cellule In plage
somme = somme + (cellule.value - moyenne) ^ 2
Next cellule
get_var_corrig = somme / (plage.Rows.Count - 1)
End Function

Function get_borne_intervale_moyenne(plage As Range, n As Variant, moy As Variant, inf As Boolean)
'
'' Cette fonction permet de d�terminer la valeur de la borne inf�rieure ou sup�rieure d'un intervalle
'' de confiance de la moyenne
'' Elle accepte en param�tres la plage o� sont pr�sentes les donn�es, la taille de la population, l'estimation de
'' la moyenne et une valeur bool�enne indiquant si on doit d�terminer la borne inf�rieure ou sup�rieure de l'intervalle de confiance
'' Pour calculer la borne (inf�rieure ou sup�rieure) on fait soit la somme, soit la diff�rence (selon la valeur
'' de la variable inf) entre l'estimation de la moyenne et le quantile d'ordre 0.95 de la loi normale multipli� par
'' la racine carr�e de l'estimation de la variance de l'estimateur de la moyenne
'' Le r�sultat du calcul sera ensuite retourn�
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
'' Cette fonction permet de d�terminer la valeur de la borne inf�rieure ou sup�rieure d'un intervalle
'' de confiance de la totale
'' Elle accepte en param�tres la plage o� sont pr�sentes les donn�es, la taille de la population, l'estimation de
'' la totale et une valeur bool�enne indiquant si on doit d�terminer la borne inf�rieure ou sup�rieure de l'intervalle de confiance
'' Pour calculer la borne (inf�rieure ou sup�rieure) on fait soit la somme, soit la diff�rence (selon la valeur
'' de la variable inf) entre l'estimation de la totale et le quantile d'ordre 0.95 de la loi normale multipli� par
'' la racine carr�e de l'estimation de la variance de l'estimateur de la totale
'' Le r�sultat du calcul sera ensuite retourn�
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
'' Cette fonction permet de d�terminer la valeur de la borne inf�rieure ou sup�rieure d'un intervalle
'' de confiance de la proportion
'' Elle accepte en param�tres la plage o� sont pr�sentes les donn�es, l'estimation de la proportion
'' et une valeur bool�enne indiquant si on doit d�terminer la borne inf�rieure ou sup�rieure de l'intervalle de confiance
'' Pour calculer la borne (inf�rieure ou sup�rieure) on fait soit la somme, soit la diff�rence (selon la valeur
'' de la variable inf) entre l'estimation de la proportion et le quantile d'ordre 0.95 de la loi normale multipli� par
'' la racine carr�e de l'estimation de la variance de l'estimateur de la proportion
'' Le r�sultat du calcul sera ensuite retourn�
n_echant = plage.Rows.Count
If inf Then
get_borne_intervale_proportion = round(proportion - 1.96 * Sqr(proportion * (1 - proportion) / n_echant), 3)
Else
get_borne_intervale_proportion = round(proportion + 1.96 * Sqr(proportion * (1 - proportion) / n_echant), 3)
End If
End Function

Function get_somme(plage As Range)
'
'' Cette fonction permet de calculer la somme des valeurs contenues dans les cellules d'une plage donn�e puis de retourner
'' cette somme
Dim somme As Variant
For Each cellule In plage
somme = somme + cellule.value
Next cellule
get_somme = somme
End Function

