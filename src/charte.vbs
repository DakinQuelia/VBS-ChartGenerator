' ########################################################
' # charte.vbs  v.1.0.0
' # 
' # Auteur: 	Dakin Quelia
' # Courriel:	dakinquelia@gmail.com
' # Site:		http://www.danielchalseche.fr.cr/
' ########################################################
Option Explicit
Dim objFS, strLine, strLine2, objFile
Dim nomSalon, dateCharte, fichier, strNewLine, strNewLine2
Const ForReading = 1
Const ForWriting = 2

' ###############################
' # Création de l'objet
' ###############################
Set objFS = CreateObject("Scripting.FileSystemObject")

' ###############################
' # Récupération des données
' ###############################
nomSalon = inputbox("Nom du salon","Salon")
dateCharte = inputbox("Date de la charte","Date de la charte")
fichier = "charte.txt"

' ###############################
' # On vérifie que les champs 
' # sont remplis
' ###############################
If Len(nomSalon) = 0 Then
	msgbox "Vous devez indiquer le nom du salon!"
Elseif Len(dateCharte) = 0 Then
	msgbox "Vous devez indiquer une date"
Else
	' Lecture du fichier
	Set objFile = objFS.OpenTextFile(fichier, ForReading)
	strLine = objFile.ReadAll
	objFile.Close
	' Ecriture dans le fichier
	Set objFile = objFS.OpenTextFile(fichier, ForWriting)
	strNewLine = Replace(strLine, "<nom du salon>", "#" + nomSalon) 
	strNewLine2 = Replace(strNewLine, "<Date du document>", dateCharte)
	objFile.WriteLine strNewLine2
	objFile.Close
End If