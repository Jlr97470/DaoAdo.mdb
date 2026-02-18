Attribute VB_Name = "mod DAO Recordsets"
Option Compare Database

' OUVERTURE/FERMETURE D'UN RECORDSET
Sub DAO_TestRecordset()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouverture du Recordset
Set db = CurrentDb
Set rst = db.OpenRecordset("tbl Propriétaires")

' Affichage
Debug.Print rst.Fields(2).Value
Debug.Print rst(2).Value
Debug.Print rst("Nom Propriétaire")
Debug.Print rst![Nom Propriétaire]

' Fermeture et libération des objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub

' ---
' PARCOURS DES PROPRIETAIRES DE LA BASE
' ---
'
Function DAO_ListerProprietaires()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouverture de la table tbl Propriétaires
Set db = Application.CurrentDb
Set rst = db.OpenRecordset("tbl Propriétaires", dbOpenDynaset)

' Parcours de la table
While rst.EOF = False
    ' Affichage de quelques informations
    Debug.Print rst("Nom Propriétaire"), rst("Prénom Propriétaire"),
    Debug.Print rst.AbsolutePosition

    ' Passage au propriétaire suivant
    rst.MoveNext
Wend

' Fermeture et libération des objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Function

' ---
' AFFICHAGE DU NOMBRE D'ENREGISTREMENTS
Sub DAO_DecompteEnregistrements()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouverture de la table Propriétaires
Set db = CurrentDb
Set rst = db.OpenRecordset("tbl Propriétaires", dbOpenDynaset)

' Affichage des résultats
MsgBox "Nombre d'enregistrements : " & rst.RecordCount
If Not rst.EOF Then rst.MoveLast
MsgBox "Nombre d'enregistrements : " & rst.RecordCount

' Repositionnement en début pour un traitement éventuel
If Not rst.BOF Then rst.MoveFirst
' ... suite du programme ici

' Fermeture et libération des objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub
' ---
' SUPPRESSION DES BIENS DE LA TABLE [TBL BIENS COPIE]
' ---
'
Function DAO_SupprimerBiens()
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim lngBiens As Long

' Décompte du nombre de biens avant suppression
lngBiens = DCount("*", "tbl Biens Copie")
MsgBox "Actuellement " & lngBiens & " biens dans la table.", vbInformation

' Ouverture de la table tbl Biens
Set db = Application.CurrentDb
Set rst = db.OpenRecordset("tbl Biens Copie")

' Parcours de la table
While rst.EOF = False
    rst.Delete

    ' Passage au bien suivant
    rst.MoveNext
Wend

' Fermeture et libération des objets
rst.Close
Set rst = Nothing
Set db = Nothing

' Décompte du nombre de biens restants
lngBiens = DCount("*", "tbl Biens Copie")
MsgBox "Après suppression, il reste " & lngBiens & " bien dans la table.", vbInformation

' Ouverture pour vérification
DoCmd.OpenTable "tbl Biens Copie"
End Function

' ---
' MODIFICATION DES BIENS DE LA BASE
' ---
'
Function DAO_ModifierBiens1()
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim intI As Integer

' Ouverture de la table tbl Biens Copie
Set db = Application.CurrentDb
Set rst = db.OpenRecordset("tbl Biens Copie")

' Parcours de la table
intI = 1
While rst.EOF = False
    ' Modification du champ Nom Bien
    rst.Edit
    rst("Nom Bien") = "Test " & intI
    rst.Update

    ' Passage au bien suivant
    intI = intI + 1
    rst.MoveNext
Wend

' Fermeture et libération des objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Function

' ---
' MODIFICATION DES BIENS DE LA BASE (VARIANTE)
' ---
'
Function DAO_ModifierBiens2()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouverture de la table tbl Biens Copie
Set db = Application.CurrentDb
Set rst = db.OpenRecordset("tbl Biens Copie")

' Parcours de la table
While rst.EOF = False
    ' Seul les biens du propriétaire n°2 sont concernés...
    If rst("Numéro Propriétaire") = 2 Then
        ' Modification du champ Statut
        rst.Edit
        rst("Statut") = "Archivé"
        rst.Update
    End If

    ' Passage au bien suivant
    rst.MoveNext
Wend

' Fermeture et libération des objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Function
' ---
' MODIFICATION DES BIENS DE LA BASE (VARIANTE SQL)
' ---
'
Function SQL_ModifierBiens()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouverture de la table tbl Biens Copie
Set db = Application.CurrentDb
Set rst = db.OpenRecordset("SELECT * FROM [tbl Biens Copie] WHERE [Numéro Propriétaire]=2;")

' Parcours de la table
While rst.EOF = False
    ' Modification du champ Nom Bien
    rst.Edit
    rst("Statut") = "Archivé"
    rst.Update

    ' Passage au bien suivant
    rst.MoveNext
Wend

' Fermeture et libération des objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Function


' ---
' AJOUT D'UN BIEN
' ---
'
Sub DAO_AjoutBien()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouvrir la table des biens
Set db = CurrentDb()
Set rst = db.OpenRecordset("tbl Biens")

' Ajouter un enregistrement
rst.AddNew
    rst("Numéro Bien") = "ABC-099"
    rst("Numéro Propriétaire") = 2
    rst("Nom Bien") = "Kerzan"
    rst("Adresse 1") = "Rue de la Rose des Vents"
    rst("Code Ville") = 25
    rst("Code Transaction") = "LOC"
    rst("Code Bien") = "APP"
    rst("Code Catégorie") = "NA"
rst.Update

' Libérer les objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub

' ---
' MODIFICATION D'UN BIEN
' ---
'
Sub DAO_ModifBien()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouvrir la table des biens
Set db = CurrentDb()
Set rst = db.OpenRecordset("SELECT * FROM [tbl Biens] WHERE [Numéro Bien]='ABC-099'")

' Y a-t-il au moins un bien ?
If rst.EOF Then
    MsgBox "Aucun bien ne répond au critère !", vbExclamation
Else
    ' Modifier l'enregistrement
    rst.Edit
        rst("Code Transaction") = "VEN"
        rst("Statut") = "A vendre"
        rst("Date de modification") = Now
    rst.Update
End If

' Libérer les objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub

' ---
' MODIFICATION D'UN BIEN PAR SQL
' ---
'
Sub SQL_ModifBien()
Dim strSQL As String

strSQL = "UPDATE [tbl Biens] SET [Code Transaction]='VEN', [Statut]='A vendre', [Date de modification]=Now " & _
    " WHERE [Numéro Bien]='ABC-099';"
CurrentDb.Execute strSQL
End Sub

' ---
' SUPPRESSION D'UN BIEN PAR RECORDSET DAO
' ---
'
Sub DAO_SuppressionBien()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouvrir une requête des biens archivés
Set db = CurrentDb()
Set rst = db.OpenRecordset("SELECT * FROM [tbl Biens Copie] WHERE [Statut]='Archivé'")

' Supprimer tous les enregistrements trouvés
While Not rst.EOF
    rst.Delete
    rst.MoveNext
Wend

' Libérer les objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub


' ---
' SUPPRESSION D'UN BIEN VIA SQL
' ---
'
Sub SQL_SuppressionBien()
Dim strSQL As String

strSQL = "DELETE * FROM [tbl Biens] WHERE [Statut]='Archivé';"
CurrentDb.Execute strSQL
End Sub

' ---
' RECHERCHE D'UN BIEN VIA DAO
' ---
Sub DAO_RechercheBien1()
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim strCritere As String

' Ouvrir la table Biens
Set db = CurrentDb
Set rst = db.OpenRecordset("tbl Biens", dbOpenDynaset)

' Lire certains enregistrements selon critères
strCritere = "([Descriptif] LIKE '*métro*') AND ([Montant]>150000)"
rst.FindFirst strCritere
While Not rst.NoMatch
    MsgBox rst("Numéro Bien") & _
        vbCrLf & rst("Descriptif") & _
        vbCrLf & rst("Montant")

    rst.FindNext strCritere
Wend

' Libérer les objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub
' ---
' RECHERCHE D'UN BIEN VIA DAO/SQL
' ---
Sub DAO_RechercheBien2()
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim strSQL As String

' Ouvrir la table Biens
Set db = CurrentDb
strSQL = "SELECT * FROM [tbl Biens] WHERE ([Descriptif] LIKE '*métro*') AND ([Montant]>150000)"
Set rst = db.OpenRecordset(strSQL, dbOpenDynaset)

' Afficher les biens répondant aux critèrs
While Not rst.EOF
    MsgBox rst("Numéro Bien") & _
        vbCrLf & rst("Descriptif") & _
        vbCrLf & rst("Montant")

    rst.MoveNext
Wend

' Libérer les objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub

' ---
' RECHERCHE D'UN BIEN VIA DAO/SEEK
' ---
Sub DAO_RechercheBien3()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouvrir la table Biens
Set db = CurrentDb
Set rst = db.OpenRecordset("tbl Biens", dbOpenTable)

' Choix de l'index de recherche
' (ne fonctione que sur un Recordset de type Table (dbOpenTable)
rst.Index = "PrimaryKey"
rst.Seek "=", "ABC-040"

' Afficher le bien répondant aux critères
MsgBox rst("Numéro Bien") & _
    vbCrLf & rst("Descriptif") & _
    vbCrLf & rst("Montant") & _
    vbCrLf & rst("Code Transaction")

' Libérer les objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub

' ---
' UTILISATION DES SIGNETS (BOOKMARKS) DAO
' ---
'
Sub DAO_TestBookmarks()
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim strSignet As String

' Ouvrir la table Propriétaires
Set db = CurrentDb
Set rst = db.OpenRecordset("tbl Propriétaires", dbOpenDynaset)
rst.FindFirst "[Nom Propriétaire]='INISAN'"
MsgBox "Propriétaire : " & rst("Nom Propriétaire")

' Mémoriser l'emplacement actuel
strSignet = rst.Bookmark

' Se déplacer sur un autre enregistrement
rst.MoveLast
MsgBox "Propriétaire : " & rst("Nom Propriétaire")

' Se replacer au niveau du signet mémorisé
rst.Bookmark = strSignet
MsgBox "Propriétaire : " & rst("Nom Propriétaire")

' Fermer et libérer les objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub
