Attribute VB_Name = "mod ADO Recordsets"
Option Compare Database
Option Explicit

' ---
' LISTE DES PROPRIETAIRES PAR ADO
' ---
Sub ADO_ListeProprietaires()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset

' Ouvrir la table Propriétaires
Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset
rst.Open "[tbl Propriétaires]", cnn, adOpenDynamic, adLockOptimistic, adCmdTable

' Afficher quelques informations
While Not rst.EOF
    Debug.Print rst("Nom Propriétaire"), rst("Prénom Propriétaire")

    rst.MoveNext
Wend

' Libérer les objets
rst.Close
Set rst = Nothing
cnn.Close
Set cnn = Nothing
End Sub


' ---
' LISTE DES PROPRIETAIRES PAR ADO
' ---
Function ADO_ListeProprietaires2()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim strSQL As String

' Ouvrir la table Propriétaires
Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset
strSQL = "SELECT * FROM [tbl Propriétaires];"
rst.Open strSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText

' Afficher quelques informations
While Not rst.EOF
    Debug.Print rst("Nom Propriétaire"), rst("Prénom Propriétaire")

    rst.MoveNext
Wend

' Libérer les objets
rst.Close
Set rst = Nothing
cnn.Close
Set cnn = Nothing
End Function

' ---
' AJOUT D'UN BIEN PAR ADO
' ---
'
Function ADO_AjoutBien()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset

' Ouvrir la table Biens
Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset
rst.Open "[tbl Biens]", cnn, adOpenDynamic, adLockOptimistic, adCmdTable

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
cnn.Close
Set cnn = Nothing
End Function

' ---
' MODIFICATION D'UN BIEN PAR ADO
' ---
'
Function ADO_ModifBien()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset

' Ouvrir la table Biens
Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset
rst.Open "SELECT * FROM [tbl Biens] WHERE [Numéro Bien]='ABC-099'", _
    cnn, adOpenDynamic, adLockOptimistic, adCmdText

If rst.EOF Then
    MsgBox "Aucun bien ne répond aux critères !", vbExclamation
Else
    ' Modifier l'enregistrement
    rst("Code Transaction") = "VEN"
    rst("Statut") = "A vendre"
    rst("Date de modification") = Now
    rst.Update
End If

' Libérer les objets
rst.Close
Set rst = Nothing
cnn.Close
Set cnn = Nothing
End Function
' ---
' SUPPRESSION D'UN BIEN PAR ADO
' ---
'
Function ADO_SuppressionBien()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset

' Ouvrir la table Biens
Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset
rst.Open "SELECT * FROM [tbl Biens Copie] WHERE [Statut]='Archivé'", _
    cnn, adOpenDynamic, adLockOptimistic

While Not rst.EOF
    rst.Delete
    rst.MoveNext
Wend

' Libérer les objets
rst.Close
Set rst = Nothing
cnn.Close
Set cnn = Nothing
End Function

' ---
' RECHERCHE D'INFORMATIONS PAR RECORDSETS ADO
' ---
'
Sub ADO_RechercheBien1()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim strCritere As String

' Ouvrir la table Biens
Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset
rst.Open "[tbl Biens]", cnn, adOpenDynamic, adLockOptimistic, adCmdTable

' Rechercher un bien par critères
strCritere = "[Descriptif] LIKE '*mer*'"
rst.Find strCritere
While Not rst.EOF
    MsgBox rst("Numéro Bien") & vbCrLf & rst("Descriptif") & _
        vbCrLf & rst("Montant")

    rst.Find strCritere, 1, adSearchForward
Wend

' Libérer les objets
rst.Close
Set rst = Nothing
cnn.Close
Set cnn = Nothing
End Sub

' ---
' RECHERCHE D'INFORMATIONS PAR RECORDSETS ADO
' ---
'
Sub ADO_RechercheBien2()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim strSQL As String

' Ouvrir la table Biens
Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset
strSQL = "SELECT * FROM [tbl Biens]" & _
    " WHERE ([Descriptif] LIKE '%Métro%')" & _
    " AND ([Montant] > 150000)"
    ' Le critère sur les dates suit la syntaxe Jet :
    '" AND ([Date de création] > #14/05/2002 15:00#)"
rst.Open strSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText

' Afficher les biens répondant aux critères
While Not rst.EOF
    MsgBox rst("Numéro Bien") & _
        vbCrLf & rst("Descriptif") & _
        vbCrLf & rst("Montant") & _
        vbCrLf & rst("Date de création")

    rst.MoveNext
Wend

' Libérer les objets
rst.Close
Set rst = Nothing
cnn.Close
Set cnn = Nothing
End Sub
