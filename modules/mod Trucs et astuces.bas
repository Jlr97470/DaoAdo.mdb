Attribute VB_Name = "mod Trucs et astuces"
Option Compare Database
Option Explicit

' ---
' LECTURE D'UN REPERTOIRE ET ECRITURE DES FICHIERS
' DANS UNE TABLE ACCESS (DAO)
' ---
Function DAO_LireRepertoire()
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim strRep As String
Dim strFichier As String

' Répertoire à analyser
strRep = "C:\Access Référence\"

' Vider la table tbl Fichiers
CurrentDb.Execute "DELETE * FROM [tbl Fichiers];"

' Ouvrir la table
Set db = CurrentDb
Set rst = db.OpenRecordset("tbl Fichiers", dbOpenDynaset)

' Ecrire les fichiers dans la table
strFichier = Dir(strRep & "*.*")
While strFichier <> ""
    rst.AddNew
        rst("Fichier") = strFichier
    rst.Update

    ' Lire le fichier suivant
    strFichier = Dir
Wend

' Fermer et libérer les objets
rst.Close
Set rst = Nothing
Set db = Nothing
End Function

' ---
' CREATION D'UNE REQUETE DANS LA BASE DE DONNEES (DAO)
' ---
' Entrée : strNomRequete <- Nom de la requête à créer
'          strSQL        <- Code SQL de la requête
'
Function DAO_CreerRequete(ByVal strNomRequete As String, _
    ByVal strSQL As String)
Dim db As DAO.Database
Dim qdf As DAO.QueryDef

' Vérifier si la requête existe en forçant
' sa lecture
On Error Resume Next
Set db = CurrentDb
Set qdf = db.QueryDefs(strNomRequete)
If Err.Number = 0 Then
    MsgBox "La requête existe déjà.", vbInformation
Else
    Set qdf = db.CreateQueryDef(strNomRequete, strSQL)
End If
Set qdf = Nothing

' Mettre à jour la fenêtre Base de données
Application.RefreshDatabaseWindow
End Function

' ---
' MODIFICATION DU CODE SQL D'UNE REQUETE EXISTANTE
' ---
Function DAO_ModifierRequete(ByVal strNomRequete As String, _
    ByVal strSQL As String)
Dim db As DAO.Database
Dim qdf As DAO.QueryDef

' Vérifier si la requête existe en forçant
' sa lecture
On Error Resume Next
Set db = CurrentDb
Set qdf = db.QueryDefs(strNomRequete)
If Err.Number <> 0 Then
    MsgBox "La requête n'existe pas.", vbInformation
Else
    Set qdf = db.QueryDefs(strNomRequete)
    MsgBox "Ancien code SQL de la requête : " & _
        vbCrLf & qdf.SQL
    qdf.SQL = strSQL
    MsgBox "Nouveau code SQL de la requête :" & _
        vbCrLf & qdf.SQL
End If
Set qdf = Nothing
End Function

' ---
' CREATION D'UNE REQUETE (VUE) ADO
' ---
'
Function ADO_CreerRequete(ByVal strNomRequete As String, _
    ByVal strSQL As String)
Dim cnn As ADODB.Connection
Dim cmd As ADODB.Command
Dim cat As ADOX.Catalog

' Initialisation du catalogue
Set cnn = CurrentProject.Connection
Set cat = New ADOX.Catalog
Set cat.ActiveConnection = CurrentProject.Connection

' Créer une vue
Set cmd = New ADODB.Command
cmd.CommandText = strSQL
cat.Views.Append strNomRequete, cmd
Set cat = Nothing
Set cnn = Nothing

' Mettre à jour la fenêtre Base de données
Application.RefreshDatabaseWindow
End Function

' ---
' OUVERTURE D'UNE REQUETE PARAMETREE (DAO)
' ---
Function DAO_RequeteParametree()
Dim db As DAO.Database
Dim qdf As DAO.QueryDef
Dim rst As DAO.Recordset

Set db = CurrentDb
Set qdf = db.QueryDefs("rqt Clients à recontacter")
qdf.Parameters("Date de début") = #3/1/2003#
qdf.Parameters("Date de fin") = #3/31/2003#
Set rst = qdf.OpenRecordset
While Not rst.EOF
    Debug.Print rst("Numéro Client"), rst("Prochain contact")
    rst.MoveNext
Wend

rst.Close
Set rst = Nothing
Set qdf = Nothing
Set db = Nothing
End Function

' ---
' GETROWS : EXEMPLE DAO
' ---
' La méthode GetRows transfère le contenu d'un Recordset
' dans un tableau de variants.
'
Sub DAO_GetRows()
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim v As Variant
Dim lngLigne As Long, lngColonne As Long

' Ouverture de la table
Set db = CurrentDb
Set rst = db.OpenRecordset("tbl Villes")

' Transfert de 5 lignes dans le tableau v
v = rst.GetRows(5)

' Affichage des lignes/colonnes
For lngLigne = 0 To UBound(v, 2)
    For lngColonne = 0 To UBound(v, 1)
        Debug.Print v(lngColonne, lngLigne),
    Next
    Debug.Print
Next

' On ferme !
rst.Close
Set rst = Nothing
Set db = Nothing
End Sub

' ---
' GETROWS : EXEMPLE ADO
' ---
' La méthode GetRows transfère le contenu d'un Recordset
' dans un tableau de variants.Sub ADOGetRows()
'
Sub ADO_GetRows()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim v As Variant
Dim lngLigne As Long, lngColonne As Long

' Ouverture de la table
Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset
rst.Open "[tbl Villes]", cnn, adOpenDynamic, _
    adLockOptimistic, adCmdTable

' Transfert de 5 lignes dans le tableau v
v = rst.GetRows(5)

' Affichage des lignes/colonnes
For lngLigne = 0 To UBound(v, 2)
    For lngColonne = 0 To UBound(v, 1)
        Debug.Print v(lngColonne, lngLigne),
    Next
    Debug.Print
Next

rst.Close
Set rst = Nothing
cnn.Close
Set cnn = Nothing
End Sub
' ---
' OUVERTURE D'UNE AUTRE BASE DE DONNEES (DAO)
' ---
'
Function DAO_ConnexionExterne()
Dim db As DAO.Database
Dim rst As DAO.Recordset

Set db = DBEngine.OpenDatabase("C:\Access Référence\Archive.mdb")
Set rst = db.OpenRecordset("tbl Villes")
While Not rst.EOF
    Debug.Print rst("Nom Ville"), rst("Code postal")
    rst.MoveNext
Wend
rst.Close
db.Close
Set rst = Nothing
Set db = Nothing
End Function
' ---
' OUVERTURE D'UNE AUTRE BASE DE DONNEES (DAO)
' ---
'
Function ADO_ConnexionExterne()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset

' Ouverture de la connexion
Set cnn = New ADODB.Connection
cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Access Référence\Archive.mdb"

' Ouverture du Recordset
Set rst = New ADODB.Recordset
rst.Open "[tbl Villes]", cnn, _
    adOpenDynamic, adLockOptimistic, adCmdTable

' Affichage des enregistrements
While Not rst.EOF
    Debug.Print rst("Nom Ville"), rst("Code postal")
    rst.MoveNext
Wend

' Fermeture et libération des objets
rst.Close
cnn.Close
Set rst = Nothing
Set cnn = Nothing
End Function

' ---
' LECTURE D'UN NUMERO AUTO VENANT D'ETRE ATTRIBUE (DAO)
' ---
'
Sub DAO_NumeroAuto()
Dim db As DAO.Database
Dim rst As DAO.Recordset

' Ouverture de la table
Set db = CurrentDb
Set rst = db.OpenRecordset("tbl Villes Copie", dbOpenDynaset)

' Ajout d'un enregistrement
rst.AddNew
    ' Le numéro est affecté et disponible
    Debug.Print rst("Code Ville")

    ' Ajout d'autres infos
    rst("Nom Ville") = "TEST"
    rst("Code postal") = "99999"
rst.Update

' Fermeture et libération des objets
rst.Close
Set rst = Nothing
Set db = Nothing

End Sub

' ---
' LECTURE D'UN NUMERO AUTO VENANT D'ETRE ATTRIBUE (ADO)
' ---
'
Sub ADO_NumeroAuto()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset

' Ouverture de la table
Set cnn = CurrentProject.Connection
Set rst = New ADODB.Recordset
rst.Open "[tbl Villes Copie]", cnn, _
    adOpenDynamic, adLockOptimistic, adCmdTable

' Ajout d'un enregistrement
rst.AddNew
    rst("Nom Ville") = "TEST"
    rst("Code postal") = "99999"
rst.Update
Debug.Print rst("Code ville")

rst.Close
Set rst = Nothing
cnn.Close
Set cnn = Nothing
End Sub
