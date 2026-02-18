Attribute VB_Name = "mod ADO Command"
Option Compare Database
Option Explicit


' ---
' EXECUTION DE COMMANDES SQL SUR LA BASE DE DONNEES
' ---
'
Function ADO_ExecuterSQL()
Dim cnn As ADODB.Connection
Dim cat As ADOX.Catalog
Dim cmd As ADODB.Command
Dim strSQL As String
Dim lngEnregistrements As Long

' Exécution d'une instruction SQL
' à partir de la connexion en cours
Set cnn = CurrentProject.Connection
strSQL = "DELETE * FROM [tbl Villes Copie]" & _
    " WHERE Left([Code postal],2)='13'"
cnn.Execute strSQL, lngEnregistrements
Debug.Print "Nombre d'enregistrements concernés : " & lngEnregistrements

' Exécution d'une instruction SQL
' à partir d'un objet Command
Set cmd = New ADODB.Command
cmd.ActiveConnection = CurrentProject.Connection
strSQL = "DELETE * FROM [tbl Villes Copie]" & _
    " WHERE Left([Code postal],2)='59'"
cmd.CommandText = strSQL
cmd.CommandType = adCmdText
cmd.Execute lngEnregistrements
Debug.Print "Nombre d'enregistrements concernés : " & lngEnregistrements

' Exécution d'une requête paramétrée
Set cat = New ADOX.Catalog
Set cat.ActiveConnection = CurrentProject.Connection
Set cmd = cat.Procedures("rqt Suppression Villes par département").Command
cmd.Parameters("Département").Value = "29"
cmd.Execute lngEnregistrements
Debug.Print "Nombre d'enregistrements concernés : " & lngEnregistrements

' Fermeture et libération des objets
Set cmd = Nothing
Set cat = Nothing
Set cnn = Nothing
End Function
