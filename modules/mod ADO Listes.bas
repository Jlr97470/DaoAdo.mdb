Attribute VB_Name = "mod ADO Listes"
Option Compare Database
Option Explicit

' ---
' LISTES DIVERSES (ADOX)
' ---
'
Function ADOX_Listes()
Dim cat As ADOX.Catalog     ' Catalogue
Dim tbl As ADOX.Table       ' Table
Dim vw As ADOX.View         ' Vue
Dim prc As ADOX.Procedure   ' Procédure stockée

' Connexion au catalogue de la base de données
Set cat = New ADOX.Catalog
cat.ActiveConnection = CurrentProject.Connection

' Liste des procédures stockées (ou requêtes Action)
For Each prc In cat.Procedures
    Debug.Print "-----"
    Debug.Print UCase(prc.Name)
    Debug.Print prc.Command.CommandText
    Debug.Print
Next
Debug.Print

' Liste des tables
Debug.Print "------"
Debug.Print "TABLES"
For Each tbl In cat.Tables
    Debug.Print tbl.Name, tbl.Type
Next
Debug.Print

' Liste des vues (ou requêtes Sélection)
Debug.Print "------"
Debug.Print "VIEWS"
For Each vw In cat.Views
    Debug.Print vw.Name
Next

Set cat = Nothing
End Function

' ---
' LISTE DETAILLEE DES TABLES VIA ADOX
' ---
Sub ADOX_DetailTables()
Dim cat As ADOX.Catalog
Dim tbl As ADOX.Table
Dim col As ADOX.Column

Set cat = New ADOX.Catalog
Set cat.ActiveConnection = CurrentProject.Connection
For Each tbl In cat.Tables
    Debug.Print UCase(tbl.Name)
    For Each col In tbl.Columns
        Debug.Print " -> "; col.Name, col.Type
    Next
    Debug.Print
Next

Set cat = Nothing
End Sub
