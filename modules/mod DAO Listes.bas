Attribute VB_Name = "mod DAO Listes"
Option Compare Database

' ---
' LISTES DIVERSES
' ---
'
Function DAO_Listes()
Dim db As DAO.Database
Dim tdf As DAO.TableDef     ' Table
Dim rel As DAO.Relation     ' Relation
Dim qdf As DAO.QueryDef     ' Requête
Dim cnt As DAO.Container    ' Container

' La base actuelle...
Set db = CurrentDb

' Liste des tables
Debug.Print "TABLES" & vbCrLf & "------"
For Each tdf In db.TableDefs
    Debug.Print tdf.Name
Next
Debug.Print

' Liste des relations
Debug.Print "RELATIONS" & vbCrLf & "---------"
For Each rel In db.Relations
    Debug.Print rel.Table & " / " & rel.ForeignTable
Next
Debug.Print

' Liste des requêtes
Debug.Print "REQUETES" & vbCrLf & "--------"
For Each qdf In db.QueryDefs
    Debug.Print qdf.Name
Next
Debug.Print

' Liste des containers
Debug.Print "CONTAINERS" & vbCrLf & "----------"
For Each cnt In db.Containers
    Debug.Print cnt.Name & " " & cnt.Documents.Count
Next
Debug.Print

' Libération des objets
Set db = Nothing
End Function

' ---
' LISTES OBTENUES A PARTIR DES CONTAINERS DAO
' ---
'
Function DAO_ListesContainers()
Dim db As DAO.Database
Dim cnt As DAO.Container    ' Container
Dim doc As DAO.Document     ' Objet de base de données

' La base actuelle...
Set db = CurrentDb

' Liste des containers
Debug.Print "CONTAINERS" & vbCrLf & "----------"
For Each cnt In db.Containers
    Debug.Print cnt.Name & " " & cnt.Documents.Count

    ' Liste des documents liés
'    For Each doc In cnt.Documents
'        Debug.Print doc.Name & " " & " " & doc.Owner & " " & doc.LastUpdated
'    Next
'    Debug.Print
Next

' Libération des objets
Set db = Nothing
End Function

' ---
' LISTES DES GROUPES ET DES UTILISATEURS DE LA BASE DE DONNEES
' ---
'
Function DAO_ListeGroupes()
Dim wrk As DAO.Workspace
Dim grp As DAO.Group
Dim usr As DAO.User

' L'espace de travail actuel...
Set wrk = DBEngine.Workspaces(0)

' Liste des groupes
Debug.Print "GROUPES" & vbCrLf & "-------"
For Each grp In wrk.Groups
    Debug.Print grp.Name & " " & grp.Users.Count

    ' Liste des utilisateurs du groupe
    For Each usr In grp.Users
        Debug.Print "> " & usr.Name
    Next
Next

' Libération des objets
Set wrk = Nothing
End Function
' ---
' LISTES DES GROUPES ET DES UTILISATEURS DE LA BASE DE DONNEES
' ---
Function DAO_ListeGroupes2()
Dim wrk As DAO.Workspace
Dim grp As DAO.Group
Dim usr As DAO.User

' L'espace de travail actuel...
Set wrk = DBEngine.Workspaces(0)

' Liste des groupes
Debug.Print "UTILISATEURS" & vbCrLf & "-------"
For Each usr In wrk.Users
    Debug.Print usr.Name

    For Each grp In usr.Groups
        Debug.Print "> " & grp.Name
    Next
Next

' Libération des objets
Set wrk = Nothing
End Function
' ---
' LISTES DETAILLEES DES TABLES
' ---
'
Function DAO_DetailTables()
Dim db As DAO.Database
Dim tdf As DAO.TableDef     ' Table
Dim fld As DAO.Field        ' Champ de table
Dim prp As DAO.Property     ' Propriété de champ
Dim intAfficherProprietes As Integer

' La base actuelle...
On Error Resume Next
Set db = CurrentDb

intAfficherProprietes = MsgBox("Souhaitez-vous le détail des propriétés ?", _
    vbQuestion + vbYesNo, "Access Référence")

' Liste des tables
Debug.Print "TABLES" & vbCrLf & "------"
For Each tdf In db.TableDefs
    If (tdf.Attributes And dbSystemObject) = 0 Then
        Debug.Print UCase(tdf.Name)

        ' Liste des champs
        For Each fld In tdf.Fields
            Debug.Print "-> "; fld.Name, fld.Type, fld.Size

            ' Liste des propriétés du champ
            If intAfficherProprietes = vbYes Then
                For Each prp In fld.Properties
                    Debug.Print "   - " & prp.Name & " = " & prp.Value
                Next
            End If
        Next
        Debug.Print
    End If
Next

' Libération des objets
Set db = Nothing
End Function
