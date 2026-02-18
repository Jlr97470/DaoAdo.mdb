Attribute VB_Name = "mod Administration"
Option Compare Database
Option Explicit
' ---
' COMPACTAGE D'UNE BASE DE DONNEES
' ---
' Entrée : strBase1 <- Chemin complet de la base à compacter
'          strBase2 <- Chemin complet de la base temporaire à créer
'
Function Compacter(ByVal strBase1 As String, _
    ByVal strBase2 As String)
Dim db As DAO.Database

Const CMP_OK = 0
Const CMP_SOURCE_INTROUVABLE = -1
Const CMP_DESTINATION_EXISTE = -2
Const CMP_SOURCE_OUVERTE = -3
Const CMP_DESTINATION_INTROUVABLE = -4

' Vérifier si la base de données de départ existe
If Dir(strBase1) = "" Then
    Compacter = CMP_SOURCE_INTROUVABLE
    Exit Function
End If

' Vérifier si la base d'arrivée existe
If Dir(strBase2) <> "" Then
    Compacter = CMP_DESTINATION_EXISTE
    Exit Function
End If

' Tenter d'ouvrir la base de départ en mode exclusif
' pour vérifier si elle est en cours d'exploitation
On Error Resume Next
Set db = OpenDatabase(strBase1, True)
If Err.Number <> 0 Then
    Compacter = CMP_SOURCE_OUVERTE
    Exit Function
End If
Set db = Nothing

' Compacter la base de données
DBEngine.CompactDatabase strBase1, strBase2

' Permuter les fichiers ancien et nouveau
If Dir(strBase2) = "" Then
    Compacter = CMP_DESTINATION_INTROUVABLE
Else
    Kill strBase1
    Name strBase2 As strBase1
    Compacter = CMP_OK
End If
End Function

' ---
' RECONNEXION DE BASES FRONTALE ET DORSALE
'
' Entrée : strChemin <- Nouvel emplacement de la base dorsale
'
Function RetablirLiaisons(ByVal strChemin As String) As Boolean
Dim db As DAO.Database
Dim tdf As DAO.TableDef

' Vérifier si la base dorsale existe
If Dir(strChemin) = "" Then
    RetablirLiaisons = False
    Exit Function
End If

Set db = CurrentDb()
For Each tdf In db.TableDefs
    If (tdf.Attributes And dbSystemObject) = 0 Then
        ' Si la propriété Connect est renseignée, la table
        ' est attachée. Refaire une liaison du type :
        ' ;DATABASE=C:\...\dorsale.mdb
        If tdf.Connect <> "" Then
            tdf.Connect = ";DATABASE=" & strChemin
            tdf.RefreshLink
        End If
    End If
Next

Set db = Nothing
RetablirLiaisons = True
End Function
