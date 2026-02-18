Attribute VB_Name = "mod DAO Bases"
Option Compare Database
Option Explicit
' ---
' INFORMATIONS FOURNIES PAR DBENGINE
' ---
Sub DAO_InfosMoteur()
' Informations sur le moteur de base de données
Debug.Print "VERSION: " & DBEngine.Version
Debug.Print "MDW: " & DBEngine.SystemDB

' Informations sur la session en cours
Debug.Print "SESSIONS: " & DBEngine.Workspaces.Count
Debug.Print "BASES DE DONNEES: " & DBEngine.Workspaces(0).Databases.Count
Debug.Print "NOM DE LA BASE 0: " & DBEngine.Workspaces(0).Databases(0).Name
End Sub

' ---
' INFORMATIONS FOURNIES PAR DBENGINE
' ---
Sub DAO_InfosMoteur2()
' Informations sur le moteur de base de données
With DBEngine
    Debug.Print "VERSION: " & .Version
    Debug.Print "MDW: " & .SystemDB

' Informations sur la session en cours
    Debug.Print "SESSIONS: " & .Workspaces.Count
    With .Workspaces(0)
        Debug.Print "BASES DE DONNEES: " & .Databases.Count
        Debug.Print "NOM DE LA BASE 0: " & .Databases(0).Name
    End With
End With
End Sub
