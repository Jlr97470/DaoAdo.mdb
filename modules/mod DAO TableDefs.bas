Attribute VB_Name = "mod DAO TableDefs"
Option Compare Database
Option Explicit

' ---
' CREATION D'UNE TABLE VIA DAO
' ---
'
Sub DAO_CreationTable()
Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim fld As DAO.Field

Set db = CurrentDb
Set tdf = New DAO.TableDef
tdf.Name = "tbl Test1"

Set fld = New DAO.Field
fld.Name = "Numéro"
fld.Type = dbLong
fld.Attributes = dbAutoIncrField
tdf.Fields.Append fld

Set fld = New DAO.Field
fld.Name = "Nom Personne"
fld.Type = dbText
fld.Size = 30
tdf.Fields.Append fld

Set fld = New DAO.Field
fld.Name = "Date de naissance"
fld.Type = dbDate
tdf.Fields.Append fld
tdf.Fields.Refresh

With db.TableDefs
    .Append tdf
    .Refresh
End With
End Sub
