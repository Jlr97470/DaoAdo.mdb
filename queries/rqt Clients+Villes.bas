Operation =1
Option =0
Begin InputTables
    Name ="tbl Villes"
    Name ="tbl Clients"
End
Begin OutputColumns
    Expression ="[tbl Clients].[Numéro Client]"
    Expression ="[tbl Clients].Titre"
    Expression ="[tbl Clients].[Nom Client]"
    Expression ="[tbl Clients].[Prénom Client]"
    Expression ="[tbl Clients].[Adresse 1]"
    Expression ="[tbl Clients].[Adresse 2]"
    Expression ="[tbl Clients].[Code Ville]"
    Expression ="[tbl Clients].[Téléphone personnel]"
    Expression ="[tbl Clients].[Téléphone mobile]"
    Expression ="[tbl Clients].[Téléphone bureau]"
    Expression ="[tbl Clients].Email"
    Expression ="[tbl Clients].[Site web]"
    Expression ="[tbl Clients].Remarques"
    Expression ="[tbl Clients].[Date de création]"
    Expression ="[tbl Clients].[Date de modification]"
    Expression ="[tbl Clients].Archiver"
    Expression ="[tbl Villes].[Code postal]"
    Expression ="[tbl Villes].[Nom Ville]"
    Expression ="[tbl Villes].Pays"
End
Begin Joins
    LeftTable ="tbl Villes"
    RightTable ="tbl Clients"
    Expression ="[tbl Villes].[Code Ville] = [tbl Clients].[Code Ville]"
    Flag =1
End
Begin OrderBy
    Expression ="[tbl Clients].[Nom Client]"
    Flag =0
    Expression ="[tbl Clients].[Prénom Client]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
Begin
    State =0
    Left =450
    Top =36
    Right =1005
    Bottom =532
    Left =-1
    Top =-1
    Right =548
    Bottom =188
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =233
        Top =7
        Right =353
        Bottom =155
        Top =0
        Name ="tbl Villes"
        Name =""
    End
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =155
        Top =0
        Name ="tbl Clients"
        Name =""
    End
End
