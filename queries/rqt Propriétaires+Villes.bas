Operation =1
Option =0
Begin InputTables
    Name ="tbl Villes"
    Name ="tbl Propriétaires"
End
Begin OutputColumns
    Expression ="[tbl Propriétaires].[Numéro Propriétaire]"
    Expression ="[tbl Propriétaires].Titre"
    Expression ="[tbl Propriétaires].[Nom Propriétaire]"
    Expression ="[tbl Propriétaires].[Prénom Propriétaire]"
    Expression ="[tbl Propriétaires].[Adresse 1]"
    Expression ="[tbl Propriétaires].[Adresse 2]"
    Expression ="[tbl Propriétaires].[Code Ville]"
    Expression ="[tbl Propriétaires].[Téléphone personnel]"
    Expression ="[tbl Propriétaires].[Téléphone mobile]"
    Expression ="[tbl Propriétaires].[Téléphone bureau]"
    Expression ="[tbl Propriétaires].Email"
    Expression ="[tbl Propriétaires].[Site web]"
    Expression ="[tbl Propriétaires].Remarques"
    Expression ="[tbl Propriétaires].[Date de création]"
    Expression ="[tbl Propriétaires].[Date de modification]"
    Expression ="[tbl Propriétaires].Archiver"
    Expression ="[tbl Villes].[Code postal]"
    Expression ="[tbl Villes].[Nom Ville]"
    Expression ="[tbl Villes].Pays"
End
Begin Joins
    LeftTable ="tbl Villes"
    RightTable ="tbl Propriétaires"
    Expression ="[tbl Villes].[Code Ville]=[tbl Propriétaires].[Code Ville]"
    Flag =1
End
Begin OrderBy
    Expression ="[tbl Propriétaires].[Nom Propriétaire]"
    Flag =0
    Expression ="[tbl Propriétaires].[Prénom Propriétaire]"
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
    Left =18
    Top =40
    Right =1002
    Bottom =434
    Left =-1
    Top =-1
    Right =973
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =155
        Top =0
        Name ="tbl Propriétaires"
        Name =""
    End
    Begin
        Left =292
        Top =6
        Right =412
        Bottom =154
        Top =0
        Name ="tbl Villes"
        Name =""
    End
End
