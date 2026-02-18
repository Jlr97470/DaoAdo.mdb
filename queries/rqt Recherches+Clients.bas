Operation =1
Option =0
Begin InputTables
    Name ="tbl Recherches"
    Name ="tbl Clients"
    Name ="tbl Villes"
End
Begin OutputColumns
    Expression ="[tbl Recherches].[Numéro Recherche]"
    Expression ="[tbl Recherches].[Numéro Client]"
    Expression ="[tbl Recherches].[Type de transaction]"
    Expression ="[tbl Recherches].[Type de bien]"
    Expression ="[tbl Recherches].Pièces"
    Expression ="[tbl Recherches].Chambres"
    Expression ="[tbl Recherches].[Surface habitable]"
    Expression ="[tbl Recherches].[Surface du terrain]"
    Expression ="[tbl Recherches].[Prix minimum]"
    Expression ="[tbl Recherches].[Prix maximum]"
    Expression ="[tbl Recherches].Remarques"
    Expression ="[tbl Clients].Titre"
    Expression ="[tbl Clients].[Nom Client]"
    Expression ="[tbl Clients].[Prénom Client]"
    Expression ="[tbl Clients].[Adresse 1]"
    Expression ="[tbl Clients].[Adresse 2]"
    Expression ="[tbl Villes].[Code postal]"
    Expression ="[tbl Villes].[Nom Ville]"
    Expression ="[tbl Villes].Pays"
End
Begin Joins
    LeftTable ="tbl Clients"
    RightTable ="tbl Recherches"
    Expression ="[tbl Clients].[Numéro Client] = [tbl Recherches].[Numéro Client]"
    Flag =1
    LeftTable ="tbl Villes"
    RightTable ="tbl Clients"
    Expression ="[tbl Villes].[Code Ville] = [tbl Clients].[Code Ville]"
    Flag =1
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
    Top =38
    Right =1002
    Bottom =432
    Left =-1
    Top =-1
    Right =977
    Bottom =274
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =155
        Top =0
        Name ="tbl Recherches"
        Name =""
    End
    Begin
        Left =216
        Top =7
        Right =336
        Bottom =236
        Top =0
        Name ="tbl Clients"
        Name =""
    End
    Begin
        Left =384
        Top =7
        Right =504
        Bottom =155
        Top =0
        Name ="tbl Villes"
        Name =""
    End
End
