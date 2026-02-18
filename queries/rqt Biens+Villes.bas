Operation =1
Option =0
Begin InputTables
    Name ="tbl Villes"
    Name ="tbl Propriétaires"
    Name ="tbl Villes"
    Alias ="tbl Villes_1"
    Name ="tbl Biens"
End
Begin OutputColumns
    Expression ="[tbl Biens].[Numéro Bien]"
    Expression ="[tbl Biens].[Numéro Propriétaire]"
    Expression ="[tbl Biens].[Nom Bien]"
    Expression ="[tbl Biens].[Adresse 1]"
    Expression ="[tbl Biens].[Adresse 2]"
    Expression ="[tbl Biens].[Code Ville]"
    Expression ="[tbl Biens].[Code Transaction]"
    Expression ="[tbl Biens].[Code Bien]"
    Expression ="[tbl Biens].[Code Catégorie]"
    Expression ="[tbl Biens].Pièces"
    Expression ="[tbl Biens].Chambres"
    Expression ="[tbl Biens].Surface"
    Expression ="[tbl Biens].[Surface Terrain]"
    Expression ="[tbl Biens].Descriptif"
    Expression ="[tbl Biens].Montant"
    Expression ="[tbl Biens].Statut"
    Expression ="[tbl Biens].[Date de création]"
    Expression ="[tbl Biens].[Date de modification]"
    Expression ="[tbl Villes].[Code postal]"
    Expression ="[tbl Villes].[Nom Ville]"
    Expression ="[tbl Villes].Pays"
    Expression ="[tbl Propriétaires].[Nom Propriétaire]"
    Expression ="[tbl Propriétaires].[Prénom Propriétaire]"
    Expression ="[tbl Propriétaires].Titre"
    Alias ="Code postal Propriétaire"
    Expression ="[tbl Villes_1].[Code postal]"
    Alias ="Nom Ville Propriétaire"
    Expression ="[tbl Villes_1].[Nom Ville]"
End
Begin Joins
    LeftTable ="tbl Propriétaires"
    RightTable ="tbl Villes_1"
    Expression ="[tbl Propriétaires].[Code Ville]=[tbl Villes_1].[Code Ville]"
    Flag =2
    LeftTable ="tbl Propriétaires"
    RightTable ="tbl Biens"
    Expression ="[tbl Propriétaires].[Numéro Propriétaire]=[tbl Biens].[Numéro Propriétaire]"
    Flag =1
    LeftTable ="tbl Villes"
    RightTable ="tbl Biens"
    Expression ="[tbl Villes].[Code Ville]=[tbl Biens].[Code Ville]"
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
    Top =104
    Right =1002
    Bottom =498
    Left =-1
    Top =-1
    Right =1017
    Bottom =180
    Left =0
    Top =6
    ColumnsShown =539
    Begin
        Left =439
        Top =7
        Right =559
        Bottom =155
        Top =6
        Name ="tbl Propriétaires"
        Name =""
    End
    Begin
        Left =216
        Top =1
        Right =336
        Bottom =149
        Top =0
        Name ="tbl Villes"
        Name =""
    End
    Begin
        Left =48
        Top =1
        Right =168
        Bottom =149
        Top =0
        Name ="tbl Biens"
        Name =""
    End
    Begin
        Left =607
        Top =7
        Right =727
        Bottom =155
        Top =0
        Name ="tbl Villes_1"
        Name =""
    End
End
