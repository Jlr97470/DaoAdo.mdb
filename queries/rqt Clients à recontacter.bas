Operation =1
Option =0
Where ="((([tbl Visites].[Prochain contact]) Between [Date de début] And [Date de fin]))"
Begin InputTables
    Name ="tbl Visites"
End
Begin OutputColumns
    Expression ="[tbl Visites].[Numéro Bien]"
    Expression ="[tbl Visites].[Numéro Client]"
    Expression ="[tbl Visites].[Date Visite]"
    Expression ="[tbl Visites].[Prochain contact]"
    Expression ="[tbl Visites].Intérêt"
    Expression ="[tbl Visites].Remarques"
End
Begin Parameters
    Name ="[Date de début]"
    Flag =8
    Name ="[Date de fin]"
    Flag =8
End
Begin OrderBy
    Expression ="[tbl Visites].[Prochain contact]"
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
    Top =38
    Right =994
    Bottom =432
    Left =-1
    Top =-1
    Right =969
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
        Name ="tbl Visites"
        Name =""
    End
End
