Operation =1
Option =0
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
Begin OrderBy
    Expression ="[tbl Visites].[Date Visite]"
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
