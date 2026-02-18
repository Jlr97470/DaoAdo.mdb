Operation =5
Option =0
Where ="((([tbl Villes Copie].[Code postal]) Like [Département] & \"%\"))"
Begin InputTables
    Name ="tbl Villes Copie"
End
Begin OutputColumns
    Expression ="[tbl Villes Copie].[Code Ville]"
    Expression ="[tbl Villes Copie].[Nom Ville]"
    Expression ="[tbl Villes Copie].[Code postal]"
    Expression ="[tbl Villes Copie].Pays"
End
Begin Parameters
    Name ="Département"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
End
Begin
    State =0
    Left =18
    Top =34
    Right =994
    Bottom =428
    Left =-1
    Top =-1
    Right =969
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =155
        Top =0
        Name ="tbl Villes Copie"
        Name =""
    End
End
