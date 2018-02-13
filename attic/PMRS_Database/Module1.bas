Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Public Function Build_Prop()
Dim MyDB As Database
Dim MyRS As Recordset

Set MyDB = CurrentDb
Set MyRS = MyDB.OpenRecordset("Area Codes")

MyDB.Execute "DELETE * FROM PropData;"

'loop thru recordset to obtain all locations
Do While Not MyRS.EOF
    MyDB.Execute "INSERT INTO PropData SELECT * FROM [" + MyRS("region") + "];"
    MyDB.Execute "UPDATE PropData SET PropData.Region = '" + MyRS("region") + "' WHERE PropData.Area = '" + MyRS("area") + "';"
    MyRS.MoveNext
Loop

End Function
