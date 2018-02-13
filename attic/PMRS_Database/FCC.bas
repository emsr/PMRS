Attribute VB_Name = "FCC"
Option Compare Database
Option Explicit

Declare Sub CalcNGSInv Lib "c:\seer_cull\CalcNGSInv.dll" _
(ByRef txlat As Single, ByRef txlon As Single, ByRef rxlat As Single, _
ByRef rxlon As Single, ByRef Datum As Long, ByRef GCD As Single, _
ByRef ForwardAzimuth As Single, ByRef BackAzimuth As Single)


Public Function Arizona_Data_Conversion()
Dim MyDB As Database
Dim MyRS As Recordset
Dim ImportIndex As Long
Dim Longitude As String
Dim Latitude As String
Dim TXLatDecimal As Single
Dim TXLonDecimal As Single
Dim RXLatDecimal As Single
Dim RXLonDecimal As Single
Dim FreeSpaceLoss As Double
Dim GCDist As Single
Dim BearTR As Single
Dim BearRT As Single


Set MyDB = CurrentDb
Set MyRS = MyDB.OpenRecordset("FCC")

MyRS.MoveLast
MyRS.MoveFirst

                For ImportIndex = 1 To MyRS.RecordCount
                    MyRS.Edit
                    
                    Longitude = "" + Left(MyRS("xlon"), 8)
        
        'Convert longitude inputs to seconds
                    If UCase(Right(Longitude, 1)) = "E" Then
                        TXLonDecimal = -((Val(Left(Longitude, 3)) * 3600) + (Val(Mid(Longitude, 4, 2)) * 60) + Val(Mid(Longitude, 6, 2))) / 3600
                    Else
                        TXLonDecimal = ((Val(Left(Longitude, 3)) * 3600) + (Val(Mid(Longitude, 4, 2)) * 60) + Val(Mid(Longitude, 6, 2))) / 3600
                    End If
        
                    MyRS("xlon") = TXLonDecimal
                                    
                    Longitude = "" + Left(MyRS("rlon"), 8)
        
        'Convert longitude inputs to seconds
                    If UCase(Right(Longitude, 1)) = "E" Then
                        RXLonDecimal = -((Val(Left(Longitude, 3)) * 3600) + (Val(Mid(Longitude, 4, 2)) * 60) + Val(Mid(Longitude, 6, 2))) / 3600
                    Else
                        RXLonDecimal = ((Val(Left(Longitude, 3)) * 3600) + (Val(Mid(Longitude, 4, 2)) * 60) + Val(Mid(Longitude, 6, 2))) / 3600
                    End If
        
                    MyRS("rlon") = RXLonDecimal
                    
                    
                    Latitude = Left(Trim(MyRS("xlat")), 8)
                
        'Convert latitude inputs to seconds
                    If UCase(Right(Latitude, 1)) = "N" Then
                        TXLatDecimal = ((Val(Left(Latitude, 2)) * 3600) + (Val(Mid(Latitude, 3, 2)) * 60) + Val(Mid(Latitude, 5, 2))) / 3600
                    Else
                        TXLatDecimal = -((Val(Left(Latitude, 2)) * 3600) + (Val(Mid(Latitude, 3, 2)) * 60) + Val(Mid(Latitude, 5, 2))) / 3600
                    End If
    
                    MyRS("xlat") = TXLatDecimal
                    
                    Latitude = Left(Trim(MyRS("rlat")), 8)
                
        'Convert latitude inputs to seconds
                    If UCase(Right(Latitude, 1)) = "N" Then
                        RXLatDecimal = ((Val(Left(Latitude, 2)) * 3600) + (Val(Mid(Latitude, 3, 2)) * 60) + Val(Mid(Latitude, 5, 2))) / 3600
                    Else
                        RXLatDecimal = -((Val(Left(Latitude, 2)) * 3600) + (Val(Mid(Latitude, 3, 2)) * 60) + Val(Mid(Latitude, 5, 2))) / 3600
                    End If
    
                    MyRS("rlat") = RXLatDecimal
                    
                    'dist returned in meters
                    Call CalcNGSInv(TXLatDecimal / 57.2958, TXLonDecimal / 57.2958, RXLatDecimal / 57.2958, RXLonDecimal / 57.2958, 1, GCDist, BearTR, BearRT)
                    
                    'freespace for MHz and meters
                    FreeSpaceLoss = 20 * (Log(MyRS("Freq")) / Log(10)) + 20 * (Log(GCDist) / Log(10)) - 27.5
                    MyRS("dbloss") = Val(Format(MyRS("dbloss") - FreeSpaceLoss, "0.0"))
                    
                    'convert dist to km
                    MyRS("dist") = Val(Format(GCDist / 1000, "0.0"))
                    
                    'convert ht from ft to meters
                    MyRS("xht") = Val(Format(MyRS("xht") * 0.3048, "0.0"))
                    MyRS("rht") = Val(Format(MyRS("rht") * 0.3048, "0.0"))
                    
                    MyRS.Update
                    MyRS.MoveNext
               
                Next
            

End Function
