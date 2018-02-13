Attribute VB_Name = "Declares"
'******************************************************************************
'
' *** COPYRIGHT:  Copyright 1999
'                 IIT Research Institute
'                 US Govt retains rights in accordance
'                 with DoD FAR Supp 252.227 - 7013.
'
'-----------------------------------------------------------------------
'
' *** CODING HISTORY:
'
'   DATE      PROGRAMMER                     DESCRIPTION
' --------   ------------                   -------------
' 4/21/99    Mike Fahler                    PMRS v. 1.00
'******************************************************************************

Public MyDatabase As Database
Public QueryResultsRS As Recordset
Public QueryString As String
Public PropResultRS As Recordset
Public StatisticsRS As Recordset

Declare Sub CalcNGSInv Lib "CalcNGSInv.dll" _
(ByRef TXLat As Single, ByRef TXLon As Single, ByRef RXLat As Single, _
ByRef RXLon As Single, ByRef Datum As Long, ByRef GCD As Single, _
ByRef ForwardAzimuth As Single, ByRef BackAzimuth As Single)

Declare Sub CalcNGSFor Lib "CalcNGSFor.dll" _
(ByRef TXLat As Single, ByRef TXLon As Single, ByRef Datum As Long, _
ByRef GCD As Single, ByRef ForwardAzimuth As Single, ByRef RXLat As Single, _
ByRef RXLon As Single, ByRef BackAzimuth As Single)

Declare Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long

'help engine declares
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_CONTENTS = &H3&
Public Const HELP_FINDER = &HB
Public Const HELP_QUIT = &H2
Public Const HELP_HELPONHELP = &H4
Public Declare Function WINHELP Lib "USER32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Public Sub End_Program()
    Dim response As Integer
    
    On Error Resume Next
    
    response = MsgBox("Are you sure you want to exit the application?", vbQuestion + vbYesNo)
    
    If response = 7 Then 'no
        Exit Sub
    End If
    
    For X = Forms.COUNT - 1 To 0 Step -1
        Unload Forms(X)
    Next
    
    MyDatabase.Close
    
    End

End Sub
