VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.0#0"; "RESIZE32.OCX"
Begin VB.Form frmPropResult 
   Caption         =   "Propagation Results"
   ClientHeight    =   5370
   ClientLeft      =   2775
   ClientTop       =   4515
   ClientWidth     =   9060
   Icon            =   "frmPropResult.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5370
   ScaleWidth      =   9060
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "C:\PMRS\TIREMTest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PropResults"
      Top             =   4725
      Width           =   3660
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3840
      _Version        =   196608
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
      Enabled         =   -1  'True
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5370
      FormDesignWidth =   9060
   End
   Begin MSDBGrid.DBGrid dbgPropResults 
      Bindings        =   "frmPropResult.frx":030A
      Height          =   4125
      Left            =   360
      OleObjectBlob   =   "frmPropResult.frx":031A
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.Label lblNumRecords 
      AutoSize        =   -1  'True
      Caption         =   "Forward"
      Height          =   195
      Left            =   6465
      TabIndex        =   2
      Top             =   4800
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Reverse"
      Height          =   195
      Left            =   1815
      TabIndex        =   1
      Top             =   4800
      Width           =   600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewQuery 
         Caption         =   "&New Query"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Propagation Comparison"
         Begin VB.Menu mnuAccess 
            Caption         =   "&Access"
         End
         Begin VB.Menu mnuExcel 
            Caption         =   "&Excel"
            Begin VB.Menu mnuExcel3 
               Caption         =   "&3.0"
            End
            Begin VB.Menu mnuExcel4 
               Caption         =   "&4.0"
            End
            Begin VB.Menu mnuExcel5 
               Caption         =   "&5.0, 7.0"
            End
            Begin VB.Menu mnuExcel97 
               Caption         =   "&97"
            End
         End
         Begin VB.Menu mnuFoxPro 
            Caption         =   "&FoxPro"
            Begin VB.Menu mnuFoxPro2 
               Caption         =   "&2.0"
            End
            Begin VB.Menu mnuFoxPro25 
               Caption         =   "2.&5"
            End
            Begin VB.Menu mnuFoxPro26 
               Caption         =   "2.&6"
            End
         End
         Begin VB.Menu mnuLotus 
            Caption         =   "&Lotus"
            Begin VB.Menu mnuLotus1 
               Caption         =   "wk&1"
            End
            Begin VB.Menu mnuLotus3 
               Caption         =   "wk&3"
            End
         End
         Begin VB.Menu mnuText 
            Caption         =   "&Comma Delimited Text"
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuStatistics 
         Caption         =   "&Statistics/Plot"
         Begin VB.Menu mnuPathLoss 
            Caption         =   "&Total Path Loss"
         End
         Begin VB.Menu mnuExcessLoss 
            Caption         =   "&Excess Loss"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On..."
      End
   End
   Begin VB.Menu mnuSortOrder 
      Caption         =   "SortOrder"
      Visible         =   0   'False
      Begin VB.Menu mnuAscending 
         Caption         =   "Sort Ascending"
      End
      Begin VB.Menu mnuSortDescending 
         Caption         =   "Sort Descending"
      End
   End
End
Attribute VB_Name = "frmPropResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim DataChanged As Boolean
Dim SortField As String
Dim SecondPrompt As Boolean

Dim FieldIndex As Integer

Dim TotalPathLoss As Boolean  'if true correlation coefficient is calculated for total path loss, if false freespace loss is excluded

Private Sub Data1_Reposition()
    
Data1.Caption = "Record " + Str(Data1.Recordset.AbsolutePosition + 1) + " of " + Str(PropResultRS.RecordCount)

End Sub


Private Sub dbgPropResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If dbgPropResults.SelStartCol > -1 Then
    If Button = 2 Then  'right click
        PopupMenu mnuSortOrder
    End If
End If

End Sub

Private Sub Form_Activate()

Dim ForCounter As Integer

For FieldIndex = 0 To dbgPropResults.Columns.COUNT - 1
    dbgPropResults.Columns.Remove (0)  'remove first
Next

ForCounter = 0

For FieldIndex = 0 To PropResultRS.Fields.COUNT - 1
    If Not (Right(Trim(PropResultRS.Fields(FieldIndex).Name), 2) = "ID" Or _
       PropResultRS.Fields(FieldIndex).Name = "Profile" Or _
       PropResultRS.Fields(FieldIndex).Name = "NumPoints") Then
        
        Add_Field_to_Grid (FieldIndex - ForCounter)
    Else
        ForCounter = ForCounter + 1
    End If
Next
    
dbgPropResults.ReBind
dbgPropResults.Refresh

End Sub

Public Sub Form_Load()

On Error GoTo errorhandler

'center form on screen
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
    
Data1.DatabaseName = App.Path + "\Tiremtest.mdb"

SortField = "[Difference(db)] DESC"

Set_Recordsets  'call sub

Data1.Refresh

errorhandler:
    Select Case Err.Number
    Case 3021  'no records selected - movefirst and movelast generate errors in this condition
        Resume Next
    Case 3061, 3075 'query failed - set variable and return to query builder
        SQLError = True
    Case Is > 0
        MsgBox "Application Error." + Chr(10) + Chr(13) + Err.Description, vbExclamation + vbOKOnly, "Application Error"
        Exit Sub
    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    mnuNewQuery_Click
    Cancel = 0
End If
End Sub

Private Sub mnuAscending_Click()
Call Sort_Output(True)
End Sub

Private Sub mnuContents_Click()
'Dim return_value
'return_value = WINHELP(hWnd, App.Path + "\gate.hlp", HELP_CONTENTS, 0&)
End Sub

Private Sub mnuExcel97_Click()
frmSQLQuery.cdlProject.DefaultExt = ".xls"
frmSQLQuery.cdlProject.Filter = "Excel 97 Worksheet (*.xls)|*.xls"
DataType = "exc97"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuExcessLoss_Click()
TotalPathLoss = False

Generate_Statistics
Load frmProfile

End Sub

Private Sub mnuExit_Click()
    
    End_Program
End Sub

Private Sub mnuHelpSearch_Click()
'Dim return_value
'return_value = WINHELP(hWnd, App.Path + "\gate.hlp", HELP_PARTIALKEY, 0&)
End Sub

Private Sub mnuNewQuery_Click()
    For X = Forms.COUNT - 1 To 0 Step -1
        If Forms(X).Name <> "frmSQLQuery" Then
            Unload Forms(X)
        End If
    Next
    frmSQLQuery.Show
    
End Sub

Private Sub mnuAccess_Click()

frmSQLQuery.cdlProject.DefaultExt = ".mdb"
frmSQLQuery.cdlProject.Filter = "Access Database (*.mdb)|*.mdb"
DataType = "acces"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuExcel3_Click()
frmSQLQuery.cdlProject.DefaultExt = ".xls"
frmSQLQuery.cdlProject.Filter = "Excel 3.0 Worksheet (*.xls)|*.xls"
DataType = "exc30"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuExcel4_Click()
frmSQLQuery.cdlProject.DefaultExt = ".xls"
frmSQLQuery.cdlProject.Filter = "Excel 4.0 Worksheet (*.xls)|*.xls"
DataType = "exc40"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuExcel5_Click()
frmSQLQuery.cdlProject.DefaultExt = ".xls"
frmSQLQuery.cdlProject.Filter = "Excel 5.0, 7.0 Worksheet (*.xls)|*.xls"
DataType = "exc50"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuFoxPro2_Click()
frmSQLQuery.cdlProject.DefaultExt = ".dbf"
frmSQLQuery.cdlProject.Filter = "FoxPro 2.0 Database (*.dbf)|*.dbf"
DataType = "fox20"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuFoxPro25_Click()
frmSQLQuery.cdlProject.DefaultExt = ".dbf"
frmSQLQuery.cdlProject.Filter = "FoxPro 2.5 Database (*.dbf)|*.dbf"
DataType = "fox25"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuFoxPro26_Click()
frmSQLQuery.cdlProject.DefaultExt = ".dbf"
frmSQLQuery.cdlProject.Filter = "FoxPro 2.6 Database (*.dbf)|*.dbf"
DataType = "fox26"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuLotus1_Click()
frmSQLQuery.cdlProject.DefaultExt = ".wk1"
frmSQLQuery.cdlProject.Filter = "Lotus 123 (*.wk1)|*.wk1"
DataType = "lot01"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuLotus3_Click()
frmSQLQuery.cdlProject.DefaultExt = ".wk3"
frmSQLQuery.cdlProject.Filter = "Lotus 123 3.0 (*.wk3)|*.wk3"
DataType = "lot03"
Set_Export_Strings  'call sub

End Sub

Private Sub mnuPathLoss_Click()
TotalPathLoss = True
Generate_Statistics
Load frmProfile

End Sub

Private Sub mnuSortDescending_Click()
    Call Sort_Output(False)
End Sub


Private Sub mnuText_Click()
frmSQLQuery.cdlProject.DefaultExt = ".txt"
frmSQLQuery.cdlProject.Filter = "Comma Delimited Text (*.txt)|*.txt"
DataType = "comma"
Set_Export_Strings  'call sub
End Sub

Public Sub Set_Recordsets()
Set PropResultRS = MyDatabase.OpenRecordset("SELECT PropData.*,PropComparison.* FROM PropData,PropComparison WHERE " _
                        + "PropData.ID = PropComparison.TestID" _
                        + " ORDER BY " + SortField)
    
Data1.RecordSource = "SELECT PropData.*,PropComparison.* FROM PropData,PropComparison WHERE " _
                        + "PropData.ID = PropComparison.TestID" _
                        + " ORDER BY " + SortField
    
PropResultRS.MoveLast
PropResultRS.MoveFirst

End Sub

Public Sub Sort_Output(Ascending As Boolean)
Dim SortIndex As Integer

If dbgPropResults.SelStartCol >= 14 Then  'mod since 2 fields from recordset are not in datagrid
    SortIndex = dbgPropResults.SelStartCol + 4
Else
    SortIndex = dbgPropResults.SelStartCol
End If

If Ascending = True Then
    SortField = "[" + PropResultRS.Fields(SortIndex).Name + "]"
Else
    SortField = "[" + PropResultRS.Fields(SortIndex).Name + "] DESC"
End If

Set_Recordsets  'call sub

Data1.Refresh
dbgPropResults.ReBind
dbgPropResults.Refresh

End Sub


Public Sub Add_Field_to_Grid(FieldID As Integer)

    dbgPropResults.Columns.Add (FieldID)
    dbgPropResults.Columns(FieldID).Caption = PropResultRS.Fields(FieldIndex).Name
    dbgPropResults.Columns(FieldID).DataField = PropResultRS.Fields(FieldIndex).Name
    dbgPropResults.Columns(FieldID).Visible = True

End Sub

Public Sub Set_Export_Strings()
    Export "PropData.*,PropComparison.*", "PropData,PropComparison", _
                "PropData.ID = PropComparison.TestID ORDER BY " + SortField

End Sub

Public Sub Generate_Statistics()
Dim StatCount As Long
Dim StatMode As String
Dim TempRS As Recordset
Dim StatIndex As Integer
Dim StDevValue As Double

'variables for correlation coefficient
Dim SumXY As Double
Dim SumX As Double
Dim SumY As Double
Dim SumX2 As Double
Dim SumY2 As Double
Dim CorrCoefNumerator As Double
Dim CorrCoefDenominator As Double
Dim CorrCoef As Double

MyDatabase.Execute "DELETE * FROM Statistics"

'execute queries to fill Mean, StDev, nad Count Fields
MyDatabase.Execute "INSERT INTO Statistics ( Mode, Mean, StDev, Count ) " _
                    + "SELECT 'LOS', Avg(PropComparison.[Difference(dB)]) AS [AvgOfDifference(dB)], " _
                    + "StDev(PropComparison.[Difference(dB)]) AS [StDevOfDifference(dB)], " _
                    + "Count(PropComparison.[Difference(dB)]) AS [CountOfDifference(dB)] " _
                    + "From PropComparison WHERE PropComparison.PropMode='LOS';"

MyDatabase.Execute "INSERT INTO Statistics ( Mode, Mean, StDev, Count ) " _
                    + "SELECT 'DIF', Avg(PropComparison.[Difference(dB)]) AS [AvgOfDifference(dB)], " _
                    + "StDev(PropComparison.[Difference(dB)]) AS [StDevOfDifference(dB)], " _
                    + "Count(PropComparison.[Difference(dB)]) AS [CountOfDifference(dB)] " _
                    + "From PropComparison WHERE PropComparison.PropMode='DIF';"

MyDatabase.Execute "INSERT INTO Statistics ( Mode, Mean, StDev, Count ) " _
                    + "SELECT 'TRO', Avg(PropComparison.[Difference(dB)]) AS [AvgOfDifference(dB)], " _
                    + "StDev(PropComparison.[Difference(dB)]) AS [StDevOfDifference(dB)], " _
                    + "Count(PropComparison.[Difference(dB)]) AS [CountOfDifference(dB)] " _
                    + "From PropComparison WHERE PropComparison.PropMode='TRO';"

MyDatabase.Execute "INSERT INTO Statistics ( Mode, Mean, StDev, Count ) " _
                    + "SELECT 'TIREM', Avg(PropComparison.[Difference(dB)]) AS [AvgOfDifference(dB)], " _
                    + "StDev(PropComparison.[Difference(dB)]) AS [StDevOfDifference(dB)], " _
                    + "Count(PropComparison.[Difference(dB)]) AS [CountOfDifference(dB)] " _
                    + "From PropComparison;"

Set StatisticsRS = MyDatabase.OpenRecordset("Statistics")
StatisticsRS.MoveFirst
Do Until StatisticsRS.EOF
    
    StatCount = StatisticsRS("Count")
    
    If StatCount = 0 Then GoTo NextRecord
    
    StatMode = StatisticsRS("Mode")
    StatisticsRS.Edit
    
'calculate/fill RMS field
    StatisticsRS("RMS") = Sqr((((StatCount - 1) / StatCount) * StatisticsRS("StDev") ^ 2) + StatisticsRS("Mean") ^ 2)
    
'calculate and fill prob distribution fields
    For StatIndex = -4 To 4
        StDevValue = StatisticsRS("Mean") + (StatIndex * StatisticsRS("StDev"))
        
        If StatisticsRS("Mode") <> "TIREM" Then
            Set TempRS = MyDatabase.OpenRecordset("SELECT Count([difference(db)]) FROM PropComparison WHERE [Difference(db)] <= " + Str(StDevValue) + " and PropMode = '" + StatMode + "'")
        Else
            Set TempRS = MyDatabase.OpenRecordset("SELECT Count([difference(db)]) FROM PropComparison WHERE [Difference(db)] <= " + Str(StDevValue))
        End If
        
        StatisticsRS(Trim(Str(StatIndex)) + " StDev") = Format(TempRS.Fields(0).Value / StatCount, "0.00")
    Next
    
'fill correlation coefficient
'fill variables using sql functions and
    If TotalPathLoss = False Then  'compare excess losses
    
        If StatisticsRS("Mode") <> "TIREM" Then
        
            Set TempRS = MyDatabase.OpenRecordset("SELECT Sum(PropComparison.[PredictedLoss(dB)]) AS SumY, " _
                        + "Sum(PropData.dbloss) AS SumX, Sum(PropComparison.[PredictedLoss(dB)]*PropComparison.[PredictedLoss(dB)]) AS SumY2, " _
                        + "Sum(PropData.dbloss*PropData.dbloss) AS SumX2, Sum(PropComparison.[PredictedLoss(dB)]*PropData.dbloss) AS SumXY " _
                        + "From PropData, PropComparison Where PropData.ID = PropComparison.TestID And PropComparison.PropMode = '" + StatMode + "'")
        
        Else
            
            Set TempRS = MyDatabase.OpenRecordset("SELECT Sum(PropComparison.[PredictedLoss(dB)]) AS SumY, " _
                        + "Sum(PropData.dbloss) AS SumX, Sum(PropComparison.[PredictedLoss(dB)]*PropComparison.[PredictedLoss(dB)]) AS SumY2, " _
                        + "Sum(PropData.dbloss*PropData.dbloss) AS SumX2, Sum(PropComparison.[PredictedLoss(dB)]*PropData.dbloss) AS SumXY " _
                        + "From PropData, PropComparison Where PropData.ID = PropComparison.TestID")
        
        End If
    
    Else
    
        If StatisticsRS("Mode") <> "TIREM" Then
        
            Set TempRS = MyDatabase.OpenRecordset("SELECT Sum(PropComparison.[TotalPathLoss(dB)]) AS SumY, " _
                        + "Sum(PropData.dbloss + PropComparison.[FreeSpaceLoss(dB)]) AS SumX, " _
                        + "Sum(PropComparison.[TotalPathLoss(dB)]*PropComparison.[TotalPathLoss(dB)]) AS SumY2, " _
                        + "Sum((PropData.dbloss + PropComparison.[FreeSpaceLoss(dB)])*(PropData.dbloss + PropComparison.[FreeSpaceLoss(dB)])) AS SumX2, " _
                        + "Sum(PropComparison.[TotalPathLoss(dB)]*(PropData.dbloss + PropComparison.[FreeSpaceLoss(dB)])) AS SumXY " _
                        + "From PropData, PropComparison Where PropData.ID = PropComparison.TestID And PropComparison.PropMode = '" + StatMode + "'")
        
        Else
            
            Set TempRS = MyDatabase.OpenRecordset("SELECT Sum(PropComparison.[TotalPathLoss(dB)]) AS SumY, " _
                        + "Sum(PropData.dbloss + PropComparison.[FreeSpaceLoss(dB)]) AS SumX, " _
                        + "Sum(PropComparison.[TotalPathLoss(dB)]*PropComparison.[TotalPathLoss(dB)]) AS SumY2, " _
                        + "Sum((PropData.dbloss + PropComparison.[FreeSpaceLoss(dB)])*(PropData.dbloss + PropComparison.[FreeSpaceLoss(dB)])) AS SumX2, " _
                        + "Sum(PropComparison.[TotalPathLoss(dB)]*(PropData.dbloss + PropComparison.[FreeSpaceLoss(dB)])) AS SumXY " _
                        + "From PropData, PropComparison Where PropData.ID = PropComparison.TestID")
        
        End If
    
    End If
        
    SumXY = TempRS("SumXY")
    SumX = TempRS("SumX")
    SumY = TempRS("SumY")
    SumX2 = TempRS("SumX2")
    SumY2 = TempRS("SumY2")

    CorrCoefNumerator = (StatCount * SumXY) - (SumX * SumY)
    CorrCoefDenominator = Sqr((StatCount * SumX2 - SumX ^ 2) * (StatCount * SumY2 - SumY ^ 2))
    CorrCoef = CorrCoefNumerator / CorrCoefDenominator
    StatisticsRS("CorCoeff") = CorrCoef
    
    StatisticsRS.Update
    
NextRecord:
    StatisticsRS.MoveNext
Loop

End Sub
