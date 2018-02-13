VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.0#0"; "RESIZE32.OCX"
Begin VB.Form frmQueryResult 
   Caption         =   "Query Results"
   ClientHeight    =   5370
   ClientLeft      =   2775
   ClientTop       =   4515
   ClientWidth     =   9060
   Icon            =   "frmQueryResult.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5370
   ScaleWidth      =   9060
   Begin VB.CommandButton cmdRunTIREM 
      Caption         =   "&Run TIREM"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "C:\PMRS\TIREMTest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PropResults"
      Top             =   4005
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
   Begin MSDBGrid.DBGrid dbgQueryResults 
      Bindings        =   "frmQueryResult.frx":030A
      Height          =   3525
      Left            =   360
      OleObjectBlob   =   "frmQueryResult.frx":031A
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.Label lblNumRecords 
      AutoSize        =   -1  'True
      Caption         =   "Forward"
      Height          =   195
      Left            =   6345
      TabIndex        =   2
      Top             =   4080
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Reverse"
      Height          =   195
      Left            =   1700
      TabIndex        =   1
      Top             =   4080
      Width           =   600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewQuery 
         Caption         =   "&New Query"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Query Results"
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
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
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
Attribute VB_Name = "frmQueryResult"
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

Private Sub cmdRunTIREM_Click()
    If QueryResultsRS.RecordCount = 0 Then
        MsgBox "No records selected.", vbExclamation + vbOKOnly, "Warning"
        Exit Sub
    End If
    
    Data1.Refresh
    frmTIREMAnalysis.Show
    Me.Hide
End Sub

Private Sub Data1_Reposition()
    
Data1.Caption = "Record " + Str(Data1.Recordset.AbsolutePosition + 1) + " of " + Str(QueryResultsRS.RecordCount)

End Sub


Private Sub dbgQueryResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If dbgQueryResults.SelStartCol > -1 Then
    If Button = 2 Then  'right click
        PopupMenu mnuSortOrder
    End If
End If

End Sub

Private Sub Form_Activate()

Dim ForCounter As Integer

For FieldIndex = 0 To dbgQueryResults.Columns.COUNT - 1
    dbgQueryResults.Columns.Remove (0)  'remove first
Next

ForCounter = 0

For FieldIndex = 0 To QueryResultsRS.Fields.COUNT - 1
    If Not (QueryResultsRS.Fields(FieldIndex).Name = "Profile" Or _
       QueryResultsRS.Fields(FieldIndex).Name = "NumPoints") Then
        
        Add_Field_to_Grid (FieldIndex - ForCounter)
    Else
        ForCounter = ForCounter + 1
    End If
Next
    
dbgQueryResults.ReBind
dbgQueryResults.Refresh

End Sub

Public Sub Form_Load()

On Error GoTo errorhandler

Set MyDatabase = OpenDatabase(App.Path + "\Tiremtest.mdb")

'center form on screen
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
    
Data1.DatabaseName = App.Path + "\Tiremtest.mdb"

SortField = "area"

Set_Recordsets  'call sub

Data1.Refresh

errorhandler:
    Select Case Err.Number
    Case 3021  'no records selected - movefirst and movelast generate errors in this condition
        MsgBox "No records met your query criteria.", vbInformation + vbOKOnly
        Resume Next
    Case 3061, 3075 'query failed - set variable and return to query builder
        SQLError = True
    Case Is > 0
        MsgBox "Application Error." + Chr(10) + Chr(13) + Err.Description, vbExclamation + vbOKOnly, "Application Error"
        Exit Sub
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmSQLQuery.Show
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

Private Sub mnuExit_Click()
    End_Program
End Sub

Private Sub mnuHelpSearch_Click()
'Dim return_value
'return_value = WINHELP(hWnd, App.Path + "\gate.hlp", HELP_PARTIALKEY, 0&)
End Sub

Private Sub mnuNewQuery_Click()
    Unload Me
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
If Len(Trim(QueryString)) > 0 Then
    Set QueryResultsRS = MyDatabase.OpenRecordset("SELECT * FROM PropData WHERE " _
                            + QueryString _
                            + " ORDER BY " + SortField)
    
    Data1.RecordSource = "SELECT * FROM PropData WHERE " + QueryString + " ORDER BY " + SortField
    
Else
    Set QueryResultsRS = MyDatabase.OpenRecordset("SELECT * FROM PropData ORDER BY " + SortField)
    Data1.RecordSource = "SELECT * FROM PropData ORDER BY " + SortField
End If


QueryResultsRS.MoveLast
QueryResultsRS.MoveFirst

End Sub

Public Sub Sort_Output(Ascending As Boolean)

Dim SortIndex As Integer

If dbgQueryResults.SelStartCol > 13 Then  'mod since 2 fields from recordset are not in datagrid
    SortIndex = dbgQueryResults.SelStartCol + 2
Else
    SortIndex = dbgQueryResults.SelStartCol
End If

If Ascending = True Then
    SortField = "[" + QueryResultsRS.Fields(SortIndex).Name + "]"
Else
    SortField = "[" + QueryResultsRS.Fields(SortIndex).Name + "] DESC"
End If

Set_Recordsets  'call sub

Data1.Refresh
dbgQueryResults.ReBind
dbgQueryResults.Refresh

End Sub


Public Sub Add_Field_to_Grid(FieldID As Integer)
    dbgQueryResults.Columns.Add (FieldID)
    dbgQueryResults.Columns(FieldID).Caption = QueryResultsRS.Fields(FieldIndex).Name
    dbgQueryResults.Columns(FieldID).DataField = QueryResultsRS.Fields(FieldIndex).Name
    dbgQueryResults.Columns(FieldID).Visible = True
End Sub

Public Sub Set_Export_Strings()
    Export "*", "PropData", QueryString
End Sub
