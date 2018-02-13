VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.0#0"; "RESIZE32.OCX"
Begin VB.Form frmSQLQuery 
   Caption         =   "Query Select Criteria"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8265
   Icon            =   "frmSQLQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlProject 
      Left            =   360
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   1800
      _Version        =   196608
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
      Enabled         =   -1  'True
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7005
      FormDesignWidth =   8265
   End
   Begin VB.ComboBox cboPolarization 
      Height          =   315
      ItemData        =   "frmSQLQuery.frx":030A
      Left            =   1920
      List            =   "frmSQLQuery.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   5760
      Width           =   1935
   End
   Begin VB.ListBox lstLocation 
      Height          =   1860
      ItemData        =   "frmSQLQuery.frx":033F
      Left            =   1920
      List            =   "frmSQLQuery.frx":036A
      Style           =   1  'Checkbox
      TabIndex        =   23
      Top             =   840
      Width           =   3495
   End
   Begin VB.ComboBox cboRXAntennaHt 
      Height          =   315
      ItemData        =   "frmSQLQuery.frx":0452
      Left            =   1920
      List            =   "frmSQLQuery.frx":046E
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtRXAntennaHt1 
      Height          =   285
      Left            =   3840
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtRXAntennaHt2 
      Height          =   285
      Left            =   5640
      TabIndex        =   18
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectRecords 
      Caption         =   "Select Records"
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox txtAntennaHt2 
      Height          =   285
      Left            =   5640
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtAntennaHt1 
      Height          =   285
      Left            =   3840
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cboAntennaHt 
      Height          =   315
      ItemData        =   "frmSQLQuery.frx":0495
      Left            =   1920
      List            =   "frmSQLQuery.frx":04B1
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtDistance2 
      Height          =   285
      Left            =   5640
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtDistance1 
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtFrequency2 
      Height          =   285
      Left            =   5640
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtFrequency1 
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ComboBox cboLinkDistance 
      Height          =   315
      ItemData        =   "frmSQLQuery.frx":04D8
      Left            =   1920
      List            =   "frmSQLQuery.frx":04F4
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ComboBox cboFrequencyOperator 
      Height          =   315
      ItemData        =   "frmSQLQuery.frx":051B
      Left            =   1920
      List            =   "frmSQLQuery.frx":0537
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Polarization"
      Height          =   195
      Left            =   360
      TabIndex        =   24
      Top             =   5760
      Width           =   810
   End
   Begin VB.Label lblRXAntennaHt 
      AutoSize        =   -1  'True
      Caption         =   "RX Antenna Height"
      Height          =   195
      Left            =   360
      TabIndex        =   22
      Top             =   5160
      Width           =   1380
   End
   Begin VB.Label lblRXAntennaHtUnits 
      AutoSize        =   -1  'True
      Caption         =   "meters"
      Height          =   195
      Left            =   7080
      TabIndex        =   21
      Top             =   5160
      Width           =   465
   End
   Begin VB.Label lblAntennaHtUnits 
      AutoSize        =   -1  'True
      Caption         =   "meters"
      Height          =   195
      Left            =   7080
      TabIndex        =   16
      Top             =   4560
      Width           =   465
   End
   Begin VB.Label lblAntennaHt 
      AutoSize        =   -1  'True
      Caption         =   "TX Antenna Height"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   4560
      Width           =   1365
   End
   Begin VB.Label lblDistanceUnits 
      AutoSize        =   -1  'True
      Caption         =   "km"
      Height          =   195
      Left            =   7080
      TabIndex        =   11
      Top             =   3960
      Width           =   210
   End
   Begin VB.Label lblFrequencyUnits 
      AutoSize        =   -1  'True
      Caption         =   "MHz"
      Height          =   195
      Left            =   7080
      TabIndex        =   8
      Top             =   3360
      Width           =   330
   End
   Begin VB.Label lblLinkDistance 
      AutoSize        =   -1  'True
      Caption         =   "Link Distance"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblFrequency 
      AutoSize        =   -1  'True
      Caption         =   "Frequency"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   750
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      Caption         =   "Select data records meeting the following criteria:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   5970
   End
   Begin VB.Label lblLocation 
      AutoSize        =   -1  'True
      Caption         =   "Location"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.Menu mnuFrequencyUnits 
      Caption         =   "FrequencyUnits"
      Visible         =   0   'False
      Begin VB.Menu mnukHz 
         Caption         =   "kHz"
      End
      Begin VB.Menu mnuMHz 
         Caption         =   "MHz"
      End
      Begin VB.Menu mnuGHz 
         Caption         =   "GHz"
      End
   End
   Begin VB.Menu mnuHeightUnits 
      Caption         =   "HeightUnits"
      Visible         =   0   'False
      Begin VB.Menu mnuMeters 
         Caption         =   "meters"
      End
      Begin VB.Menu mnuFeet 
         Caption         =   "feet"
      End
   End
   Begin VB.Menu mnuDistanceUnits 
      Caption         =   "DistanceUnits"
      Visible         =   0   'False
      Begin VB.Menu mnuKM 
         Caption         =   "km"
      End
      Begin VB.Menu mnuStatuteMI 
         Caption         =   "st. mi."
      End
      Begin VB.Menu mnuNauticalMI 
         Caption         =   "nmi"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmSQLQuery"
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

Option Explicit

Dim UnitLabel As Control

Dim FrequencyLow As Single
Dim FrequencyHigh As Single
Dim FrequencyFactor As Single

Dim LinkDistanceLow As Single
Dim LinkDistanceHigh As Single
Dim LinkDistanceFactor As Single

Dim AntennaHtLow As Single
Dim AntennaHtHigh As Single
Dim AntennaHtFactor As Single

Dim RXAntennaHtLow As Single
Dim RXAntennaHtHigh As Single
Dim RXAntennaHtFactor As Single

Private Sub cboAntennaHt_Click()
Select Case cboAntennaHt.ListIndex
Case 0  'all - no select
    txtAntennaHt1.Visible = False
    txtAntennaHt2.Visible = False
    lblAntennaHtUnits.Visible = False
Case 7  'between
    txtAntennaHt1.Visible = True
    txtAntennaHt2.Visible = True
    lblAntennaHtUnits.Visible = True
    lblAntennaHtUnits.Left = 7080
Case Else
    txtAntennaHt1.Visible = True
    txtAntennaHt2.Visible = False
    lblAntennaHtUnits.Visible = True
    lblAntennaHtUnits.Left = 5160
End Select

End Sub

Private Sub cboFrequencyOperator_Click()
Select Case cboFrequencyOperator.ListIndex
Case 0  'all - no select
    txtFrequency1.Visible = False
    txtFrequency2.Visible = False
    lblFrequencyUnits.Visible = False
Case 7  'between
    txtFrequency1.Visible = True
    txtFrequency2.Visible = True
    lblFrequencyUnits.Visible = True
    lblFrequencyUnits.Left = 7080
Case Else
    txtFrequency1.Visible = True
    txtFrequency2.Visible = False
    lblFrequencyUnits.Visible = True
    lblFrequencyUnits.Left = 5160
End Select
End Sub

Private Sub cboLinkDistance_Click()
Select Case cboLinkDistance.ListIndex
Case 0  'all - no select
    txtDistance1.Visible = False
    txtDistance2.Visible = False
    lblDistanceUnits.Visible = False
Case 7  'between
    txtDistance1.Visible = True
    txtDistance2.Visible = True
    lblDistanceUnits.Visible = True
    lblDistanceUnits.Left = 7080
Case Else
    txtDistance1.Visible = True
    txtDistance2.Visible = False
    lblDistanceUnits.Visible = True
    lblDistanceUnits.Left = 5160
End Select

End Sub

Private Sub cboRXAntennaHt_Click()
Select Case cboRXAntennaHt.ListIndex
Case 0  'all - no select
    txtRXAntennaHt1.Visible = False
    txtRXAntennaHt2.Visible = False
    lblRXAntennaHtUnits.Visible = False
Case 7  'between
    txtRXAntennaHt1.Visible = True
    txtRXAntennaHt2.Visible = True
    lblRXAntennaHtUnits.Visible = True
    lblRXAntennaHtUnits.Left = 7080
Case Else
    txtRXAntennaHt1.Visible = True
    txtRXAntennaHt2.Visible = False
    lblRXAntennaHtUnits.Visible = True
    lblRXAntennaHtUnits.Left = 5160
End Select

End Sub

Private Sub cmdSelectRecords_Click()
Dim AdditionalQueryString As String
Dim AreaTableName As String
Dim i As Integer

QueryString = "" 'initialize

'check if location of interest = all, if not perform location query
If lstLocation.SelCount = 0 Then
    MsgBox "You must select an area of interest.", vbExclamation + vbOKOnly, "Warning"
    lstLocation.SetFocus
    Exit Sub
End If

If lstLocation.Selected(0) = False Then
    For i = 0 To lstLocation.ListCount - 1
        If lstLocation.Selected(i) = True Then
            If Len(Trim(QueryString)) > 0 Then
                QueryString = QueryString + " or PropData.Region = '" + lstLocation.List(i) + "'"
            Else
                QueryString = "PropData.Region = '" + lstLocation.List(i) + "'"
            End If
        End If
    Next
End If

'if multiple locations selected the or clause needs to be bracketed
If Len(Trim(QueryString)) > 0 Then
    QueryString = "(" + QueryString + ")"
End If

'frequency query
If cboFrequencyOperator.ListIndex <> 0 Then
    'check/load variables
    Select Case lblFrequencyUnits.Caption
    Case "kHz"
        FrequencyFactor = 0.001
    Case "MHz"
        FrequencyFactor = 1
    Case "GHz"
        FrequencyFactor = 1000
    End Select
    
    If IsNumeric(txtFrequency1.Text) And Val(txtFrequency1.Text) > 0 Then
        FrequencyLow = Val(txtFrequency1.Text) * FrequencyFactor
    Else
        MsgBox "Frequency must be a numeric greater than 0.", vbExclamation + vbOKOnly, "Warning"
        txtFrequency1.SetFocus
        Exit Sub
    End If

    If cboFrequencyOperator.ListIndex = 7 Then 'between, check upper freq
        If IsNumeric(txtFrequency2.Text) And Val(txtFrequency2.Text) > 0 Then
            FrequencyHigh = Val(txtFrequency2.Text) * FrequencyFactor
        Else
            MsgBox "Frequency must be a numeric greater than 0.", vbExclamation + vbOKOnly, "Warning"
            txtFrequency2.SetFocus
            Exit Sub
        End If
        
        AdditionalQueryString = "PropData.freq BETWEEN" + Str(FrequencyLow) + " AND" + Str(FrequencyHigh)
        
    Else
        AdditionalQueryString = "PropData.freq " + cboFrequencyOperator.Text + Str(FrequencyLow)
    End If
    
    If Len(Trim(QueryString)) > 0 Then
        QueryString = QueryString + " and " + AdditionalQueryString
    Else
        QueryString = AdditionalQueryString
    End If
    
End If

'link distance query
If cboLinkDistance.ListIndex <> 0 Then
    Select Case lblDistanceUnits.Caption
    Case "km"
        LinkDistanceFactor = 1
    Case "st. mi."
        LinkDistanceFactor = 1.609344
    Case "nmi"
        FrequencyFactor = 1.852
    End Select
    
    'check/load variables
    If IsNumeric(txtDistance1.Text) And Val(txtDistance1.Text) > 0 Then
        LinkDistanceLow = Val(txtDistance1.Text) * LinkDistanceFactor
    Else
        MsgBox "Link Distance must be a numeric greater than 0.", vbExclamation + vbOKOnly, "Warning"
        txtDistance1.SetFocus
        Exit Sub
    End If

    If cboLinkDistance.ListIndex = 7 Then 'between, check upper freq
        If IsNumeric(txtDistance2.Text) And Val(txtDistance2.Text) > 0 Then
            LinkDistanceHigh = Val(txtDistance2.Text) * LinkDistanceFactor
        Else
            MsgBox "Link Distance must be a numeric greater than 0.", vbExclamation + vbOKOnly, "Warning"
            txtDistance2.SetFocus
            Exit Sub
        End If
        
        AdditionalQueryString = "PropData.dist BETWEEN" + Str(LinkDistanceLow) + " AND" + Str(LinkDistanceHigh)
    Else
        AdditionalQueryString = "PropData.dist " + cboLinkDistance.Text + Str(LinkDistanceLow)
    End If
    
    If Len(Trim(QueryString)) > 0 Then
        QueryString = QueryString + " and " + AdditionalQueryString
    Else
        QueryString = AdditionalQueryString
    End If
    
End If

'TX Antenna Ht query
If cboAntennaHt.ListIndex <> 0 Then
    
    Select Case lblAntennaHtUnits.Caption
    Case "meters"
        AntennaHtFactor = 1
    Case "feet"
        AntennaHtFactor = 0.3048006
    End Select
    
    'check/load variables
    If IsNumeric(txtAntennaHt1.Text) And Val(txtAntennaHt1.Text) > 0 Then
        AntennaHtLow = Val(txtAntennaHt1.Text) * AntennaHtFactor
    Else
        MsgBox "TX Antenna Height must be a numeric greater than 0.", vbExclamation + vbOKOnly, "Warning"
        txtAntennaHt1.SetFocus
        Exit Sub
    End If

    If cboAntennaHt.ListIndex = 7 Then 'between, check upper freq
        If IsNumeric(txtAntennaHt2.Text) And Val(txtAntennaHt2.Text) > 0 Then
            AntennaHtHigh = Val(txtAntennaHt2.Text) * AntennaHtFactor
        Else
            MsgBox "TX Antenna Height must be a numeric greater than 0.", vbExclamation + vbOKOnly, "Warning"
            txtAntennaHt2.SetFocus
            Exit Sub
        End If
        
        AdditionalQueryString = "PropData.xht BETWEEN" + Str(AntennaHtLow) + " AND" + Str(AntennaHtHigh)
    Else
        AdditionalQueryString = "PropData.xht " + cboAntennaHt.Text + Str(AntennaHtLow)
    End If
    
    If Len(Trim(QueryString)) > 0 Then
        QueryString = QueryString + " and " + AdditionalQueryString
    Else
        QueryString = AdditionalQueryString
    End If
    
End If

'RX Antenna Ht query
If cboRXAntennaHt.ListIndex <> 0 Then
    
    Select Case lblRXAntennaHtUnits.Caption
    Case "meters"
        RXAntennaHtFactor = 1
    Case "feet"
        RXAntennaHtFactor = 0.3048006
    End Select
    
    'check/load variables
    If IsNumeric(txtRXAntennaHt1.Text) And Val(txtRXAntennaHt1.Text) > 0 Then
        RXAntennaHtLow = Val(txtRXAntennaHt1.Text) * RXAntennaHtFactor
    Else
        MsgBox "RX Antenna Height must be a numeric greater than 0.", vbExclamation + vbOKOnly, "Warning"
        txtRXAntennaHt1.SetFocus
        Exit Sub
    End If

    If cboRXAntennaHt.ListIndex = 7 Then 'between, check upper freq
        If IsNumeric(txtRXAntennaHt2.Text) And Val(txtRXAntennaHt2.Text) > 0 Then
            RXAntennaHtHigh = Val(txtRXAntennaHt2.Text) * RXAntennaHtFactor
        Else
            MsgBox "RX Antenna Height must be a numeric greater than 0.", vbExclamation + vbOKOnly, "Warning"
            txtRXAntennaHt2.SetFocus
            Exit Sub
        End If
        
        AdditionalQueryString = "PropData.rht BETWEEN" + Str(RXAntennaHtLow) + " AND" + Str(RXAntennaHtHigh)
    Else
        AdditionalQueryString = "PropData.rht " + cboRXAntennaHt.Text + Str(RXAntennaHtLow)
    End If
    
    If Len(Trim(QueryString)) > 0 Then
        QueryString = QueryString + " and " + AdditionalQueryString
    Else
        QueryString = AdditionalQueryString
    End If
    
End If

'polarization
If cboPolarization.ListIndex = 0 Then  'like polarizations
    AdditionalQueryString = "PropData.xpol = PropData.rpol"
Else
    AdditionalQueryString = "PropData.xpol <> PropData.rpol"
End If
    
If Len(Trim(QueryString)) > 0 Then
    QueryString = QueryString + " and " + AdditionalQueryString
Else
    QueryString = AdditionalQueryString
End If

frmSQLQuery.Hide
frmQueryResult.Show

End Sub

Private Sub Form_Load()
    Dim MyDB As Database
    Dim MyRS As Recordset
    
'open database and recordset to load locations into list box
    Set MyDB = OpenDatabase(App.Path + "\tiremtest.mdb")
    Set MyRS = MyDB.OpenRecordset("Area Codes")
    
    lstLocation.Clear
    lstLocation.AddItem "All"
    
    MyRS.MoveFirst
    
'loop thru recordset to obtain all locations
    Do While Not MyRS.EOF
        lstLocation.AddItem MyRS("region")
        MyRS.MoveNext
    Loop
        
'set list box defaults
    lstLocation.Selected(0) = True
    cboFrequencyOperator.Text = cboFrequencyOperator.List(0)
    cboLinkDistance.Text = cboLinkDistance.List(0)
    cboAntennaHt.Text = cboAntennaHt.List(0)
    cboRXAntennaHt.Text = cboRXAntennaHt.List(0)
    cboPolarization = cboPolarization.List(0)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        End_Program
        Cancel = 0
    End If
End Sub

Private Sub lblAntennaHtUnits_Click()
    Set UnitLabel = lblAntennaHtUnits
    PopupMenu mnuHeightUnits
End Sub

Private Sub lblDistanceUnits_Click()
    Set UnitLabel = lblDistanceUnits
    PopupMenu mnuDistanceUnits
End Sub

Private Sub lblFrequencyUnits_Click()
    Set UnitLabel = lblFrequencyUnits
    PopupMenu mnuFrequencyUnits
End Sub

Private Sub lblRXAntennaHtUnits_Click()
    Set UnitLabel = lblRXAntennaHtUnits
    PopupMenu mnuHeightUnits
End Sub

Private Sub lstLocation_Click()
Dim i As Integer

If lstLocation.ListIndex = 0 Then
    For i = 1 To lstLocation.ListCount - 1
        lstLocation.Selected(i) = False
    Next
Else
    lstLocation.Selected(0) = False
End If

End Sub

Private Sub mnuExit_Click()
    End_Program
End Sub

Private Sub mnuFeet_Click()
    UnitLabel.Caption = "feet"
End Sub

Private Sub mnuGHz_Click()
    UnitLabel.Caption = "GHz"
End Sub

Private Sub mnukHz_Click()
    UnitLabel.Caption = "kHz"
End Sub

Private Sub mnuKM_Click()
    UnitLabel.Caption = "km"
End Sub

Private Sub mnuMeters_Click()
    UnitLabel.Caption = "meters"
End Sub

Private Sub mnuMHz_Click()
    UnitLabel.Caption = "MHz"
End Sub

Private Sub mnuNauticalMI_Click()
    UnitLabel.Caption = "nmi"
End Sub

Private Sub mnuStatuteMI_Click()
    UnitLabel.Caption = "st. mi."
End Sub
