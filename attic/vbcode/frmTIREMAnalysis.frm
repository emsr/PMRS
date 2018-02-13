VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.0#0"; "RESIZE32.OCX"
Object = "{F8E39C0D-D176-101B-AC32-04022400DC29}#2.2#0"; "SYLVMA32.OCX"
Begin VB.Form frmTIREMAnalysis 
   Caption         =   "TIREM Inputs"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   Icon            =   "frmTIREMAnalysis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin SylvmapLib.SylvMap SylvMap1 
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
      _Version        =   131074
      _ExtentX        =   1931
      _ExtentY        =   1085
      _StockProps     =   105
      MapStackArraySize=   1
      MSNumberOfSymbols000=   1
      BeginProperty SYLabelFont000000 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SYSymbolMetafile000000=   "frmTIREMAnalysis.frx":030A
      FirstPicture    =   "frmTIREMAnalysis.frx":0326
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3000
      _Version        =   196608
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
      Enabled         =   -1  'True
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7605
      FormDesignWidth =   6675
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   22
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Frame fraClimatic 
      Caption         =   "Enter Climatic Information"
      Height          =   3135
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   5655
      Begin VB.ComboBox cboGroundTypes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmTIREMAnalysis.frx":2B76
         Left            =   3480
         List            =   "frmTIREMAnalysis.frx":2B95
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtSeaLevelRefractivity 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmTIREMAnalysis.frx":2C7D
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtConductivity 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmTIREMAnalysis.frx":2C83
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtPermittivity 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmTIREMAnalysis.frx":2C88
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtHumidity 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmTIREMAnalysis.frx":2C8B
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblHumidity 
         AutoSize        =   -1  'True
         Caption         =   "Humidity"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   510
         Width           =   600
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CCIR Ground Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   19
         Top             =   1520
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sea-Level Atmospheric Refractivity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   18
         Top             =   1035
         Width           =   2565
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N-Units"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4680
         TabIndex        =   17
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label lblConductivity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conductivity Of Earth Surface"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   16
         Top             =   2000
         Width           =   2160
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4680
         TabIndex        =   15
         Top             =   1995
         Width           =   270
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relative Permittivity Of Earth Surface"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   14
         Top             =   2480
         Width           =   2655
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "g/m3"
         Height          =   195
         Left            =   4680
         TabIndex        =   13
         Top             =   555
         Width           =   375
      End
   End
   Begin VB.Frame fraTopo 
      Caption         =   "Topographic Extraction Inputs"
      Height          =   2775
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin VB.ComboBox cboInterpolation 
         Height          =   315
         ItemData        =   "frmTIREMAnalysis.frx":2C90
         Left            =   3600
         List            =   "frmTIREMAnalysis.frx":2C9D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1395
      End
      Begin VB.ComboBox cboSpacing 
         Height          =   315
         ItemData        =   "frmTIREMAnalysis.frx":2CBF
         Left            =   3600
         List            =   "frmTIREMAnalysis.frx":2CD5
         TabIndex        =   2
         Text            =   "15"
         Top             =   1200
         Width           =   1410
      End
      Begin VB.ComboBox cboDatumCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmTIREMAnalysis.frx":2CEE
         Left            =   2400
         List            =   "frmTIREMAnalysis.frx":2D19
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label lbl3DInterpolation 
         Caption         =   "Interpolation Method"
         Height          =   450
         Left            =   705
         TabIndex        =   6
         Top             =   570
         Width           =   2520
      End
      Begin VB.Label lblSpacing 
         Caption         =   "Profile Spacing (in seconds).  Blank defaults to spacing of topo file."
         Height          =   615
         Left            =   705
         TabIndex        =   5
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Datum Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   720
         TabIndex        =   4
         Top             =   2040
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmTIREMAnalysis"
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

Dim ResultRecordset As Recordset

'input variable to TIREM
Dim TXLat As Single
Dim TXLatRad As Single
Dim TXLon As Single
Dim TXLonRad As Single
Dim Txantht As Single

Dim RXLat As Single
Dim RXLatRad As Single
Dim RXLon As Single
Dim RXLonRad As Single
Dim Rxantht As Single

Dim POLARZ As String * 4

Dim Spacing As Integer
Dim Datum As Long

Dim groundtype As Long
Dim SeaRefract As Single
Dim CONDUC As Single
Dim PERMIT As Single
Dim REFRAC As Single
Dim INTERP As Long    ' interpolation method
Dim HUMID As Single

'WOTLRET
Dim SPACNG As Single  ' profile spacing in meters
Dim spheroid As Long
Dim ecc As Single
Dim ERRCODE As Long
Dim MJAxis As Single
Dim Flat As Single
Dim PRFERR As Long

'Const MJAxis As Single = 6378137# ' SEMI-MAJOR AXIS OF THE EARTH
'Const Flat As Single = 0.00335281 ' FLATTENING OF THE EARTH
Dim MNAxis As Single
Const TOPFIL As String = "            "
Const ERROPT As String = "E   " ' RETURN WITH ERROR CONDITION SET
Const DELELV As Single = -1# ' DEFAULT ELEVATION

' TOPGET Constants:
  
Const TNAME As String * 12 = "            "
Const NAMLST As String * 12 = TNAME
Const WOTLType As String * 4 = "R   " ' Real Radians.
Const DUMMY1 As Long = 0
'   Const COUNT As Long = 0
   
' TOPGET Arguments:
Dim WOTRER As Long ' TOPGET error return: 0 => OK   Non-Zero => Error.
Dim NSSP As Long ' LATITUDE SPACING (SECONDS).
Dim NegLong As Single

'tirem output
Dim ALPHAE As Single 'effective angle in radians
Dim BETAE As Single  'effective angle in radians
Dim HORZTX As Long   'profile point for tx horizon
Dim HORZRX As Long   'profile point for rx horizon
Dim TXANG As Single  'tx take off angle in radians
Dim RXANG As Single  'rx take off angle in radians
Dim THET00 As Single 'scattering angle in radians
Dim TOTDIF As Single 'total diffraction loss in db
Dim TOTTRO As Single 'total troposcatter loss in db
Dim ABLOSS As Single 'absolute loss in db
Dim Mode As String * 4 ' MODE INDICATOR: LINE OF SIGHT, DIFFRACTION, or TROPO SCATTER from Tirem Dll
Dim PRLoss As Single     ' TOTAL PATH LOSS (BASIC TRANSMISSION LOSS) IN DB from Tirem Dll
Dim FSPLSS As Single     ' FREE SPACE LOSS IN DB from Tirem DLL

'path length/bearing
Dim BearIE As Single
Dim BearEI As Single
Dim PTHLEN As Single
Dim BearIE_deg As Single
Dim BearEI_deg As Single

Dim ErrorCounter As Long
Dim ErrorFound As String
Dim InterferenceCounter As Long
Dim AnalysisCounter As Long
Dim InteractionCounter As Long
Dim PercentComplete As Integer

Dim Difference As Single
Dim PredictedLoss As Single
Dim MeasuredLoss As Single
Dim TestID As Long

Private Function CheckValidityOfInputs()

'topo inputs loaded to variables
    If IsNumeric(cboSpacing.Text) And Val(cboSpacing.Text) > 0 Then
        Spacing = Val(cboSpacing.Text)
    Else
        MsgBox "Spacing must be a number greater than 0.", vbExclamation, "Warning"
        cboSpacing.SetFocus
        CheckValidityOfInputs = 1
        Exit Function
    End If
   
    SPACNG = Spacing * 30  'convert to meters
    INTERP = cboInterpolation.ItemData(cboInterpolation.ListIndex)
    Datum = cboDatumCode.ListIndex

'Climatic inputs loaded to variables
    If IsNumeric(txtHumidity.Text) And Val(txtHumidity.Text) > 0 Then
        HUMID = Val(txtHumidity.Text)
    Else
        MsgBox "Humidity must be a number greater than 0.", vbExclamation, "Warning"
        txtHumidity.SetFocus
        CheckValidityOfInputs = 1
        Exit Function
    End If
    
    If IsNumeric(txtSeaLevelRefractivity.Text) And Val(txtSeaLevelRefractivity.Text) > 0 Then
        SeaRefract = Val(txtSeaLevelRefractivity.Text)
    Else
        MsgBox "Sea-Level Refractivity must be a number greater than 0.", vbExclamation, "Warning"
        txtSeaLevelRefractivity.SetFocus
        CheckValidityOfInputs = 1
        Exit Function
    End If
    
    If IsNumeric(txtConductivity.Text) And Val(txtConductivity.Text) > 0 Then
        CONDUC = txtConductivity.Text
    Else
        MsgBox "Conductivity must be a number greater than 0.", vbExclamation, "Warning"
        txtConductivity.SetFocus
        CheckValidityOfInputs = 1
        Exit Function
    End If
    
    If IsNumeric(txtPermittivity.Text) And Val(txtPermittivity.Text) > 0 Then
        PERMIT = txtPermittivity.Text
    Else
        MsgBox "Pemittivity must be a number greater than 0.", vbExclamation, "Warning"
        txtPermittivity.SetFocus
        CheckValidityOfInputs = 1
        Exit Function
    End If

End Function


Private Sub cboGroundTypes_Click()
groundtype = Val(cboGroundTypes.ListIndex)

'temporarily set prop freq = to receiver freq
PROPFQ = QueryResultsRS("freq")

Call CalcGrConst(PROPFQ, groundtype, PERMIT, CONDUC)

txtConductivity.Text = CONDUC
txtPermittivity.Text = PERMIT

If cboGroundTypes.ListIndex = 0 Then 'groundtype = none
    txtConductivity.Enabled = True
    txtPermittivity.Enabled = True
Else
    txtConductivity.Enabled = False
    txtPermittivity.Enabled = False
End If


End Sub

Private Sub cmdCancel_Click()
    frmQueryResult.Show
    Unload Me
End Sub

Private Sub cmdRun_Click()

On Error GoTo errorhandler

If CheckValidityOfInputs = 1 Then
    Exit Sub
End If

Screen.MousePointer = 11

Me.Hide

MyDatabase.Execute "DELETE * FROM PropComparison"
Set ResultRecordset = MyDatabase.OpenRecordset("PropComparison")

frmStatusBar.Show
DoEvents

'initialize
AnalysisCounter = 0 'for status bar
InterferenceCounter = 0 'number of records exceeding thresh
ErrorCounter = 0 'number of records not processed due to lack of data

InteractionCounter = QueryResultsRS.RecordCount

QueryResultsRS.MoveFirst

'set major and minor axis for WOTL and TIREM call, wont vary per tx site so no reason to repeat inside tx loop
Call Datum2Axis(Datum, spheroid, MJAxis, MNAxis, Flat, ecc, ERRCODE)

Do Until QueryResultsRS.EOF

    AnalysisCounter = AnalysisCounter + 1
    
    InitializeOutputVariables

    If IsNull(QueryResultsRS("NumPoints")) Then
        
        'Load queryresult recordset to variables
        If LoadDB_to_Variables = 1 Then 'call sub
            GoTo UpdateResultDB
        End If
                    
    'pathlength in meters
        CalculatePathLength  'call sub
        
    'check frequency limits
        If PROPFQ < 1 Or PROPFQ > 20000 Then
            ErrorFound = "Frequency Out of Range"
            ErrorCounter = ErrorCounter + 1
            GoTo UpdateResultDB
        End If
        
        Calculate_Propagation_Loss  'call sub
        
    'check for errors calculating prop loss
        If PRFERR > 0 Then
            ErrorFound = "Error Extracting Elevation Data"
            ErrorCounter = ErrorCounter + 1
            GoTo UpdateResultDB
        Else
            If PRLoss = 0 Then
                ErrorFound = "Error Calculating Path Loss"
                ErrorCounter = ErrorCounter + 1
                GoTo UpdateResultDB
            End If
        End If
    
    Else
        
        NUMELV = QueryResultsRS("NumPoints")
        
        'Load queryresult recordset to variables
        If LoadProfile_to_Variables = 1 Then 'call sub
            GoTo UpdateResultDB
        End If
    
        Calculate_Profile_Propagation_Loss  'call sub
        
    'check for errors calculating prop loss
        If PRLoss = 0 Then
            ErrorFound = "Error Calculating Path Loss"
            ErrorCounter = ErrorCounter + 1
            GoTo UpdateResultDB
        End If
    
    End If
    
    PredictedLoss = PRLoss - FSPLSS
    Difference = PredictedLoss - MeasuredLoss
    
UpdateResultDB:
    Write_to_ResultDatabase
        
    QueryResultsRS.MoveNext  'go to next record

    Update_StatusBar_Percent
    
Loop

Unload frmStatusBar

Screen.MousePointer = 0

Analysis_Summary
    
frmPropResult.Show

Unload Me

errorhandler:
    Select Case Err.Number
    Case 94  'invalid use for Null(loading a null field to a variable)
        ErrorFound = "Missing required data"
        Err.Clear  'reset error flag
        ErrorCounter = ErrorCounter + 1
        Write_to_ResultDatabase
        Resume Next
        
    Case 317
        Resume Next
    Case 11
        Resume
    Case Is > 0
        MsgBox Err.Description, Err.Number
        Unload frmStatusBar
        Screen.MousePointer = 0
        Me.Show
        Exit Sub
'        Resume
    End Select
    
End Sub
Public Sub Calculate_Propagation_Loss()
'initialize error constants
    PRFERR = 0
        
'mod groundconst value for frequency
    If groundtype <> 0 Then
        Call CalcGrConst(PROPFQ, groundtype, PERMIT, CONDUC)
    End If

'redim the elevation and distance arrays to the maximum allowed
    ReDim HPRFL(MXNELV)
    ReDim XPRFL(MXNELV)

'extract the path profile            Call GetPathProfile
    Call GetProfile(TXLatRad, TXLonRad, RXLatRad, RXLonRad, SPACNG, MJAxis, Flat, _
                Datum, TOPFIL, INTERP, ERROPT, DELELV, _
                MXNELV, XPRFL(1), HPRFL(1), NUMELV, PRFERR)
            
    If PRFERR > 0 Then
        Exit Sub
    End If

'make the arrays so they only contain the actual number of returned points to avoid emptys
    ReDim Preserve HPRFL(NUMELV)
    ReDim Preserve XPRFL(NUMELV)

    Call CalculatePropagationLoss(REFRAC, PERMIT, CONDUC, Mode, PRLoss, FSPLSS, _
                  ALPHAE, BETAE, HORZTX, HORZRX, _
                  TXANG, RXANG, THET00, TOTDIF, TOTTRO, ABLOSS)
    

End Sub


Private Sub Form_Load()

cboGroundTypes.ListIndex = 4
cboDatumCode.ListIndex = 0
cboInterpolation.ListIndex = 2
cboSpacing.Text = "15"

End Sub
Public Sub CalculatePathLength()
 
'calculate the path length and the bearings
Call CalcNGSInv(TXLatRad, TXLonRad, RXLatRad, RXLonRad, Datum, PTHLEN, BearEI, BearIE)

TXLonRad = -TXLonRad
RXLonRad = -RXLonRad

BearEI_deg = BearEI * 57.29578
BearIE_deg = BearIE * 57.29578
    
End Sub

Public Sub Update_StatusBar_Percent()
    
    PercentComplete = (AnalysisCounter / InteractionCounter) * 100
    If PercentComplete / 5 = Int(PercentComplete / 5) Then
        frmStatusBar.ProgressBar1.Value = PercentComplete
        frmStatusBar.lblStatus.Caption = Str(PercentComplete) + " Percent Complete"
        DoEvents
    End If

End Sub

Public Sub Analysis_Summary()
    
    MsgBox "Number Of Interactions Processed:  " + Str(InteractionCounter) + Chr(13) + Chr(10) _
           + Chr(13) + Chr(10) + "Number of Interactions Not Analyzed Due to Missing Data:  " _
           + Str(ErrorCounter) + Chr(13) + Chr(10), vbExclamation + vbOKOnly, "Analysis Results"
           
End Sub

Public Function LoadDB_to_Variables()
On Error GoTo errorhandler
    
    LoadDB_to_Variables = 0
    
    TestID = QueryResultsRS("ID")
    TXLat = QueryResultsRS("xlat")
    TXLon = -QueryResultsRS("xlon")
    
    TXLatRad = TXLat / 57.29578
    TXLonRad = TXLon / 57.29578
    
    RXLat = QueryResultsRS("rlat")
    RXLon = -QueryResultsRS("rlon")
    
    RXLatRad = RXLat / 57.29578
    RXLonRad = RXLon / 57.29578
    
    PROPFQ = QueryResultsRS("freq")
    
    If IsNull(QueryResultsRS("xpol")) Then
        ErrorFound = "Missing Polarization"
        POLARZ = "V   "
    Else
        POLARZ = QueryResultsRS("xpol")
    End If
    
    Txantht = QueryResultsRS("xht")
    Rxantht = QueryResultsRS("rht")
    
    MeasuredLoss = QueryResultsRS("dbloss")

errorhandler:
    Select Case Err.Number
    Case 94  'invalid use of null
        LoadDB_to_Variables = 1
        ErrorFound = "Missing required data"
        ErrorCounter = ErrorCounter + 1
        Err.Clear
        Exit Function
    Case Is > 0
        LoadDB_to_Variables = 1
        ErrorFound = Err.Description
        ErrorCounter = ErrorCounter + 1
        Exit Function
    End Select

End Function
Public Sub CalculatePropagationLoss(REFRAC As Single, PERMIT As Single, CONDUC As Single, Mode As String, PRLoss As Single, FSPLSS As Single, ALPHAE As Single, BETAE As Single, HORZTX As Long, HORZRX As Long, _
                              TXANG As Single, RXANG As Single, THET00 As Single, TOTDIF As Single, TOTTRO As Single, ABLOSS As Single)

On Error GoTo errorhandler

EXTNSN = 0 ' False
        
'set the last point in the distance array to the GCD to avoid errors in path profile
XPRFL(NUMELV) = PTHLEN

Call NSubS2(HPRFL(1), HPRFL(NUMELV), Txantht, Rxantht, _
             SeaRefract, REFRAC)

Call TiremAnalysis(Txantht, Rxantht, PROPFQ, NUMELV, HPRFL(1), XPRFL(1), _
            EXTNSN, REFRAC, CONDUC, PERMIT, HUMID, POLARZ, _
            VRSION, Mode, PRLoss, FSPLSS, TOTTRO, TOTDIF, ABLOSS, _
            THET00, TXANG, RXANG, ALPHAE, BETAE, HORZTX, HORZRX)
      
errorhandler:
    If Err.Number > 0 Then
        Resume Next
    End If
    
End Sub


Public Sub Write_to_ResultDatabase()
    InterferenceCounter = InterferenceCounter + 1
    
    ResultRecordset.AddNew
'write uniq entries based on analysis type
    ResultRecordset("TestID") = TestID
    
    If PTHLEN = 99999 Then
        ResultRecordset("PathLength(km)") = Null
    Else
        ResultRecordset("PathLength(km)") = PTHLEN / 1000
    End If

    If PRLoss = 999 Or PRLoss = 0 Then
        ResultRecordset("totalpathloss(db)") = Null
    Else
        ResultRecordset("totalpathloss(db)") = PRLoss
    End If
    
    If FSPLSS = 999 Or FSPLSS = 0 Then
        ResultRecordset("FreeSpaceLoss(db)") = Null
    Else
        ResultRecordset("FreeSpaceLoss(db)") = FSPLSS
    End If
    
    If PredictedLoss = 999 Or (PredictedLoss = 0 And PRLoss = 999) Then
        ResultRecordset("PredictedLoss(db)") = Null
    Else
        ResultRecordset("PredictedLoss(db)") = PredictedLoss
    End If
    
    If Mode = "" Then
        ResultRecordset("PropMode") = Null
    Else
        ResultRecordset("PropMode") = Mode
    End If
    
    If Difference = 99999 Then
        ResultRecordset("Difference(dB)") = Null
    Else
        ResultRecordset("Difference(dB)") = Difference
    End If
    
    ResultRecordset("Error") = ErrorFound
    
    ResultRecordset.Update
    
    Err.Clear  'clear any pre-existing errors
    
End Sub

Public Sub InitializeOutputVariables()
    ErrorFound = "None"
    PTHLEN = 99999
    Difference = 99999
    PredictedLoss = 999
    Mode = ""
    PRLoss = 999    ' TOTAL PATH LOSS (BASIC TRANSMISSION LOSS) IN DB from Tirem Dll
    FSPLSS = 999    ' FREE SPACE LOSS IN DB from Tirem DLL
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    cmdCancel_Click
    Cancel = 0
End If
End Sub

Public Function LoadProfile_to_Variables()
On Error GoTo errorhandler
    
    LoadProfile_to_Variables = 0
    
    TestID = QueryResultsRS("ID")
    
    PROPFQ = QueryResultsRS("freq")
    
    If IsNull(QueryResultsRS("xpol")) Then
        ErrorFound = "Missing Polarization"
        POLARZ = "V   "
    Else
        POLARZ = QueryResultsRS("xpol")
    End If
    
    Txantht = QueryResultsRS("xht")
    Rxantht = QueryResultsRS("rht")
    
    MeasuredLoss = QueryResultsRS("dbloss")

    PTHLEN = QueryResultsRS("dist") * 1000 'convert to meters from km

errorhandler:
    Select Case Err.Number
    Case 94  'invalid use of null
        LoadProfile_to_Variables = 1
        ErrorFound = "Missing required data"
        ErrorCounter = ErrorCounter + 1
        Err.Clear
        Exit Function
    Case Is > 0
        LoadProfile_to_Variables = 1
        ErrorFound = Err.Description
        ErrorCounter = ErrorCounter + 1
        Exit Function
    End Select

End Function

Public Sub Calculate_Profile_Propagation_Loss()
'mod groundconst value for frequency
    If groundtype <> 0 Then
        Call CalcGrConst(PROPFQ, groundtype, PERMIT, CONDUC)
    End If

'redim the elevation and distance arrays to the maximum allowed
    ReDim HPRFL(NUMELV)
    ReDim XPRFL(NUMELV)

'load profile from QueryResultsRS
'extract profile from BLOB
    Dim MyBlob As String  'BLOBS AS STRING
    Dim MyVertexArray As Variant

    MyBlob = QueryResultsRS("Profile")
    
'Convert to array
    MyVertexArray = SylvMap1.BlobToVertexArray(MyBlob, NUMELV)
    
    Dim PointCounter As Long
    
    For PointCounter = 1 To NUMELV
        XPRFL(PointCounter) = MyVertexArray(PointCounter - 1, 0)
        HPRFL(PointCounter) = MyVertexArray(PointCounter - 1, 1)
'        Debug.Print XPRFL(PointCounter), HPRFL(PointCounter)
    Next
    
    Call CalculatePropagationLoss(REFRAC, PERMIT, CONDUC, Mode, PRLoss, FSPLSS, _
                  ALPHAE, BETAE, HORZTX, HORZRX, _
                  TXANG, RXANG, THET00, TOTDIF, TOTTRO, ABLOSS)
    
End Sub
