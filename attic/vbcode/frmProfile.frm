VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{5A721583-5AF0-11CE-8384-0020AF2337F2}#1.0#0"; "VCFI32.OCX"
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.0#0"; "RESIZE32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmProfile 
   AutoRedraw      =   -1  'True
   Caption         =   "Frequency Occupancy Histogram"
   ClientHeight    =   8445
   ClientLeft      =   1530
   ClientTop       =   1890
   ClientWidth     =   12585
   Icon            =   "frmProfile.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8445
   ScaleWidth      =   12585
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   14
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "New Query"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Export Graph"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Print"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Title"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Footnote"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Legend"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   ""
            ImageIndex      =   33
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Help"
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VCIFiLib.VtChart chtProfilePoints 
      Height          =   5655
      Left            =   240
      OleObjectBlob   =   "frmProfile.frx":030A
      TabIndex        =   1
      Top             =   840
      Width           =   12060
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4440
      _Version        =   196608
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
      Enabled         =   -1  'True
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   8445
      FormDesignWidth =   12585
   End
   Begin MSComDlg.CommonDialog cdlExport 
      Left            =   0
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DialogTitle     =   "Export Chart"
      Filter          =   "Bitmap (*.bmp)|*.bmp |Windows Metafile (*.wmf)|*.wmf"
      Flags           =   2
   End
   Begin VCIF1Lib.F1Book spdStatistics 
      Height          =   1550
      Left            =   240
      OleObjectBlob   =   "frmProfile.frx":30CD
      TabIndex        =   2
      Top             =   6720
      Width           =   12060
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   33
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":377F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":3A99
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":3DB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":40CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":43E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":44F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":460B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":4925
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":4C3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":4F59
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":5273
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":558D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":58A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":5BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":5EDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":61F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":650F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":6829
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":6B43
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":6C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":6D67
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":7081
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":739B
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":76B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":79CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":7CE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":8003
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":831D
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":8637
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":8951
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":8C6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":8D7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProfile.frx":9097
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Query"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Chart"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintData 
         Caption         =   "P&rint"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuChartEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditSelection 
         Caption         =   "Edit Se&lection"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAxis 
         Caption         =   "A&xis"
         Begin VB.Menu mnuAxisGeneral 
            Caption         =   "&General"
         End
         Begin VB.Menu mnuAxisLabel 
            Caption         =   "&Label"
         End
         Begin VB.Menu mnuAxisTitle 
            Caption         =   "&Title"
         End
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "&Series"
         Begin VB.Menu mnuSeriesGeneral 
            Caption         =   "&General"
         End
         Begin VB.Menu mnuSeriesLabel 
            Caption         =   "&Label"
         End
      End
      Begin VB.Menu mnuProfileTitle 
         Caption         =   "&Title"
      End
      Begin VB.Menu mnuFootnote 
         Caption         =   "&Footnote"
      End
      Begin VB.Menu mnuLegend 
         Caption         =   "Le&gend"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy to Clipboard"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste from Clipboard"
      End
   End
   Begin VB.Menu mnuSelectPart 
      Caption         =   "SelectPart"
      Visible         =   0   'False
      Begin VB.Menu mnuSelection 
         Caption         =   "Edit Se&lection"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLabel 
         Caption         =   "La&bel"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuProfileRefresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On..."
      End
   End
End
Attribute VB_Name = "frmProfile"
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
    
    Dim Index As Long

    Dim NumberofBins As Long
    
'To determine selected part of chart
    Dim PartType As Integer
    Dim SeriesID As Integer
    Dim DataPointID As Integer
    Dim Axis As Integer
    Dim Unused As Integer
    
    Dim DataPointSelected As Integer
    
    
Private Sub chtProfilePoints_ChartActivated(MouseFlags As Integer, Cancel As Integer)
    DoEvents
'Dont let user select the chart (modification ruins chart)
    Cancel = True
End Sub

Private Sub chtProfilePoints_ChartSelected(MouseFlags As Integer, Cancel As Integer)
    DoEvents
    frmProfile.mnuEditSelection.Enabled = False

End Sub

Private Sub chtProfilePoints_DblClick()
    
    DoEvents
    chtProfilePoints.ActivateSelectionDialog

End Sub

Private Sub chtProfilePoints_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next  'just in case a portion of the chart was not selected
DoEvents

'Return selected portion of chart
    chtProfilePoints.GetSelectedPart PartType, SeriesID, DataPointID, Axis, Unused
    
'Select chart if nothing selected
    If PartType = 0 Then
        chtProfilePoints.SelectPart VtChPartTypeChart, 1, 1, 1, 1
    End If
        
    If Button = 1 Then  'Left mouse button
        
'If a profile point is selected, display coordinates
        If PartType <> VtChPartTypePoint And SeriesID <> 1 Then
            fraCoordinates.Visible = False
            BarSelected = False
            If Forms.COUNT > 2 Then
                frmDataDisplay.Reload_Data
            End If
        End If
    
        Select Case PartType
        
        Case VtChPartTypeChart
            frmProfile.mnuAxis.Enabled = False
            frmProfile.mnuEditSelection.Enabled = False
            frmProfile.mnuSeries.Enabled = False
            
        Case VtChPartTypeSeries, VtChPartTypeSeriesLabel
              
            frmProfile.mnuAxis.Enabled = False
            frmProfile.mnuSeries.Enabled = True
            frmProfile.mnuEditSelection.Enabled = True
            
        Case VtChPartTypeAxis, VtChPartTypeAxisLabel, VtChPartTypeAxisTitle
            
            frmProfile.mnuAxis.Enabled = True
            frmProfile.mnuSeries.Enabled = False
            frmProfile.mnuEditSelection.Enabled = True
            
        Case Else
            frmProfile.mnuEditSelection.Enabled = True
       
        End Select
    
    Else  'Right mouse button, determine floating menu options
        Select Case PartType
        
        Case VtChPartTypeChart  'Chart menu
        
            frmProfile.mnuLabel.Visible = False
            frmProfile.mnuSelection.Visible = False
            frmProfile.PopupMenu mnuChartEdit
            
        Case VtChPartTypeSeries  'Series menu
            frmProfile.mnuLabel.Visible = True
            frmProfile.mnuSelection.Visible = True
            
            frmProfile.PopupMenu mnuSelectPart
            
        Case VtChPartTypeAxis  'Axis menu
            frmProfile.mnuLabel.Visible = True
            frmProfile.mnuSelection.Visible = True
            
            frmProfile.PopupMenu mnuSelectPart
            
        Case VtChPartTypePoint  'Point selected menu
            frmProfile.mnuLabel.Visible = True
            frmProfile.mnuSelection.Visible = True
            
            frmProfile.PopupMenu mnuSelectPart
            
        Case Else
            frmProfile.mnuLabel.Visible = False
            frmProfile.mnuSelection.Visible = True
            
            frmProfile.PopupMenu mnuSelectPart
            
        End Select
        
    End If

End Sub

Private Sub chtProfilePoints_PlotSelected(MouseFlags As Integer, Cancel As Integer)
    
    DoEvents
'Dont let user select the plot (plot modification ruins chart)
    Cancel = True

'Select Chart instead
    chtProfilePoints.SelectPart VtChPartTypeChart, 1, 1, 1, 1

End Sub

Private Sub chtProfilePoints_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
On Error GoTo errorhandler
    
  DoEvents
    
  DataPointSelected = DataPoint  'store in the event user switches chart type, need to know datapoint to update txtoccupancy
  
  fraCoordinates.Visible = True
'  dbgEnvirTX.Visible = True
  chtProfilePoints.RowLabelIndex = 1
  chtProfilePoints.Row = DataPoint
  lblDistance.Caption = "Frequency (" + UnitsType + ")"
  txtFrequency.Text = Format(Str(Val(chtProfilePoints.RowLabel)), "fixed")
  
  chtProfilePoints.Column = 1
  
  txtOccupancy.Text = Format(Str(chtProfilePoints.Data), "fixed")
    
  BarSelected = True
  frmDataDisplay.Form_Load
  
errorhandler:
    If Err.Number = 6 Then  'overflow error
        Resume
    Else
        Exit Sub
    End If
    
End Sub


Private Sub mnuAxisGeneral_Click()
    chtProfilePoints.SelectPart VtChPartTypeAxis, SeriesID, 1, 1, 1
    chtProfilePoints.ActivateSelectionDialog
End Sub

Private Sub mnuAxisLabel_Click()
    chtProfilePoints.SelectPart VtChPartTypeAxisLabel, SeriesID, 1, 1, 1
    chtProfilePoints.ActivateSelectionDialog

End Sub

Private Sub mnuAxisTitle_Click()
    chtProfilePoints.SelectPart VtChPartTypeAxisTitle, SeriesID, 1, 1, 1
    chtProfilePoints.ActivateSelectionDialog

End Sub


Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuCopy_Click()
    
'copy chart to clipboard
    chtProfilePoints.EditCopy

'Select Chart
    chtProfilePoints.SelectPart VtChPartTypeChart, 1, 1, 1, 1
    
End Sub


Private Sub mnuEditChart_Click()
    
    mnuEditChartData_Click

End Sub

Public Sub mnuEditChartData_Click()

'Open Edit Chart Dialog Box
    chtProfilePoints.EditChartData
    
    ReDim GridArray(1 To chtProfilePoints.RowCount, 1 To 1)
    chtProfilePoints.CopyDataToArray 1, 1, chtProfilePoints.RowCount, 1, GridArray

End Sub

Private Sub Form_Load()
    
On Error GoTo errorhandler

'center form on screen
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
    
'Turn off menu options
frmProfile.mnuAxis.Enabled = False
frmProfile.mnuSeries.Enabled = False
    
Load_Occupancy

Load_Spreadsheet

'turn on footnote, set text, font, location
'    chtProfilePoints.Footnote.Location.Visible = True
'    chtProfilePoints.Footnote.Location.LocationType = VtChLocationTypeBottomLeft
'    chtProfilePoints.Footnote.TextLayout.HorzAlignment = VtHorizontalAlignmentLeft
'    chtProfilePoints.Footnote.VtFont.Size = 9
'    chtProfilePoints.Footnote.Text = "Minimum Frequency  = " + Str(Format(StartFreq, "fixed")) + " " + UnitsType + _
                                    Chr(13) + Chr(10) + "Maximum Frequency  = " + Str(Format(EndFreq, "fixed")) + " " + UnitsType + _
                                    Chr(13) + Chr(10) + "Channel Width  = " + Str(Format(ChannelWidth, "fixed")) + " " + UnitsType

'Select footnote
chtProfilePoints.SelectPart VtChPartTypeChart, 1, 1, 1, 1

'set form caption
frmProfile.Caption = "TIREM Statistics"

frmProfile.Show

errorhandler:
    Select Case Err.Number
    Case 13  'type mismatch when loading datagrid - database field empty
        Resume Next
    Case 3015  'no index
        Resume Next
    Case Is > 0
        MsgBox Err.Description
        Exit Sub
    End Select
    
End Sub

Private Sub mnuEditSelection_Click()
    
    chtProfilePoints.ActivateSelectionDialog

End Sub

Private Sub mnuEnvDatabase_Click()

frmDataDisplay.Form_Load

End Sub

Private Sub mnuExit_Click()
    Unload frmProfile
End Sub

Private Sub mnuExport_Click()
    
'Detect pressing of Cancel button
    On Error GoTo ExportError
    
    cdlExport.filename = ""  'required for NT
    
    cdlExport.ShowSave
    
    cdlExport.Flags = cdlOFNOverwritePrompt
    
    If Right(cdlExport.FileTitle, 3) = "bmp" Then '.bmp file
        chtProfilePoints.WritePictureToFile cdlExport.filename, VtPictureTypeBMP, VtPictureOptionTextAsCurves
    Else  '.wmf file
        chtProfilePoints.WritePictureToFile cdlExport.filename, VtPictureTypeWMF, VtPictureOptionTextAsCurves
    End If
    
    Exit Sub

ExportError:
    Exit Sub
    
End Sub

Private Sub mnuFileNew_Click()
    For X = Forms.COUNT - 1 To 0 Step -1
        If Forms(X).Name <> "frmSQLQuery" Then
            Unload Forms(X)
        End If
    Next
    frmSQLQuery.Show
    
End Sub
Private Sub mnuFootnote_Click()
    
    chtProfilePoints.SelectPart VtChPartTypeFootnote, 1, 1, 1, 1
    chtProfilePoints.ActivateSelectionDialog

End Sub


Private Sub mnuHelpContents_Click()
    Dim return_value
    return_value = WINHELP(hWnd, App.Path + "\spoc.hlp", HELP_CONTENTS, 0&)

End Sub

Private Sub mnuHelpSearch_Click()
    Dim return_value
    return_value = WINHELP(hWnd, App.Path + "\spoc.hlp", HELP_PARTIALKEY, 0&)

End Sub

Private Sub mnuLabel_Click()
    
    Select Case PartType
        
    Case VtChPartTypeSeries  'series selected
    
        chtProfilePoints.SelectPart VtChPartTypeSeriesLabel, SeriesID, 1, 1, 1
        chtProfilePoints.ActivateSelectionDialog
                        
    Case VtChPartTypeAxis  'axis selected
            
        chtProfilePoints.SelectPart VtChPartTypeAxisLabel, SeriesID, 1, 1, 1
        chtProfilePoints.ActivateSelectionDialog
            
    Case VtChPartTypePoint  'point selected
            
        chtProfilePoints.SelectPart VtChPartTypePointLabel, SeriesID, DataPointID, 1, 1
        chtProfilePoints.ActivateSelectionDialog
            
    End Select
            
    
End Sub

Private Sub mnuLegend_Click()
    
    chtProfilePoints.SelectPart VtChPartTypeLegend, 1, 1, 1, 1
    chtProfilePoints.ActivateSelectionDialog

End Sub


Private Sub mnuOccupancyDisplay_Click()
On Error GoTo errorhandler
    
    mnuOccupancyDisplay.Checked = True
    mnuPowerDisplay.Checked = False
    lblElevation.Caption = "Occupancy:"
    
    Load_Occupancy
    chtProfilePoints.Plot.Axis(1, 1).AxisTitle.Text = YAxisTitle

    chtProfilePoints.Row = DataPointSelected
    txtOccupancy.Text = Format(Str(chtProfilePoints.Data), "fixed")

errorhandler:
    Select Case Err.Number
    Case 13  'type mismatch when loading datagrid - database field empty, load datagrid with 0 as place holder
        chtProfilePoints.Data = 0
        Resume Next
    Case Is > 0
        MsgBox Err.Description
    End Select
 End Sub

Private Sub mnuPaste_Click()
On Error GoTo errorhandler
    
    chtProfilePoints.EditPaste

errorhandler:
    Exit Sub
End Sub

Private Sub mnuPowerDisplay_Click()
On Error GoTo errorhandler

    mnuPowerDisplay.Checked = True
    mnuOccupancyDisplay.Checked = False
    lblElevation.Caption = "Max Power(KM):"
    
    If PowerOn = True Then
        chtProfilePoints.Row = 1
        chtProfilePoints.Column = 1
    

        PowerOutRS.MoveFirst
        
        For Index = 1 To PowerOutRS.RecordCount
            chtProfilePoints.Data = PowerOutRS("power")
            PowerOutRS.MoveNext
        Next
    End If

    chtProfilePoints.Plot.Axis(1, 1).AxisTitle.Text = "Max Power (KW)"
    
    chtProfilePoints.Row = DataPointSelected
    txtOccupancy.Text = Format(Str(chtProfilePoints.Data), "fixed")

errorhandler:
    Select Case Err.Number
    Case 13  'type mismatch when loading datagrid - database field empty, load datagrid with 0 as place holder
        chtProfilePoints.Data = 0
        Resume Next
    Case Is > 0
        MsgBox Err.Description
    End Select
    
End Sub

Private Sub mnuPrintChart_Click()
    chtProfilePoints.PrintInformation.LayoutForPrinter = True
'Open Printer Dialog Box
    chtProfilePoints.PrintSetupDialog

End Sub

Private Sub mnuPrintData_Click()
Dim WinHandle As Long
Dim ID As Long

On Error GoTo errorhandler
    chtProfilePoints.WritePictureToFile App.Path + "\temp.wmf", VtPictureTypeWMF, VtPictureOptionTextAsCurves
    
    WinHandle = GetMetaFile(App.Path + "\temp.wmf")
    
    spdStatistics.ObjNewPicture 0.01, 4, 14.99, 31.97, ID, WinHandle, 8, 1, 1
    
    spdStatistics.PrintLandscape = True

    spdStatistics.FilePageSetupDlg  'page setup dialog
    spdStatistics.FilePrint True  'printer setup dialog
   
    spdStatistics.ObjSetSelection ID
    spdStatistics.EditClear F1ClearAll
    
    Kill (App.Path + "\temp.wmf")
errorhandler:
    Select Case Err.Number
    Case 53 'file doesn't exist
        Exit Sub
    Case Is > 0
        MsgBox Err.Description
        Exit Sub
    End Select
    
End Sub

'Private Sub mnuPrintDataFile_Click()
'On Error GoTo FileSaveError
'
'    Dim fName As String
'    Dim fType As Integer
'
'    PrintDataSetup  'call sub
'
'    frmProfilePrint.spdPrint.SaveFileDlg "", fName, fType  'export/save file of type ftype dialog box
'    frmProfilePrint.spdPrint.Write fName, fType  'write to file
'
'    Unload frmProfilePrint
'
'FileSaveError:
'    Unload frmProfilePrint
'    chtProfilePoints.Repaint = True
'    Exit Sub
'
'End Sub


Private Sub mnuProfileRefresh_Click()
    chtProfilePoints.Repaint = True
End Sub


Private Sub mnuProfileTitle_Click()

'Select Title
    chtProfilePoints.SelectPart VtChPartTypeTitle, 1, 1, 1, 1
    chtProfilePoints.ActivateSelectionDialog

End Sub

Private Sub mnuSelection_Click()
    mnuEditSelection_Click
End Sub

Private Sub mnuSeriesGeneral_Click()
    chtProfilePoints.SelectPart VtChPartTypeSeries, SeriesID, 1, 1, 1
    chtProfilePoints.ActivateSelectionDialog

End Sub

Private Sub mnuSeriesLabel_Click()
    chtProfilePoints.SelectPart VtChPartTypeSeriesLabel, SeriesID, 1, 1, 1
    chtProfilePoints.ActivateSelectionDialog

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As Button)
Select Case Button.Index
    Case 1
        mnuFileNew_Click
    Case 2
        mnuExport_Click
    Case 3
        mnuPrintData_Click
    Case 5
        mnuProfileTitle_Click
    Case 6
        mnuFootnote_Click
    Case 7
        mnuLegend_Click
    Case 9
        mnuCopy_Click
    Case 10
        mnuPaste_Click
    Case 12
        mnuProfileRefresh_Click
    Case 14
        mnuHelpContents_Click
    End Select

End Sub

Private Sub VideoSoftElastic1_RealignFrame()
    chtProfilePoints.Repaint = True

End Sub

Private Sub VideoSoftElastic1_ResizeChildren()
    chtProfilePoints.Repaint = True

End Sub


Public Sub PrintDataSetup()
'load data into array
    ReDim GridArray(1 To chtProfilePoints.RowCount, 1 To 1)
    chtProfilePoints.CopyDataToArray 1, 1, chtProfilePoints.RowCount, 1, GridArray

'Load Profile data into Formula One Spreadsheet
    For ColumnIndex = 1 To 2
        
        For RowIndex = 1 To chtProfilePoints.RowCount
            frmProfilePrint.spdPrint.NumberRC(RowIndex, ColumnIndex) = GridArray(RowIndex, ColumnIndex)
        Next
    
    Next

'Add title
    Dim title1 As String
    frmProfilePrint.spdPrint.PrintHeader = FileNameProfile
    
'Label Columns 1 through 4
    frmProfilePrint.spdPrint.ColWidth(1) = 3584
    frmProfilePrint.spdPrint.ColWidth(2) = 4096
    
    frmProfilePrint.spdPrint.SetHdrSelection False, False, True
    frmProfilePrint.spdPrint.SetAlignment F1HAlignCenter, True, F1VAlignCenter, 0
    frmProfilePrint.spdPrint.HdrHeight = 540
    
    If UnitsTypeProfile = "meters" Then
        frmProfilePrint.spdPrint.ColText(1) = "Distance" + Chr(10) + Chr(13) + "(km)"
        frmProfilePrint.spdPrint.ColText(2) = "Elevation" + Chr(10) + Chr(13) + "(meters)"
    Else
        frmProfilePrint.spdPrint.ColText(1) = "Distance" + Chr(10) + Chr(13) + "(miles)"
        frmProfilePrint.spdPrint.ColText(2) = "Elevation" + Chr(10) + Chr(13) + "(feet)"
    End If
    
    frmProfilePrint.spdPrint.SetSelection 1, 1, chtProfilePoints.RowCount, 1
    frmProfilePrint.spdPrint.FormatFixed2

End Sub

Public Sub Load_Occupancy()
    Dim FieldIndex As Integer
    Dim ColumnCounter As Integer
    
On Error GoTo errorhandler

    ColumnCounter = 0
    
    chtProfilePoints.Row = 1
    chtProfilePoints.Column = 2
    
    StatisticsRS.MoveFirst
    
    Do Until StatisticsRS.EOF
        ColumnCounter = ColumnCounter + 1
        
        chtProfilePoints.Row = 1
        chtProfilePoints.Column = 2 * ColumnCounter
        
        For FieldIndex = 6 To StatisticsRS.Fields.COUNT - 1
            chtProfilePoints.Data = StatisticsRS.Fields(FieldIndex).Value
        Next
        
        StatisticsRS.MoveNext
    Loop

errorhandler:
    Select Case Err.Number
    Case 13  'null
        Resume Next
    Case Is > 0
        MsgBox Err.Description
    End Select
    
End Sub

Public Sub Load_Spreadsheet()
    Dim FieldIndex As Integer
    Dim RowCounter As Integer
    
    On Error GoTo errorhandler
    
    RowCounter = 0
    
    StatisticsRS.MoveFirst
    
    Do Until StatisticsRS.EOF
        RowCounter = RowCounter + 1

'1st field is text
        spdStatistics.TextRC(RowCounter, 1) = StatisticsRS.Fields(0).Value
        
        For FieldIndex = 1 To StatisticsRS.Fields.COUNT - 1
            spdStatistics.NumberRC(RowCounter, FieldIndex + 1) = StatisticsRS.Fields(FieldIndex).Value
        Next
        
        StatisticsRS.MoveNext
    Loop
    
errorhandler:
    Select Case Err.Number
    Case 94  'null
        Resume Next
    Case Is > 0
        MsgBox Err.Description
    End Select
    
End Sub
