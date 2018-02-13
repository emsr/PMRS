VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5A721583-5AF0-11CE-8384-0020AF2337F2}#1.0#0"; "VCFI32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "VSOCX32.OCX"
Begin VB.Form frmAnalysisWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analysis Wizard"
   ClientHeight    =   5265
   ClientLeft      =   6540
   ClientTop       =   120
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5265
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin VsOcxLib.VideoSoftIndexTab tabAnalysisWizard 
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   75
      Width           =   6735
      _Version        =   327680
      _ExtentX        =   11880
      _ExtentY        =   7646
      _StockProps     =   102
      Caption         =   "1|2|3|4|5|6|7|8|9|10"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      CurrTab         =   2
      AutoScroll      =   -1  'True
      ShowFocusRect   =   0   'False
      TabsPerPage     =   10
      DogEars         =   0   'False
      TabHeight       =   100
      New3D           =   -1  'True
      MouseIcon       =   "frmAnalysisWizard.frx":0000
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic3 
         Height          =   4140
         Left            =   -1.49955e5
         TabIndex        =   4
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         BevelOuter      =   0
         Picture         =   "frmAnalysisWizard.frx":001C
         MouseIcon       =   "frmAnalysisWizard.frx":0038
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic10 
            Height          =   3855
            Left            =   240
            TabIndex        =   92
            Top             =   120
            Width           =   6165
            _Version        =   327680
            _ExtentX        =   10874
            _ExtentY        =   6800
            _StockProps     =   70
            Caption         =   "Enter bounds for geographic area of interest"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ConvInfo        =   1418783674
            BevelOuter      =   6
            Style           =   1
            QuickPaint      =   -1  'True
            Picture         =   "frmAnalysisWizard.frx":0054
            MouseIcon       =   "frmAnalysisWizard.frx":0070
            Begin VB.TextBox txtWestLong 
               Height          =   360
               Left            =   480
               TabIndex        =   96
               Top             =   1680
               Width           =   1650
            End
            Begin VB.TextBox txtEastLong 
               Height          =   360
               Left            =   4005
               TabIndex        =   95
               Top             =   1680
               Width           =   1650
            End
            Begin VB.TextBox txtSouthLat 
               Height          =   360
               Left            =   2280
               TabIndex        =   94
               Top             =   2655
               Width           =   1575
            End
            Begin VB.TextBox txtNorthLat 
               Height          =   360
               Left            =   2280
               TabIndex        =   93
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label lbl3DNELongitude 
               AutoSize        =   -1  'True
               Caption         =   "Longitude (ddd mm ss H)"
               Height          =   195
               Left            =   4005
               TabIndex        =   100
               Top             =   2160
               Width           =   1755
            End
            Begin VB.Label lbl3DNELatitude 
               AutoSize        =   -1  'True
               Caption         =   "Latitude (dd mm ss H)"
               Height          =   195
               Left            =   2310
               TabIndex        =   99
               Top             =   1200
               Width           =   1530
            End
            Begin VB.Label lbl3DSWLongitude 
               AutoSize        =   -1  'True
               Caption         =   "Longitude (ddd mm ss H)"
               Height          =   195
               Left            =   480
               TabIndex        =   98
               Top             =   2160
               Width           =   1755
            End
            Begin VB.Label lbl3DSWLatitude 
               AutoSize        =   -1  'True
               Caption         =   "Latitude (dd mm ss H)"
               Height          =   195
               Left            =   2310
               TabIndex        =   97
               Top             =   3120
               Width           =   1530
            End
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic5 
         Height          =   4140
         Left            =   -1.49805e5
         TabIndex        =   5
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         Picture         =   "frmAnalysisWizard.frx":008C
         MouseIcon       =   "frmAnalysisWizard.frx":00A8
         Begin VB.Frame fraTopo 
            Caption         =   "Topographic Extraction Inputs"
            Height          =   2775
            Left            =   480
            TabIndex        =   34
            Top             =   480
            Width           =   5655
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
               ItemData        =   "frmAnalysisWizard.frx":00C4
               Left            =   2400
               List            =   "frmAnalysisWizard.frx":00EF
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Top             =   2040
               Width           =   2655
            End
            Begin VB.ComboBox cboSpacing 
               Height          =   315
               ItemData        =   "frmAnalysisWizard.frx":020D
               Left            =   3600
               List            =   "frmAnalysisWizard.frx":0223
               TabIndex        =   36
               Text            =   "15"
               Top             =   1200
               Width           =   1410
            End
            Begin VB.ComboBox cboInterpolation 
               Height          =   315
               ItemData        =   "frmAnalysisWizard.frx":023C
               Left            =   3600
               List            =   "frmAnalysisWizard.frx":0249
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   600
               Width           =   1395
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
               TabIndex        =   56
               Top             =   2040
               Width           =   870
            End
            Begin VB.Label lblSpacing 
               Caption         =   "Profile Spacing (in seconds).  Blank defaults to spacing of topo file."
               Height          =   615
               Left            =   705
               TabIndex        =   38
               Top             =   1200
               Width           =   2655
            End
            Begin VB.Label lbl3DInterpolation 
               Caption         =   "Interpolation Method"
               Height          =   450
               Left            =   705
               TabIndex        =   37
               Top             =   570
               Width           =   2520
            End
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic9 
         Height          =   4140
         Left            =   45
         TabIndex        =   6
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         Picture         =   "frmAnalysisWizard.frx":026B
         MouseIcon       =   "frmAnalysisWizard.frx":0287
         Begin VB.Frame fraClimatic 
            Caption         =   "Enter Climatic Information"
            Height          =   3135
            Left            =   720
            TabIndex        =   22
            Top             =   480
            Width           =   5175
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
               Left            =   3240
               MultiLine       =   -1  'True
               TabIndex        =   110
               Text            =   "frmAnalysisWizard.frx":02A3
               Top             =   480
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
               Left            =   3240
               MultiLine       =   -1  'True
               TabIndex        =   26
               Text            =   "frmAnalysisWizard.frx":02A8
               Top             =   2400
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
               Left            =   3240
               MultiLine       =   -1  'True
               TabIndex        =   25
               Text            =   "frmAnalysisWizard.frx":02AB
               Top             =   1920
               Width           =   1095
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
               Left            =   3240
               MultiLine       =   -1  'True
               TabIndex        =   24
               Text            =   "frmAnalysisWizard.frx":02B0
               Top             =   960
               Width           =   1095
            End
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
               ItemData        =   "frmAnalysisWizard.frx":02B6
               Left            =   3240
               List            =   "frmAnalysisWizard.frx":02D5
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "g/m3"
               Height          =   195
               Left            =   4440
               TabIndex        =   111
               Top             =   550
               Width           =   375
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
               TabIndex        =   33
               Top             =   2480
               Width           =   2655
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
               Left            =   4440
               TabIndex        =   32
               Top             =   1995
               Width           =   270
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
               TabIndex        =   31
               Top             =   2000
               Width           =   2160
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
               Left            =   4440
               TabIndex        =   30
               Top             =   1020
               Width           =   525
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
               TabIndex        =   29
               Top             =   1035
               Width           =   2565
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
               TabIndex        =   28
               Top             =   1520
               Width           =   1335
            End
            Begin VB.Label lblHumidity 
               AutoSize        =   -1  'True
               Caption         =   "Humidity"
               Height          =   195
               Left            =   360
               TabIndex        =   27
               Top             =   510
               Width           =   600
            End
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic12 
         Height          =   4140
         Left            =   1.50045e5
         TabIndex        =   7
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         Picture         =   "frmAnalysisWizard.frx":03BD
         MouseIcon       =   "frmAnalysisWizard.frx":03D9
         Begin VB.OptionButton optInterFreeSpace 
            Caption         =   "Free Space"
            Height          =   372
            Left            =   1320
            TabIndex        =   17
            Top             =   1560
            Width           =   3252
         End
         Begin VB.OptionButton optInterSEM 
            Caption         =   "Smooth Earth Model (SEM)"
            Height          =   372
            Left            =   1320
            TabIndex        =   16
            Top             =   2040
            Width           =   3372
         End
         Begin VB.OptionButton optInterTIREM 
            Caption         =   "Terrain Integrated Rough Earth Model (TIREM)"
            Height          =   372
            Left            =   1320
            TabIndex        =   15
            Top             =   2520
            Width           =   3732
         End
         Begin VB.Label Label3 
            Caption         =   "Enter method to calculate propagation loss for the interfering signal:"
            Height          =   375
            Left            =   840
            TabIndex        =   21
            Top             =   1200
            Width           =   5175
         End
         Begin VB.Label Label2 
            Caption         =   "I = Pt + Gt + Gr -"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   20
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label4 
            Caption         =   "Lp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3360
            TabIndex        =   19
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "- FDR - Ah"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4080
            TabIndex        =   18
            Top             =   360
            Width           =   1815
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic11 
         Height          =   4140
         Left            =   1.50195e5
         TabIndex        =   9
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         Picture         =   "frmAnalysisWizard.frx":03F5
         MouseIcon       =   "frmAnalysisWizard.frx":0411
         Begin VB.Frame fraBandwidth 
            Caption         =   "Roll-Off Information"
            Height          =   1095
            Left            =   720
            TabIndex        =   39
            Top             =   2520
            Width           =   5055
            Begin VB.TextBox txtEnvirRollOff 
               Height          =   315
               Left            =   3120
               TabIndex        =   47
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label67 
               Caption         =   "Environmental Equipment Roll-off (dB/Decade)"
               Height          =   375
               Left            =   360
               TabIndex        =   48
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.OptionButton optQuickFDR 
            Caption         =   "Quick FDR"
            Height          =   372
            Left            =   2640
            TabIndex        =   11
            Top             =   1680
            Width           =   1452
         End
         Begin VB.OptionButton optNoFDR 
            Caption         =   "No FDR"
            Height          =   375
            Left            =   2640
            TabIndex        =   10
            Top             =   1275
            Width           =   1095
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "- Ah"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5160
            TabIndex        =   112
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "I = Pt + Gt + Gr - Lp  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   840
            TabIndex        =   14
            Top             =   240
            Width           =   3225
         End
         Begin VB.Label Label6 
            Caption         =   "Enter method to calculate frequency dependent rejection (FDR):"
            Height          =   375
            Left            =   840
            TabIndex        =   13
            Top             =   840
            Width           =   5175
         End
         Begin VB.Label lblFDR 
            AutoSize        =   -1  'True
            Caption         =   "- FDR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3960
            TabIndex        =   12
            Top             =   240
            Width           =   1245
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic17 
         Height          =   4140
         Left            =   1.50345e5
         TabIndex        =   40
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         Picture         =   "frmAnalysisWizard.frx":042D
         MouseIcon       =   "frmAnalysisWizard.frx":0449
         Begin VB.Frame Frame7 
            Caption         =   "Analysis Options"
            Height          =   3135
            Left            =   600
            TabIndex        =   41
            Top             =   480
            Width           =   5295
            Begin VB.TextBox txtRXThreshold 
               Height          =   315
               Left            =   3000
               TabIndex        =   107
               Top             =   2400
               Width           =   1095
            End
            Begin VB.TextBox txtMaxDistance 
               Height          =   315
               Left            =   3000
               TabIndex        =   57
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtMaxFDR 
               Height          =   315
               Left            =   3000
               TabIndex        =   45
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox txtMinDistance 
               Height          =   315
               Left            =   3000
               TabIndex        =   42
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label lblRXThreshold 
               Caption         =   "Receiver Interference Threshold"
               Height          =   255
               Left            =   480
               TabIndex        =   109
               Top             =   2430
               Width           =   2295
            End
            Begin VB.Label lblRXThresholdUnits 
               AutoSize        =   -1  'True
               Caption         =   "dB"
               Height          =   195
               Left            =   4320
               TabIndex        =   108
               Top             =   2430
               Width           =   195
            End
            Begin VB.Label lblMaxFDRUnits 
               Caption         =   "dB"
               Height          =   255
               Left            =   4320
               TabIndex        =   68
               Top             =   1230
               Width           =   615
            End
            Begin VB.Label lblMaxAllowableUnits 
               Caption         =   "st. mi."
               Height          =   255
               Left            =   4320
               TabIndex        =   67
               Top             =   630
               Width           =   615
            End
            Begin VB.Label lblClosestApproachUnits 
               Caption         =   "st. mi."
               Height          =   255
               Left            =   4320
               TabIndex        =   66
               Top             =   1830
               Width           =   735
            End
            Begin VB.Label lblMaxDistance 
               AutoSize        =   -1  'True
               Caption         =   "Maximum Analysis Distance"
               Height          =   195
               Left            =   480
               TabIndex        =   58
               Top             =   600
               Width           =   1950
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Maximum Allowable FDR"
               Height          =   195
               Left            =   480
               TabIndex        =   44
               Top             =   1200
               Width           =   1755
            End
            Begin VB.Label Label46 
               Caption         =   "Minimum Distance for Mobiles"
               Height          =   255
               Left            =   480
               TabIndex        =   43
               Top             =   1830
               Width           =   2295
            End
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic20 
         Height          =   4140
         Left            =   1.50495e5
         TabIndex        =   49
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         Picture         =   "frmAnalysisWizard.frx":0465
         MouseIcon       =   "frmAnalysisWizard.frx":0481
         Begin VB.TextBox txtRadius 
            Height          =   315
            Left            =   3120
            TabIndex        =   104
            Top             =   2730
            Width           =   1410
         End
         Begin VB.TextBox txtMobileAntHt 
            Height          =   315
            Left            =   3120
            TabIndex        =   101
            Top             =   3300
            Width           =   1410
         End
         Begin VB.Frame Frame6 
            Caption         =   "Stepping Increment"
            Height          =   2055
            Left            =   480
            TabIndex        =   50
            Top             =   360
            Width           =   5655
            Begin VB.ComboBox cboLatSpacing 
               Height          =   315
               ItemData        =   "frmAnalysisWizard.frx":049D
               Left            =   3465
               List            =   "frmAnalysisWizard.frx":04B9
               TabIndex        =   52
               Text            =   "15"
               Top             =   480
               Width           =   1410
            End
            Begin VB.ComboBox cboLongSpacing 
               Height          =   315
               ItemData        =   "frmAnalysisWizard.frx":04DB
               Left            =   3465
               List            =   "frmAnalysisWizard.frx":04F7
               TabIndex        =   51
               Text            =   "15"
               Top             =   1320
               Width           =   1410
            End
            Begin VB.Label Label49 
               Caption         =   "Latitude Spacing (in seconds).  Blank defaults to spacing of topo file."
               Height          =   615
               Left            =   600
               TabIndex        =   54
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label Label48 
               Caption         =   "Longitude Spacing (in seconds).  Blank defaults to spacing of topo file."
               Height          =   660
               Left            =   600
               TabIndex        =   53
               Top             =   1275
               Width           =   2700
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Label lblEnterRadius 
            AutoSize        =   -1  'True
            Caption         =   "Enter Radius of Interest"
            Height          =   195
            Left            =   1080
            TabIndex        =   106
            Top             =   2790
            Width           =   1665
         End
         Begin VB.Label lblUnits 
            AutoSize        =   -1  'True
            Caption         =   "nmi"
            Height          =   195
            Left            =   5040
            TabIndex        =   105
            Top             =   2790
            Width           =   240
         End
         Begin VB.Label lblMobileAntHt 
            AutoSize        =   -1  'True
            Caption         =   "Antenna Height of Mobile"
            Height          =   195
            Left            =   1080
            TabIndex        =   103
            Top             =   3345
            Width           =   1800
         End
         Begin VB.Label lblAntHtUnits 
            AutoSize        =   -1  'True
            Caption         =   "meters"
            Height          =   195
            Left            =   4920
            TabIndex        =   102
            Top             =   3345
            Width           =   465
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic21 
         Height          =   4140
         Left            =   1.50645e5
         TabIndex        =   59
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         Picture         =   "frmAnalysisWizard.frx":0519
         MouseIcon       =   "frmAnalysisWizard.frx":0535
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   5040
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   3720
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.ListBox lstIntroRX 
            Height          =   1230
            ItemData        =   "frmAnalysisWizard.frx":0551
            Left            =   3480
            List            =   "frmAnalysisWizard.frx":0553
            TabIndex        =   65
            Top             =   820
            Width           =   2655
         End
         Begin VB.ListBox lstIntroTX 
            Height          =   1230
            ItemData        =   "frmAnalysisWizard.frx":0555
            Left            =   480
            List            =   "frmAnalysisWizard.frx":0557
            TabIndex        =   64
            Top             =   820
            Width           =   2655
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete Link"
            Height          =   375
            Left            =   5040
            TabIndex        =   63
            Top             =   2880
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add Link"
            Height          =   375
            Left            =   5040
            TabIndex        =   62
            Top             =   2280
            Width           =   1095
         End
         Begin MSDBGrid.DBGrid dbgIntroLinks 
            Bindings        =   "frmAnalysisWizard.frx":0559
            Height          =   1575
            Left            =   480
            OleObjectBlob   =   "frmAnalysisWizard.frx":0569
            TabIndex        =   60
            Top             =   2280
            Width           =   4215
         End
         Begin VB.Label lblIntroRX 
            AutoSize        =   -1  'True
            Caption         =   "Introduced Receivers"
            Height          =   195
            Left            =   3480
            TabIndex        =   82
            Top             =   580
            Width           =   1530
         End
         Begin VB.Label lblIntroTX 
            AutoSize        =   -1  'True
            Caption         =   "Introduced Transmitters"
            Height          =   195
            Left            =   480
            TabIndex        =   81
            Top             =   580
            Width           =   1665
         End
         Begin VB.Label lblEstablishIntroLinks 
            AutoSize        =   -1  'True
            Caption         =   "Establish Introduced Desired Links:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   480
            TabIndex        =   61
            Top             =   150
            Width           =   3675
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
         Height          =   4140
         Left            =   1.50795e5
         TabIndex        =   69
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         Picture         =   "frmAnalysisWizard.frx":0F42
         MouseIcon       =   "frmAnalysisWizard.frx":0F5E
         Begin VB.Frame fraHarmonicInfo 
            Caption         =   "Harmonic Information"
            Height          =   1335
            Left            =   1080
            TabIndex        =   75
            Top             =   2520
            Width           =   4455
            Begin VB.TextBox txtHarmonicAttenuation 
               Height          =   285
               Left            =   2280
               TabIndex        =   78
               Top             =   850
               Width           =   1215
            End
            Begin VB.ComboBox cboHarmonicOrder 
               Height          =   315
               ItemData        =   "frmAnalysisWizard.frx":0F7A
               Left            =   3120
               List            =   "frmAnalysisWizard.frx":0F87
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "dB"
               Height          =   195
               Left            =   3720
               TabIndex        =   80
               Top             =   850
               Width           =   195
            End
            Begin VB.Label lblHarmonicAttenuation 
               Caption         =   "Harmonic Attenuation"
               Height          =   255
               Left            =   360
               TabIndex        =   79
               Top             =   850
               Width           =   1695
            End
            Begin VB.Label lblHarmonicOrder 
               Caption         =   "Enter highest order to be considered"
               Height          =   255
               Left            =   360
               TabIndex        =   77
               Top             =   360
               Width           =   2655
            End
         End
         Begin VB.OptionButton optNoHarmonic 
            Caption         =   "No Harmonics"
            Height          =   375
            Left            =   1680
            TabIndex        =   71
            Top             =   1515
            Width           =   1695
         End
         Begin VB.OptionButton optHarmonic 
            Caption         =   "Consider Harmonics"
            Height          =   375
            Left            =   1680
            TabIndex        =   70
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblHarmonicEqu 
            AutoSize        =   -1  'True
            Caption         =   "- Ah"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5280
            TabIndex        =   74
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label10 
            Caption         =   "Should this analysis consider harmonics (Ah)?"
            Height          =   375
            Left            =   1680
            TabIndex        =   73
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label Label9 
            Caption         =   "I = Pt + Gt + Gr - Lp - Ls - FDR "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   72
            Top             =   360
            Width           =   4815
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic2 
         Height          =   4140
         Left            =   1.50945e5
         TabIndex        =   83
         Top             =   150
         Width           =   6645
         _Version        =   327680
         _ExtentX        =   11721
         _ExtentY        =   7303
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         BorderWidth     =   0
         ChildSpacing    =   0
         Picture         =   "frmAnalysisWizard.frx":0F94
         MouseIcon       =   "frmAnalysisWizard.frx":0FB0
         Begin VB.OptionButton optGalactic 
            Caption         =   "Galactic"
            Height          =   375
            Left            =   360
            TabIndex        =   89
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton optQuietRural 
            Caption         =   "Quiet Rural"
            Height          =   375
            Left            =   360
            TabIndex        =   88
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton optRural 
            Caption         =   "Rural"
            Height          =   375
            Left            =   360
            TabIndex        =   87
            Top             =   1770
            Width           =   1095
         End
         Begin VB.OptionButton optResidential 
            Caption         =   "Residential"
            Height          =   375
            Left            =   360
            TabIndex        =   86
            Top             =   2250
            Width           =   1455
         End
         Begin VB.OptionButton optBusiness 
            Caption         =   "Business"
            Height          =   375
            Left            =   360
            TabIndex        =   85
            Top             =   2730
            Width           =   1335
         End
         Begin VCIFiLib.VtChart VtChart1 
            Height          =   3855
            Left            =   1560
            OleObjectBlob   =   "frmAnalysisWizard.frx":0FCC
            TabIndex        =   90
            Top             =   480
            Width           =   5295
         End
         Begin VB.Label Label14 
            Caption         =   "Enter type of background noise:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   91
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   240
            TabIndex        =   84
            Top             =   240
            Width           =   45
         End
      End
   End
   Begin VB.Label Label33 
      Caption         =   "Emission Roll-off (dB/Decade)"
      Height          =   375
      Left            =   1320
      TabIndex        =   46
      Top             =   3600
      Width           =   2415
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
   Begin VB.Menu mnuLatUnits 
      Caption         =   "LatUnits"
      Visible         =   0   'False
      Begin VB.Menu mnuDMS 
         Caption         =   "dd mm ss H"
      End
      Begin VB.Menu mnuDD 
         Caption         =   "decimal degrees"
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
   Begin VB.Menu mnuLongUnits 
      Caption         =   "LongUnits"
      Visible         =   0   'False
      Begin VB.Menu mnuLongDMS 
         Caption         =   "ddd mm ss H"
      End
      Begin VB.Menu mnuLongDD 
         Caption         =   "decimal degrees"
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
End
Attribute VB_Name = "frmAnalysisWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'
' *** COPYRIGHT:  Copyright 1998
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
' 8/15/98    Mike Fahler                    GATE v. 1.00
'******************************************************************************

Option Explicit

'recordsets for i/n analysis and as required
Dim EnvirAnalysisRecordset As Recordset
Dim IntroAnalysisRecordset As Recordset

'additional recordsets for s/i analysis
Dim IntroTXRecordset As Recordset
Dim IntroRXRecordset As Recordset
Dim IntroLinkRS As Recordset

'result db
Dim ResultRecordset As Recordset

'variables for unit conversion
Dim UnitString As String
Dim UnitLabel As Control
Dim UnitText As Control

Dim EnvirName As String
Dim EnvirID As Long
Dim EnvirLat As Single
Dim EnvirLong As Single
Dim EnvirLatRad As Single
Dim EnvirLongRad As Single
Dim EnvirAntennaHt As Single
Dim EnvirAntMotion As String
Dim EnvirPolar As String * 1
Dim EnvirCurve As Double
Dim EnvirRollOff As Double
Dim EnvirFreqMin As Single
Dim EnvirFreqMax As Single
Dim EnvirSingleFreq As Double
Dim EnvirRadiusofMobility As Single
Dim EnvirDegOffAxis As Single
Dim EnvirOffAxisGain As Single
Dim EnvirPointingAngle As Single
Dim EnvirMBGain As Single
Dim EnvirCrossPolar As Boolean  'if MB coupling then true

'tx uniq
Dim TPOWER As Single
Dim Modulation As String

Dim IntroName As String
Dim IntroID As Long
Dim IntroLat As Single
Dim IntroLong As Single
Dim IntroLatRad As Single
Dim IntroLongRad As Single
Dim IntroLatDD As Single
Dim IntroLongDD As Single
Dim IntroAntennaHt As Single
Dim IntroFreq As Double
Dim IntroCurve As Double
Dim IntroRollOff As Double
Dim IntroPolar As String * 1
Dim IntroRadiusofMobility As Single
Dim IntroDegOffAxis As Single
Dim IntroOffAxisGain As Single
Dim IntroPointingAngle As Single
Dim IntroMBGain As Single
Dim IntroCrossPolar As Boolean  'if MB coupling then true

'rx uniq
Dim RXThreshold As Single
Dim NoiseFigure As Single

'passed to subs to distinguish between intro tx and rx
Const RX As Integer = 1
Const TX As Integer = 2

'off-axis angles for antenna gain
Dim Q As Single
Dim MinAngle As Single
Dim MaxAngle As Single
Dim MinOffAxis As Single  'angle between IntroPointingAngle and minangle
Dim MaxOffAxis As Single  'angle between IntroPointingAngle and maxangle
Dim RelativeGain As Single
Dim StandardDev As Single
Dim STATError As Long

Dim Spacing As Integer
Dim LatSpacing As Double
Dim LongSpacing As Double
Dim LongSpacingDD As Double
Dim LatSpacingDD As Double
Dim Datum As Long

Dim TabStepArray() As Integer
Dim TabStepCounter As Integer
Dim NumberOfTabs As Integer

Dim groundtype As Long
Dim SeaRefract As Single
Dim CONDUC As Single
Dim PERMIT As Single
Dim REFRAC As Single

Dim LatNERad As Double
Dim LatSWRad As Double
Dim LongNERad As Double
Dim LongSWRad As Double
Dim SignalStrength As Single

'WOTLRET
Dim SPACNG As Single  ' profile spacing in meters
Dim spheroid As Long
Dim ecc As Single
Dim ERRCODE As Long
Dim MJAxis As Single
Dim Flat As Single

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

Dim TANTHT As Single
Dim RANTHT As Single
Dim HUMID As Single
Dim EnvirELEV As Single
Dim RELEV As Single

Dim SiteELEV As Single ' The elevation at the specific lat/lon from topget.

Dim POLARZ As String * 4
Dim INTERP As Long    ' interpolation method

'sem declarations
Dim SEMMODE As String * 4
Dim SEMPRLoss As Single
Dim SEMFSPLSS As Single

Dim FreqSep As Single

Dim InterPropagationModel As Integer

'quick fdr output
Dim FDR As Double
Dim OTR As Double
Dim OFR As Double
Dim FDRError As Long
Dim QuickFDR As Boolean  'flag to determine if Quick FDR is performed

'harmonics
Dim HarmonicOrder As Integer  '1=no harmonics
Dim HarmonicAttenuation As Single
Dim FDRHarmonic As Double
Dim HarmonicPropFreq As Double
Dim HarmonicBandwidth As Double
Dim Harmonic As Integer  'stored to result db

'polarization
Dim CrossPolar As Integer

'noise
Dim C_Constant As Single
Dim D_Constant As Single
Dim NoiseAnalysis As Boolean

'run options/control
Dim MaxAllowableFDR As Single
Dim MinDistanceMobiles As Single
Dim MaxDistance As Single

Dim PRFERR As Long       ' for PRFILE dll call - error return 0 => OK and 1 => Error.

'LOS polygons
Dim VertexArray(4, 2) As Double
Dim MobileAntHt As Single
Dim Circle_Radius As Single
Dim LowScale As Long

'analysis frequencies
Dim PROPFQ As Single
Dim LOSFreq As Single

Dim Mode As String * 4 ' MODE INDICATOR: LINE OF SIGHT, DIFFRACTION, or TROPO SCATTER from Tirem Dll
Dim PropMode As String  'model(tirem or sem)+_+ mode
Dim PRLoss As Single     ' TOTAL PATH LOSS (BASIC TRANSMISSION LOSS) IN DB from Tirem Dll
Dim FSPLSS As Single     ' FREE SPACE LOSS IN DB from Tirem DLL

'path length/bearing
Dim BearIE As Single
Dim BearEI As Single
Dim PTHLEN As Single
Dim PTHLENMobile As Single
Dim BearIE_deg As Single
Dim BearEI_deg As Single
Dim SlantRange As Single

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

Dim InterferenceLevel As Single
Dim NoiseLevel As Single

Dim ErrorFound As String
Dim AnalysisComment As String

Dim CalculatingDesSignal As Boolean  'indicates proploss for the desired signal for S/I is being calculated

Dim PathLengthLimit As Long
Dim BandWidth3dB As Single
Dim UseThisRecordForNextStep As Boolean

Dim SymbolCounter As Integer
Dim GridCounter As Long
Dim TotalGridPoints As Long

Dim DeltaAngle As Single

Dim SiteElevation As Single

Dim response As Integer  'msgbox answer

Dim TotalLatRad As Double
Dim TotalLongRad As Double

'analysis counters for status, analysis info
Dim NumberOfRows As Long
Dim NumberOfColumns As Long
Dim ColumnIndex As Long
Dim RowIndex As Long
Dim NumberOfGridPoints As Long
Dim NumberOfIntro As Long
Dim NumberOfEnvir As Long
Dim AnalysisCounter As Long  'number of records analyzed
Dim ErrorCounter As Long  'number of records not processed due to lack of data
Dim InterferenceCounter As Long  'number of records exceeding thresh
Dim InteractionCounter As Long

'status bar update
Dim PercentComplete As Integer

Private Sub cboGroundTypes_Click()

groundtype = Val(cboGroundTypes.ListIndex)

'temporarily set prop freq = to receiver freq
Select Case AnalysisType
Case 1, 2
    PROPFQ = IntroRXSiteRecordset("freq(mhz)")
Case 3
    PROPFQ = IntroTXSiteRecordset("freq(mhz)")
Case 5
    PROPFQ = LOSFreq
End Select

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


Private Sub cmdAdd_Click()
On Error GoTo errorhandler

If lstIntroTX.ListIndex = -1 Or lstIntroRX.ListIndex = -1 Then 'item not selected
    MsgBox "You must select both an introduced transmitter and receiver.", vbOKOnly, "Input Required"
    Exit Sub
End If

Data1.Recordset.AddNew
'add to database
Data1.Recordset("TX_Name") = lstIntroTX.List(lstIntroTX.ListIndex)
Data1.Recordset("RX_Name") = lstIntroRX.List(lstIntroRX.ListIndex)

Data1.Recordset.Update

Data1.Refresh

dbgIntroLinks.ReBind

cmdDelete.Enabled = True

errorhandler:
    Select Case Err.Number

    Case 3022  'tx/rx pair already exists
        MsgBox "TX/RX Pair already exists.  Please select another pair.", vbExclamation + vbOKOnly, "Warning"
        Data1.Recordset.MoveFirst
        dbgIntroLinks.ReBind
    Case Is > 0
        MsgBox Err.Description, vbOKOnly, "Warning"
    End Select

End Sub

Private Sub cmdBack_Click()

cmdNext.Enabled = True
cmdFinish.Enabled = False

TabStepCounter = TabStepCounter - 1

'if tirem not selected, -1's are inserted to indicate tabs should be skipped

If TabStepArray(TabStepCounter) = -1 Then
    TabStepCounter = TabStepCounter - 1
    
    If TabStepArray(TabStepCounter) = -1 Then
        TabStepCounter = TabStepCounter - 1
    End If
    
End If

tabAnalysisWizard.CurrTab = TabStepArray(TabStepCounter)

If TabStepCounter = 1 Then
    cmdBack.Enabled = False
End If

End Sub

Private Sub cmdCancel_Click()
DeleteOldAnalysis = False

Select Case AnalysisType
Case 1, 2
    TXCustomFilterString = ""
    CustomTXQuery = False
    frmProjectFeatures.Refresh_TXSiteLayer
Case 3
    RXCustomFilterString = ""
    CustomRXQuery = False
    frmProjectFeatures.Refresh_RXSiteLayer
End Select

Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next

Data1.Recordset.Delete
Data1.Refresh

dbgIntroLinks.ReBind
dbgIntroLinks.Refresh

If Data1.Recordset.RecordCount = 0 Then
    cmdDelete.Enabled = False
End If

End Sub

Public Sub cmdFinish_Click()

Dim DataMissing As Boolean  'for LOS contour
Dim GeneratedError As Boolean
Dim MyBlob As Variant
Dim GalacticNoise As Single
Dim ManMadeNoise As Single

DataMissing = True  'initialize
CalculatingDesSignal = False
On Error GoTo errorhandler

If CheckValidityOfCurrentTabInputs = 1 Then
    Exit Sub
End If

response = MsgBox("Save analysis parameters for next run?", vbQuestion + vbYesNo, "Save Inputs")

If response = 6 Then  'yes
    tabAnalysisWizard.CurrTab = 0

    Dim FormControl As Control
    
'write to analysis.ini
    Open App.Path + "\analysis.ini" For Output As #1 ' Open file to write.
    
    For Each FormControl In frmAnalysisWizard
        If TypeOf FormControl Is TextBox Then
            Write #1, FormControl.Text
        End If
        
        If TypeOf FormControl Is OptionButton Then
            Write #1, FormControl.Value
        End If
        
        If TypeOf FormControl Is Label Then
            Write #1, FormControl.Caption
        End If
        
        If TypeOf FormControl Is ComboBox Then
            If FormControl.Style = 2 Then  'dropdown list
                Write #1, FormControl.ListIndex
            Else
                Write #1, FormControl.Text
            End If
        End If
    Next
    Close #1    ' Close file.

End If

Screen.MousePointer = 11

Me.Hide

frmStatusBar.Show
DoEvents

'initialize
AnalysisCounter = 0 'for status bar
InterferenceCounter = 0 'number of records exceeding thresh
ErrorCounter = 0 'number of records not processed due to lack of data

Select Case AnalysisType

Case 1  'env tx versus intro rx I/N
    Set ResultRecordset = MyDatabase.OpenRecordset("pnt_introrx_results")
    Set EnvirAnalysisRecordset = MyDatabase.OpenRecordset("pnt_txsites")
    Set IntroAnalysisRecordset = MyDatabase.OpenRecordset("pnt_introrxsites")
    
    IntroAnalysisRecordset.MoveLast
    NumberOfIntro = IntroAnalysisRecordset.RecordCount
    
    EnvirAnalysisRecordset.MoveLast
    NumberOfEnvir = EnvirAnalysisRecordset.RecordCount
    
    InteractionCounter = NumberOfIntro * NumberOfEnvir
    
    IntroAnalysisRecordset.MoveFirst

'set major and minor axis for WOTL and TIREM call, wont vary per tx site so no reason to repeat inside tx loop
    Call Datum2Axis(Datum, spheroid, MJAxis, MNAxis, Flat, ecc, ERRCODE)

    Do Until IntroAnalysisRecordset.EOF

        'Load Intro DB to variables
        IntroDB_to_Variables (RX)  'call sub
        
        EnvirAnalysisRecordset.MoveFirst
        
        Do Until EnvirAnalysisRecordset.EOF
            'reinitialize
            Initialize_Common_Variables
            
            AnalysisCounter = AnalysisCounter + 1
            
            'Load envir DB to variables
            If EnvirDB_to_Variables(TX) = 1 Then 'call sub
                GoTo UpdateResultDB
            End If
                        
    'pathlength in meters
            CalculatePathLength  'call sub
            
            Calculate_PathLength_Mobiles  'call sub
            
            If PTHLENMobile > MaxDistance * 1000 Then 'discard
                GoTo NextRecord  'analysis for this record over
            End If
    
            EnvirSingleFreq = Range_to_Single_Freq(EnvirFreqMin, EnvirFreqMax, IntroFreq)
           
            If QuickFDR = True Then
                
                EnvirCurve = EnvirAnalysisRecordset("bdw_3db(khz)")
                
                If GeneratedError = True Then
                    GeneratedError = False
                    GoTo NextRecord
                End If
                
                If EnvirCurve <= 0 Then
                    ErrorFound = "Missing TX Bandwidth"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateResultDB
                End If
                
                Call CALCQUICKFDR(EnvirSingleFreq * 1000, IntroFreq * 1000, _
                        EnvirCurve, IntroCurve, EnvirRollOff, _
                        IntroRollOff, FDR, OTR, OFR, FDRError)
             
                If FDRError < 0 Then
                    ErrorFound = "Error Calculating FDR"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateResultDB
                End If
                
                If FDR > MaxAllowableFDR Then
                    FDR = MaxAllowableFDR
                End If

                FreqSep = Abs(IntroFreq - EnvirSingleFreq)
                
                Harmonic = 1
    
    'calculate harmonics
                If HarmonicOrder > 1 Then 'consider harmonics
    'compare FDR and harmonic attenuation to get worst case interferer for each env tx
                    Calculate_IntroRX_HarmonicAttenuation
                End If
            
            Else
                FDR = 0
            End If
    
            PROPFQ = IntroFreq  'set to rx tuned frequency
            
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
                Select Case PRLoss
                Case 0
                    ErrorFound = "Error Calculating Path Loss"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateResultDB
                Case 999  'free-space error
                    ErrorFound = "Link not LOS - Free Space NA"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateResultDB
                End Select
            End If
            
            Calculate_Envir_OffAxisGain
    
            Calculate_Intro_OffAxisGain
    
            Calculate_CrossPolar 'call sub
            
    'interference level - harmonic attenuation is included in FDR
            InterferenceLevel = EnvirAnalysisRecordset("power(dbm)") + EnvirOffAxisGain + IntroOffAxisGain - PRLoss - FDR - CrossPolar
                    
    'noise for receiver BW in khz
            If PROPFQ < 250 = True Then
                GalacticNoise = 10 ^ ((52 - 23 * (Log(PROPFQ) / Log(10))) / 10)
                
                If C_Constant = 52 Then 'galactic noise
                    ManMadeNoise = 0
                Else  'man-made noise
                    ManMadeNoise = 10 ^ ((C_Constant - D_Constant * (Log(PROPFQ) / Log(10))) / 10)
                End If
            
                NoiseLevel = 10 * (Log(IntroCurve) / Log(10)) - 144 _
                                + 10 * (Log(10 ^ (NoiseFigure / 10) - 1 + GalacticNoise + ManMadeNoise) / Log(10))
            Else
                NoiseLevel = 10 * (Log(IntroCurve) / Log(10)) - 144 + 10 * (Log(10 ^ (NoiseFigure / 10) - 1) / Log(10))
            End If
            
            If InterferenceLevel - NoiseLevel > RXThreshold Then
    
    'update recordsets
UpdateResultDB:
                Write_to_ResultDatabase
                
            End If  'if threshold exceeded
            
NextRecord:
            EnvirAnalysisRecordset.MoveNext  'go to next record
    
            Update_StatusBar_Percent
            
        Loop
        
        IntroAnalysisRecordset.MoveNext  'go to next record
    
    Loop

    Call Draw_Multi_Symbol("pnt_introrx_results.error", TXSiteLayer, "Square", 9408399, 8388608, 4, 15)
    
    Analysis_Summary
    
    IntroRXINPerformed = True
    IntroRXSIPerformed = False
    
    If EnvirTXLink = True Then
        frmProjectFeatures.Delete_Feature_Data "pln_link_designator WHERE rrlntype = 1;" 'delete assosiated links
        EnvirTXLink = False
        frmLinkCriteria.Form_Load
        frmLinkCriteria.chkRXInterferenceLinks.Value = 1
        frmLinkCriteria.Determine_Links
    End If
    
    RX_Analysis_ON  'call sub, set button bar
    
Case 2  'Environmental TX versus Intro RX S/I
    Dim LinkTXID As Long
    Dim LinkRXID As Long

'open recordsets
    Set IntroLinkRS = MyDatabase.OpenRecordset("introlinks")
    
    IntroLinkRS.MoveLast
    IntroLinkRS.MoveFirst
    
    Do Until IntroLinkRS.EOF
        
        LinkTXID = IntroLinkRS("TX_ID")
        
        LinkRXID = IntroLinkRS("RX_ID")

        Set IntroAnalysisRecordset = MyDatabase.OpenRecordset _
                ("SELECT * FROM pnt_introrxsites WHERE pnt_introrxsites.pointid = " + Str(LinkRXID))

'Load Intro DB to variables
        IntroDB_to_Variables (RX)  'call sub

'save pertinent info to environmental variables
        EnvirLatRad = IntroLatRad
        EnvirLongRad = IntroLongRad
        EnvirAntennaHt = IntroAntennaHt
        EnvirMBGain = IntroMBGain
        EnvirRadiusofMobility = IntroRadiusofMobility
        EnvirELEV = SiteElevation
        EnvirName = IntroName
        
        Set IntroAnalysisRecordset = MyDatabase.OpenRecordset _
                ("SELECT * FROM pnt_introtxsites WHERE pnt_introtxsites.pointid = " + Str(LinkTXID))

'Load Intro DB to variables
        IntroDB_to_Variables (TX)  'call sub
        
'calculate signal strength
    'pathlength in meters
        CalculatePathLength  'call sub
            
        'use furthest separation for mobiles
        PTHLENMobile = PTHLEN + (EnvirRadiusofMobility + IntroRadiusofMobility) * 1000
        
        PROPFQ = IntroFreq
            
'set major and minor axis for WOTL and TIREM call, wont vary per tx site so no reason to repeat inside tx loop
        Call Datum2Axis(Datum, spheroid, MJAxis, MNAxis, Flat, ecc, ERRCODE)
        
        CalculatingDesSignal = True
        
        'check frequency limits
        If PROPFQ < 1 Or PROPFQ > 20000 Then
            MsgBox "Frequency out of range for " + IntroName + " and " + EnvirName + ".  Delete this link, and re-run.", vbOKOnly, "Error"
            Screen.MousePointer = 0
            Unload frmStatusBar
            DoEvents
            cmdBack.Enabled = False
            tabAnalysisWizard.CurrTab = 7
            Me.Show
            
            Exit Sub
        End If
            
        Calculate_Propagation_Loss  'call sub
        
        If PRFERR > 0 Or PRLoss = 0 Or PRLoss = 999 Then
            MsgBox "Error calculating received signal for " + IntroName + " and " + EnvirName + ".  Delete this link, and re-run.", vbOKOnly, "Error"
            Screen.MousePointer = 0
            Unload frmStatusBar
            DoEvents
            cmdBack.Enabled = False
            tabAnalysisWizard.CurrTab = 7
            Me.Show
            
            Exit Sub
        End If
        
        SignalStrength = TPOWER + IntroMBGain + EnvirMBGain - PRLoss
                
        IntroLinkRS.Edit
        IntroLinkRS("Signal") = SignalStrength
        IntroLinkRS.Update
        
        IntroLinkRS.MoveNext
        
    Loop

'now calculate Interference
    CalculatingDesSignal = False
   
    Set ResultRecordset = MyDatabase.OpenRecordset("pnt_introrx_results")
    Set EnvirAnalysisRecordset = MyDatabase.OpenRecordset("pnt_txsites")
    Set IntroAnalysisRecordset = MyDatabase.OpenRecordset _
            ("SELECT * FROM pnt_introrxsites,introlinks WHERE pnt_introrxsites.pointid = introlinks.rx_id")
    
    IntroAnalysisRecordset.MoveLast
    NumberOfIntro = IntroAnalysisRecordset.RecordCount
    
    EnvirAnalysisRecordset.MoveLast
    NumberOfEnvir = EnvirAnalysisRecordset.RecordCount
    
    InteractionCounter = NumberOfIntro * NumberOfEnvir
    
    IntroAnalysisRecordset.MoveFirst

'set major and minor axis for WOTL and TIREM call, wont vary per tx site so no reason to repeat inside tx loop
    Call Datum2Axis(Datum, spheroid, MJAxis, MNAxis, Flat, ecc, ERRCODE)

    Do Until IntroAnalysisRecordset.EOF

        'Load Intro DB to variables
        IntroDB_to_Variables (RX)  'call sub
        
        SignalStrength = IntroAnalysisRecordset("Signal")
        
        EnvirAnalysisRecordset.MoveFirst
        
        Do Until EnvirAnalysisRecordset.EOF
            'reinitialize
            Initialize_Common_Variables
            
            AnalysisCounter = AnalysisCounter + 1
            
            'Load to variables
            If EnvirDB_to_Variables(TX) = 1 Then 'call sub
                GoTo UpdateSIResultDB
            End If
    
    'pathlength in meters
            CalculatePathLength  'call sub
            
            Calculate_PathLength_Mobiles  'call sub
            
            If PTHLENMobile > MaxDistance * 1000 Then 'discard
                GoTo NextSIRecord  'analysis for this record over
            End If
            
            EnvirSingleFreq = Range_to_Single_Freq(EnvirFreqMin, EnvirFreqMax, IntroFreq)
           
            If QuickFDR = True Then
                
                EnvirCurve = EnvirAnalysisRecordset("bdw_3db(khz)")

                If GeneratedError = True Then
                    GeneratedError = False
                    GoTo NextSIRecord
                End If
                
                If EnvirCurve <= 0 Then
                    ErrorFound = "Missing TX Bandwidth"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateSIResultDB
                End If
                
                Call CALCQUICKFDR(EnvirSingleFreq * 1000, IntroFreq * 1000, _
                        EnvirCurve, IntroCurve, EnvirRollOff, _
                        IntroRollOff, FDR, OTR, OFR, FDRError)
             
                If FDRError < 0 Then
                    ErrorFound = "Error Calculating FDR"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateSIResultDB
                End If
                
                If FDR > MaxAllowableFDR Then
                    FDR = MaxAllowableFDR
                End If

                FreqSep = Abs(IntroFreq - EnvirSingleFreq)
                
                Harmonic = 1
    
    'calculate harmonics
                If HarmonicOrder > 1 Then 'consider harmonics
    'compare FDR and harmonic attenuation to get worst case interferer for each env tx
                    Calculate_IntroRX_HarmonicAttenuation
                End If
            
            Else
                FDR = 0
            End If
    
            PROPFQ = IntroFreq  'set to rx tuned frequency
            
    'check frequency limits
            If PROPFQ < 1 Or PROPFQ > 20000 Then
                ErrorFound = "Frequency Out of Range"
                ErrorCounter = ErrorCounter + 1
                GoTo UpdateSIResultDB
            End If
            
            Calculate_Propagation_Loss  'call sub
            
'check for errors calculating prop loss
            If PRFERR > 0 Then
                ErrorFound = "Error Extracting Elevation Data"
                ErrorCounter = ErrorCounter + 1
                GoTo UpdateSIResultDB
            Else
                Select Case PRLoss
                Case 0
                    ErrorFound = "Error Calculating Path Loss"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateSIResultDB
                Case 999  'free-space error
                    ErrorFound = "Link not LOS - Free Space NA"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateSIResultDB
                End Select
            End If
            
            Calculate_Envir_OffAxisGain
    
            Calculate_Intro_OffAxisGain
    
            Calculate_CrossPolar 'call sub
            
    'interference level - harmonic attenuation is included in FDR
            InterferenceLevel = EnvirAnalysisRecordset("power(dbm)") + EnvirOffAxisGain + IntroOffAxisGain - PRLoss - FDR - CrossPolar
            
            If SignalStrength - InterferenceLevel < RXThreshold Then
    
    'update recordsets
UpdateSIResultDB:
                Write_to_ResultDatabase
                
            End If  'if threshold exceeded
            
NextSIRecord:
            EnvirAnalysisRecordset.MoveNext  'go to next record
    
            Update_StatusBar_Percent
            
        Loop
        
        IntroAnalysisRecordset.MoveNext  'go to next record
    
    Loop
    
    Call Draw_Multi_Symbol("pnt_introrx_results.error", TXSiteLayer, "Square", 9408399, 8388608, 4, 15)
    
    Analysis_Summary
    
    IntroRXSIPerformed = True
    IntroRXINPerformed = False

    If EnvirTXLink = True Then
        frmProjectFeatures.Delete_Feature_Data "pln_link_designator WHERE rrlntype = 1;" 'delete assosiated links
        EnvirTXLink = False
        frmLinkCriteria.Form_Load
        frmLinkCriteria.chkRXInterferenceLinks.Value = 1
        frmLinkCriteria.Determine_Links
    End If
    
    If IntroLink = True Then
        frmProjectFeatures.Delete_Feature_Data "pln_link_designator WHERE rrlntype = 3;" 'delete assosiated links
        IntroLink = False
        frmLinkCriteria.Form_Load
        frmLinkCriteria.chkIntroDesiredLinks.Value = 1
        frmLinkCriteria.Determine_Links
    End If
    
    RX_Analysis_ON  'call sub, set button bar

Case 3  'Intro TX versus environmental rx I/N
    
    Set ResultRecordset = MyDatabase.OpenRecordset("pnt_introtx_results")
    Set EnvirAnalysisRecordset = MyDatabase.OpenRecordset("pnt_rxsites")
    Set IntroAnalysisRecordset = MyDatabase.OpenRecordset("pnt_introtxsites")
    
    IntroAnalysisRecordset.MoveLast
    NumberOfIntro = IntroAnalysisRecordset.RecordCount
    
    EnvirAnalysisRecordset.MoveLast
    NumberOfEnvir = EnvirAnalysisRecordset.RecordCount
    
    InteractionCounter = NumberOfIntro * NumberOfEnvir
    
    IntroAnalysisRecordset.MoveFirst

'set major and minor axis for WOTL and TIREM call, wont vary per tx site so no reason to repeat inside tx loop
    Call Datum2Axis(Datum, spheroid, MJAxis, MNAxis, Flat, ecc, ERRCODE)

    Do Until IntroAnalysisRecordset.EOF

        'Load Intro DB to variables
        IntroDB_to_Variables (TX)  'call sub
        
        EnvirAnalysisRecordset.MoveFirst
        
        Do Until EnvirAnalysisRecordset.EOF
            'reinitialize
            Initialize_Common_Variables
            
            NoiseLevel = 0  'initialize analysis uniq variables
            
            AnalysisCounter = AnalysisCounter + 1
            
            'Load Intro DB to variables
            If EnvirDB_to_Variables(RX) = 1 Then 'call sub
                GoTo UpdateIntroTXResultDB
            End If
    
    'pathlength in meters
            CalculatePathLength  'call sub
            
            Calculate_PathLength_Mobiles  'call sub
            
            If PTHLENMobile > MaxDistance * 1000 Then 'discard
                GoTo NextIntroTXRecord  'analysis for this record over
            End If
            
            EnvirSingleFreq = Range_to_Single_Freq(EnvirFreqMin, EnvirFreqMax, IntroFreq)
           
            EnvirCurve = EnvirAnalysisRecordset("bdw_3db(khz)")

            If EnvirCurve <= 0 Then
                ErrorFound = "Missing RX Bandwidth"
                ErrorCounter = ErrorCounter + 1
                GoTo UpdateIntroTXResultDB
            End If
                
            If GeneratedError = True Then
                GeneratedError = False
                GoTo NextIntroTXRecord
            End If
            
            If QuickFDR = True Then
                
                Call CALCQUICKFDR(IntroFreq * 1000, EnvirSingleFreq * 1000, _
                        IntroCurve, EnvirCurve, IntroRollOff, _
                        EnvirRollOff, FDR, OTR, OFR, FDRError)
             
                If FDRError < 0 Then
                    ErrorFound = "Error Calculating FDR"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateIntroTXResultDB
                End If
                
                If FDR > MaxAllowableFDR Then
                    FDR = MaxAllowableFDR
                End If

                FreqSep = Abs(IntroFreq - EnvirSingleFreq)
                
                Harmonic = 1
    
    'calculate harmonics
                If HarmonicOrder > 1 Then 'consider harmonics
    'compare FDR and harmonic attenuation to get worst case interferer for each env tx
                    Calculate_IntroTX_HarmonicAttenuation
                End If
            
            Else
                FDR = 0
            End If
    
            PROPFQ = EnvirSingleFreq
            
    'check frequency limits
            If PROPFQ < 1 Or PROPFQ > 20000 Then
                ErrorFound = "Frequency Out of Range"
                ErrorCounter = ErrorCounter + 1
                GoTo UpdateIntroTXResultDB
            End If
            
            Calculate_Propagation_Loss  'call sub
            
'check for errors calculating prop loss
            If PRFERR > 0 Then
                ErrorFound = "Error Extracting Elevation Data"
                ErrorCounter = ErrorCounter + 1
                GoTo UpdateIntroTXResultDB
            Else
                Select Case PRLoss
                Case 0
                    ErrorFound = "Error Calculating Path Loss"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateIntroTXResultDB
                Case 999  'free-space error
                    ErrorFound = "Link not LOS - Free Space NA"
                    ErrorCounter = ErrorCounter + 1
                    GoTo UpdateIntroTXResultDB
                End Select
            End If
            
            Calculate_Envir_OffAxisGain
    
            Calculate_Intro_OffAxisGain
    
            Calculate_CrossPolar 'call sub
            
'interference level - harmonic attenuation is included in FDR
            InterferenceLevel = TPOWER + EnvirOffAxisGain + IntroOffAxisGain - PRLoss - FDR - CrossPolar
                    
'noise for receiver BW in khz
            If PROPFQ < 250 Then
                GalacticNoise = 10 ^ ((52 - 23 * (Log(PROPFQ) / Log(10))) / 10)
                
                If C_Constant = 52 Then 'galactic noise
                    ManMadeNoise = 0
                Else  'man-made noise
                    ManMadeNoise = 10 ^ ((C_Constant - D_Constant * (Log(PROPFQ) / Log(10))) / 10)
                End If
            
                NoiseLevel = 10 * (Log(EnvirCurve) / Log(10)) - 144 _
                                + 10 * (Log(10 ^ (NoiseFigure / 10) - 1 + GalacticNoise + ManMadeNoise) / Log(10))
            
            Else
                NoiseLevel = 10 * (Log(EnvirCurve) / Log(10)) - 144 + 10 * (Log(10 ^ (NoiseFigure / 10) - 1) / Log(10))
            End If
            
            If InterferenceLevel - NoiseLevel > RXThreshold Then
    
'update recordsets
UpdateIntroTXResultDB:
                
            Write_to_ResultDatabase
        
        End If  'if threshold exceeded
        
NextIntroTXRecord:
            EnvirAnalysisRecordset.MoveNext  'go to next record
    
            Update_StatusBar_Percent
            
        Loop
        
        IntroAnalysisRecordset.MoveNext  'go to next record
    
    Loop
    
    Call Draw_Multi_Symbol("pnt_introtx_results.error", RXSiteLayer, "Triangle", 9408399, 128, 4, 15)
    
    Analysis_Summary
    
    IntroTXINPerformed = True
    
    If IntroTXLink = True Then
        frmProjectFeatures.Delete_Feature_Data "pln_link_designator WHERE rrlntype = 2;" 'delete assosiated links
        IntroTXLink = False
        frmLinkCriteria.Form_Load
        frmLinkCriteria.chkTXInterferenceLinks.Value = 1
        frmLinkCriteria.Determine_Links
    End If
    
    TX_Analysis_ON  'call sub, set button bar

'Case 4  'elevation contour
'
'Set ContourRecordset = MyDatabase.OpenRecordset("pgn_Contour")
'
'LatNERad = 144000
'LatSWRad = 126000
'LongNERad = -270000
'LongSWRad = -288000
'LatSpacing = 60
'LongSpacing = 60
'
''calculate total lat and long in seconds
'TotalLatRad = Abs(LatNERad - LatSWRad)
'TotalLongRad = Abs(LongNERad - LongSWRad)
'
''calculate number of rows and columns
'NumberOfRows = (TotalLatRad / LatSpacing) + 1
'NumberOfColumns = (TotalLongRad / LongSpacing) + 1
'
'NumberOfGridPoints = NumberOfRows * NumberOfColumns
'
'HighElevation = -99999
'LowElevation = 99999
'
'    'initialize column start point
'    IntroLong = LongSWRad
'
'    For ColumnIndex = 1 To NumberOfColumns
'
''initialize row start point
'        IntroLat = LatNERad
'
'        For RowIndex = 1 To NumberOfRows
'
'
'            NegLong = -IntroLong ' TOPGET uses East negative.
'
'            Call TOPGET((IntroLat / 3600) / 57.29578, (NegLong / 3600) / 57.29578, WOTLType, 4, _
'                      TNAME, NAMLST, COUNT, 900, _
'                      DUMMY1, 0, WOTRER, SiteELEV)
'
'            'Call GetSiteElevation
'
''find greatest elevation
'            If SiteELEV > HighElevation Then
'                HighElevation = SiteELEV
'            Else
'                If RXELEV < LowElevation Then
'                    LowElevation = SiteELEV
'                End If
'            End If
'
''load database with signal strengths
''create database of grid points
''fill in vertex array
'            VertexArray(0, 0) = (IntroLong / 3600) + ((LongSpacing / 2) / 3600)
'            VertexArray(1, 0) = (IntroLong / 3600) - ((LongSpacing / 2) / 3600)
'            VertexArray(2, 0) = (IntroLong / 3600) - ((LongSpacing / 2) / 3600)
'            VertexArray(3, 0) = (IntroLong / 3600) + ((LongSpacing / 2) / 3600)
'
'            VertexArray(0, 1) = (IntroLat / 3600) + ((LatSpacing / 2) / 3600)
'            VertexArray(1, 1) = (IntroLat / 3600) + ((LatSpacing / 2) / 3600)
'            VertexArray(2, 1) = (IntroLat / 3600) - ((LatSpacing / 2) / 3600)
'            VertexArray(3, 1) = (IntroLat / 3600) - ((LatSpacing / 2) / 3600)
'
'            MyBlob = frmProjectFeatures.mapSelection.VertexArrayToBlob(VertexArray, 4)
'
''fill data fields
'            ContourRecordset.AddNew
'
''set equal for sylvan layer join - allows queries to only show intended links from tx's matching criteria
'            ContourRecordset("numvert") = 4
'            ContourRecordset("boundlowx") = VertexArray(1, 0)
'            ContourRecordset("boundlowy") = VertexArray(2, 1)
'            ContourRecordset("bounduppx") = VertexArray(0, 0)
'            ContourRecordset("bounduppy") = VertexArray(0, 1)
'            ContourRecordset("centerx") = IntroLong / 3600
'            ContourRecordset("centery") = IntroLat / 3600
'            ContourRecordset("elevation") = SiteELEV
'
'            'convert vertices to blob
'            ContourRecordset("VERTICES") = MyBlob
'
'            'update recordsets
'            ContourRecordset.Update
'
'            IntroLat = IntroLat - LatSpacing
'        Next
'
'        IntroLong = IntroLong + LongSpacing
'    Next
'
'SymbolCounter = 0
'
'If LowElevation < 0 Then
'    LowScale = 0
'Else
'    LowScale = LowElevation
'End If
'
''set symbol layer values
'frmProjectFeatures.mapSelection.LayerIndex = ElevationContourLayer
'For X = frmProjectFeatures.mapSelection.NumberOfSymbols - 1 To 2 Step -1 '0 and 1 are set
'    SymbolCounter = SymbolCounter + 1
'    frmProjectFeatures.mapSelection.SymbolIndex = X
'    frmProjectFeatures.mapSelection.LayerSymbolExpression = "pgn_contour.elevation"
'    frmProjectFeatures.mapSelection.LayerLinkType = 1
'    frmProjectFeatures.mapSelection.SymbolNumericRangeFrom = (HighElevation) - (((HighElevation - LowScale) / 6) * SymbolCounter)
'    frmProjectFeatures.mapSelection.SymbolNumericRangeTo = (HighElevation) - (((HighElevation - LowScale) / 6) * (SymbolCounter - 1))
'Next
'
'If LowElevation <= 0 Then
'    frmProjectFeatures.mapSelection.SymbolIndex = 2
'    frmProjectFeatures.mapSelection.SymbolNumericRangeFrom = 1
'End If
''delete layers not required
'If LowElevation > 0 Then
'    frmProjectFeatures.mapSelection.SymbolIndex = 1
'    frmProjectFeatures.mapSelection.DeleteSymbol
'    frmProjectFeatures.mapSelection.SymbolIndex = 0
'    frmProjectFeatures.mapSelection.DeleteSymbol
'Else
'    If LowElevation = 0 Then
'        frmProjectFeatures.mapSelection.SymbolIndex = 0
'        frmProjectFeatures.mapSelection.DeleteSymbol
'    End If
'End If
'
'ElevationContourLayerOn
'ElevationContour = True

Case 5
'set conduc and permit since doesn't matter for LOS calculation
    CONDUC = 0.002
    PERMIT = 15

    Set ContourRecordset = MyDatabase.OpenRecordset("pgn_Contour")
    
    If LOSTable = "pnt_introtxsites" Then
    'Load Intro DB to variables
        Set IntroAnalysisRecordset = MyDatabase.OpenRecordset _
                    ("SELECT * FROM pnt_introtxsites WHERE pnt_introtxsites.pointid=" + Str(LOSPointID))
        IntroDB_to_Variables (TX)  'call sub
    Else
        Set IntroAnalysisRecordset = MyDatabase.OpenRecordset _
                    ("SELECT * FROM pnt_introrxsites WHERE pnt_introrxsites.pointid=" + Str(LOSPointID))
        IntroDB_to_Variables (RX)  'call sub
    End If
    
    PROPFQ = IntroFreq
    
    'set major and minor axis for WOTL and TIREM call, wont vary per tx site so no reason to repeat inside tx loop
    Call Datum2Axis(Datum, spheroid, MJAxis, MNAxis, Flat, ecc, ERRCODE)
    
    EnvirAntennaHt = MobileAntHt
    
    'calculate total lat and long in seconds
    TotalLatRad = Abs(LatNERad - LatSWRad)
    TotalLongRad = Abs(LongNERad - LongSWRad)
        
    'calculate number of rows and columns
    NumberOfRows = (TotalLatRad / LatSpacing) + 1
    NumberOfColumns = (TotalLongRad / LongSpacing) + 1
    
    NumberOfGridPoints = NumberOfRows * NumberOfColumns
        
    'initialize column start point
    EnvirLongRad = LongSWRad
    
    AnalysisCounter = 0
    
        For ColumnIndex = 1 To NumberOfColumns
        
            AnalysisCounter = AnalysisCounter + 1
    
    'initialize row start point
            EnvirLatRad = LatNERad
            
            For RowIndex = 1 To NumberOfRows
                
            ' get the overall path length
                CalculatePathLength
                
                If PTHLEN > Circle_Radius Then
                    GoTo incrementrecord
                End If
                
'redim the elevation and distance arrays to the maximum allowed
                ReDim HPRFL(MXNELV)
                ReDim XPRFL(MXNELV)
                
'extract the path profile            Call GetPathProfile
                Call PRFILE(EnvirLatRad, EnvirLongRad, IntroLatRad, IntroLongRad, SPACNG, MJAxis, Flat, _
                            Datum, TOPFIL, INTERP, ERROPT, DELELV, _
                            MXNELV, XPRFL(1), HPRFL(1), NUMELV, PRFERR)
                
                If PRFERR = 1 Then
                    Mode = "DATA"  'no data
                    GoTo patherror
                End If
    
    'make the arrays so they only contain the actual number of returned points to avoid emptys
                ReDim Preserve HPRFL(NUMELV)
                ReDim Preserve XPRFL(NUMELV)
    
                Call CalculatePropagationLoss(REFRAC, PERMIT, CONDUC, Mode, PRLoss, FSPLSS, _
                                              ALPHAE, BETAE, HORZTX, HORZRX, _
                                              TXANG, RXANG, THET00, TOTDIF, TOTTRO, ABLOSS)
                
                
                If PRLoss = 0 Then
                    Mode = "DATA"  'no data
                End If
    
patherror:
                If Mode = "LOS " Or Mode = "DATA" Then
                
                    IntroLongDD = EnvirLongRad * 57.29578
                    IntroLatDD = EnvirLatRad * 57.29578
    'load database with signal strengths
    'create database of grid points
    'fill in vertex array
                    VertexArray(0, 0) = IntroLongDD + LongSpacingDD
                    VertexArray(1, 0) = IntroLongDD - LongSpacingDD
                    VertexArray(2, 0) = IntroLongDD - LongSpacingDD
                    VertexArray(3, 0) = IntroLongDD + LongSpacingDD
                
                    VertexArray(0, 1) = IntroLatDD + LatSpacingDD
                    VertexArray(1, 1) = IntroLatDD + LatSpacingDD
                    VertexArray(2, 1) = IntroLatDD - LatSpacingDD
                    VertexArray(3, 1) = IntroLatDD - LatSpacingDD
                
                    MyBlob = frmProjectFeatures.mapSelection.VertexArrayToBlob(VertexArray, 4)
    
    'fill data fields
                    ContourRecordset.AddNew
    
    'set equal for sylvan layer join - allows queries to only show intended links from tx's matching criteria
                    ContourRecordset("numvert") = 4
                    ContourRecordset("boundlowx") = VertexArray(1, 0)
                    ContourRecordset("boundlowy") = VertexArray(2, 1)
                    ContourRecordset("bounduppx") = VertexArray(0, 0)
                    ContourRecordset("bounduppy") = VertexArray(0, 1)
                    ContourRecordset("centerx") = IntroLongDD
                    ContourRecordset("centery") = IntroLatDD
                    ContourRecordset("mode") = Mode
                
                'convert vertices to blob
                    ContourRecordset("VERTICES") = MyBlob
    
                'update recordsets
                    ContourRecordset.Update
                    
                End If
incrementrecord:
                EnvirLatRad = EnvirLatRad - LatSpacing
                
            Next
            PercentComplete = (AnalysisCounter / NumberOfColumns) * 100
            frmStatusBar.ProgressBar1.Value = PercentComplete
            frmStatusBar.lblStatus.Caption = Str(PercentComplete) + " Percent Complete"
            DoEvents
            
            EnvirLongRad = EnvirLongRad + LongSpacing
        Next
    
        Dim ContourRS As Recordset
        
        Set ContourRS = MyDatabase.OpenRecordset("SELECT * FROM pgn_contour WHERE mode='DATA'")
        ContourRS.MoveLast
                
        frmProjectFeatures.mapSelection.LayerIndex = LOSContourLayer
        
        If DataMissing = True Then  'add symbol indicating data is missing
            If frmProjectFeatures.mapSelection.NumberOfSymbols = 1 Then  'first pass thru
                frmProjectFeatures.mapSelection.AddSymbol
                frmProjectFeatures.mapSelection.SymbolIndex = 1
                frmProjectFeatures.mapSelection.SymbolName = "Missing Topo Data"
                frmProjectFeatures.mapSelection.SymbolPolygonBorderColor = 65535  'yellow
                frmProjectFeatures.mapSelection.SymbolPolygonBorderType = "SOLID"
                frmProjectFeatures.mapSelection.SymbolPolygonColor = 65535  'yellow
                frmProjectFeatures.mapSelection.SymbolPolygonHatch = "SOLID"  'yellow
                frmProjectFeatures.mapSelection.LayerSymbolExpression = "pgn_contour.mode"  'yellow
                frmProjectFeatures.mapSelection.LayerLinkType = 0  'string
                frmProjectFeatures.mapSelection.SymbolStringEquals = "DATA"  'string
            
                frmProjectFeatures.mapSelection.SymbolIndex = 0
                frmProjectFeatures.mapSelection.SymbolName = "Line-of-Sight"
                frmProjectFeatures.mapSelection.LayerSymbolExpression = "pgn_contour.mode"  'yellow
                frmProjectFeatures.mapSelection.LayerLinkType = 0  'string
                frmProjectFeatures.mapSelection.SymbolStringEquals = "LOS "  'string
            End If
        Else
            If frmProjectFeatures.mapSelection.NumberOfSymbols > 1 Then  'first pass thru
                frmProjectFeatures.mapSelection.SymbolIndex = 1
                frmProjectFeatures.mapSelection.DeleteSymbol
                frmProjectFeatures.mapSelection.SymbolIndex = 0
                frmProjectFeatures.mapSelection.SymbolName = ""
            End If
        End If
    
    LOSContourLayerOn
    LOSContour = True
    
End Select

Dim handle As Long
            
ChangesSaved = False

Unload frmStatusBar

Set_Layer_Filter frmProjectFeatures  'call sub

'set status bar graphics
frmProjectFeatures.Update_Status_Bar

'update data display forms
frmProjectFeatures.Update_Data_Display_Forms  'call sub

frmProjectFeatures.mapSelection.DrawMap

DeleteOldAnalysis = False  're-initialize

Unload Me

errorhandler:
    Select Case Err.Number
    Case 94  'invalid use for Null(loading a null field to a variable)
        ErrorFound = "Missing required data"
        Err.Clear  'reset error flag
        ErrorCounter = ErrorCounter + 1
        Write_to_ResultDatabase
        GeneratedError = True
        Resume Next
        
'        Select Case AnalysisType
'        Case 1
'            GoTo UpdateResultDB
'        Case 2
'            GoTo UpdateSIResultDB
'        Case 3
'            GoTo UpdateIntroTXResultDB
'        End Select
    
    Case 317
        Resume Next
    Case 11
        Resume
    Case 3021  'attempted movelast in LOSContour with no records
        DataMissing = False
        Resume Next
    Case Is > 0
        MsgBox Err.Description, Err.Number
        Unload frmStatusBar
        Screen.MousePointer = 0
        tabAnalysisWizard.CurrTab = TabStepArray(1)
        Me.Show
        Exit Sub
'        Resume
    End Select
    
End Sub

Private Sub cmdNext_Click()

If CheckValidityOfCurrentTabInputs = 1 Then
    Exit Sub
End If

cmdBack.Enabled = True

TabStepCounter = TabStepCounter + 1

'if tirem not selected, -1's are inserted to indicate tabs should be skipped

If TabStepArray(TabStepCounter) = -1 Then
    TabStepCounter = TabStepCounter + 1
    
    If TabStepArray(TabStepCounter) = -1 Then
        TabStepCounter = TabStepCounter + 1
    End If
    
End If

tabAnalysisWizard.CurrTab = TabStepArray(TabStepCounter)

If TabStepCounter = NumberOfTabs Then
    cmdNext.Enabled = False
    cmdFinish.Enabled = True
End If

End Sub

Private Sub Form_Load()

Dim NoiseRecordset As Recordset

On Error GoTo errorhandler

PerformingLOS = False 'reset

tabAnalysisWizard.TabHeight = 1

NumberOfTabs = 10
ReDim TabStepArray(1 To NumberOfTabs)

'load defaults from ini file
Dim FormControl As Control
Dim Contents As Variant

'write to analysis.ini
Open App.Path + "\analysis.ini" For Input As #1 ' Open file to read.

tabAnalysisWizard.CurrTab = 0

For Each FormControl In frmAnalysisWizard
    If TypeOf FormControl Is TextBox Then
        Input #1, Contents
        FormControl.Text = Contents
    End If

    If TypeOf FormControl Is OptionButton Then
        Input #1, Contents
        FormControl.Value = Contents
    End If

    If TypeOf FormControl Is Label Then
        Input #1, Contents
        FormControl.Caption = Contents
    End If

    If TypeOf FormControl Is ComboBox Then
        Input #1, Contents
        If FormControl.Style = 2 Then  'dropdown list
            FormControl.ListIndex = Contents
        Else
            FormControl.Text = Contents
        End If
    End If

Next
Close #1    ' Close file.

If cboGroundTypes.ListIndex = 0 Then 'groundtype = none
    txtConductivity.Enabled = True
    txtPermittivity.Enabled = True
Else
    txtConductivity.Enabled = False
    txtPermittivity.Enabled = False
End If

TabStepCounter = 1

Select Case AnalysisType
Case 1  'env tx versus intro rx I/N
    Frame7.Height = 2415
    lblRXThreshold.Visible = False
    txtRXThreshold.Visible = False
    lblRXThresholdUnits.Visible = False
    
    frmAnalysisWizard.Icon = frmProjectFeatures.ImageList1.ListImages(40).Picture
    
    Set NoiseRecordset = MyDatabase.OpenRecordset("SELECT Min([freq(mhz)]) AS [LowFreq]" _
                        + " FROM pnt_introrxsites;")

    If NoiseRecordset("lowfreq") <= 250 Then 'consider noise
        NoiseAnalysis = True
        NumberOfTabs = 7
        ReDim TabStepArray(1 To NumberOfTabs)
        TabStepArray(6) = 9
        TabStepArray(7) = 5
        
        tabAnalysisWizard.CurrTab = 9  'noise
    
    Else
        NoiseAnalysis = False
        NumberOfTabs = 6
        ReDim TabStepArray(1 To NumberOfTabs)
        TabStepArray(6) = 5
    End If
    
    TabStepArray(1) = 3
    TabStepArray(2) = 2
    TabStepArray(3) = 1
    TabStepArray(4) = 8
    TabStepArray(5) = 4

Case 2  'env tx versus intro rx S/I
    Frame7.Height = 2415
    lblRXThreshold.Visible = False
    txtRXThreshold.Visible = False
    lblRXThresholdUnits.Visible = False

    frmAnalysisWizard.Icon = frmProjectFeatures.ImageList1.ListImages(41).Picture

'load currtab=7
    tabAnalysisWizard.CurrTab = 7

    lstIntroTX.Clear
    lstIntroTX.Clear

    Set IntroTXRecordset = MyDatabase.OpenRecordset("pnt_introtxsites")
    Set IntroRXRecordset = MyDatabase.OpenRecordset("pnt_introrxsites")

    IntroTXRecordset.MoveLast
    IntroTXRecordset.MoveFirst

    IntroRXRecordset.MoveLast
    IntroRXRecordset.MoveFirst

    For X = 1 To IntroTXRecordset.RecordCount
        lstIntroTX.AddItem IntroTXRecordset("name")
        IntroTXRecordset.MoveNext
    Next

    For X = 1 To IntroRXRecordset.RecordCount
        lstIntroRX.AddItem IntroRXRecordset("name")
        IntroRXRecordset.MoveNext
    Next

    Data1.DatabaseName = App.Path + "\Contours.mdb"
    Data1.RecordSource = "introlinks"
    Data1.Refresh
    
    Data1.Recordset.MoveLast
    
    If Data1.Recordset.RecordCount = 0 Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
    End If
    
    Dim FieldIndex As Integer
    
    For FieldIndex = 0 To dbgIntroLinks.Columns.COUNT - 1
        dbgIntroLinks.Columns.Remove (0)
    Next

    For FieldIndex = 0 To 1  'only want first 2 fields
        dbgIntroLinks.Columns.Add (FieldIndex)
        dbgIntroLinks.Columns(FieldIndex).Caption = Data1.Recordset.Fields(FieldIndex).Name
        dbgIntroLinks.Columns(FieldIndex).DataField = Data1.Recordset.Fields(FieldIndex).Name
        dbgIntroLinks.Columns(FieldIndex).Visible = True
    Next

    NoiseAnalysis = False
    NumberOfTabs = 7
    ReDim TabStepArray(1 To NumberOfTabs)
    
    TabStepArray(1) = 7
    TabStepArray(2) = 3
    TabStepArray(3) = 2
    TabStepArray(4) = 1
    TabStepArray(5) = 8
    TabStepArray(6) = 4
    TabStepArray(7) = 5

Case 3  'intro tx versus envir rx I/N
    
    frmAnalysisWizard.Icon = frmProjectFeatures.ImageList1.ListImages(42).Picture

    Set NoiseRecordset = MyDatabase.OpenRecordset("SELECT Min([freq_min(mhz)]) AS [LowFreq]" _
                        + " FROM pnt_rxsites;")

    If NoiseRecordset("lowfreq") <= 250 Then 'consider noise
        NoiseAnalysis = True
        NumberOfTabs = 7
        ReDim TabStepArray(1 To NumberOfTabs)
        TabStepArray(6) = 9
        TabStepArray(7) = 5
        
        tabAnalysisWizard.CurrTab = 9  'noise
    
    Else
        NoiseAnalysis = False
        NumberOfTabs = 6
        ReDim TabStepArray(1 To NumberOfTabs)
        TabStepArray(6) = 5
    End If
    
    TabStepArray(1) = 3
    TabStepArray(2) = 2
    TabStepArray(3) = 1
    TabStepArray(4) = 8
    TabStepArray(5) = 4

Case 4  'elevations
    'disable inputs that are not required
    tabAnalysisWizard.CurrTab = 2
    lblHumidity.Enabled = False
    txtHumidity.Enabled = False
    Label11.Enabled = False
    cboGroundTypes.Enabled = False
    lblConductivity.Enabled = False
    txtConductivity.Enabled = False
    Label16.Enabled = False
    txtPermittivity.Enabled = False
        
    NumberOfTabs = 3
    ReDim TabStepArray(1 To NumberOfTabs)
    
    TabStepArray(1) = 2
    TabStepArray(2) = 4
    TabStepArray(3) = 16

Case 5  'LOS contour
    
    frmAnalysisWizard.Icon = frmProjectFeatures.ImageList1.ListImages(44).Picture
    
    NumberOfTabs = 3
    ReDim TabStepArray(1 To NumberOfTabs)
    
    TabStepArray(1) = 2
    TabStepArray(2) = 1
    TabStepArray(3) = 6

End Select
    
'set required tabs for prop model selected
Select Case InterPropagationModel
Case 1  'tirem
    optInterTIREM_Click
Case 2  'sem
    optInterSEM_Click
Case 3  'free space
    optInterFreeSpace_Click
End Select

tabAnalysisWizard.CurrTab = TabStepArray(TabStepCounter)

'center form on screen
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

Me.Show

cmdBack.Enabled = False
cmdFinish.Enabled = False

If DesireHelpPrompt = True Then
    HelpPrompt "Navigate through the GATE Analysis Wizard using the Next/Back buttons.  Enter the information required on each screen.  " _
                + "Click on the Finish button to start the analysis." + Chr(13) + Chr(10) + Chr(13) + Chr(10) _
                + "Note:  it is suggested that the Maintenance Wizard be run be populate missing environmental data prior to performing an analysis."
End If

errorhandler:
    Select Case Err.Number
    Case 3021  'no record selected (data1.recordset.movelast)
        Resume Next
    Case Is > 0
        MsgBox Err.Description
        Exit Sub
    End Select
    
End Sub

Public Sub CalculatePathLength()
 
'calculate the path length and the bearings
Call CalcNGSInv(EnvirLatRad, EnvirLongRad, IntroLatRad, IntroLongRad, Datum, PTHLEN, BearEI, BearIE)

EnvirLongRad = -EnvirLongRad
IntroLongRad = -IntroLongRad

BearEI_deg = BearEI * 57.29578
BearIE_deg = BearIE * 57.29578
    
End Sub

Public Sub GetPathProfile()

Call PRFILE(EnvirLatRad, EnvirLongRad, IntroLatRad, IntroLongRad, SPACNG, MJAxis, Flat, _
                Datum, TOPFIL, INTERP, ERROPT, DELELV, _
                MXNELV, XPRFL(1), HPRFL(1), NUMELV, PRFERR)
                
End Sub

Public Sub CalculatePropagationLoss(REFRAC As Single, PERMIT As Single, CONDUC As Single, Mode As String, PRLoss As Single, FSPLSS As Single, ALPHAE As Single, BETAE As Single, HORZTX As Long, HORZRX As Long, _
                              TXANG As Single, RXANG As Single, THET00 As Single, TOTDIF As Single, TOTTRO As Single, ABLOSS As Single)

On Error GoTo errorhandler

EXTNSN = 0 ' False
        
'set the last point in the distance array to the GCD to avoid errors in path profile
XPRFL(NUMELV) = PTHLEN
'set height to rx site elevation
HPRFL(NUMELV) = SiteElevation  'intro site elevation
'set environmental site elevation

If CalculatingDesSignal = True Then
    HPRFL(1) = EnvirELEV
End If

Call NSubS2(HPRFL(1), SiteElevation, EnvirAntennaHt, IntroAntennaHt, _
             SeaRefract, REFRAC)

Call TiremAnalysis(EnvirAntennaHt, IntroAntennaHt, PROPFQ, NUMELV, HPRFL(1), XPRFL(1), _
            EXTNSN, REFRAC, CONDUC, PERMIT, HUMID, POLARZ, _
            VRSION, Mode, PRLoss, FSPLSS, TOTTRO, TOTDIF, ABLOSS, _
            THET00, TXANG, RXANG, ALPHAE, BETAE, HORZTX, HORZRX)
      
PropMode = "TIREM_" + Mode

errorhandler:
    If Err.Number > 0 Then
        Resume Next
    End If
    
End Sub

Public Function CheckValidityOfCurrentTabInputs() As Integer
On Error GoTo errorhandler

CheckValidityOfCurrentTabInputs = 0

Select Case tabAnalysisWizard.CurrTab
Case 0  'tab 1
'tab 1 inputs loaded to variables
        LatNERad = txtNorthLat.Text * 3600
        LatSWRad = txtSouthLat.Text * 3600
        LongNERad = txtEastLong.Text * 3600
        LongSWRad = txtWestLong.Text * 3600
Case 1
'tab 2 inputs loaded to variables
    If IsNumeric(cboSpacing.Text) And Val(cboSpacing.Text) > 0 Then
        Spacing = Val(cboSpacing.Text)
    Else
        MsgBox "Spacing must be a number greater than 0.", vbExclamation, "Warning"
        cboSpacing.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If
   
    SPACNG = Spacing * 30  'convert to meters
    INTERP = cboInterpolation.ItemData(cboInterpolation.ListIndex)
    Datum = cboDatumCode.ListIndex

Case 2
'tab 3 inputs loaded to variables
    If IsNumeric(txtHumidity.Text) And Val(txtHumidity.Text) > 0 Then
        HUMID = Val(txtHumidity.Text)
    Else
        MsgBox "Humidity must be a number greater than 0.", vbExclamation, "Warning"
        txtHumidity.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If
    
    If IsNumeric(txtSeaLevelRefractivity.Text) And Val(txtSeaLevelRefractivity.Text) > 0 Then
        SeaRefract = Val(txtSeaLevelRefractivity.Text)
    Else
        MsgBox "Sea-Level Refractivity must be a number greater than 0.", vbExclamation, "Warning"
        txtSeaLevelRefractivity.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If
    
    If IsNumeric(txtConductivity.Text) And Val(txtConductivity.Text) > 0 Then
        CONDUC = txtConductivity.Text
    Else
        MsgBox "Conductivity must be a number greater than 0.", vbExclamation, "Warning"
        txtConductivity.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If
    
    If IsNumeric(txtPermittivity.Text) And Val(txtPermittivity.Text) > 0 Then
        PERMIT = txtPermittivity.Text
    Else
        MsgBox "Pemittivity must be a number greater than 0.", vbExclamation, "Warning"
        txtPermittivity.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If
    
Case 4
'tab 5 inputs loaded to variables
    If IsNumeric(txtEnvirRollOff.Text) And Val(txtEnvirRollOff.Text) > 0 Then
        EnvirRollOff = txtEnvirRollOff.Text
    Else
        MsgBox "Roll-off must be a number greater than 0.", vbExclamation, "Warning"
        txtEnvirRollOff.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If
    
Case 5
'tab 6 inputs loaded to variables
    If IsNumeric(txtMaxDistance.Text) And Val(txtMaxDistance.Text) > 0 Then
        Select Case Trim(lblMaxAllowableUnits)
        Case "km"
            MaxDistance = txtMaxDistance.Text
        Case "st. mi."
            MaxDistance = Val(txtMaxDistance.Text) * 1.609344
        Case "nmi"
            MaxDistance = Val(txtMaxDistance.Text) * 1.852
        End Select
        
    Else
        MsgBox "Maximum allowable distance must be a number greater than 0.", vbExclamation, "Warning"
        txtMaxDistance.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If
    
    If IsNumeric(txtMaxFDR.Text) And Val(txtMaxFDR.Text) > 0 Then
        MaxAllowableFDR = txtMaxFDR.Text
    Else
        MsgBox "Maximum FDR must be a number greater than 0.", vbExclamation, "Warning"
        txtMaxFDR.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If
    
    If IsNumeric(txtMinDistance.Text) And Val(txtMinDistance.Text) > 0 Then
        
        Select Case Trim(lblClosestApproachUnits)
        Case "km"
            MinDistanceMobiles = txtMinDistance.Text
        Case "st. mi."
            MinDistanceMobiles = Val(txtMinDistance.Text) * 1.609344
        Case "nmi"
            MinDistanceMobiles = Val(txtMinDistance.Text) * 1.852
        End Select
        
    Else
        MsgBox "Minimum Distance must be a number greater than 0.", vbExclamation, "Warning"
        txtMinDistance.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If

    If AnalysisType = 3 Then  'introtx vs envirrx
        
        If IsNumeric(txtRXThreshold.Text) Then
            RXThreshold = txtRXThreshold.Text
        Else
            MsgBox "Receiver Threshold must be numeric.", vbExclamation, "Warning"
            txtRXThreshold.SetFocus
            CheckValidityOfCurrentTabInputs = 1
            Exit Function
        End If
    
    End If
    
Case 6
    Dim EndPointLat As Single
    Dim EndPointLong As Single
    Dim BackAzimuth As Single
    
    LatSpacing = (Val(cboLatSpacing.Text) / 3600) / 57.29578
    LongSpacing = (Val(cboLongSpacing.Text) / 3600) / 57.29578
    LongSpacingDD = (LongSpacing / 2) * 57.29578
    LatSpacingDD = (LatSpacing / 2) * 57.29578

    If IsNumeric(txtRadius.Text) And Val(txtRadius.Text) > 0 Then
        
        Select Case Trim(lblUnits)  'convert to meters
        Case "km"
            Circle_Radius = Val(txtRadius.Text) * 1000
        Case "st. mi."
            Circle_Radius = Val(txtRadius.Text) * 1609.344
        Case "nmi"
            Circle_Radius = Val(txtRadius.Text) * 1852
        End Select
        
    Else
        MsgBox "Radius must be a number greater than 0.", vbExclamation, "Warning"
        txtRadius.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If
        
    Call CalcNGSFor((PositionY / 57.2958), (PositionX / 57.2958), Datum, Circle_Radius, 0, EndPointLat, EndPointLong, BackAzimuth)
    LatNERad = EndPointLat
        
    Call CalcNGSFor((PositionY / 57.2958), (PositionX / 57.2958), Datum, Circle_Radius, 1.571, EndPointLat, EndPointLong, BackAzimuth)
    LongNERad = EndPointLong
    
    Call CalcNGSFor((PositionY / 57.2958), (PositionX / 57.2958), Datum, Circle_Radius, 3.1416, EndPointLat, EndPointLong, BackAzimuth)
    LatSWRad = EndPointLat
    
    Call CalcNGSFor((PositionY / 57.2958), (PositionX / 57.2958), Datum, Circle_Radius, 4.7124, EndPointLat, EndPointLong, BackAzimuth)
    LongSWRad = EndPointLong
    
    If IsNumeric(txtMobileAntHt.Text) Then
        If UCase(lblAntHtUnits) = "FEET" Then  'convert to meters
            MobileAntHt = Val(txtMobileAntHt.Text) * 0.304800609
        Else
            MobileAntHt = Val(txtMobileAntHt.Text)
        End If
                
    Else
        MsgBox "Mobile antenna height must be a number.", vbExclamation, "Warning"
        txtMobileAntHt.SetFocus
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If

Case 7
'tab 8 inputs loaded to variables
    If Data1.Recordset.RecordCount = 0 Then
        MsgBox "You must define a desired introduced link.", vbExclamation, "Warning"
        CheckValidityOfCurrentTabInputs = 1
        Exit Function
    End If

'update tx and rx ids
    Dim SQLQuery As String
    
    SQLQuery = "UPDATE introlinks,pnt_introtxsites SET introlinks.tx_id = pnt_introtxsites.pointid WHERE introlinks.tx_name = pnt_introtxsites.name;"
    MyDatabase.Execute SQLQuery
    
    SQLQuery = "UPDATE introlinks,pnt_introrxsites SET introlinks.rx_id = pnt_introrxsites.pointid WHERE introlinks.rx_name = pnt_introrxsites.name;"
    MyDatabase.Execute SQLQuery

Case 8
'tab 9 - harmonics
    If optHarmonic.Value = True Then
        HarmonicOrder = Val(cboHarmonicOrder.Text)
        If IsNumeric(txtHarmonicAttenuation.Text) And Val(txtHarmonicAttenuation.Text) > 0 Then
            HarmonicAttenuation = Val(txtHarmonicAttenuation.Text)
        Else
            MsgBox "Harmonic attenuation must be a numeric greater than 0.", vbExclamation, "Warning"
            CheckValidityOfCurrentTabInputs = 1
            Exit Function
        End If
    
        optQuickFDR = True  'must apply FDR if harmonics considered
        
    Else
        HarmonicOrder = 1
    End If

End Select

errorhandler:
    Select Case Err.Number
    Case 13
        Resume Next
    Case Is > 0
        MsgBox Err.Description
        
    End Select
    
    
End Function

Private Sub Form_Unload(Cancel As Integer)
frmProjectFeatures.Show
End Sub

Private Sub lblAntHtUnits_Click()
    Set UnitLabel = lblAntHtUnits
    PopupMenu mnuHeightUnits
End Sub

Private Sub lblClosestApproachUnits_Click()
    Set UnitLabel = lblClosestApproachUnits
    PopupMenu mnuDistanceUnits
End Sub

Private Sub lblMaxAllowableUnits_Click()
    Set UnitLabel = lblMaxAllowableUnits
    PopupMenu mnuDistanceUnits
End Sub

Private Sub lblUnits_Click()
    Set UnitLabel = lblUnits
    PopupMenu mnuDistanceUnits
End Sub

Private Sub mnuFeet_Click()
    UnitLabel.Caption = "feet"
End Sub

Private Sub mnuKM_Click()
    UnitLabel.Caption = "km"
End Sub

Private Sub mnuMeters_Click()
    UnitLabel.Caption = "meters"
End Sub

Private Sub mnuNauticalMI_Click()
    UnitLabel.Caption = "nmi"
End Sub

Private Sub mnuStatuteMI_Click()
    UnitLabel.Caption = "st. mi."
End Sub

Private Sub optBusiness_Click()
    C_Constant = 76.8
    D_Constant = 27.7
    
    Dim Series As Object
    
    For Each Series In VtChart1.Plot.SeriesCollection
        Series.Pen.VtColor.Set 96, 96, 96
    Next

    VtChart1.Plot.SeriesCollection.Item(1).Pen.VtColor.Set 255, 0, 0
    
End Sub

Private Sub optGalactic_Click()
    
    C_Constant = 52
    D_Constant = 23
    
    Dim Series As Object
    
    For Each Series In VtChart1.Plot.SeriesCollection
        Series.Pen.VtColor.Set 96, 96, 96
    Next

    VtChart1.Plot.SeriesCollection.Item(9).Pen.VtColor.Set 255, 0, 0
    
End Sub

Private Sub optHarmonic_Click()
    fraHarmonicInfo.Enabled = True
End Sub

Private Sub optInterFreeSpace_Click()
Select Case AnalysisType
Case 1, 3 'env tx versus intro rx I/N
'insert -1 so tirem tabs are skipped
    TabStepArray(2) = -1
    TabStepArray(3) = -1
  
Case 2
    TabStepArray(3) = -1
    TabStepArray(4) = -1
End Select

InterPropagationModel = 3  'free space

End Sub

Private Sub optInterSEM_Click()
Select Case AnalysisType
Case 1, 3 'env tx versus intro rx I/N and vice-versa
    TabStepArray(2) = 2
    TabStepArray(3) = -1

Case 2  'env tx versus intro rx S/I
    TabStepArray(3) = 2
    TabStepArray(4) = -1

End Select

InterPropagationModel = 2  'SEM

End Sub

Private Sub optInterTIREM_Click()
Select Case AnalysisType
Case 1, 3 'env tx versus intro rx and vice-versa I/N
    TabStepArray(2) = 2
    TabStepArray(3) = 1
Case 2 'env tx versus intro rx and vice-versa S/I
    TabStepArray(3) = 2
    TabStepArray(4) = 1
End Select

InterPropagationModel = 1  'tirem

End Sub


Private Sub optNESWCorner_Click()

txtNorthLat.Visible = True
txtSouthLat.Visible = True
txtEastLong.Visible = True
txtWestLong.Visible = True
lbl3DNELatitude.Visible = True
lbl3DNELongitude.Visible = True
lbl3DSWLatitude.Visible = True
lbl3DSWLongitude.Visible = True

End Sub

Private Sub optNoFDR_Click()
If HarmonicOrder > 1 Then  'must apply FDR if harmonics considered
    MsgBox "FDR must be applied if harmonics are considered.", vbExclamation + vbOKOnly, "Input Violation"
    optQuickFDR.Value = True
Else
    QuickFDR = False
    fraBandwidth.Enabled = False
End If
End Sub

Private Sub optNoHarmonic_Click()
    fraHarmonicInfo.Enabled = False
    HarmonicOrder = 1
End Sub

Private Sub optQuickFDR_Click()
    QuickFDR = True
    fraBandwidth.Enabled = True
End Sub

Private Sub optQuietRural_Click()
    C_Constant = 53
    D_Constant = 28.6
    
    Dim Series As Object
    
    For Each Series In VtChart1.Plot.SeriesCollection
        Series.Pen.VtColor.Set 96, 96, 96
    Next

    VtChart1.Plot.SeriesCollection.Item(7).Pen.VtColor.Set 255, 0, 0
    
End Sub

Public Sub GetSiteElevation(Lat As Single, Lon As Single)
   
 ' Use WotRet.dll
   
   NegLong = -Lon ' TOPGET uses East negative.
 
   Call TOPGET(Lat, NegLong, WOTLType, INTERP, _
                  TNAME, NAMLST, COUNT, CLng(SPACNG), _
                  DUMMY1, Datum, WOTRER, SiteELEV)
                     
End Sub

Private Sub optResidential_Click()
    C_Constant = 72.5
    D_Constant = 27.7
    
    Dim Series As Object
    
    For Each Series In VtChart1.Plot.SeriesCollection
        Series.Pen.VtColor.Set 96, 96, 96
    Next

    VtChart1.Plot.SeriesCollection.Item(3).Pen.VtColor.Set 255, 0, 0
    
End Sub

Private Sub optRural_Click()
    C_Constant = 67.2
    D_Constant = 27.7
    
    Dim Series As Object
    
    For Each Series In VtChart1.Plot.SeriesCollection
        Series.Pen.VtColor.Set 96, 96, 96
    Next

    VtChart1.Plot.SeriesCollection.Item(5).Pen.VtColor.Set 255, 0, 0
    
End Sub

Public Function Range_to_Single_Freq(FreqMin As Single, FreqMax As Single, SingleFreq As Double) As Double

'decide what environmental frequency to use
        If FreqMax = 0 Or FreqMax < FreqMin Then
            Range_to_Single_Freq = FreqMin
        Else
            If SingleFreq <= FreqMax And SingleFreq >= FreqMin Then
                Range_to_Single_Freq = SingleFreq  'assume on-tune
            Else
                If Abs(SingleFreq - FreqMax) < Abs(SingleFreq - FreqMin) Then
                    Range_to_Single_Freq = FreqMax
                Else
                    Range_to_Single_Freq = FreqMin
                End If
            End If
        End If

End Function

Public Sub Quick_FDR()
   Dim Transmitter_Frequency_In_Khz As Double
   Dim Receiver_Frequency_In_Khz As Double
   Dim Transmitter_3dB_Bandwidth_In_Khz As Double
   Dim Receiver_3dB_Bandwidth_In_Khz As Double
   Dim Left_Fall_Off_Slope_In_dB_Per_Decade As Double
   Dim Left_Selectivity_Slope_In_dB_Per_Decade As Double
   
   Call CALCQUICKFDR( _
      Transmitter_Frequency_In_Khz, _
      Receiver_Frequency_In_Khz, _
      Transmitter_3dB_Bandwidth_In_Khz, _
      Receiver_3dB_Bandwidth_In_Khz, _
      Left_Fall_Off_Slope_In_dB_Per_Decade, _
      Left_Selectivity_Slope_In_dB_Per_Decade, _
      FDR, OTR, OFR, Error)

End Sub

Public Sub Calculate_CrossPolar()

    'cross-polarization
            Select Case Right(Trim(IntroPolar), 1)
            Case "V"
                If EnvirPolar = "H" Then
                    CrossPolar = 6
                Else
                    CrossPolar = 0
                End If
            Case "H"
                If EnvirPolar = "V" Then
                    CrossPolar = 6
                Else
                    CrossPolar = 0
                End If
            
            Case "L"
                If EnvirPolar = "R" Then
                    CrossPolar = 6
                Else
                    CrossPolar = 0
                End If
            
            Case "R"
                If EnvirPolar = "L" Then
                    CrossPolar = 6
                Else
                    CrossPolar = 0
                End If
            
            Case Else
                CrossPolar = 0
            
            End Select

End Sub
            
Public Sub Set_Intro_Site_Elevation(ElevationOption As Integer)
    Select Case ElevationOption
    Case 1
        Call GetSiteElevation(IntroLatRad, IntroLongRad)
        SiteElevation = SiteELEV
     
    Case 3  'max of user entered or topo file
        Call GetSiteElevation(IntroLatRad, IntroLongRad)
        
        If SiteElevation < SiteELEV Then
            SiteElevation = SiteELEV
        End If
    End Select

End Sub

Public Sub Calculate_PathLength_Mobiles()
   
    PTHLENMobile = PTHLEN - (EnvirRadiusofMobility + IntroRadiusofMobility) * 1000
    
    If PTHLENMobile < MinDistanceMobiles * 1000 Then
        PTHLENMobile = MinDistanceMobiles * 1000
    End If
            
End Sub

Public Sub IntroDB_to_Variables(IntroType As Integer)
    
    IntroPolar = Right(Trim(IntroAnalysisRecordset("ant_pol")), 1)
   
    If IntroType = RX Then  'extract rx uniq info
        RXThreshold = IntroAnalysisRecordset("threshold")
        NoiseFigure = IntroAnalysisRecordset("noise_fig")
    
    Else  'tx uniq info
        If IntroPolar <> "H" Then
            POLARZ = "V   "  'input to tirem and sem - all other polariz will have some V component (also worst-case)
        Else
            POLARZ = "H   "
        End If
    
        TPOWER = IntroAnalysisRecordset("Power(dBm)")
    
    End If
    
    IntroID = IntroAnalysisRecordset("pointid")
    IntroName = IntroAnalysisRecordset("name")
    IntroLat = IntroAnalysisRecordset("ycoord")
    IntroLong = IntroAnalysisRecordset("xcoord")
    IntroLatRad = IntroLat / 57.29578
    IntroLongRad = IntroLong / 57.29578
    IntroAntennaHt = IntroAnalysisRecordset("ANT_HGT(m)")
    IntroFreq = IntroAnalysisRecordset("FREQ(MHz)")
    IntroCurve = IntroAnalysisRecordset("bdw_3db(khz)")
    IntroRollOff = IntroAnalysisRecordset("rolloff(db/dec)")
    IntroMBGain = IntroAnalysisRecordset("ANT_GAIN(dBi)")
    
    If Not IsNull(IntroAnalysisRecordset("site_elev(m)")) Then
        SiteElevation = IntroAnalysisRecordset("site_elev(m)")
    Else
        SiteElevation = 0
    End If
    
    If AnalysisType = 5 Then 'los, ingore mobile
        Call Set_Intro_Site_Elevation(IntroAnalysisRecordset("elev_option")) 'call sub
    Else
        If Trim(IntroAnalysisRecordset("FMICDE")) = "M" Then
            IntroRadiusofMobility = IntroAnalysisRecordset("radmob(km)")
        Else
            IntroRadiusofMobility = 0
    'check site elevation option for introduced equipment, if fixed and prop model is TIREM
            If InterPropagationModel = 1 Then 'tirem
                Call Set_Intro_Site_Elevation(IntroAnalysisRecordset("elev_option")) 'call sub
            End If
        End If
    End If
    
End Sub

Public Function EnvirDB_to_Variables(EnvirType As Integer) As Integer
On Error GoTo errorhandler
    
    EnvirDB_to_Variables = 0
    
    EnvirLat = EnvirAnalysisRecordset("ycoord")
    EnvirLong = EnvirAnalysisRecordset("xcoord")
    EnvirLatRad = EnvirLat / 57.29578
    EnvirLongRad = EnvirLong / 57.29578
    EnvirFreqMin = EnvirAnalysisRecordset("freq_min(MHz)")
    EnvirAntennaHt = EnvirAnalysisRecordset("ant_hgt(m)")
    
    If EnvirType = RX Then 'extract rx uniq info
        If IsNull(EnvirAnalysisRecordset("ant_polcde")) Then
            EnvirPolar = ""
        Else
            EnvirPolar = Right(Trim(EnvirAnalysisRecordset("ant_polcde")), 1)
        End If
            
'noisefig
        NoiseFigure = EnvirAnalysisRecordset("noise_fig")
    Else  'extract tx uniq info
        If IsNull(EnvirAnalysisRecordset("ant_polcde")) Then
            POLARZ = "V   "
        Else
            EnvirPolar = Right(Trim(EnvirAnalysisRecordset("ant_polcde")), 1)
           
            If EnvirPolar <> "H" Then
                POLARZ = "V   "  'input to tirem and sem - all other polariz will have some V component (also worst-case)
            Else
                POLARZ = "H   "
            End If
        
        End If
    End If
    
    If IsNull(EnvirAnalysisRecordset("freq_max(MHz)")) Then
        EnvirFreqMax = 0
    Else
        EnvirFreqMax = EnvirAnalysisRecordset("freq_max(MHz)")
    End If
    
    If IsNull(EnvirAnalysisRecordset("rad_mob(km)")) Then
        EnvirRadiusofMobility = 0
    Else
        EnvirRadiusofMobility = EnvirAnalysisRecordset("rad_mob(km)")
        If EnvirRadiusofMobility < 0 Then
            EnvirDB_to_Variables = 1
            ErrorFound = "Radius of Mobility is negative"
            ErrorCounter = ErrorCounter + 1
        End If
    End If

    Exit Function

errorhandler:
    Select Case Err.Number
    Case 94  'invalid use of null
        EnvirDB_to_Variables = 1
        ErrorFound = "Missing required data"
        ErrorCounter = ErrorCounter + 1
        Err.Clear
        Exit Function
    Case Is > 0
        EnvirDB_to_Variables = 1
        ErrorFound = Err.Description
        ErrorCounter = ErrorCounter + 1
        Exit Function
    End Select
    
End Function

Public Sub Calculate_IntroRX_HarmonicAttenuation()

On Error GoTo errorhandler

    For X = 2 To HarmonicOrder
'if BW not null examine modulation, otherwise multiply by harmonic order
        If Not IsNull(EnvirAnalysisRecordset("Modulation")) Then
            Modulation = Trim(EnvirAnalysisRecordset("Modulation"))
            Select Case UCase(Left(Modulation, 1))
            Case "G"  'BPSK
                If Mid(Modulation, 2, 1) = "1" And (X = 2 Or X = 4) Then
                    HarmonicBandwidth = 0.001
                Else
                    HarmonicBandwidth = EnvirCurve
                End If
            Case "F"  'FM,FSK,MSK
                HarmonicBandwidth = EnvirCurve * X
            
            Case "H", "J", "R" 'SSB
                HarmonicBandwidth = EnvirCurve * Sqr(X)
            
            Case Else
                HarmonicBandwidth = EnvirCurve
            End Select
        Else
            HarmonicBandwidth = EnvirCurve * X
        End If
        
        HarmonicPropFreq = Range_to_Single_Freq(EnvirFreqMin * X, EnvirFreqMax * X, IntroFreq)
        
        Call CALCQUICKFDR(HarmonicPropFreq * 1000, IntroFreq * 1000, _
                    HarmonicBandwidth, IntroCurve, EnvirRollOff, _
                    IntroRollOff, FDRHarmonic, OTR, OFR, FDRError)

        If FDRError < 0 Then
            FDRHarmonic = MaxAllowableFDR
        End If

        If FDRHarmonic + HarmonicAttenuation < FDR Then
            FDR = FDRHarmonic + HarmonicAttenuation
            EnvirSingleFreq = HarmonicPropFreq
            FreqSep = Abs(EnvirSingleFreq - IntroFreq)
            Harmonic = X
        End If
    Next

errorhandler:
    Select Case Err.Number
    Case 11
        Resume
    Case Else
        Resume Next
    End Select
    
End Sub

Public Sub Calculate_IntroTX_HarmonicAttenuation()
On Error GoTo errorhandler

    For X = 2 To HarmonicOrder
        
        HarmonicPropFreq = IntroFreq * X
        
        EnvirSingleFreq = Range_to_Single_Freq(EnvirFreqMin, EnvirFreqMax, HarmonicPropFreq)

'if not null examine modulation, otherwise multiply by harmonic order
        Select Case UCase(Trim(Modulation))
        Case "BPSK"
                If X = 2 Or X = 4 Then
                    HarmonicBandwidth = 0.001
                Else
                    HarmonicBandwidth = IntroCurve
                End If
        Case "SSB"
            HarmonicBandwidth = IntroCurve * Sqr(X)
        
        Case "FM VOICE", "FSK", "MSK"
            HarmonicBandwidth = IntroCurve * X
            
        Case Else
            HarmonicBandwidth = IntroCurve
        
        End Select
        
        Call CALCQUICKFDR(HarmonicPropFreq * 1000, EnvirSingleFreq * 1000, _
                    HarmonicBandwidth, EnvirCurve, IntroRollOff, _
                    EnvirRollOff, FDRHarmonic, OTR, OFR, FDRError)
    
        If FDRError < 0 Then
            FDRHarmonic = MaxAllowableFDR
        End If

        If FDRHarmonic + HarmonicAttenuation < FDR Then
            FDR = FDRHarmonic + HarmonicAttenuation
            IntroFreq = HarmonicPropFreq
            FreqSep = Abs(EnvirSingleFreq - IntroFreq)
            Harmonic = X
        End If
    Next

errorhandler:
    Select Case Err.Number
    Case 11
        Resume
    Case Else
        Resume Next
    End Select
    
End Sub

Public Sub Calculate_Propagation_Loss()
'initialize error constants
    PRFERR = 0
        
'mod groundconst value for frequency
    If groundtype <> 0 Then
        Call CalcGrConst(PROPFQ, groundtype, PERMIT, CONDUC)
    End If

'determine mode of interference propagation
    Select Case InterPropagationModel
    
    Case 1 'TIREM
    
        If EnvirRadiusofMobility = 0 And IntroRadiusofMobility = 0 Then  'fixed/tirem
'redim the elevation and distance arrays to the maximum allowed
            ReDim HPRFL(MXNELV)
            ReDim XPRFL(MXNELV)

'extract the path profile            Call GetPathProfile
            Call PRFILE(EnvirLatRad, EnvirLongRad, IntroLatRad, IntroLongRad, SPACNG, MJAxis, Flat, _
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
            
        Else  'SEM
            Call NSubS2(0, 0, EnvirAntennaHt, IntroAntennaHt, _
                             SeaRefract, REFRAC)

            Call SEMDLL(EnvirAntennaHt, IntroAntennaHt, PROPFQ, 0, PTHLENMobile, _
                REFRAC, CONDUC, PERMIT, HUMID, POLARZ, _
                2, VRSION, SEMMODE, SEMPRLoss, SEMFSPLSS)

            PRLoss = SEMPRLoss
            PropMode = "SEM_" + SEMMODE
        End If
        
    Case 2  'SEM
        Call NSubS2(0, 0, EnvirAntennaHt, IntroAntennaHt, _
                    SeaRefract, REFRAC)
        
        Call SEMDLL(EnvirAntennaHt, IntroAntennaHt, PROPFQ, 0, PTHLENMobile, _
            REFRAC, CONDUC, PERMIT, HUMID, POLARZ, _
            2, VRSION, SEMMODE, SEMPRLoss, SEMFSPLSS)

        PRLoss = SEMPRLoss
        PropMode = "SEM_" + SEMMODE
    
    Case 3  'free space
        'check if LOS
        If (PTHLENMobile * 0.00062137119) <= Sqr(2 * IntroAntennaHt * 3.2808399) + Sqr(2 * EnvirAntennaHt * 3.2808399) Then
            SlantRange = Sqr((Abs(IntroAntennaHt - EnvirAntennaHt)) ^ 2 + PTHLENMobile ^ 2)
            PRLoss = 20 * (Log(PROPFQ) / Log(10)) + 20 * (Log(SlantRange) / Log(10)) - 27.5
            PropMode = "FREE_SPACE"
        Else
            PRLoss = 999
        End If
        
    End Select

End Sub

Public Sub Calculate_Envir_OffAxisGain()
    EnvirCrossPolar = False  'initialize
    
    EnvirMBGain = EnvirAnalysisRecordset("ant_gain(dbi)")

    If EnvirRadiusofMobility = 0 Then

     'determine degrees off-axis
        Select Case UCase(Right(Trim(EnvirAnalysisRecordset("ANTMOTCDE")), 1))
    
        Case "O", "S"
            EnvirOffAxisGain = EnvirMBGain
        
        Case "D", "H", "F"
            If Not IsNull(EnvirAnalysisRecordset("ANTAZMBEG")) Then
                
                EnvirPointingAngle = EnvirAnalysisRecordset("ANTAZMBEG")
            
                If IntroRadiusofMobility > 0 Then
                    
                    If PTHLEN <= IntroRadiusofMobility * 1000 Then 'overlap, use mb gain
                        
                        EnvirOffAxisGain = EnvirMBGain
                    
                    Else  'calculate off-axis angle
                        
                        Q = (IntroRadiusofMobility * 1000) / PTHLEN
                        DeltaAngle = Atn(Q / Sqr(-Q * Q + 1)) * 57.29578
                        
                        MinAngle = BearEI_deg - DeltaAngle
                        MaxAngle = BearEI_deg + DeltaAngle
                        
                        If EnvirPointingAngle >= MinAngle And EnvirPointingAngle <= MaxAngle Then
                            
                            EnvirOffAxisGain = EnvirMBGain
                        
                        Else
                            
                            MinOffAxis = Abs(EnvirPointingAngle - MinAngle)
                            
                            If MinOffAxis > 180 Then
                                MinOffAxis = 360 - MaxOffAxis
                            End If
                                
                            MaxOffAxis = Abs(EnvirPointingAngle - MaxAngle)
                            
                            If MaxOffAxis > 180 Then
                                MaxOffAxis = 360 - MaxOffAxis
                            End If
                                
                            If MaxOffAxis < MinOffAxis Then
                                EnvirDegOffAxis = MaxOffAxis
                            Else
                                EnvirDegOffAxis = MinOffAxis
                            End If
                        
                            Call Calculate_FixedOffAxis_Gain(EnvirMBGain, EnvirDegOffAxis, EnvirOffAxisGain)
                        
                        End If
                       
                    End If
                    
                Else
                
                    If Abs(EnvirPointingAngle - BearEI_deg) <= 180 Then
                        EnvirDegOffAxis = Abs(EnvirPointingAngle - BearEI_deg)
                    Else
                        EnvirDegOffAxis = 360 - Abs(EnvirPointingAngle - BearEI_deg)
                    End If
        
                    Call Calculate_FixedOffAxis_Gain(EnvirMBGain, EnvirDegOffAxis, EnvirOffAxisGain)
               
                End If
            
            Else
                EnvirOffAxisGain = EnvirMBGain
                If EnvirOffAxisGain > 10 Then
                    AnalysisComment = "Conservative Prediction - Directional Antenna w/o Pointing Angle, MB Gain used."
                End If
            End If
            
        Case "E", "N", "R", "T"
            EnvirOffAxisGain = -10
    
        Case Else
            EnvirOffAxisGain = EnvirMBGain
    
        End Select

    Else
        EnvirOffAxisGain = EnvirMBGain
                
        Select Case UCase(Right(Trim(EnvirAnalysisRecordset("ANTMOTCDE")), 1))
    
        Case "D", "H", "F", "E", "N", "R", "T"
            If EnvirOffAxisGain > 10 Then
                AnalysisComment = "Conservative Prediction - Mobile High Gain Antenna, MB Gain Used."
            End If
        End Select
    End If  'if transmitter is fixed

End Sub

Public Sub Write_to_ResultDatabase()
    InterferenceCounter = InterferenceCounter + 1
    
    ResultRecordset.AddNew
'write uniq entries based on analysis type
    Select Case AnalysisType
    Case 1
        ResultRecordset("TXID") = EnvirAnalysisRecordset("pointid")
        ResultRecordset("TX_Nomen") = EnvirAnalysisRecordset("nomen")
        ResultRecordset("TX_OPID") = EnvirAnalysisRecordset("OPID")
        
        If EnvirLat = 999 Then
            ResultRecordset("TXLat") = Null
            Exit Sub
        Else
            ResultRecordset("TXLat") = EnvirLat
        End If
        
        If EnvirLong = 999 Then
            ResultRecordset("TXLong") = Null
            Exit Sub
        Else
            ResultRecordset("TXLong") = EnvirLong
        End If
        
        If BearEI_deg = 999 Then
            ResultRecordset("bearing_TxRx(deg)") = Null
        Else
            ResultRecordset("bearing_TxRx(deg)") = BearEI_deg
        End If
        
        If EnvirOffAxisGain = 999 Then
            ResultRecordset("TXAntGain(dBi)") = Null
        Else
            ResultRecordset("TXAntGain(dBi)") = EnvirOffAxisGain
        End If
        
        ResultRecordset("RXID") = IntroID
        ResultRecordset("RX_Nomen") = IntroName
        ResultRecordset("RXLat") = IntroLat
        ResultRecordset("RXLong") = IntroLong
        
        If IntroOffAxisGain = 999 Then
            ResultRecordset("RXAntGain(dBi)") = Null
        Else
            ResultRecordset("RXAntGain(dBi)") = IntroOffAxisGain
        End If
        
        If NoiseLevel = 0 Then
            ResultRecordset("Noise(dBm)") = Null
            
            If InterferenceLevel = 999 Then
                ResultRecordset("I/N(dB)") = Null
                ResultRecordset("Delta_Threshold(dB)") = Null
                ResultRecordset("Interference(dBm)") = Null
            Else
                ResultRecordset("Interference(dBm)") = InterferenceLevel
            End If
            
        Else
            ResultRecordset("Noise(dBm)") = NoiseLevel
            
            If InterferenceLevel = 999 Then
                ResultRecordset("I/N(dB)") = Null
                ResultRecordset("Delta_Threshold(dB)") = Null
                ResultRecordset("Interference(dBm)") = Null
            Else
                ResultRecordset("Interference(dBm)") = InterferenceLevel
                ResultRecordset("I/N(dB)") = InterferenceLevel - NoiseLevel
                ResultRecordset("Delta_Threshold(dB)") = (InterferenceLevel - NoiseLevel) - RXThreshold
            End If
            
        End If
        
    Case 2
        ResultRecordset("TXID") = EnvirAnalysisRecordset("pointid")
        ResultRecordset("TX_Nomen") = EnvirAnalysisRecordset("nomen")
        ResultRecordset("TX_OPID") = EnvirAnalysisRecordset("OPID")
        
        If EnvirLat = 999 Then
            ResultRecordset("TXLat") = Null
            Exit Sub
        Else
            ResultRecordset("TXLat") = EnvirLat
        End If
        
        If EnvirLong = 999 Then
            ResultRecordset("TXLong") = Null
            Exit Sub
        Else
            ResultRecordset("TXLong") = EnvirLong
        End If
        
        If BearEI_deg = 999 Then
            ResultRecordset("bearing_TxRx(deg)") = Null
        Else
            ResultRecordset("bearing_TxRx(deg)") = BearEI_deg
        End If
    
        If EnvirOffAxisGain = 999 Then
            ResultRecordset("TXAntGain(dBi)") = Null
        Else
            ResultRecordset("TXAntGain(dBi)") = EnvirOffAxisGain
        End If
        
        ResultRecordset("RXID") = IntroID
        ResultRecordset("RX_Nomen") = IntroName
        ResultRecordset("RXLat") = IntroLat
        ResultRecordset("RXLong") = IntroLong
        
        If IntroOffAxisGain = 999 Then
            ResultRecordset("RXAntGain(dBi)") = Null
        Else
            ResultRecordset("RXAntGain(dBi)") = IntroOffAxisGain
        End If
        
        If InterferenceLevel = 999 Then
            ResultRecordset("S/I(dB)") = Null
            ResultRecordset("Delta_Threshold(dB)") = Null
            ResultRecordset("Interference(dBm)") = Null
        Else
            ResultRecordset("Interference(dBm)") = InterferenceLevel
            ResultRecordset("S/I(dB)") = SignalStrength - InterferenceLevel
            ResultRecordset("Delta_Threshold(dB)") = -((SignalStrength - InterferenceLevel) - RXThreshold)
        End If
             
        ResultRecordset("Signal(dBm)") = SignalStrength
    
    Case 3
        ResultRecordset("TXID") = IntroID
        ResultRecordset("TX_Nomen") = IntroName
        ResultRecordset("TXLat") = IntroLat
        ResultRecordset("TXLong") = IntroLong
        
        If BearIE_deg = 999 Then
            ResultRecordset("bearing_TxRx(deg)") = Null
        Else
            ResultRecordset("bearing_TxRx(deg)") = BearIE_deg
        End If
    
        If IntroOffAxisGain = 999 Then
            ResultRecordset("TXAntGain(dBi)") = Null
        Else
            ResultRecordset("TXAntGain(dBi)") = IntroOffAxisGain
        End If
        
        ResultRecordset("RXID") = EnvirAnalysisRecordset("pointid")
        ResultRecordset("RX_Nomen") = EnvirAnalysisRecordset("nomen")
        ResultRecordset("RX_OPID") = EnvirAnalysisRecordset("OPID")
        
        If EnvirLat = 999 Then
            ResultRecordset("RXLat") = Null
'            Exit Sub
        Else
            ResultRecordset("RXLat") = EnvirLat
        End If
        
        If EnvirLong = 999 Then
            ResultRecordset("RXLong") = Null
'            Exit Sub
        Else
            ResultRecordset("RXLong") = EnvirLong
        End If
        
        If EnvirOffAxisGain = 999 Then
            ResultRecordset("RXAntGain(dBi)") = Null
        Else
            ResultRecordset("RXAntGain(dBi)") = EnvirOffAxisGain
        End If
        
        If NoiseLevel = 0 Then
            ResultRecordset("Noise(dBm)") = Null
            
            If InterferenceLevel = 999 Then
                ResultRecordset("I/N(dB)") = Null
                ResultRecordset("Delta_Threshold(dB)") = Null
                ResultRecordset("Interference(dBm)") = Null
            Else
                ResultRecordset("Interference(dBm)") = InterferenceLevel
            End If
            
        Else
            ResultRecordset("Noise(dBm)") = NoiseLevel
            
            If InterferenceLevel = 999 Then
                ResultRecordset("I/N(dB)") = Null
                ResultRecordset("Delta_Threshold(dB)") = Null
                ResultRecordset("Interference(dBm)") = Null
            Else
                ResultRecordset("Interference(dBm)") = InterferenceLevel
                ResultRecordset("I/N(dB)") = InterferenceLevel - NoiseLevel
                ResultRecordset("Delta_Threshold(dB)") = (InterferenceLevel - NoiseLevel) - RXThreshold
            End If
            
        End If
        
    Case 5
        GoTo UpdateRecordset
    End Select
    
'common fields
    If PTHLENMobile = 999 Then
        ResultRecordset("distance(km)") = Null
    Else
        ResultRecordset("distance(km)") = PTHLENMobile / 1000
    End If
    
    If PRLoss = 999 Then
        ResultRecordset("pathloss(db)") = Null
    Else
        ResultRecordset("pathloss(db)") = PRLoss
    End If
    
    If PROPFQ = 0 Then
        ResultRecordset("Prop_Freq(MHz)") = Null
    Else
        ResultRecordset("Prop_Freq(MHz)") = PROPFQ
    End If
    
    ResultRecordset("Prop_Mode") = PropMode
    
    If FDR = 999 Then
        ResultRecordset("FDR(dB)") = Null
    Else
        ResultRecordset("FDR(dB)") = FDR
    End If
    
    If FreqSep = 999 Then
        ResultRecordset("FreqSep(MHz)") = Null
    Else
        ResultRecordset("FreqSep(MHz)") = FreqSep
    End If
    
    If Harmonic = 999 Then
        ResultRecordset("Harmonic") = Null
    Else
        ResultRecordset("Harmonic") = Harmonic
    End If
    
    If CrossPolar = 999 Then
        ResultRecordset("Polar_Loss(dB)") = Null
    Else
        ResultRecordset("Polar_Loss(dB)") = CrossPolar
    End If
    
    ResultRecordset("Error") = ErrorFound
    
    If Len(AnalysisComment) > 0 Then
        ResultRecordset("Comment") = AnalysisComment
    End If
        
UpdateRecordset:
    ResultRecordset.Update
    
    Err.Clear  'clear any pre-existing errors
    
End Sub

Public Sub Calculate_Intro_OffAxisGain()
'rx antenna gain
    If IntroRadiusofMobility = 0 Then 'fixed
    
        Select Case UCase(Trim(IntroAnalysisRecordset("ANT_MOTION")))
    
        Case "O", "R"
            IntroOffAxisGain = IntroMBGain
        Case "D"
            IntroPointingAngle = IntroAnalysisRecordset("ant_azimuth")
            
            If EnvirRadiusofMobility > 0 Then
                If PTHLEN <= EnvirRadiusofMobility * 1000 Then 'overlap, use mb gain
                    IntroOffAxisGain = IntroMBGain
                Else  'calculate off-axis angle
                    
                    Q = (EnvirRadiusofMobility * 1000) / PTHLEN
                    DeltaAngle = Atn(Q / Sqr(-Q * Q + 1)) * 57.29578
                    
                    MinAngle = BearIE_deg - DeltaAngle
                    MaxAngle = BearIE_deg + DeltaAngle
                    
                    If IntroPointingAngle >= MinAngle And IntroPointingAngle <= MaxAngle Then
                        
                        IntroOffAxisGain = IntroMBGain
                    
                    Else
                        
                        MinOffAxis = Abs(IntroPointingAngle - MinAngle)
                        
                        If MinOffAxis > 180 Then
                            MinOffAxis = 360 - MaxOffAxis
                        End If
                            
                        MaxOffAxis = Abs(IntroPointingAngle - MaxAngle)
                        
                        If MaxOffAxis > 180 Then
                            MaxOffAxis = 360 - MaxOffAxis
                        End If
                            
                        If MaxOffAxis < MinOffAxis Then
                            IntroDegOffAxis = MaxOffAxis
                        Else
                            IntroDegOffAxis = MinOffAxis
                        End If
                    
                        Call Calculate_FixedOffAxis_Gain(IntroMBGain, IntroDegOffAxis, IntroOffAxisGain)
    
                    End If
                   
                End If
                
                   
            Else
            
                If Abs(IntroPointingAngle - BearIE_deg) <= 180 Then
                    IntroDegOffAxis = Abs(IntroPointingAngle - BearIE_deg)
                Else
                    IntroDegOffAxis = 360 - Abs(IntroPointingAngle - BearIE_deg)
                End If
    
                Call Calculate_FixedOffAxis_Gain(IntroMBGain, IntroDegOffAxis, IntroOffAxisGain)
            
            End If
            
        End Select
            
    Else
        IntroOffAxisGain = IntroMBGain
    
    End If 'if fixed

End Sub

Public Sub Draw_Multi_Symbol(ResultDBField As String, EnvirLayer As Integer, _
             SymbolType As String, Color0 As Long, Color1 As Long, _
             PointSize As Integer, Thickness As Integer)
    
    frmProjectFeatures.mapSelection.LayerIndex = EnvirLayer
    
    If ErrorCounter > 0 = True Then 'add symbol indicating rec not processed due to missing data
        If frmProjectFeatures.mapSelection.NumberOfSymbols = 1 Then  'first pass thru
            frmProjectFeatures.mapSelection.AddSymbol
            frmProjectFeatures.mapSelection.SymbolIndex = 1
            frmProjectFeatures.mapSelection.SymbolName = "EMI"
            frmProjectFeatures.mapSelection.SymbolPointSymbol = SymbolType
            frmProjectFeatures.mapSelection.SymbolPointColor = Color1
            frmProjectFeatures.mapSelection.SymbolPointSize = PointSize
            frmProjectFeatures.mapSelection.SymbolPointThickness = Thickness
            frmProjectFeatures.mapSelection.LayerSymbolExpression = ResultDBField
            frmProjectFeatures.mapSelection.LayerLinkType = 0  'string
            frmProjectFeatures.mapSelection.SymbolStringEquals = "None"
        
            frmProjectFeatures.mapSelection.SymbolIndex = 0
            frmProjectFeatures.mapSelection.SymbolName = "Missing Data"
            frmProjectFeatures.mapSelection.SymbolPointSymbol = SymbolType
            frmProjectFeatures.mapSelection.SymbolPointColor = Color0
            frmProjectFeatures.mapSelection.SymbolPointSize = PointSize
            frmProjectFeatures.mapSelection.SymbolPointThickness = Thickness
            frmProjectFeatures.mapSelection.LayerSymbolExpression = ResultDBField
            frmProjectFeatures.mapSelection.LayerLinkType = 0  'string
            frmProjectFeatures.mapSelection.SymbolStringEquals = ""  'string
        End If
    Else
        frmProjectFeatures.mapSelection.SymbolIndex = 0
        
        If frmProjectFeatures.mapSelection.NumberOfSymbols > 1 And _
                frmProjectFeatures.mapSelection.SymbolName = "Missing Data" Then  'first pass thru
            
            frmProjectFeatures.mapSelection.SymbolIndex = 1
            frmProjectFeatures.mapSelection.DeleteSymbol
            frmProjectFeatures.mapSelection.SymbolIndex = 0
            frmProjectFeatures.mapSelection.SymbolName = ""
            frmProjectFeatures.mapSelection.SymbolPointColor = Color1
            frmProjectFeatures.mapSelection.LayerSymbolExpression = ""
        End If
    End If

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
           + "Number of Interactions Culled/Removed:  " + Str(InteractionCounter - InterferenceCounter) + Chr(13) + Chr(10) _
           + Chr(13) + Chr(10) + "Number of Remaining Interactions:  " + Str(InterferenceCounter) + Chr(13) + Chr(10) _
           + "     Number of Interactions Exceeding Threshold:  " + Str(InterferenceCounter - ErrorCounter) + Chr(13) + Chr(10) _
           + "     Number of Interactions Not Analyzed Due to Missing Data:  " + Str(ErrorCounter), vbExclamation + vbOKOnly, "Analysis Results"
           
End Sub

Public Sub Calculate_FixedOffAxis_Gain(MainGain As Single, OffAxis As Single, OffAxisGain As Single)
'off-axis angle in degrees
    
    If MainGain > 9.33 Then 'statgain
        
        Call STATGAIN(MainGain, OffAxis, RelativeGain, StandardDev, STATError)

'actual gain returned
        OffAxisGain = RelativeGain
    
    Else
    
        Call GNWOLF(MainGain, (OffAxis / 57.2958), RelativeGain)
        
'relative gain returned as neg
        OffAxisGain = MainGain + RelativeGain
    
    End If
End Sub

Public Sub Initialize_Common_Variables()
    ErrorFound = "None"
    Mode = ""
    PropMode = ""
    AnalysisComment = ""
    
    EnvirLat = 999
    EnvirLong = 999
    EnvirOffAxisGain = 999
    IntroOffAxisGain = 999
    InterferenceLevel = 999
    PTHLEN = 999
    PTHLENMobile = 999
    BearEI_deg = 999
    PRLoss = 999
    FDR = 999
    FreqSep = 999
    Harmonic = 999
    CrossPolar = 999
    PROPFQ = 0
    
End Sub
