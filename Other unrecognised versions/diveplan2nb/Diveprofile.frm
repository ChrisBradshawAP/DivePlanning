VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Diveprofile 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Diveplan"
   ClientHeight    =   8160
   ClientLeft      =   165
   ClientTop       =   540
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Diveprofile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11880
   Begin VB.OptionButton Option2 
      Caption         =   "Meters"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   16
      Top             =   3045
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Feet"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   15
      Top             =   3045
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pan >"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   2
      Left            =   8760
      TabIndex        =   26
      Top             =   3060
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pan <"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   1
      Left            =   7680
      TabIndex        =   25
      Top             =   3060
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zoom"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   0
      Left            =   6600
      TabIndex        =   24
      Top             =   3060
      Width           =   975
   End
   Begin VB.CommandButton cmdsetting 
      Caption         =   "Setting"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5280
      TabIndex        =   18
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton cmdplan 
      Caption         =   "Plan List"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   240
      TabIndex        =   14
      Top             =   3060
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4320
      TabIndex        =   13
      Top             =   3060
      Width           =   855
   End
   Begin VB.CommandButton Cmdprint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3360
      TabIndex        =   12
      Top             =   3060
      Width           =   855
   End
   Begin VB.CommandButton cmddetails 
      Caption         =   "details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2400
      TabIndex        =   11
      Top             =   3060
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2775
      Left            =   7125
      TabIndex        =   3
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   6500
      Cols            =   51
      FixedCols       =   0
      RowHeightMin    =   30
      BackColorBkg    =   16777215
      GridLines       =   0
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame8 
      Caption         =   "Decompression Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   7040
      TabIndex        =   49
      Top             =   0
      Width           =   4800
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1780
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   3760
      Width           =   3015
   End
   Begin VB.Frame Frame6 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   60
      TabIndex        =   46
      Top             =   3525
      Width           =   3255
   End
   Begin VB.CommandButton cmdremove 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   305
      Left            =   960
      TabIndex        =   39
      Top             =   2640
      Width           =   795
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   305
      Left            =   255
      TabIndex        =   38
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdheup 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   190
      Left            =   1650
      TabIndex        =   36
      Top             =   1785
      Width           =   190
   End
   Begin VB.CommandButton cmd02up 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   190
      Left            =   1650
      TabIndex        =   35
      Top             =   1305
      Width           =   190
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tissue Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   60
      TabIndex        =   31
      Top             =   5700
      Width           =   11735
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   2025
         Left            =   120
         OleObjectBlob   =   "Diveprofile.frx":030A
         TabIndex        =   32
         Top             =   240
         Width           =   11535
      End
   End
   Begin VB.CommandButton cmddepthup 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   190
      Left            =   1650
      Picture         =   "Diveprofile.frx":1D99
      TabIndex        =   27
      Top             =   390
      Width           =   190
   End
   Begin MSComDlg.CommonDialog dlgchart 
      Left            =   840
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
   End
   Begin VB.TextBox txttime 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   10
      Top             =   885
      Width           =   795
   End
   Begin VB.TextBox txto2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   9
      Top             =   1350
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      TabIndex        =   6
      Top             =   3060
      Width           =   975
   End
   Begin VB.TextBox txtdepth 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Top             =   420
      Width           =   795
   End
   Begin VB.TextBox txthe 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   795
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   6960
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   10
      RTSEnable       =   -1  'True
   End
   Begin VB.CommandButton cmddepthdown 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1650
      TabIndex        =   28
      Top             =   585
      Width           =   195
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dive Plan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   60
      TabIndex        =   41
      Top             =   0
      Width           =   6975
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Height          =   2775
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton Command4 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   190
            Left            =   1470
            TabIndex        =   66
            Top             =   2185
            Width           =   190
         End
         Begin VB.CommandButton CMDPPO2PLUS 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   190
            Left            =   1470
            TabIndex        =   65
            Top             =   2000
            Width           =   190
         End
         Begin VB.TextBox TXTPPO2 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   660
            TabIndex        =   64
            Top             =   2040
            Width           =   795
         End
         Begin VB.CommandButton cmdhedown 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   190
            Left            =   1470
            TabIndex        =   62
            Top             =   1730
            Width           =   190
         End
         Begin VB.CommandButton cmdo2down 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   190
            Left            =   1470
            TabIndex        =   61
            Top             =   1260
            Width           =   190
         End
         Begin VB.CommandButton cmdtimedown 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   190
            Left            =   1470
            TabIndex        =   60
            Top             =   790
            Width           =   190
         End
         Begin VB.CommandButton cmdtimeup 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   190
            Left            =   1470
            TabIndex        =   59
            Top             =   600
            Width           =   190
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Clear All"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   305
            Left            =   1680
            TabIndex        =   48
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0FFC0&
            Caption         =   "bar"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1780
            TabIndex        =   67
            Top             =   2060
            Width           =   615
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0FFC0&
            Caption         =   "PPO2 :"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            TabIndex        =   63
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0FFC0&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   57
            Top             =   1640
            Width           =   495
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0FFC0&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   56
            Top             =   1170
            Width           =   615
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0FFC0&
            Caption         =   "minutes"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   55
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0FFC0&
            Caption         =   "meters"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   54
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0FFC0&
            Caption         =   "He :"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0FFC0&
            Caption         =   "O2 :"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Time :"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            TabIndex        =   51
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Depth :"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   50
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   15
         Left            =   0
         TabIndex        =   45
         Top             =   2760
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   2775
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   50
         Cols            =   5
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox textchange 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   58
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   7080
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Current Dive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3240
      TabIndex        =   42
      Top             =   3525
      Width           =   4135
      Begin MSChart20Lib.MSChart MSChart3 
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "Diveprofile.frx":9336
         TabIndex        =   43
         Top             =   240
         Width           =   3960
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dive History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7320
      TabIndex        =   0
      Top             =   3525
      Width           =   4495
      Begin VB.TextBox Text5 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5500
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   4300
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   4300
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4300
         TabIndex        =   21
         Text            =   "Text4"
         Top             =   4300
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   20
         Text            =   "Text3"
         Top             =   4300
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1700
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   4300
         Width           =   960
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Draw Dive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   1
         Top             =   4440
         Width           =   1095
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "Diveprofile.frx":AC30
         TabIndex        =   17
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.FileListBox File2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6360
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   5760
      Top             =   5520
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Depth"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2880
      TabIndex        =   40
      Top             =   75
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2520
      TabIndex        =   37
      Top             =   75
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "He :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   34
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "O2 :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   33
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Time :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   30
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Depth :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Diveprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim DB As Database
Dim Y(20) As Integer
Dim X As Integer
Dim fp As UserDocument
Dim previouspoint
Dim profilefound As Integer
Dim maxdprofile As Integer
Dim CB As Byte
Dim C As String
Dim c1 As Byte
Dim i As Integer
Dim G As Integer
Dim LP As Long
Dim wp(400, 12) As String
Dim xp(400) As Integer
Dim yp(400) As Integer
Dim yp1(400) As Integer
Dim yp2(400) As Integer
Dim yp3(400) As Integer
Dim yp4(400) As Integer
Dim yp5(400) As Integer
Dim yp6(400) As Integer
Dim yp7(400) As Integer
Dim yp8(400) As Integer
Dim yp9(400) As Integer
Dim xpmax As Integer
Dim ypmax1 As Integer
Dim ypmax2 As Integer
Dim ypmax3 As Integer
Dim ypmax4 As Integer
Dim ypmax5 As Integer
Dim ypmax6 As Integer
Dim ypmax7 As Integer
Dim ypmax8 As Integer
Dim ypmax9 As Integer
Dim j As Integer
Dim txt3(30) As String
Dim tempstartdate As String
Dim tempfinishdate As String
Dim txt2(20) As String
'Dim #1 As Integer
Dim hOutFile As Integer
'
Dim F1 As String
Dim T(4) As String
Dim T2(4) As String
Dim S As String
Dim TS As String
Dim K As Integer
Dim H As String
Dim W(4) As Boolean
Dim zxt As Double
Dim CT As Integer
Dim Auto_run As Integer

Dim zoom As Integer
Dim pan As Integer

Private Sub Check1_Click()
  plotgraph
End Sub
Private Sub Check2_Click()
   plotgraph
End Sub
Private Sub Check3_Click()
  plotgraph
End Sub
Private Sub Check4_Click()
 plotgraph
End Sub
Private Sub Check5_Click()
   plotgraph
End Sub
Private Sub Check6_Click()
   plotgraph
End Sub
Private Sub Check7_Click()
  plotgraph
End Sub

Private Sub Check8_Click()
   plotgraph
End Sub
Private Sub Check9_Click()
   plotgraph
End Sub
Private Sub Check10_Click()
 plotgraph
End Sub
Private Sub Check11_Click()
 plotgraph
End Sub
Private Sub Combo1_Change()

End Sub




Private Sub cmd02up_Click()
txto2 = txto2 + 1
End Sub

Private Sub cmdadd_Click()
MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
K = MSFlexGrid3.Rows
MSFlexGrid3.Row = K - 1
MSFlexGrid3.Col = 0
MSFlexGrid3.Text = K - 1
MSFlexGrid3.Col = 1
MSFlexGrid3.Text = txtdepth
MSFlexGrid3.Col = 2
MSFlexGrid3.Text = txttime
MSFlexGrid3.Col = 3
MSFlexGrid3.Text = txto2
MSFlexGrid3.Col = 4
MSFlexGrid3.Text = txthe
End Sub

Private Sub cmdclose_Click()
Unload Me
rbmain.Show
End Sub

Private Sub cmddepthdown_Click()
txtdepth = txtdepth - 1
End Sub

Private Sub cmddepthup_Click()
txtdepth = txtdepth + 1
End Sub

Private Sub cmdgas_Click()
Unload Me
rbgas.Show
End Sub

Private Sub cmdgo_Click()
If Option1 = True Then
   displaydefaulted = "Feet"
Else
   displaydefaulted = "Meter"
End If
Unload Me
rbinterface.Show
End Sub

Private Sub cmdhedown_Click()
txthe = txthe - 1
End Sub

Private Sub cmdheup_Click()
txthe = txthe + 1
End Sub

Private Sub cmdo2down_Click()
txto2 = txto2 - 1
End Sub

Private Sub cmdsave_Click()
 On Error GoTo ErrorHandler2
    CommonDialog1.Action = 2
 filefilter = "Text Files (*.csv)|*.csv|All Files (*.*)|*.*"
 CommonDialog1.Filter = filefilter
 Open CommonDialog1.FileName For Output As #1
 MSFlexGrid1.Row = 0
 
    For j = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Col = j
        rowtext = MSFlexGrid1.Text
        cOMPTEXT = cOMPTEXT + (rowtext + ",")
    Next j
    Print #1, cOMPTEXT
    cOMPTEXT = ""
          Select Case displaydefaulted
             Case "Feet"
                displaydefault = "Feet"
             Case "Meter"
                displaydefault = "Meter"
             Case Else
                SQL = "SELECT * FROM Display "
                Set RS4 = DB.OpenRecordset(SQL)
                displaydefault = RS4("display")
              End Select
          
    RS.MoveFirst
    Do Until RS.EOF
      For j = 0 To RS.Fields.Count - 1
         If IsNull(RS(j)) Then
            rowtext = ""
         Else
            rowtext = CStr(RS(j))
         End If
         If displaydefault = "Feet" Then
            If j = 2 Or j = 9 Or j = 7 Or j = 15 Or j = 13 Then
               cOMPTEXT = Trim(cOMPTEXT)
            Else
               cOMPTEXT = Trim(cOMPTEXT)
               cOMPTEXT = cOMPTEXT + (rowtext) & ","
            End If
         Else
            If j = 3 Or j = 10 Or j = 8 Or j = 16 Or j = 14 Then
              cOMPTEXT = Trim(cOMPTEXT)
               'cOMPTEXT = cOMPTEXT + (rowtext) & ","
            Else
                cOMPTEXT = Trim(cOMPTEXT)
               cOMPTEXT = cOMPTEXT + (rowtext) & ","
              
            
            End If
         End If
        Next j
           Print #1, cOMPTEXT
        cOMPTEXT = ""
        RS.MoveNext
    Loop
    Close #1

ErrorHandler2:
    If Err = 32755 Then
        MsgBox "Data is not saved to a file....!!"
        Exit Sub
    End If

End Sub

Private Sub cmdsavegraph_Click()
 On Error GoTo saverr
  Dim strsavefile As String
  With dlgchart ' CommonDialog object
    .Filter = "Pictures (*.bmp)|*.bmp"
    .DefaultExt = "bmp"
    .CancelError = True
    .ShowSave
    strsavefile = .FileName
    If strsavefile = "" Then Exit Sub
  End With
  MSChart1.EditCopy
  SavePicture Clipboard.GetData, strsavefile
  Exit Sub
saverr:
'  MsgBox Err.Description
End Sub

Private Sub cmdtimedown_Click()
txttime = txttime - 1
End Sub

Private Sub cmdtimeup_Click()
txttime = txttime + 1
End Sub

Private Sub Cmdtissue_Click()
Unload Me
rbtissue.Show
End Sub

Private Sub cmdup_Click()

End Sub

Private Sub Command1_Click()
  Unload Me
  rbdetails.Show
  
 End Sub



Private Sub cleartext2()
Dim ind As Integer
 For ind = 0 To 18
        txt2(ind) = ""
 Next ind
End Sub

Private Sub fMain()

  Dim hOutFile As Integer

  hOutFile = FreeFile
  Open "mydata.csv" For Output As hOutFile

  Print #hOutFile, "xcv""xcvxc"

  Close hOutFile

End Sub
Private Sub cleartext3()
Dim ind As Integer
 For ind = 0 To 12
        txt2(ind) = ""
 Next ind
End Sub

Private Sub Command2_Click(Index As Integer) 'nick
  Select Case Index
    Case 0
    If Totalcount > 100 Then
      If zoom = 10 Then
        zoom = 1
        pan = 1
        Command2(1).Enabled = False
        Command2(2).Enabled = False
      Else
        zoom = 10
        pan = 1
        Command2(1).Enabled = True
        Command2(2).Enabled = True
      End If
    End If
    Case 1
      pan = pan - 1
      If pan < 1 Then pan = 1
    Case 2
      pan = pan + 1
      If pan > 10 Then pan = 10
  End Select
  If zoom = 10 Then
    If pan = 1 Then Command2(1).Enabled = False Else Command2(1).Enabled = True
    If pan = 10 Then Command2(2).Enabled = False Else Command2(2).Enabled = True
  End If
  Command2(0).Caption = "Zoom " + CStr(zoom)
  plotgraph
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
fgactivate = "0"
colselected = "false"
cols0activated = "false"
initialgrid
txtdepth = 0
txttime = 0
txto2 = 0
txthe = 0

Dim p As Integer


End Sub
Private Sub initialgrid()
MSFlexGrid3.Cols = 5
MSFlexGrid3.Col = 0
MSFlexGrid3.Rows = 1
MSFlexGrid3.Row = 0
MSFlexGrid3.Text = "No."
MSFlexGrid3.Col = 1
MSFlexGrid3.Text = "Depth"
MSFlexGrid3.Col = 2
MSFlexGrid3.Text = "Minutes"
MSFlexGrid3.Col = 3
MSFlexGrid3.Text = "O2"
MSFlexGrid3.Col = 4
MSFlexGrid3.Text = "He"
MSFlexGrid3.ColWidth(0) = 565
MSFlexGrid3.ColWidth(1) = 820
MSFlexGrid3.ColWidth(2) = 820
MSFlexGrid3.ColWidth(3) = 820
MSFlexGrid3.ColWidth(4) = 820
End Sub
Private Sub Check_input()
End Sub
Function checkpfexist(ByVal tempname As String) As Boolean
SQL = "SELECT COUNT(*) FROM pfindex "
SQL = SQL & " WHERE "
SQL = SQL & " itemname ='" & Trim(tempname) & "'"
Set RS3 = DB.OpenRecordset(SQL)
If RS3.Fields(0) = 0 Then
    checkpfexist = False
Else
    checkpfexist = True
End If

Set RS3 = Nothing
End Function

Function plotgraph()

End Function


Function determinexaxis()
 If zoom = 10 Then
  totalseconds = Val(Totalcount) * Val(txtinterval)
  totalseconds = totalseconds / zoom ' nick
  totalseconds = ((pan - 1) * totalseconds)
  totalbreak = totalseconds ' / 4
  genbreak = Format$(totalbreak, "#0")
  minutesbreak = genbreak / 60
  minutesbreak = minutesbreak - 0.499
  minutesbreak = Format$(minutesbreak, "#0")
  secondremainder = Val(genbreak) - Val(minutesbreak * 60)
  If Val(minutesbreak) > 60 Then
    hourbreak = minutesbreak / 60
    hourbreak = Format$(hourbreak, "#0")
    hourremainder = minutesbreak - (hourbreak * 60)
    If hourremainder < 10 Then
      hourremainder = "0" & hourremainder
    End If
  End If
  If minutesbreak < 10 Then
     minutesbreak = "0" & minutesbreak
  End If
  If secondremainder < 10 Then
     secondremainder = "0" & secondremainder
  End If
  If hourremainder = "" Then
    hourremainder = "00"
  End If
  Text1.Text = hourremainder & ":" & minutesbreak & ":" & secondremainder
 Else
   Text1.Text = "00:00:00"
 End If
    
    totalseconds = Val(Totalcount) * Val(txtinterval)
    If zoom = 10 Then
      totalseconds = totalseconds / zoom ' nick
      totalseconds = totalseconds / 4 + ((pan - 1) * totalseconds)
      totalbreak = totalseconds
    Else
      totalbreak = totalseconds / 4
    End If
    genbreak = Format$(totalbreak, "#0")
    minutesbreak = genbreak / 60
    minutesbreak = minutesbreak - 0.499
  minutesbreak = Format$(minutesbreak, "#0")
  secondremainder = Val(genbreak) - Val(minutesbreak * 60)
  If Val(minutesbreak) > 60 Then
    hourbreak = minutesbreak / 60
    hourbreak = Format$(hourbreak, "#0")
    hourremainder = minutesbreak - (hourbreak * 60)
    If hourremainder < 10 Then
      hourremainder = "0" & hourremainder
    End If
  End If
  If minutesbreak < 10 Then
     minutesbreak = "0" & minutesbreak
  End If
  If secondremainder < 10 Then
     secondremainder = "0" & secondremainder
  End If
  If hourremainder = "" Then
    hourremainder = "00"
  End If
  Text2.Text = hourremainder & ":" & minutesbreak & ":" & secondremainder
  
  'text3
     totalseconds = Val(Totalcount) * Val(txtinterval)
     If zoom = 10 Then
       totalseconds = totalseconds / zoom ' nick
       totalseconds = totalseconds / 2 + ((pan - 1) * totalseconds)
       'rtotalseconds = Val(totalseconds) - 0.499
       totalbreak = totalseconds
     Else
       rtotalseconds = Val(totalseconds) - 0.499
       totalbreak = rtotalseconds / 2
     End If
    genbreak = Format$(totalbreak, "#0")
    minutesbreak = genbreak / 60
    minutesbreak = minutesbreak - 0.499
    minutesbreak = Format$(minutesbreak, "#0")
    secondremainder = genbreak - (minutesbreak * 60)
    If Val(minutesbreak) > 60 Then
    hourbreak = minutesbreak / 60
    hourbreak = Format$(hourbreak, "#0")
    hourremainder = minutesbreak - (hourbreak * 60)
    If hourremainder < 10 Then
      hourremainder = "0" & hourremainder
    End If
  End If
  If minutesbreak < 10 Then
     minutesbreak = "0" & minutesbreak
  End If
  If secondremainder < 10 Then
     secondremainder = "0" & secondremainder
  End If
  If hourremainder = "" Then
    hourremainder = "00"
  End If
  Text3.Text = hourremainder & ":" & minutesbreak & ":" & secondremainder
  
  'text4
    totalseconds = Val(Totalcount) * Val(txtinterval)
    If zoom = 10 Then
      totalseconds = totalseconds / zoom ' nick
      totalseconds = (totalseconds * 3 / 4) + ((pan - 1) * totalseconds)
      'rtotalseconds = Val(totalseconds) - 0.499
      'totalbreak = rtotalseconds / 4
      totalbreak = totalseconds
    Else
      rtotalseconds = Val(totalseconds) - 0.499
      totalbreak = rtotalseconds / 4
      totalbreak = totalbreak * 3
    End If
    genbreak = Format$(totalbreak, "#0")
    minutesbreak = genbreak / 60
    minutesbreak = minutesbreak - 0.499
    minutesbreak = Format$(minutesbreak, "#0")
     secondremainder = genbreak - (minutesbreak * 60)
    If Val(minutesbreak) > 60 Then
    hourbreak = minutesbreak / 60
    hourbreak = hourbreak - 0.499
    hourbreak = Format$(hourbreak, "#0")
    minutesbreak = minutesbreak - (hourbreak * 60)
    If hourbreak < 10 Then
      hourbreak = "0" & hourbreak
    End If
  End If
  If minutesbreak < 10 Then
     minutesbreak = "0" & minutesbreak
  End If
  If secondremainder < 10 Then
     secondremainder = "0" & secondremainder
  End If
  If hourbreak = "" Then
    hourbreak = "00"
  End If
  Text4.Text = hourbreak & ":" & minutesbreak & ":" & secondremainder
  
  'text5
    totalseconds = Val(Totalcount) * Val(txtinterval)
    If zoom = 10 Then
      totalseconds = totalseconds / zoom ' nick
      totalseconds = totalseconds + ((pan - 1) * totalseconds)
    End If
    minutesbreak = totalseconds / 60
    minutesbreak = minutesbreak - 0.499
    minutesbreak = Format$(minutesbreak, "#0")
    secondremainder = totalseconds - (minutesbreak * 60)
    If Val(minutesbreak) > 60 Then
    hourbreak = minutesbreak / 60
    hourbreak = hourbreak - 0.499
    hourbreak = Format$(hourbreak, "#0")
    minutesbreak = minutesbreak - (hourbreak * 60)
    If hourbreak < 10 Then
      hourbreak = "0" & hourbreak
    End If
  End If
  If minutesbreak < 10 Then
     minutesbreak = "0" & minutesbreak
  End If
  If secondremainder < 10 Then
     secondremainder = "0" & secondremainder
  End If
  If hourbreak = "" Then
    hourbreak = "00"
  End If
  Text5.Text = hourbreak & ":" & minutesbreak & ":" & secondremainder
End Function

Private Sub Picture1_Click()
txtdepth = txtdepth + 1
End Sub

Private Sub Picture2_Click()
txtdepth = txtdepth - 1
End Sub

Private Sub Picture3_Click()
txttime = txttime + 1
End Sub

Private Sub Picture4_Click()
txttime = txttime - 1
End Sub

Private Sub Picture5_Click()
txto2 = txto2 + 1
End Sub

Private Sub Picture7_Click()
txto2 = txto2 - 1
End Sub

Private Sub Picture8_Click()
txthe = txthe + 1
End Sub

Private Sub Picture9_Click()
txthe = txthe - 1
End Sub

Private Sub Textchange_GotFocus()
MSFlexGrid3.Text = textchange.Text
  If cols0activated <> "true" Then
   ChangeCellText
  End If
End Sub
Public Sub ChangeCellText() ' Move Textbox to active cell.
   textchange.Move MSFlexGrid3.Left + MSFlexGrid3.CellLeft, _
   MSFlexGrid3.Top + MSFlexGrid3.CellTop, _
   MSFlexGrid3.CellWidth, MSFlexGrid3.CellHeight
   textchange.SetFocus
   textchange.ZOrder 0
End Sub

Private Sub Text2_LostFocus()
 If UsingMouse = True Then
      UsingMouse = False
      Exit Sub
   End If
   
   If MSFlexGrid3.Col <= MSFlexGrid3.Cols - 2 Then
      MSFlexGrid3.Col = MSFlexGrid3.Col + 1
      ChangeCellText
   Else
      If MSFlexGrid3.Row + 1 < MSFlexGrid3.Rows Then
        MSFlexGrid3.Row = MSFlexGrid3.Row + 1
        MSFlexGrid3.Col = 1
        ChangeCellText
      End If
   End If
End Sub
Private Sub MSFlexGrid3_MouseDown(Button As Integer, Shift As Integer, _
      X As Single, Y As Single)
      rowindentified = MSFlexGrid3.Row
      UsingMouse = True
     fgactivate = "1"
     textchange.Text = MSFlexGrid3.Text
  If MSFlexGrid3.Col = 0 Then
     colselected = "true"
     cols0activated = "true"
    ' MsgBox cols0activated
     For p = 0 To 4
        MSFlexGrid3.Row = rowindentified
        MSFlexGrid3.Col = p
        MSFlexGrid3.CellForeColor = vbWhite
        MSFlexGrid3.CellBackColor = vbBlue
        fgactivate = "0"
     Next p
    ' MsgBox "test"
     cmdremove.SetFocus
     
  Else
     colselected = "false"
     cols0activated = "false"
  End If
  If cols0activated <> "true" Then
        ChangeCellText
     End If
  'MsgBox MSFlexGrid3.Col
End Sub
Private Sub MSFlexGrid3_LeaveCell()
' Assign textbox value to grid
If colselected <> "true" Then
  If fgactivate = "1" Then
     MSFlexGrid3.Text = textchange.Text
     textchange.Text = ""
  End If
End If
'
End Sub

