VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Splanmain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Dive Series Main List"
   ClientHeight    =   9675
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11970
   Icon            =   "seqdpmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9375
      Left            =   -1080
      TabIndex        =   14
      Top             =   0
      Width           =   12135
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5400
         MouseIcon       =   "seqdpmain.frx":2CFA
         MousePointer    =   99  'Custom
         Picture         =   "seqdpmain.frx":3004
         ScaleHeight     =   375
         ScaleWidth      =   480
         TabIndex        =   24
         ToolTipText     =   "Click to Plan new dive"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   365
         Left            =   6720
         MouseIcon       =   "seqdpmain.frx":39A6
         MousePointer    =   99  'Custom
         Picture         =   "seqdpmain.frx":3CB0
         ScaleHeight     =   360
         ScaleWidth      =   465
         TabIndex        =   23
         ToolTipText     =   "Click to edit selected dive"
         Top             =   2400
         Width           =   470
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   365
         Left            =   8280
         MouseIcon       =   "seqdpmain.frx":4536
         MousePointer    =   99  'Custom
         Picture         =   "seqdpmain.frx":4840
         ScaleHeight     =   360
         ScaleWidth      =   465
         TabIndex        =   22
         ToolTipText     =   "Click to delete selected dive"
         Top             =   2400
         Width           =   465
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   480
         Picture         =   "seqdpmain.frx":50C6
         ScaleHeight     =   615
         ScaleWidth      =   375
         TabIndex        =   21
         ToolTipText     =   "Click to show gas settings"
         Top             =   9480
         Width           =   375
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   8280
         MouseIcon       =   "seqdpmain.frx":59F0
         MousePointer    =   99  'Custom
         Picture         =   "seqdpmain.frx":5CFA
         ScaleHeight     =   465
         ScaleWidth      =   525
         TabIndex        =   19
         ToolTipText     =   "Click to deletes selected series"
         Top             =   2280
         Width           =   525
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   6720
         MouseIcon       =   "seqdpmain.frx":69E4
         MousePointer    =   99  'Custom
         Picture         =   "seqdpmain.frx":6CEE
         ScaleHeight     =   465
         ScaleWidth      =   525
         TabIndex        =   18
         ToolTipText     =   "Click to edit selected dive series"
         Top             =   2280
         Width           =   525
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3120
         Left            =   2640
         TabIndex        =   27
         Top             =   3120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5503
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         RowHeightMin    =   30
         BackColor       =   16777215
         BackColorFixed  =   12632256
         BackColorBkg    =   14737632
         GridLines       =   0
         ScrollBars      =   2
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3135
         Left            =   2640
         TabIndex        =   28
         Top             =   3120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         RowHeightMin    =   30
         BackColor       =   14737632
         BackColorFixed  =   14737632
         BackColorBkg    =   14737632
         GridLines       =   0
         ScrollBars      =   2
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FFFFFF&
         Height          =   9015
         Left            =   1080
         Picture         =   "seqdpmain.frx":7A44
         ScaleHeight     =   8955
         ScaleWidth      =   10575
         TabIndex        =   31
         Top             =   0
         Width           =   10635
         Begin MSComDlg.CommonDialog CommonDialog3 
            Left            =   960
            Top             =   7200
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   480
            Top             =   7200
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   445
            Left            =   2280
            MouseIcon       =   "seqdpmain.frx":6C316
            MousePointer    =   99  'Custom
            Picture         =   "seqdpmain.frx":6C620
            ScaleHeight     =   450
            ScaleWidth      =   525
            TabIndex        =   43
            ToolTipText     =   "View Dive Series list"
            Top             =   2280
            Visible         =   0   'False
            Width           =   520
         End
         Begin VB.PictureBox Picture10 
            BorderStyle     =   0  'None
            Height          =   370
            Left            =   2280
            MouseIcon       =   "seqdpmain.frx":6D376
            MousePointer    =   99  'Custom
            Picture         =   "seqdpmain.frx":6D680
            ScaleHeight     =   375
            ScaleWidth      =   465
            TabIndex        =   42
            ToolTipText     =   "View Dive Plan List"
            Top             =   2280
            Visible         =   0   'False
            Width           =   470
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   465
            Left            =   4320
            MouseIcon       =   "seqdpmain.frx":6DF06
            MousePointer    =   99  'Custom
            Picture         =   "seqdpmain.frx":6E210
            ScaleHeight     =   465
            ScaleWidth      =   525
            TabIndex        =   37
            ToolTipText     =   "Click to Plan a new seires"
            Top             =   2280
            Width           =   525
         End
         Begin VB.Label lbldiveplan 
            BackStyle       =   0  'Transparent
            Caption         =   "View Dive Plan"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1800
            MouseIcon       =   "seqdpmain.frx":6EF66
            MousePointer    =   99  'Custom
            TabIndex        =   41
            ToolTipText     =   "View Dive Plan list"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Delete This Series"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            MouseIcon       =   "seqdpmain.frx":6F270
            MousePointer    =   99  'Custom
            TabIndex        =   40
            ToolTipText     =   "Delete the selected series"
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblodl 
            BackStyle       =   0  'Transparent
            Caption         =   "Edit This Series"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            MouseIcon       =   "seqdpmain.frx":6F57A
            MousePointer    =   99  'Custom
            TabIndex        =   39
            ToolTipText     =   "Click to Edit the selected Series"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblmdu 
            BackStyle       =   0  'Transparent
            Caption         =   "Make Series"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            MouseIcon       =   "seqdpmain.frx":6F884
            MousePointer    =   99  'Custom
            TabIndex        =   38
            ToolTipText     =   "Create a new seires"
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label lbldiveseries 
            BackStyle       =   0  'Transparent
            Caption         =   "View Dive Series"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   1800
            MouseIcon       =   "seqdpmain.frx":6FB8E
            MousePointer    =   99  'Custom
            TabIndex        =   36
            ToolTipText     =   "View Dive Seires List"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Delete This Dive"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6840
            MouseIcon       =   "seqdpmain.frx":6FE98
            MousePointer    =   99  'Custom
            TabIndex        =   35
            ToolTipText     =   "Delete the selected dive"
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Edit This Dive"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   5280
            MouseIcon       =   "seqdpmain.frx":701A2
            MousePointer    =   99  'Custom
            TabIndex        =   34
            ToolTipText     =   "Click to Edit the selected Dive"
            Top             =   2055
            Width           =   1455
         End
         Begin VB.Label lbllo 
            BackStyle       =   0  'Transparent
            Caption         =   "Plan New Dive"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   3960
            MouseIcon       =   "seqdpmain.frx":704AC
            MousePointer    =   99  'Custom
            TabIndex        =   33
            ToolTipText     =   "Plan a New Dive"
            Top             =   2055
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   $"seqdpmain.frx":707B6
            DataField       =   $"seqdpmain.frx":70851
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1095
            Left            =   4080
            TabIndex        =   32
            Top             =   6240
            Width           =   2415
         End
         Begin VB.Image Image1 
            Height          =   9075
            Left            =   0
            Picture         =   "seqdpmain.frx":708EC
            Stretch         =   -1  'True
            Top             =   -600
            Width           =   10560
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "    Dive List"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   1200
         Width           =   5895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "    Dive Series List"
         Height          =   255
         Left            =   3120
         TabIndex        =   29
         Top             =   5160
         Width           =   6015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dive Planning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   25
         Top             =   195
         Width           =   4935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gas Default"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   8880
         Width           =   2655
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblexit 
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   17
         ToolTipText     =   "Closes down this software"
         Top             =   9840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Series Planning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   16
         Top             =   4440
         Width           =   4335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Gas Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         TabIndex        =   15
         ToolTipText     =   "Click to show gas settings"
         Top             =   9720
         Width           =   1095
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3135
         Index           =   0
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   3375
         Index           =   0
         Left            =   120
         Top             =   4680
         Width           =   10695
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000009&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1455
         Index           =   0
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   4320
         Width           =   10695
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   855
         Index           =   2
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   9360
         Width           =   2535
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   2
         Left            =   120
         Top             =   9240
         Width           =   2775
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FFFFC0&
         BorderColor     =   &H80000009&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1335
         Index           =   2
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   8760
         Width           =   2775
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3255
         Index           =   1
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   2415
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000009&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1455
         Index           =   1
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   10695
      End
   End
   Begin VB.CommandButton cmddeldive 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   6360
      Picture         =   "seqdpmain.frx":D51BE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8400
      Width           =   1380
   End
   Begin VB.CommandButton cmdEditSeq 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   3000
      Picture         =   "seqdpmain.frx":D9C60
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8400
      Width           =   1380
   End
   Begin VB.CommandButton cmdseqplan 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   2400
      Picture         =   "seqdpmain.frx":DE702
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8280
      Width           =   1380
   End
   Begin VB.CommandButton cmdgasprofile 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   9225
      Picture         =   "seqdpmain.frx":E31A4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   1380
   End
   Begin VB.CommandButton cmdnewplan 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   5280
      Picture         =   "seqdpmain.frx":E7C46
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Width           =   1380
   End
   Begin VB.CommandButton cmdelete 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   3600
      Picture         =   "seqdpmain.frx":EC6E8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   1380
   End
   Begin VB.CommandButton cmdgo 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   7680
      Picture         =   "seqdpmain.frx":F118A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8400
      Width           =   1380
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Main Menu"
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
      Left            =   9360
      TabIndex        =   5
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files (*.csv)|*.csv"
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save as csv"
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
      Left            =   7320
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   5760
      Top             =   5520
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
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   8280
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File2 
      Height          =   675
      Left            =   6240
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtsearch 
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   1380
      Left            =   4560
      TabIndex        =   13
      Top             =   8235
      Width           =   6135
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   3240
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnudiveplan 
      Caption         =   "Dive &Planning"
      Begin VB.Menu mnuviewlibplan 
         Caption         =   "&View Dive Plan"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnunewplan 
         Caption         =   "&New Dive Plan"
         Shortcut        =   ^N
      End
      Begin VB.Menu Mnuplanprofile 
         Caption         =   "&Edit Dive Plan"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnugasprofile 
         Caption         =   "&Gas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudpprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnusavecsv2 
         Caption         =   "Save in &CSV"
      End
      Begin VB.Menu mnusort 
         Caption         =   "&Sort"
         Begin VB.Menu mnusortplanno 
            Caption         =   "By &Plan No"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnusortpdate 
            Caption         =   "By Plan &Date"
         End
         Begin VB.Menu mnusortmaxdepth 
            Caption         =   "By &Max Depth"
         End
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   "&Dive Series"
      Enabled         =   0   'False
      Begin VB.Menu mnuviewdiveseires 
         Caption         =   "&View Dive Series"
      End
      Begin VB.Menu mnuseqnewpl 
         Caption         =   "&New Dive Series"
      End
      Begin VB.Menu mnupopprofile 
         Caption         =   "&Edit Series"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&Delete Series"
      End
      Begin VB.Menu mnusavecsv 
         Caption         =   "Save in &CSV"
      End
   End
   Begin VB.Menu mnuUnits 
      Caption         =   "&Units"
      Begin VB.Menu mnuUnitsMeters 
         Caption         =   "&Meters"
         Checked         =   -1  'True
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuUnitsFeet 
         Caption         =   "&Feet"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnugetting 
         Caption         =   "&Getting Started"
      End
      Begin VB.Menu mnuwalk 
         Caption         =   "&Walk Through"
      End
      Begin VB.Menu Mnutips 
         Caption         =   "&Tips"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuend 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "Splanmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim DB As Database
Dim Y(20) As Integer
Dim X As Integer
Dim fp As UserDocument
'Dim fpp As CFileBinaryReadable
'Dim FileMgr As New CFileManager
'Dim TxtFile As CFileTextReadable
Dim profilefound As Integer
Dim maxdprofile As Integer
Dim CB As Byte
Dim C As String
Dim c1 As Byte
Dim i As Integer
Dim G As Integer
Dim LP As Long

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
Dim k As Integer
Dim H As Integer
Dim w(4) As Boolean
Dim zxt As Double
Dim CT As Integer
Dim Auto_run As Integer



Private Sub CMDCLOSE_Click()
Unload Me
previousform = "SEQLIST"
'main.Show

End Sub

Private Sub cmddelete_Click()
If MSFlexGrid1.CellBackColor = vbBlue Then
   MSFlexGrid1.Col = 0
   tempserialno = MSFlexGrid1.Text
   MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
   MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor
   ans = MsgBox("Are you sure you want to deleted the selected record(s)?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
      Select Case ans
         Case vbYes
            SQL = "DELETE FROM main "
            SQL = SQL & "WHERE seqdiveidmain = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            Me.MousePointer = vbNormal
         Case vbNo
            Me.MousePointer = vbNormal
            Exit Sub
      End Select
    On Error GoTo errorhandle:
errorhandle:
   If Err.Number <> 0 Then
   End If
    MsgBox Error$

End If
End Sub

Private Sub cmddeldive_Click()
If Trim(rowindentified) <> "" Then
MSFlexGrid2.Row = rowindentified
MSFlexGrid2.Col = 0
tempserialno = MSFlexGrid2.Text
tempseqduplicate = False
   If MSFlexGrid2.CellBackColor = vbBlue Then
      SQL = "select * FROM seqdplist "
      SQL = SQL & "WHERE seqdiveid = '" & tempserialno & "'"
      Set RS = DB.OpenRecordset(SQL)
      While RS.EOF = False
          tempseqduplicate = True
          RS.MoveNext
      Wend
      If tempseqduplicate = True Then
         Title = "Error deleting dive"
         MsgBox "Cannot delete dive, as being used in dive series! " & Chr(13) & "Delete series or delete dive from other series first.", 48, Title
      Else
         ans = MsgBox("Are you sure you want to delete this Dive from the database ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
         Select Case ans
         Case vbYes
            Screen.MousePointer = 11
            SQL = "select * FROM seqdplist "
            SQL = SQL & "WHERE seqdiveid = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            While RS.EOF = False
                tempseq = RS("seqdiveidseq")
                RS.Delete
                RS.MoveNext
            Wend
            SQL = "select * FROM seqdpmain "
            SQL = SQL & "WHERE diveplanid = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            While RS.EOF = False
               RS.Delete
               RS.MoveNext
            Wend
            RS.Close
            SQL = "select * FROM seqdpprofile "
            SQL = SQL & "WHERE dpprofileid = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            While RS.EOF = False
               RS.Delete
               RS.MoveNext
            Wend
       
            SQL = "select * FROM dpmaingaslist "
            SQL = SQL & "WHERE dpmainid = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            While RS.EOF = False
               RS.Delete
               RS.MoveNext
            Wend
'            MsgBox "Dive deleted from database!"
            Screen.MousePointer = 0
         Case vbNo
            Me.MousePointer = 0
            Exit Sub
         End Select
      End If
      validateserialno
   Else
      Title = "Error deleting dive"
      MsgBox "No dive selected to delete !", 48, Title
   End If
Else
   Title = "Error deleting dive"
   MsgBox "No dive selected to delete !", 48, Title

End If
reloadgrid2
End Sub

Private Sub cmdEditSeq_Click()
  mnupopprofile_Click
End Sub

Private Sub cmdelete_Click()
mnudelete_Click
End Sub

Private Sub cmdgasprofile_Click()
mnugasprofile_Click
End Sub


Private Sub cmdgo_Click()
Mnuplanprofile_Click
End Sub



Private Sub cmdnewplan_Click()
mnunewplan_Click
End Sub

Private Sub cmdsave_Click()
 On Error GoTo ErrorHandler2
    CommonDialog1.Action = 2
 If Right(CommonDialog1.FileName, 3) = "csv" Then
 Open CommonDialog1.FileName For Output As #1
 MSFlexGrid1.Row = 0
 
    For j = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Col = j
        rowtext = MSFlexGrid1.Text
        comptext = comptext + (rowtext + ",")
    Next j
    Print #1, comptext
    comptext = ""
    RS.MoveFirst
    Do Until RS.EOF
        For j = 0 To RS.Fields.Count - 1
            If IsNull(RS(j)) Then
               rowtext = ""
            Else
               rowtext = CStr(RS(j))
            End If
             rowtext = Trim(rowtext)
             comptext = comptext + (rowtext) & ","
                       
        Next j
           Print #1, comptext
        comptext = ""
        RS.MoveNext
    Loop
    Close #1
    MsgBox "Data saved to CSV file....!!"
 Else
    MsgBox "File wrong extention !"
 End If
ErrorHandler2:
    If Err = 32755 Then
        MsgBox "Data is not saved to a file....!!"
        Exit Sub
    End If

End Sub

Private Sub Command1_Click()
  Unload Planprofile2
  'Planprofile2.Show 'rbdetails.Show
  
 End Sub

Private Sub Command3_Click()
End Sub

Private Sub cleartext2()
Dim ind As Integer
 For ind = 0 To 18
        txt2(ind) = ""
 Next ind
End Sub
Private Sub purgerecord()
  SQL = "SELECT * FROM seqdpmain "
  SQL = SQL & " order by diveplanid"
  Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      temppurgeno = RS3("diveplanid")
      SQL = "SELECT count(*) FROM seqdpprofile "
      SQL = SQL & " WHERE dpprofileID = '" & temppurgeno & "' "
      Set RS4 = DB.OpenRecordset(SQL)
        If RS4.Fields(0) = 0 Then
           SQL = "select * FROM seqdpmain "
           SQL = SQL & "WHERE DiveplanID = '" & temppurgeno & "'"
           Set RS5 = DB.OpenRecordset(SQL)
           While RS5.EOF = False
              RS5.Delete
              RS5.MoveNext
           Wend
           RS5.Close
           Exit Sub
        End If
      RS3.MoveNext
    Wend
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

Private Sub Command4_Click()
  Open F1 For Binary As #2
  Text9.Text = F1
  MSComm1.Output = "L"
  For I1 = 0 To 10
    Cls
    S = MSComm1.Input
    text8.Text = S
  Next I1
  'wait
  For I1 = 0 To 2816
    Get #2, , c1
    MSComm1.Output = Chr$(c1)
    Cls
    text8.Text = Chr$(c1)
  Next I1
  MSComm1.Output = vbCrLf
  Close #2
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdseqplan_Click()
previousform = "SEQLIST"
rowindentified = ""
SQL = "select * FROM dpserialno "
Set RS = DB.OpenRecordset(SQL)
tempseqdiveno2 = RS("seqdiveserialno")
  tempseqdiveno = Right(tempseqdiveno2, 8)
  newseqdiveno = CInt(tempseqdiveno) + 1
  tempseqdiveno = CInt(tempseqdiveno) + 1
  lengthsn = Len(tempseqdiveno)
  Select Case lengthsn
  Case 1
     tempseqdiveno = "TM0000000" & tempseqdiveno
     newseqdiveno = "SM0000000" & newseqdiveno
  Case 2
     tempseqdiveno = "TM000000" & tempseqdiveno
     newseqdiveno = "SM000000" & newseqdiveno
  Case 3
     tempseqdiveno = "TM00000" & tempseqdiveno
     newseqdiveno = "SM00000" & newseqdiveno
  Case 4
     tempseqdiveno = "TM0000" & tempseqdiveno
     newseqdiveno = "SM0000" & newseqdiveno
  Case 5
     tempseqdiveno = "TM000" & tempseqdiveno
     newseqdiveno = "SM000" & newseqdiveno
  Case 6
     tempseqdiveno = "TM00" & tempseqdiveno
     newseqdiveno = "SM00" & newseqdiveno
  Case 7
     tempseqdiveno = "TM0" & tempseqdiveno
     newseqdiveno = "SM0" & newseqdiveno
  Case 8
     tempseqdiveno = "TM" & tempseqdiveno
     newseqdiveno = "SM" & newseqdiveno
 End Select
 tempchoice = "NSP"
 Unload Me
 frmseqdive.Show
End Sub

Private Sub Form_Activate()
Unload Planprofile2
Unload frmgasprofile2
mnuviewlibplan.Enabled = False
mnuviewdiveseires = True
mnuseqnewpl.Enabled = False
mnupopprofile.Enabled = False
mnudelete.Enabled = False
mnusavecsv.Enabled = False


End Sub

Private Sub Form_GotFocus()
Unload Planprofile2
Unload frmgasprofile2

End Sub

Private Sub Form_Load()
'mnuUnitsFeet_Click
doneupdateonce = 0
systemversion = "Lite"
'systemversion = "Pro"
  MsgBox "DO NOT DIVE USING ANY TABLES GENERATED BY THIS SOFTWARE. BETA TESTING ONLY"
  Set DB = OpenDatabase(App.Path & "/planmain.mdb")
  If systemstarted = False Then
     cleandatabase
  End If
  purgerecord
  validateserialno
  systemstarted = True
  SQL = "SELECT * FROM seqdpmain"
  SQL = SQL & " order by diveplanid "
  Set RS5 = DB.OpenRecordset(SQL) 'nick changed to RS5
  If RS5.EOF = True And RS5.BOF = True Then
     MsgBox "No Library Dives detected, please create Dives"
  End If
  SQL = "SELECT * FROM dpserialno"
  Set RS = DB.OpenRecordset(SQL)
  'RS.Edit
  SQL = "SELECT * FROM dpserialno"
  Set RS = DB.OpenRecordset(SQL)
  'RS("dpunits") = "Feet"
  If IsNull(RS("dpunits")) Then
    RS.AddNew
    RS.Update
    RS.MoveFirst
    RS.Edit
    RS!dpunits = "Feet" 'feetormeter_string
    RS.Update
  End If
  
  feetormeter_string = RS("dpunits")
  If InStr(1, feetormeter_string, "Feet") Then
    mnuUnitsFeet_Click
  Else
    mnuUnitsMeters_Click
  End If

rowindentified = MSFlexGrid1.Rows - 1
rowindentified = MSFlexGrid2.Rows - 1
MSFlexGrid1_Click
MSFlexGrid2_Click

  SQL = "SELECT * FROM dpserialno"
  Set RS = DB.OpenRecordset(SQL)
  If IsNull(RS!buhl) Then
    buhl_mode = 2
    RS.Edit
    RS!buhl = CStr(buhl_mode)
    RS.Update
  Else
    buhl_mode = CInt(RS!buhl)
  End If
  buhl_mode = 2
If buhl_mode = 2 Then Picture11.Visible = False Else Picture11.Visible = True
'errorhandler:
 ' MsgBox "error"
Unload Planprofile2
Unload frmgasprofile2
Unload Planprofile
If MSFlexGrid1.Rows > 1 Then
   If MSFlexGrid1.Rows < 12 Then MSFlexGrid1.TopRow = 1 Else MSFlexGrid1.TopRow = MSFlexGrid1.Rows - 8
   If MSFlexGrid2.Rows < 12 Then MSFlexGrid2.TopRow = 1 Else MSFlexGrid2.TopRow = MSFlexGrid2.Rows - 8
End If
MSFlexGrid1.Visible = False
End Sub

Private Sub saveseqdpmain()
SQL = "select * FROM seqdplist "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempsediveid = RS3("seqdiveidmain")
   If tempsediveid Like "T*" Then
      tempdpid2 = Right(tempsediveid, 9)
      tempdpid2 = "S" & tempdpid2
      RS3.Edit
      RS3!seqdiveidmain = tempdpid2
      RS3.Update
   End If
   RS3.MoveNext
Wend
SQL = "select * FROM dpserialno "
Set RS3 = DB.OpenRecordset(SQL)
tempsediveid = RS3("seqdiveserialno")
If tempsediveid Like "T*" Then
      tempdpid2 = Right(tempsediveid, 9)
      tempdpid2 = "S" & tempdpid2
      RS3.Edit
      RS3!seqdiveserialno = tempdpid2
      RS3.Update
End If
End Sub
Private Sub deleteseqdpmain()
SQL = "select * FROM seqdplist "
 SQL = SQL & "order by seqdiveidmain "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("seqdiveidmain")
   If tempdpid Like "T*" Then
      RS3.Delete
   End If
   RS3.MoveNext
Wend
SQL = "select * FROM dpserialno "
Set RS3 = DB.OpenRecordset(SQL)
tempdpid1 = Right$(tempdpid, 8)
tempdpid1 = tempdpid1 - 1
lengthsn = Len(tempdpid1)
  Select Case lengthsn
  Case 1
     tempdpid1 = "SM0000000" & tempdpid1
  Case 2
     tempdpid1 = "SM000000" & tempdpid1
  Case 3
    tempdpid1 = "SM00000" & tempdpid1
  Case 4
    tempdpid1 = "SM0000" & tempdpid1
  Case 5
    tempdpid1 = "SM000" & tempdpid1
  Case 6
    tempdpid1 = "SM00" & tempdpid1
  Case 7
    tempdpid1 = "SM0" & tempdpid1
  Case 8
    tempdpid1 = "SM" & tempdpid1
 End Select
RS3.Edit
RS3!seqdiveserialno = tempdpid1
RS3.Update
RS3.Close
End Sub


Private Sub mnuaddinfo_Click()
End Sub

Private Sub mnupoopgas_Click()
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then Me.WindowState = 0
If Me.WindowState = 0 Then
  Me.Width = 10755
  Me.Height = 9275
  
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'previousform = "SEQLIST"
'main.Show
  SQL = "SELECT * FROM dpserialno"
  Set RS = DB.OpenRecordset(SQL)
  RS.Edit
'  RS!inc_depth = CStr(inc_depth)
  RS!buhl = CStr(buhl_mode)
  RS.Update
End Sub

Private Sub Label1_Click()
  cmddeldive_Click
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label4_Click()
  cmdgasprofile_Click
End Sub

Private Sub Label8_Click()
  cmdgo_Click
End Sub

Private Sub Label9_Click()
  mnudelete_Click
End Sub

Private Sub lbldiveplan_Click()
MSFlexGrid1.Visible = False
MSFlexGrid2.Visible = True
mnuviewlibplan.Enabled = False
mnuviewdiveseires = True
mnuseqnewpl.Enabled = False
mnupopprofile.Enabled = False
mnudelete.Enabled = False
mnusavecsv.Enabled = False
lbllo.Visible = True
Label8.Visible = True
Label1.Visible = True
Picture1.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = True
Picture6.Visible = True
Picture8.Visible = True
If buhl_mode = 2 Then Picture11.Visible = False Else Picture11.Visible = True
Picture10.Visible = False
Label9.Visible = False
lblodl.Visible = False
lblmdu.Visible = False
lbllo.Visible = True
Label8.Visible = True
Label1.Visible = True
mnuseqnewpl.Enabled = False
mnupopprofile.Enabled = False
mnudelete.Enabled = False
mnusavecsv.Enabled = False
mnunewplan.Enabled = True
Mnuplanprofile.Enabled = True
mnugasprofile.Enabled = True
mnusort.Enabled = True
lbldiveplan.Visible = False
lbldiveseries.Visible = True
Label5.Visible = True
'Label7.Visible = False
MSFlexGrid2_Click
End Sub

Private Sub lbldiveseries_Click()
MSFlexGrid1.Visible = True
MSFlexGrid2.Visible = False
mnuviewlibplan.Enabled = True
mnuviewdiveseires = False
mnuseqnewpl.Enabled = True
mnupopprofile.Enabled = True
mnudelete.Enabled = True
mnusavecsv.Enabled = True
lbllo.Visible = False
Label8.Visible = False
Label1.Visible = False
Picture1.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = False
Picture8.Visible = False

Picture10.Visible = True
Picture11.Visible = False

Label9.Visible = True
lblodl.Visible = True
lblmdu.Visible = True
lbllo.Visible = False
Label8.Visible = False
Label1.Visible = False
mnuseqnewpl.Enabled = True
mnupopprofile.Enabled = True
mnudelete.Enabled = True
mnusavecsv.Enabled = True
mnunewplan.Enabled = False
Mnuplanprofile.Enabled = False
mnugasprofile.Enabled = False
mnusort.Enabled = False
lbldiveplan.Visible = True
lbldiveseries.Visible = False
Label5.Visible = False
'Label7.Visible = True
MSFlexGrid1_Click
End Sub

Private Sub lbllo_Click()
  cmdnewplan_Click
End Sub

Private Sub lblmdu_Click()
  If MSFlexGrid2.Rows < 2 Then
     MsgBox "No library dive available, please create some dive plan first"
  Else
     cmdseqplan_Click
  End If
End Sub

Private Sub lblodl_Click()
  cmdEditSeq_Click
End Sub

Private Sub mnuabout_Click()
frmSplash.Show
End Sub

Private Sub mnudelete_Click()
If Trim(rowindentified) <> "" Then
   MSFlexGrid1.Row = rowindentified
   If MSFlexGrid1.CellBackColor = vbBlue Then
      ans = MsgBox("Are you sure you want to deleted the selected record(s)?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
      Select Case ans
         Case vbYes
            SQL = "select * FROM seqdplist "
            SQL = SQL & "WHERE seqdiveidmain = '" & tempseqdiveno & "'"
            SQL = SQL & "order by seqdiveidseq "
            Set RS = DB.OpenRecordset(SQL)
            While RS.EOF = False
               RS.Delete
               RS.MoveNext
            Wend
           
            Me.MousePointer = vbNormal
            'Unload Me
            'Splanmain.Show
         Case vbNo
            Me.MousePointer = vbNormal
            Exit Sub
            
      End Select
      initialise_all
      lbldiveseries_Click
   End If
   validateserialno
Else
 Title = "Error to delete the sequential plan"
   MsgBox "No Plan selected to delete !", 48, Title

End If
End Sub

Private Sub mnudpprint_Click()
Dim comptext As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim p As Integer
Dim q As Integer
On Error GoTo ErrorHandler:

  CommonDialog3.ShowPrinter

    Text1.Text = ""
    comptext = "DO NOT DIVE USING THESE TABLES. BETA SOFTWARE TESTING ONLY"
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = "Dive Plan List as on : " & Format(Now, "dd/mmm/yyyy")
    Text1.Text = Text1.Text + comptext + vbCrLf
    For k = 0 To MSFlexGrid2.Rows - 1
      For i = 0 To MSFlexGrid2.Cols - 1
         MSFlexGrid1.Row = k
         MSFlexGrid1.Col = i
         rowtext = MSFlexGrid1.Text
         comptext = comptext + (rowtext + vbTab)
      Next i
     Text1.Text = Text1.Text + comptext + vbCrLf
     comptext = ""
    Next k
    Printer.Print Text1.Text
    Text1.Text = ""
    Printer.NewPage
    Printer.EndDoc
ErrorHandler:
   MsgBox "Printer error !!"

End Sub

Private Sub mnuend_Click()
End
End Sub

Private Sub mnugasprofile_Click()
  MSFlexGrid2.Col = 0
 tempdiveserialno = MSFlexGrid2.Text
  'MSFlexGrid1.Col = 2
 tempserialno = MSFlexGrid2.Text
 oldserialno = tempserialno
 MSFlexGrid2.CellForeColor = MSFlexGrid1.ForeColor
 MSFlexGrid2.CellBackColor = MSFlexGrid1.BackColor '
 Unload Me
 tempchoice = "GSP"
 previousform = "SEQLIST"
 frmgasprofile2.Show
End Sub

Private Sub mnugetting_Click()
SQL = "SELECT * FROM dpSERIALNO"
  Set RS = DB.OpenRecordset(SQL)
  RS.Edit
  If mnugetting.Checked = True Then
     mnugetting.Checked = False
     RS!getstarted = "Close"
     RS.Update
  Else
     RS!getstarted = "Open"
     mnugetting.Checked = True
     RS.Update
     frmGetS.Show
     Unload Splanmain
     
  End If
'
'  If mnugetting.Checked = True Then
'    RS!getstarted = "Open"
'  Else
'    RS!getstarted = "Close"
'  End If
 ' RS.Update
    
End Sub

Private Sub mnunewplan_Click()
On Error Resume Next
  Unload Me
  previousform = "SEQLIST"
  tempchoice = "NPP"
  frmgasprofile2.Show
  Unload frmgasprofile2
End Sub

Private Sub Mnuplanprofile_Click()
If Trim(rowindentified) <> "" Then
   tempchoice = "SPP"
   Unload Me
   previousform = "SEQLIST"
   Planprofile2.Show
Else
  Title = "Error to pop the pfofile."
   MsgBox "No plan selected to display the profile !", 48, Title
End If
End Sub

Private Sub mnupopprofile_Click()
If Trim(rowindentified) <> "" Then
   'tempseqdiveno
   Unload Me
   tempchoice = "SPP"
   previousform = "SEQLIST"
   frmseqdive.Show
Else
   Title = "Error to pop the pfofile."
   MsgBox "No plan selected to display the profile for !", 48, Title
End If
End Sub

Private Sub mnupoptissue_Click()
End Sub

Private Sub mnusavecsv_Click()
 On Error GoTo ErrorHandler2
 CommonDialog1.Action = 2
 filefilter = "Excel Files (*.csv)|*.csv" '|All Files (*.*)|*.*"
 CommonDialog1.Filter = filefilter
  If Right(CommonDialog1.FileName, 3) = "csv" Then
 Open CommonDialog1.FileName For Output As #1
 MSFlexGrid1.Row = 0
 comptext = "Dive Series List "
 Print #1, comptext
 comptext = ""
 Print #1, comptext
 For k = 0 To MSFlexGrid1.Rows - 1
    For j = 0 To MSFlexGrid1.Cols - 1
       MSFlexGrid1.Col = j
       MSFlexGrid1.Row = k
       rowtext = MSFlexGrid1.Text
       comptext = comptext + (rowtext + ",")
    Next j
    Print #1, comptext
    comptext = ""
 Next k
    RS.MoveFirst
    comptext = ""
    Print #1, comptext
    Print #1, comptext
    Print #1, comptext
    comptext = "Dive Plan List "
    Print #1, comptext
    comptext = ""
    Print #1, comptext
    For k = 0 To MSFlexGrid2.Rows - 1
       For j = 0 To MSFlexGrid2.Cols - 1
          MSFlexGrid2.Col = j
          MSFlexGrid2.Row = k
          rowtext = MSFlexGrid2.Text
          comptext = comptext + (rowtext + ",")
       Next j
       Print #1, comptext
       comptext = ""
    Next k
    Print #1, comptext
    Close #1
    MsgBox "Data saved to CSV file....!!"
End If
ErrorHandler2:
    If Err = 32755 Then
        MsgBox "Data is not saved to a file....!!"
        Exit Sub
    End If

End Sub

Private Sub MNUSAVEDB_Click()

End Sub

Private Sub mnusavecsv2_Click()
Dim comptext As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim p As Integer
Dim q As Integer

 On Error GoTo ErrorHandler2
    CommonDialog2.Action = 2
 Open CommonDialog2.FileName For Output As #1
  comptext = "DO NOT DIVE USING THESE TABLES. BETA SOFTWARE TESTING ONLY"
   Print #1, comptext
  comptext = "Dive Plan list as on : " & Format(Now, "dd/mmm/yyyy")
   Print #1, comptext
   comptext = ""
   Print #1, comptext
   comptext = ""
   For k = 0 To MSFlexGrid2.Rows - 1
      For j = 0 To MSFlexGrid2.Cols - 1
         MSFlexGrid2.Row = k
         MSFlexGrid2.Col = j
         rowtext = MSFlexGrid2.Text
         comptext = comptext + (rowtext + ",")
      Next j
    Print #1, comptext
    comptext = ""
   Next k

  Close #1
  MsgBox "Data saved to CSV file....!!"

ErrorHandler2:
    If Err = 32755 Then
        MsgBox "Data is not saved to a file....!!"
        Exit Sub
    End If
End Sub

Private Sub mnuseqnewpl_Click()
SQL = "select * FROM dpserialno "
Set RS = DB.OpenRecordset(SQL)
tempseqdiveno2 = RS("seqdiveserialno")
  tempseqdiveno = Right(tempseqdiveno2, 8)
  newseqdiveno = CInt(tempseqdiveno) + 1
  tempseqdiveno = CInt(tempseqdiveno) + 1
  lengthsn = Len(tempseqdiveno)
  Select Case lengthsn
  Case 1
     tempseqdiveno = "TM0000000" & tempseqdiveno
     newseqdiveno = "SM0000000" & newseqdiveno
  Case 2
     tempseqdiveno = "TM000000" & tempseqdiveno
     newseqdiveno = "SM000000" & newseqdiveno
  Case 3
     tempseqdiveno = "TM00000" & tempseqdiveno
     newseqdiveno = "SM00000" & newseqdiveno
  Case 4
     tempseqdiveno = "TM0000" & tempseqdiveno
     newseqdiveno = "SM0000" & newseqdiveno
  Case 5
     tempseqdiveno = "TM000" & tempseqdiveno
     newseqdiveno = "SM000" & newseqdiveno
  Case 6
     tempseqdiveno = "TM00" & tempseqdiveno
     newseqdiveno = "SM00" & newseqdiveno
  Case 7
     tempseqdiveno = "TM0" & tempseqdiveno
     newseqdiveno = "SM0" & newseqdiveno
  Case 8
     tempseqdiveno = "TM" & tempseqdiveno
     newseqdiveno = "SM" & newseqdiveno
 End Select
 frmseqdive.Show
End Sub

Private Sub mnusortmaxdepth_Click()
 Dim divelast As String
Dim w As Integer
'initialise_all
mnusortpdate.Checked = False
mnusortplanno.Checked = False
mnusortmaxdepth.Checked = True
Label5.Caption = "These are singles dives. They assume no previous dive history. To add these dives into a mission sequence, use the Dive Series Planning features below."
divelast = "last"
MSFlexGrid2.Rows = 1
MSFlexGrid2.Cols = 9
MSFlexGrid2.Col = 0
MSFlexGrid2.Row = 0
MSFlexGrid2.Text = "Plan No."
MSFlexGrid2.Col = 1
MSFlexGrid2.Text = "Plan Date"
MSFlexGrid2.Col = 2
MSFlexGrid2.Text = "Max Depth"
MSFlexGrid2.Col = 3
MSFlexGrid2.Text = "Bottom mins"
MSFlexGrid2.Col = 4
MSFlexGrid2.Text = "Gas ID"
MSFlexGrid2.Col = 5
MSFlexGrid2.Text = "PPO2"
MSFlexGrid2.Col = 6
MSFlexGrid2.Text = "Open/Closed"
MSFlexGrid2.Col = 7
MSFlexGrid2.Text = "O2"
MSFlexGrid2.Col = 8
MSFlexGrid2.Text = "Helium"
MSFlexGrid2.ColWidth(0) = 1250
MSFlexGrid2.ColWidth(1) = 1140
MSFlexGrid2.ColWidth(2) = 940
MSFlexGrid2.ColWidth(3) = 750
MSFlexGrid2.ColWidth(4) = 0 '1040
MSFlexGrid2.ColWidth(5) = 840
MSFlexGrid2.ColWidth(6) = 1140
MSFlexGrid2.ColWidth(7) = 550
MSFlexGrid2.ColWidth(8) = 550
SQL = "SELECT * FROM seqdpmain"
SQL = SQL & " order by maxdepth  "
Set RS5 = DB.OpenRecordset(SQL) 'nick changed to RS5
 If RS5.EOF = True Then
     Exit Sub
     MsgBox "No Plan detected, please create some plan first"
     Splanmain.Show
 Else
     RS5.MoveFirst
     While RS5.EOF = False
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        k = MSFlexGrid2.Rows
        MSFlexGrid2.Row = k - 1
        MSFlexGrid2.Col = 0
        MSFlexGrid2.Text = RS5("diveplanid")
        tempdpid = RS5("diveplanid")
        MSFlexGrid2.Col = 1
        tempdate = RS5("plandate")
        MSFlexGrid2.Text = Format(tempdate, "dd/mmm/yyyy")
        MSFlexGrid2.Col = 2
        tempdepth = RS5("maxdepth")
        MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring) + "   "
         SQL = "SELECT * FROM seqdpprofile"
        SQL = SQL & " where dpprofileid = '" & Trim(tempdpid) & "'  "
        Set RS4 = DB.OpenRecordset(SQL) 'nick changed to RS5
        RS4.MoveFirst
        MSFlexGrid2.Col = 4
        MSFlexGrid2.Text = RS4("gasid")
        MSFlexGrid2.Col = 5
        MSFlexGrid2.Text = Format(RS4("po2"), "0.00") + "   "
        MSFlexGrid2.Col = 6
        MSFlexGrid2.Text = " " + RS4("dpcircuit")
        MSFlexGrid2.Col = 7
        MSFlexGrid2.Text = RS4("dpo2") + "   "
        MSFlexGrid2.Col = 8
        MSFlexGrid2.Text = RS4("dphe") + "   "
        MSFlexGrid2.Col = 3
        While RS4.EOF = False
           If MSFlexGrid2.Text = "" Then
              MSFlexGrid2.Text = "0"
           Else
              MSFlexGrid2.Text = CStr(CInt(MSFlexGrid2.Text) + CInt(RS4("duration")))
           End If
           RS4.MoveNext
        Wend

        RS5.MoveNext
      Wend
  End If
  
  
For k = 1 To MSFlexGrid2.Rows - 1
  For p = 0 To 0
    MSFlexGrid2.Row = k
    MSFlexGrid2.Col = p
         checkslected = True
         If MSFlexGrid2.Row = 1 Then
            defaultcolor = &HE0E0E0
         Else
            If defaultcolor = &HE0E0E0 Then
                  defaultcolor = &HFFFFFF
            Else
                  defaultcolor = &HE0E0E0
            End If
         End If
         For H = 0 To 8
            MSFlexGrid2.Row = k
            MSFlexGrid2.Col = H
            If defaultcolor = &HE0E0E0 Then
               MSFlexGrid2.CellBackColor = &HE0E0E0    '&H00E0E0E0&
               MSFlexGrid2.CellForeColor = vbBlack
            Else
               MSFlexGrid2.CellBackColor = &HFFFFFF
               MSFlexGrid2.CellForeColor = vbBlack
            End If
          Next H
  Next p
  MSFlexGrid2.Col = 3
  MSFlexGrid2.Text = MSFlexGrid2.Text + "mins"
Next k

If MSFlexGrid1.Rows > 1 Then
   If MSFlexGrid1.Rows < 12 Then MSFlexGrid1.TopRow = 1 Else MSFlexGrid1.TopRow = MSFlexGrid1.Rows - 8
   If MSFlexGrid2.Rows < 12 Then MSFlexGrid2.TopRow = 1 Else MSFlexGrid2.TopRow = MSFlexGrid2.Rows - 8
End If
MSFlexGrid2_Click
End Sub

Private Sub mnusortpdate_Click()

Dim divelast As String
'initialise_all
mnusortpdate.Checked = True
mnusortplanno.Checked = False
mnusortmaxdepth.Checked = False
Label5.Caption = "These are singles dives. They assume no previous dive history. To add these dives into a mission sequence, use the Dive Series Planning features below."
divelast = "last"
MSFlexGrid2.Rows = 1
MSFlexGrid2.Cols = 9
MSFlexGrid2.Col = 0
MSFlexGrid2.Row = 0
MSFlexGrid2.Text = "Plan No."
MSFlexGrid2.Col = 1
MSFlexGrid2.Text = "Plan Date"
MSFlexGrid2.Col = 2
MSFlexGrid2.Text = "Max Depth"
MSFlexGrid2.Col = 3
MSFlexGrid2.Text = "Bottom mins"
MSFlexGrid2.Col = 4
MSFlexGrid2.Text = "Gas ID"
MSFlexGrid2.Col = 5
MSFlexGrid2.Text = "PPO2"
MSFlexGrid2.Col = 6
MSFlexGrid2.Text = "Open/Closed"
MSFlexGrid2.Col = 7
MSFlexGrid2.Text = "O2"
MSFlexGrid2.Col = 8
MSFlexGrid2.Text = "Helium"
MSFlexGrid2.ColWidth(0) = 1250
MSFlexGrid2.ColWidth(1) = 1140
MSFlexGrid2.ColWidth(2) = 940
MSFlexGrid2.ColWidth(3) = 750
MSFlexGrid2.ColWidth(4) = 0 '1040
MSFlexGrid2.ColWidth(5) = 840
MSFlexGrid2.ColWidth(6) = 1140
MSFlexGrid2.ColWidth(7) = 550
MSFlexGrid2.ColWidth(8) = 550
SQL = "SELECT * FROM seqdpmain "
SQL = SQL & "ORDER by plandate  "
Set RS4 = DB.OpenRecordset(SQL)
If RS4.EOF = True Then
     Exit Sub
     MsgBox "No Plan detected, please create some plan first"
     Splanmain.Show
Else
   RS4.MoveFirst
   While RS4.EOF = False
      tempdpid = RS4("Diveplanid")
      MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
      k = MSFlexGrid2.Rows
      MSFlexGrid2.Row = k - 1
      MSFlexGrid2.Col = 0
      MSFlexGrid2.Text = RS4("diveplanid")
      MSFlexGrid2.Col = 1
      'MSFlexGrid2.Text = RS4("plandate")
      MSFlexGrid2.Text = Format(RS4("plandate"), "dd/mmm/yyyy")
      
      SQL = "SELECT * FROM seqdpprofile"
      SQL = SQL & " where Dpprofileid = '" & Trim(tempdpid) & "' order by Dpnumseq "
      Set RS5 = DB.OpenRecordset(SQL) 'nick changed to RS5
      RS5.MoveFirst
      tempdepth = "0"
      While RS5.EOF = False
        If IsNumeric(RS5("depth")) = False Then
          tempdepth = "0.1"
        End If
        If tempdepth < RS5("depth") Then
          tempdepth = RS5("depth")
        End If
        
        MSFlexGrid2.Col = 3
           If MSFlexGrid2.Text = "" Then
              MSFlexGrid2.Text = "0"
           Else
              MSFlexGrid2.Text = CStr(CInt(MSFlexGrid2.Text) + CInt(RS5("duration")))
           End If
           MSFlexGrid2.Col = 4
        MSFlexGrid2.Text = RS5("gasid")
        MSFlexGrid2.Col = 5
        MSFlexGrid2.Text = Format(RS5("po2"), "0.00") + "   "
        MSFlexGrid2.Col = 6
        MSFlexGrid2.Text = " " + RS5("dpcircuit")
        MSFlexGrid2.Col = 7
        MSFlexGrid2.Text = RS5("dpo2") + "   "
        MSFlexGrid2.Col = 8
        MSFlexGrid2.Text = RS5("dphe") + "   "
      RS5.MoveNext
      Wend
      MSFlexGrid2.Col = 2
      MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring) + "   "
    '    MSFlexGrid2.Col = 3
    '    MSFlexGrid2.Text = RS4("duration")
    '    MSFlexGrid2.Col = 4
    '    MSFlexGrid2.Text = RS5("gasid")
    '    MSFlexGrid2.Col = 5
    '    MSFlexGrid2.Text = Format(RS5("po2"), "0.00") + "   "
    '    MSFlexGrid2.Col = 6
    '    MSFlexGrid2.Text = " " + RS5("dpcircuit")
    '    MSFlexGrid2.Col = 7
    '    MSFlexGrid2.Text = RS5("dpo2") + "   "
    '    MSFlexGrid2.Col = 8
    '    MSFlexGrid2.Text = RS5("dphe") + "   "
    '    divelast = tempdpid
     
     RS4.MoveNext
  Wend
For k = 1 To MSFlexGrid2.Rows - 1
  For p = 0 To 0
    MSFlexGrid2.Row = k
    MSFlexGrid2.Col = p
         checkslected = True
         If MSFlexGrid2.Row = 1 Then
            defaultcolor = &HE0E0E0
         Else
            If defaultcolor = &HE0E0E0 Then
                  defaultcolor = &HFFFFFF
            Else
                  defaultcolor = &HE0E0E0
            End If
         End If
         For H = 0 To 8
            MSFlexGrid2.Row = k
            MSFlexGrid2.Col = H
            If defaultcolor = &HE0E0E0 Then
               MSFlexGrid2.CellBackColor = &HE0E0E0    '&H00E0E0E0&
               MSFlexGrid2.CellForeColor = vbBlack
            Else
               MSFlexGrid2.CellBackColor = &HFFFFFF
               MSFlexGrid2.CellForeColor = vbBlack
            End If
          Next H
  Next p
  MSFlexGrid2.Col = 3
  MSFlexGrid2.Text = MSFlexGrid2.Text + "mins"
Next k
End If
If MSFlexGrid1.Rows > 1 Then
   If MSFlexGrid1.Rows < 12 Then MSFlexGrid1.TopRow = 1 Else MSFlexGrid1.TopRow = MSFlexGrid1.Rows - 8
   If MSFlexGrid2.Rows < 12 Then MSFlexGrid2.TopRow = 1 Else MSFlexGrid2.TopRow = MSFlexGrid2.Rows - 8
End If
MSFlexGrid2_Click
End Sub

Private Sub mnusortplanno_Click()
Dim divelast As String
'initialise_all
Label5.Caption = "These are singles dives. They assume no previous dive history. To add these dives into a mission sequence, use the Dive Series Planning features below."
divelast = "last"
MSFlexGrid2.Rows = 1
MSFlexGrid2.Cols = 9
MSFlexGrid2.Col = 0
MSFlexGrid2.Row = 0
MSFlexGrid2.Text = "Plan No."
MSFlexGrid2.Col = 1
MSFlexGrid2.Text = "Plan Date"
MSFlexGrid2.Col = 2
MSFlexGrid2.Text = "Max Depth"
MSFlexGrid2.Col = 3
MSFlexGrid2.Text = "Bottom mins"
MSFlexGrid2.Col = 4
MSFlexGrid2.Text = "Gas ID"
MSFlexGrid2.Col = 5
MSFlexGrid2.Text = "PPO2"
MSFlexGrid2.Col = 6
MSFlexGrid2.Text = "Open/Closed"
MSFlexGrid2.Col = 7
MSFlexGrid2.Text = "O2"
MSFlexGrid2.Col = 8
MSFlexGrid2.Text = "Helium"
MSFlexGrid2.ColWidth(0) = 1250
MSFlexGrid2.ColWidth(1) = 1140
MSFlexGrid2.ColWidth(2) = 940
MSFlexGrid2.ColWidth(3) = 750
MSFlexGrid2.ColWidth(4) = 0 '1040
MSFlexGrid2.ColWidth(5) = 840
MSFlexGrid2.ColWidth(6) = 1140
MSFlexGrid2.ColWidth(7) = 550
MSFlexGrid2.ColWidth(8) = 550
SQL = "SELECT * FROM seqdpprofile"
SQL = SQL & " order by Dpprofileid "
Set RS5 = DB.OpenRecordset(SQL) 'nick changed to RS5
 If RS5.EOF = True Then
     Exit Sub
     MsgBox "No Plan detected, please create some plan first"
     Splanmain.Show
 Else
  RS5.MoveFirst
     
     While RS5.EOF = False
     tempdpid = RS5("Dpprofileid")
     If tempdpid = divelast Then
        MSFlexGrid2.Col = 2
        If IsNumeric(RS5("depth")) = False Then
          tempdepth = "0.1"
        End If
        If tempdepth < RS5("depth") Then
          tempdepth = RS5("depth")
          If IsNumeric(tempdepth) = False Then
            tempdepth = "0.1"
          End If
          MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring) + "   "
        End If
        MSFlexGrid2.Col = 3
        MSFlexGrid2.Text = CStr(CInt(MSFlexGrid2.Text) + CInt(RS5("duration")))
     Else
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        k = MSFlexGrid2.Rows
        MSFlexGrid2.Row = k - 1
        MSFlexGrid2.Col = 0
        If IsNull(tempdpid) Then Else MSFlexGrid2.Text = tempdpid
        MSFlexGrid2.Col = 1
        SQL = "SELECT * FROM seqdpmain "
        SQL = SQL & "WHERE DiveplanID ='" & Trim(tempdpid) & "'  "
        Set RS4 = DB.OpenRecordset(SQL)
        tempdate = RS4("plandate")
        MSFlexGrid2.Text = Format(tempdate, "dd/mmm/yyyy")
        MSFlexGrid2.Col = 2
        tempdepth = RS5("depth")
        If IsNumeric(tempdepth) = False Then
          tempdepth = "0.1"
        End If
        MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring) + "   "
        MSFlexGrid2.Col = 3
        MSFlexGrid2.Text = RS5("duration")
        MSFlexGrid2.Col = 4
        MSFlexGrid2.Text = RS5("gasid")
        MSFlexGrid2.Col = 5
        MSFlexGrid2.Text = Format(RS5("po2"), "0.00") + "   "
        MSFlexGrid2.Col = 6
        MSFlexGrid2.Text = " " + RS5("dpcircuit")
        MSFlexGrid2.Col = 7
        MSFlexGrid2.Text = RS5("dpo2") + "   "
        MSFlexGrid2.Col = 8
        MSFlexGrid2.Text = RS5("dphe") + "   "
        divelast = tempdpid
     End If
     RS5.MoveNext
  Wend
For k = 1 To MSFlexGrid2.Rows - 1
  For p = 0 To 0
    MSFlexGrid2.Row = k
    MSFlexGrid2.Col = p
         checkslected = True
         If MSFlexGrid2.Row = 1 Then
            defaultcolor = &HE0E0E0
         Else
            If defaultcolor = &HE0E0E0 Then
                  defaultcolor = &HFFFFFF
            Else
                  defaultcolor = &HE0E0E0
            End If
         End If
         For H = 0 To 8
            MSFlexGrid2.Row = k
            MSFlexGrid2.Col = H
            If defaultcolor = &HE0E0E0 Then
               MSFlexGrid2.CellBackColor = &HE0E0E0    '&H00E0E0E0&
               MSFlexGrid2.CellForeColor = vbBlack
            Else
               MSFlexGrid2.CellBackColor = &HFFFFFF
               MSFlexGrid2.CellForeColor = vbBlack
            End If
          Next H
  Next p
  MSFlexGrid2.Col = 3
  MSFlexGrid2.Text = MSFlexGrid2.Text + "mins"
Next k
End If
If MSFlexGrid1.Rows > 1 Then
   If MSFlexGrid1.Rows < 12 Then MSFlexGrid1.TopRow = 1 Else MSFlexGrid1.TopRow = MSFlexGrid1.Rows - 8
   If MSFlexGrid2.Rows < 12 Then MSFlexGrid2.TopRow = 1 Else MSFlexGrid2.TopRow = MSFlexGrid2.Rows - 8
End If
MSFlexGrid2_Click
End Sub

Private Sub Mnutips_Click()
  frmtips.Show
End Sub

Private Sub mnuUnitsFeet_Click()
  mnuUnitsFeet.Checked = True
  mnuUnitsMeters.Checked = False
  feetormeter_factor = 3.280839
  psiorbar_factor = 14.7
  feetormeter_string = "Feet"
  feetormeter_shortstring = "ft"
  feetormeter_feeton = 1
  feetormeter_decostep = 3.048
 ' initialise_all
  SQL = "SELECT * FROM dpserialno"
  Set RS = DB.OpenRecordset(SQL)
  RS.Edit
  RS!dpunits = feetormeter_string
  RS.Update
  reloadgrid2
  inc_depth = 3#
End Sub

Private Sub mnuUnitsMeters_Click()
  mnuUnitsFeet.Checked = False
  mnuUnitsMeters.Checked = True
  feetormeter_factor = 1#
  psiorbar_factor = 1#
  feetormeter_string = "Meters"
  feetormeter_shortstring = "m"
  feetormeter_feeton = 0
  feetormeter_decostep = 3
  'initialise_all
  SQL = "SELECT * FROM dpserialno"
  Set RS = DB.OpenRecordset(SQL)
  RS.Edit
  RS!dpunits = feetormeter_string
  RS.Update
  RS.Close
  reloadgrid2
  inc_depth = 1#
End Sub

Private Sub mnuviewdiveseires_Click()
lbldiveseries_Click
End Sub

Private Sub mnuviewlibplan_Click()
lbldiveplan_Click
End Sub

Private Sub mnuwalk_Click()
frmintro.Show
End Sub

Private Sub MSFlexGrid1_Click()
If MSFlexGrid1.Rows > 1 Then
checkslected = False
rowindentified = MSFlexGrid1.Row
For k = 0 To Totalcount
  For p = 0 To 0
    MSFlexGrid1.Row = k
    MSFlexGrid1.Col = p
    If MSFlexGrid1.CellBackColor = vbBlue Then
      If checkslected = False Then
         checkslected = True
         If MSFlexGrid1.Row = 1 Then
            defaultcolor = &HE0E0E0
         Else
            temptext = MSFlexGrid1.Text
            MSFlexGrid1.Row = MSFlexGrid1.Row - 1
            If Trim(temptext) <> "" Then
               If MSFlexGrid1.CellBackColor = &HFFFFFF Then
                  defaultcolor = &HE0E0E0
               Else
                  defaultcolor = &HFFFFFF
               End If
            Else
               If MSFlexGrid1.CellBackColor = &HFFFFFF Then
                  defaultcolor = &HFFFFFF
               Else
                  defaultcolor = &HE0E0E0
               End If
            End If
         End If
         For H = 0 To 6
            MSFlexGrid1.Row = k
            MSFlexGrid1.Col = H
            If defaultcolor = &HE0E0E0 Then
               MSFlexGrid1.CellBackColor = &HE0E0E0    '&H00E0E0E0&
               MSFlexGrid1.CellForeColor = vbBlack
            Else
               MSFlexGrid1.CellBackColor = &HFFFFFF
               MSFlexGrid1.CellForeColor = vbBlack
            End If
          Next H
       Else
         For H = 0 To 6
            MSFlexGrid1.Row = k
            MSFlexGrid1.Col = H
            If defaultcolor = &HE0E0E0 Then
               MSFlexGrid1.CellBackColor = &HE0E0E0
               MSFlexGrid1.CellForeColor = vbBlack
            Else
               MSFlexGrid1.CellBackColor = &HFFFFFF
               MSFlexGrid1.CellForeColor = vbBlack
            End If
          Next H
       End If
    End If
  Next p
Next k
'MSFlexGrid1.Col = 2
'MSFlexGrid1.Row = rowindentified
'tempserialno = MSFlexGrid1.Text
MSFlexGrid1.Col = 0
MSFlexGrid1.Row = rowindentified
checkslected = False
While checkslected = False
   MSFlexGrid1.Col = 0
   If Trim(MSFlexGrid1.Text) <> "" Then
      checkslected = True
   Else
      MSFlexGrid1.Row = MSFlexGrid1.Row - 1
   End If
   If checkslected = True Then
      MSFlexGrid1.Col = 0
      tempseqdiveno = MSFlexGrid1.Text
   End If
Wend
For p = 0 To 6
    MSFlexGrid1.Col = p
    MSFlexGrid1.Row = rowindentified
    'MsgBox MSFlexGrid1.Text
    MSFlexGrid1.CellForeColor = vbWhite
    MSFlexGrid1.CellBackColor = vbBlue
Next
'rowchecked = rowchecked + 1
'For p = rowchecked To MSFlexGrid1.Rows - 1
'   MSFlexGrid1.Row = p
'   MSFlexGrid1.Col = 0
'   tempsctext = MSFlexGrid1.Text
'     If tempsctext = "" Then
'      For K = 0 To 6
'        MSFlexGrid1.Col = K
'        MSFlexGrid1.CellForeColor = vbWhite
'        MSFlexGrid1.CellBackColor = vbBlue
'      Next K
'    Else
'    p = MSFlexGrid1.Rows - 1
'   End If
'Next p
'Else
 '  MsgBox "No Sequential Dive found, please insert some Dive(s) !"
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
 tempchoice = "SPP"
 oldserialno = tempseqdiveno
 previousform = "SEQLIST"
 Unload Me
 frmseqdive.Show
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
'    PopupMenu mnupopup
End If
End Sub

Private Sub initial_heading()
MSFlexGrid1.Col = 0
MSFlexGrid1.Row = 0
MSFlexGrid1.ColWidth(0) = 1800
MSFlexGrid1.Text = "Series No"
MSFlexGrid1.Col = 1
MSFlexGrid1.Row = 0
MSFlexGrid1.ColWidth(1) = 400
MSFlexGrid1.Text = "#"
MSFlexGrid1.Col = 2
MSFlexGrid1.Row = 0
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.Text = "Dive Plan No"
MSFlexGrid1.Col = 3
MSFlexGrid1.Row = 0
MSFlexGrid1.ColWidth(3) = 1100
MSFlexGrid1.Text = "Surface Intval"
MSFlexGrid1.Col = 4
MSFlexGrid1.Row = 0
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.Text = "Max Depth"
MSFlexGrid1.Col = 5
MSFlexGrid1.Row = 0
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.Text = "Safety Factor"
MSFlexGrid1.Col = 6
MSFlexGrid1.Row = 0
MSFlexGrid1.ColWidth(0) = 1400
MSFlexGrid1.Text = "Atmospheric"
End Sub
Private Sub initialise_all()
'Set DB = OpenDatabase(App.Path & "/planmain.mdb")
Screen.MousePointer = 11
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 30
tempsnfound = "False"
MSFlexGrid1.Rows = 1
SQL = "select * FROM seqdplist "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("seqdiveidmain")
   If tempdpid Like "T*" Then
      tempsnfound = "True"
   End If
   RS3.MoveNext
Wend
If tempsnfound = "True" Then
Title = "Error on System Validation.."
ans = MsgBox("You have Dive plan that was not saved, " & Chr(13) & "Press No will remove all previous unsaved plans !" & Chr(13) & "Do you want to save now ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
   saveseqdpmain
   MsgBox "Dive plan Saved !!"
Case vbNo
   deleteseqdpmain
End Select
End If
initial_heading
lblmdu.Visible = False
Picture1.Visible = False
Label1.Visible = False
Picture3.Visible = False
lblodl.Visible = False
Picture4.Visible = False
Label9.Visible = False
mnuseqnewpl.Enabled = False
mnupopprofile.Enabled = False
mnudelete.Enabled = False
mnusavecsv.Enabled = False
'Label7.Visible = True
SQL = "SELECT * FROM seqdplist "
SQL = SQL & "order by seqdiveidmain,seqdiveidseq "
Set RS = DB.OpenRecordset(SQL)
Splanmain.Caption = "ProPlanner Dive Planning Main List"
MSFlexGrid1.FontSize = 7
MSFlexGrid1.FontBold = False
If RS.BOF And RS.EOF Then
  Screen.MousePointer = 0
    SQL = "SELECT * FROM seqdpmain"
    SQL = SQL & " order by diveplanid "
    Set RS5 = DB.OpenRecordset(SQL) 'nick changed to RS5
    If RS5.EOF = True And RS5.BOF = True Then
   '    MsgBox "No Library Dives detected, please create Dives"
       cmdseqplan.Visible = False
       cmdEditSeq.Visible = False
       cmdgo.Visible = False
       Mnuplanprofile.Enabled = False
       cmdgasprofile.Visible = False
       mnugasprofile.Enabled = False
       cmdelete.Visible = False
       cmdSave.Enabled = False
       cmddeldive.Visible = False
       Picture6.Visible = False
       Label8.Visible = False
       Picture8.Visible = False
       
       Exit Sub
    Else
       cmdseqplan.Visible = True
       cmdEditSeq.Visible = False
       cmdgo.Visible = False
       Mnuplanprofile.Enabled = True
       cmdgasprofile.Visible = True
       mnugasprofile.Enabled = True
       cmdelete.Visible = False
       cmdSave.Enabled = True
       cmddeldive.Visible = True
       Picture6.Visible = True
       Label8.Visible = True
       Picture8.Visible = True
       Label1.Visible = True
       Exit Sub
    End If
    
       
Else
   cmdgo.Visible = True
   Mnuplanprofile.Enabled = True
   cmdgasprofile.Visible = True
   mnugasprofile.Enabled = True
   cmdelete.Visible = True
   cmddeldive.Visible = True
   cmdSave.Enabled = True
   cmdEditSeq.Visible = True
   Picture6.Visible = True
   Label8.Visible = True
   Picture8.Visible = True
   Label1.Visible = True
End If

    RS.MoveFirst
       numrow = 0
       While RS.EOF = False
          numrow = RS.RecordCount
          numrow = numrow + 1
          RS.MoveNext
       Wend
          RS.MoveFirst
          For i = 1 To numrow - 1
             If RS.EOF Then
                Exit For
             End If
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                MSFlexGrid1.Row = i
                
                For j = 0 To RS.Fields.Count - 1
                MSFlexGrid1.CellForeColor = &H0&
                   MSFlexGrid1.Col = j
                   If IsNull(RS(j)) Then
                      MSFlexGrid1.Text = ""
                   Else
                      TEMPVALUE = CStr(RS(j))
                      If j = 3 Then TEMPVALUE = TEMPVALUE + "hrs"
                      If j = 4 Then TEMPVALUE = Format(CDbl(TEMPVALUE) * feetormeter_factor, "###0" & feetormeter_shortstring)
                      If j = 5 Then TEMPVALUE = TEMPVALUE + "%"
                      If j = 6 Then TEMPVALUE = TEMPVALUE + "mBar"
                      If Val(TEMPVALUE) < 1 And Val(TEMPVALUE) > 0 Then
                         TEMPVALUE = "0" & TEMPVALUE
                         MSFlexGrid1.Text = TEMPVALUE
                      Else
                         MSFlexGrid1.Text = TEMPVALUE
                         If j = 0 Then
                            tempsmserial = TEMPVALUE
                          
                            If tempsmchanged = tempsmserial Then
                               smchanged = False
                               MSFlexGrid1.Text = ""
                            Else
                               smchanged = True
                               tempsmchanged = TEMPVALUE
                            End If
                          If smchanged = True Then
                             If test1 = &HE0E0E0 Then
                                MSFlexGrid1.CellBackColor = &HFFFFFF
                                test1 = &HFFFFFF
                             Else
                                MSFlexGrid1.CellBackColor = &HE0E0E0
                                test1 = &HE0E0E0
                             End If
                          Else
                             If test1 = &HE0E0E0 Then
                                MSFlexGrid1.CellBackColor = &HE0E0E0
                             Else
                                MSFlexGrid1.CellBackColor = &HFFFFFF
                             End If
                          End If
                       Else
                          If test1 = &HE0E0E0 Then
                             MSFlexGrid1.CellBackColor = &HE0E0E0
                          Else
                             MSFlexGrid1.CellBackColor = &HFFFFFF
                          End If
                          
                       End If
                    End If
                  End If
                 
              Next j
              RS.MoveNext
            Next i
            MSFlexGrid1.Rows = numrow
    Totalcount = numrow - 1
     Screen.MousePointer = 0
If MSFlexGrid1.Rows < 12 Then MSFlexGrid1.TopRow = 1 Else MSFlexGrid1.TopRow = MSFlexGrid1.Rows - 8
If MSFlexGrid2.Rows < 12 Then MSFlexGrid2.TopRow = 1 Else MSFlexGrid2.TopRow = MSFlexGrid2.Rows - 8
End Sub

Private Sub MSFlexGrid2_Click()
If MSFlexGrid2.Rows > 1 Then
 checkslected = False
rowindentified = MSFlexGrid2.Row
For k = 0 To MSFlexGrid2.Rows - 1
  For p = 0 To 0
    MSFlexGrid2.Row = k
    MSFlexGrid2.Col = p
    If MSFlexGrid2.CellBackColor = vbBlue Then
      If checkslected = False Then
         checkslected = True
         If MSFlexGrid2.Row = 1 Then
            defaultcolor = &HE0E0E0
         Else
            temptext = MSFlexGrid2.Text
            MSFlexGrid2.Row = MSFlexGrid2.Row - 1
            If Trim(temptext) <> "" Then
               If MSFlexGrid2.CellBackColor = &HFFFFFF Then
                  defaultcolor = &HE0E0E0
               Else
                  defaultcolor = &HFFFFFF
               End If
            Else
               If MSFlexGrid2.CellBackColor = &HFFFFFF Then
                  defaultcolor = &HFFFFFF
               Else
                  defaultcolor = &HE0E0E0
               End If
            End If
         End If
         For H = 0 To 8
            MSFlexGrid2.Row = k
            MSFlexGrid2.Col = H
            If defaultcolor = &HE0E0E0 Then
               MSFlexGrid2.CellBackColor = &HE0E0E0    '&H00E0E0E0&
               MSFlexGrid2.CellForeColor = vbBlack
            Else
               MSFlexGrid2.CellBackColor = &HFFFFFF
               MSFlexGrid2.CellForeColor = vbBlack
            End If
          Next H
       Else
         For H = 0 To 8
            MSFlexGrid2.Row = k
            MSFlexGrid2.Col = H
            If defaultcolor = &HE0E0E0 Then
               MSFlexGrid2.CellBackColor = &HE0E0E0
               MSFlexGrid2.CellForeColor = vbBlack
            Else
               MSFlexGrid2.CellBackColor = &HFFFFFF
               MSFlexGrid2.CellForeColor = vbBlack
            End If
          Next H
       End If
    End If
  Next p
Next k
MSFlexGrid2.Col = 0
MSFlexGrid2.Row = rowindentified
tempserialno = MSFlexGrid2.Text
'While checkslected = False
'   MSFlexGrid2.Col = 0
'   If Trim(MSFlexGrid2.Text) <> "" Then
'      checkslected = True
'   Else
'      MSFlexGrid2.Row = MSFlexGrid2.Row - 1
'   End If
'   If checkslected = True Then
'      MSFlexGrid2.Col = 0
'      tempseqdiveno = MSFlexGrid2.Text
'   End If
'Wend
For p = 0 To 8
    MSFlexGrid2.Col = p
    MSFlexGrid2.Row = rowindentified
    MSFlexGrid2.CellForeColor = vbWhite
    MSFlexGrid2.CellBackColor = vbBlue
Next
  MSFlexGrid2.Col = 0
  SQL = "SELECT * FROM seqdpprofile"
  SQL = SQL & " where dpprofileid = '" & MSFlexGrid2.Text & "' "
  SQL = SQL & " order by dpnumseq "
  Set RS = DB.OpenRecordset(SQL)
    RS.MoveFirst
    Label5.Caption = "Dive " + MSFlexGrid2.Text + ":" + vbCrLf
    Label5.Alignment = 2
  While RS.EOF = False
'     Label5(1).Caption = Label5(1).Caption + RS("dpnumseq") + " "
     Label5.Caption = Label5.Caption + Format(CStr(CDbl(RS("depth")) * feetormeter_factor), "###") + feetormeter_shortstring + " "
     Label5.Caption = Label5.Caption + RS("duration") + "mins "
     Label5.Caption = Label5.Caption + RS("dpcircuit") + vbCrLf
'     Label5(1).Caption = Label5(1).Caption + RS("dpo2")
'     Label5(1).Caption = Label5(1).Caption + RS("dphe")
'     Label5(1).Caption = Label5(1).Caption + RS("po2")
'     Label5(1).Caption = Label5(1).Caption + RS("gasid")
     RS.MoveNext
   Wend
'Else
'   MsgBox "No Dive found, please insert some Dive(s) !"
End If

End Sub

Private Sub MSFlexGrid2_DblClick()
Mnuplanprofile_Click
'MSFlexGrid2_Click
End Sub

Private Sub Picture1_Click()
  cmdseqplan_Click
End Sub

Private Sub Picture10_Click()
lbldiveplan_Click

End Sub

Private Sub Picture11_Click()
lbldiveseries_Click
End Sub

Private Sub Picture2_Click()
  cmdgasprofile_Click
End Sub

Private Sub Picture3_Click()
  cmdEditSeq_Click
End Sub

Private Sub Picture4_Click()
  mnudelete_Click
End Sub

Private Sub Picture5_Click()
  cmdnewplan_Click
End Sub

Private Sub Picture6_Click()
  cmdgo_Click
End Sub

Private Sub Picture8_Click()
  cmddeldive_Click
End Sub

Private Sub reloadgrid2()
Dim divelast As String
initialise_all
Label5.Caption = "These are singles dives. They assume no previous dive history. To add these dives into a mission sequence, use the Dive Series Planning features below."
divelast = "last"
MSFlexGrid2.Rows = 1
MSFlexGrid2.Cols = 9
MSFlexGrid2.Col = 0
MSFlexGrid2.Row = 0
MSFlexGrid2.Text = "Plan No."
MSFlexGrid2.Col = 1
MSFlexGrid2.Text = "Plan Date"
MSFlexGrid2.Col = 2
MSFlexGrid2.Text = "Max Depth"
MSFlexGrid2.Col = 3
MSFlexGrid2.Text = "Bottom mins"
MSFlexGrid2.Col = 4
MSFlexGrid2.Text = "Gas ID"
MSFlexGrid2.Col = 5
MSFlexGrid2.Text = "PPO2"
MSFlexGrid2.Col = 6
MSFlexGrid2.Text = "Open/Closed"
MSFlexGrid2.Col = 7
MSFlexGrid2.Text = "O2"
MSFlexGrid2.Col = 8
MSFlexGrid2.Text = "Helium"
MSFlexGrid2.ColWidth(0) = 1250
MSFlexGrid2.ColWidth(1) = 1140
MSFlexGrid2.ColWidth(2) = 940
MSFlexGrid2.ColWidth(3) = 750
MSFlexGrid2.ColWidth(4) = 0 '1040
MSFlexGrid2.ColWidth(5) = 840
MSFlexGrid2.ColWidth(6) = 1140
MSFlexGrid2.ColWidth(7) = 550
MSFlexGrid2.ColWidth(8) = 550
SQL = "SELECT * FROM seqdpprofile"
SQL = SQL & " order by Dpprofileid "
Set RS5 = DB.OpenRecordset(SQL) 'nick changed to RS5
 If RS5.EOF = True Then
     Exit Sub
     MsgBox "No Plan detected, please create some plan first"
     Splanmain.Show
 Else
  RS5.MoveFirst
     
     While RS5.EOF = False
     'tempdpid = RS5("Dpprofileid")
      If IsNull(RS5("Dpprofileid")) Then
        tempdpid = "SP00000001"
        RS5.Edit
        RS5!Dpprofileid = tempdid
        RS5.Update
      Else
        tempdpid = RS5("Dpprofileid")
      End If
     If tempdpid = divelast Then
        MSFlexGrid2.Col = 2
        If IsNumeric(RS5("depth")) = False Then
          tempdepth = "0.1"
        End If
        If tempdepth < RS5("depth") Then
          tempdepth = RS5("depth")
          If IsNumeric(tempdepth) = False Then
            tempdepth = "0.1"
          End If
          MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring) + "   "
        End If
        MSFlexGrid2.Col = 3
        MSFlexGrid2.Text = CStr(CInt(MSFlexGrid2.Text) + CInt(RS5("duration")))
     Else
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        k = MSFlexGrid2.Rows
        MSFlexGrid2.Row = k - 1
        MSFlexGrid2.Col = 0
        If IsNull(tempdpid) Then Else MSFlexGrid2.Text = tempdpid
        MSFlexGrid2.Col = 1
        SQL = "SELECT * FROM seqdpmain "
        SQL = SQL & "WHERE DiveplanID ='" & Trim(tempdpid) & "'  "
        Set RS4 = DB.OpenRecordset(SQL)
        tempdate = RS4("plandate")
        MSFlexGrid2.Text = Format(tempdate, "dd/mmm/yyyy")
        MSFlexGrid2.Col = 2
        tempdepth = RS5("depth")
        If IsNumeric(tempdepth) = False Then
          tempdepth = "0.1"
        End If
        MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring) + "   "
        MSFlexGrid2.Col = 3
        MSFlexGrid2.Text = RS5("duration")
        MSFlexGrid2.Col = 4
        MSFlexGrid2.Text = RS5("gasid")
        MSFlexGrid2.Col = 5
        MSFlexGrid2.Text = Format(RS5("po2"), "0.00") + "   "
        MSFlexGrid2.Col = 6
        MSFlexGrid2.Text = " " + RS5("dpcircuit")
        MSFlexGrid2.Col = 7
        MSFlexGrid2.Text = RS5("dpo2") + "   "
        MSFlexGrid2.Col = 8
        MSFlexGrid2.Text = RS5("dphe") + "   "
        divelast = tempdpid
     End If
     RS5.MoveNext
  Wend
For k = 1 To MSFlexGrid2.Rows - 1
  For p = 0 To 0
    MSFlexGrid2.Row = k
    MSFlexGrid2.Col = p
         checkslected = True
         If MSFlexGrid2.Row = 1 Then
            defaultcolor = &HE0E0E0
         Else
            If defaultcolor = &HE0E0E0 Then
                  defaultcolor = &HFFFFFF
            Else
                  defaultcolor = &HE0E0E0
            End If
         End If
         For H = 0 To 8
            MSFlexGrid2.Row = k
            MSFlexGrid2.Col = H
            If defaultcolor = &HE0E0E0 Then
               MSFlexGrid2.CellBackColor = &HE0E0E0    '&H00E0E0E0&
               MSFlexGrid2.CellForeColor = vbBlack
            Else
               MSFlexGrid2.CellBackColor = &HFFFFFF
               MSFlexGrid2.CellForeColor = vbBlack
            End If
          Next H
  Next p
  MSFlexGrid2.Col = 3
  MSFlexGrid2.Text = MSFlexGrid2.Text + "mins"
Next k
End If
If MSFlexGrid1.Rows > 1 Then
   If MSFlexGrid1.Rows < 12 Then MSFlexGrid1.TopRow = 1 Else MSFlexGrid1.TopRow = MSFlexGrid1.Rows - 8
   If MSFlexGrid2.Rows < 12 Then MSFlexGrid2.TopRow = 1 Else MSFlexGrid2.TopRow = MSFlexGrid2.Rows - 8
End If
MSFlexGrid2_Click
End Sub

