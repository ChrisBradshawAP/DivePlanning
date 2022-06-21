VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmseqdive 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Dive Series"
   ClientHeight    =   9585
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12660
   FillColor       =   &H00800000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmseqdive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Height          =   3375
      Left            =   12720
      TabIndex        =   245
      Top             =   1680
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   5953
      _Version        =   393216
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox safetytext 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   126
      Text            =   "0"
      ToolTipText     =   "Enter Safety Factor"
      Top             =   8760
      Width           =   615
   End
   Begin VB.TextBox atmtext 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   125
      Text            =   "1000"
      ToolTipText     =   "Enter expected Atmospheric pressure in mBar"
      Top             =   8400
      Width           =   615
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      Picture         =   "frmseqdive.frx":2CFA
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   123
      ToolTipText     =   "Edit selected Dive as new dive number"
      Top             =   680
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1080
      Picture         =   "frmseqdive.frx":369C
      ScaleHeight     =   345
      ScaleWidth      =   495
      TabIndex        =   122
      ToolTipText     =   "Edit this Dive"
      Top             =   740
      Width           =   495
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1680
      Picture         =   "frmseqdive.frx":3F22
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   121
      ToolTipText     =   "Delete this Dive"
      Top             =   720
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   1095
      Left            =   13680
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   1931
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   14737632
      BackColorBkg    =   8421376
      BorderStyle     =   0
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
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   325
      Index           =   0
      Left            =   10560
      Picture         =   "frmseqdive.frx":47A8
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   116
      ToolTipText     =   "Returns to previous screen"
      Top             =   120
      Visible         =   0   'False
      Width           =   335
   End
   Begin VB.CommandButton cmdinsert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6480
      Picture         =   "frmseqdive.frx":4E1A
      TabIndex        =   107
      ToolTipText     =   "Insert the plan above the selected plan"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdmodify 
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4440
      Picture         =   "frmseqdive.frx":66BC
      TabIndex        =   106
      ToolTipText     =   "Modifies the contents of the selected plan "
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5520
      Picture         =   "frmseqdive.frx":81B6
      TabIndex        =   105
      ToolTipText     =   "Insert the plan into the last position of the list"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtplanno 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtinterval 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      MaxLength       =   4
      TabIndex        =   13
      Top             =   3600
      Width           =   495
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Decompression Result"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3665
      Left            =   12840
      TabIndex        =   109
      Top             =   3840
      Width           =   5950
   End
   Begin VB.CommandButton cmdremove 
      BackColor       =   &H000000FF&
      Caption         =   "Remove from Series"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   104
      ToolTipText     =   "Remove selected dive from the series"
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   1
      Left            =   9840
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1125
      Width           =   520
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   2
      Left            =   9840
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1320
      Width           =   520
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   3
      Left            =   9840
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1515
      Width           =   520
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   4
      Left            =   9840
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1725
      Width           =   520
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   5
      Left            =   9840
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1920
      Width           =   520
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   6
      Left            =   9840
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   2115
      Width           =   520
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   7
      Left            =   9840
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2325
      Width           =   520
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   8
      Left            =   9840
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2520
      Width           =   520
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   9
      Left            =   9840
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2715
      Width           =   520
   End
   Begin VB.TextBox txtmaxdft 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   0
      Left            =   9840
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   915
      Width           =   520
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   855
      Left            =   16440
      TabIndex        =   9
      Top             =   7440
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   855
      Left            =   14520
      TabIndex        =   8
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dive Series Details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   4080
      TabIndex        =   0
      Top             =   11040
      Width           =   11895
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10200
         TabIndex        =   3
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton Cmdcreate 
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9480
         TabIndex        =   2
         ToolTipText     =   "Create New Sequential dive "
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdgenerate 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8520
         TabIndex        =   1
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   2295
         Left            =   6120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Text            =   "frmseqdive.frx":99B8
         Top             =   840
         Width           =   4815
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Height          =   3280
         Left            =   4320
         TabIndex        =   5
         Top             =   480
         Width           =   930
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "SAVE"
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CMDDELETE 
         Caption         =   "Command1"
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2490
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   4392
      _Version        =   393216
      Rows            =   200
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   14737632
      GridLines       =   0
      ScrollBars      =   2
      BorderStyle     =   0
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   2640
      TabIndex        =   110
      Top             =   4680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   10
      TabsPerRow      =   10
      TabHeight       =   520
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dive 1"
      TabPicture(0)   =   "frmseqdive.frx":99BE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblzerodepg(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbldepg(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbltimeg(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblzerotimeg(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblunitsg(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblminsg(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "decoresultgridlite(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "decoresultgrid(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Picture1(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Dive 2"
      TabPicture(1)   =   "frmseqdive.frx":99DA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1(1)"
      Tab(1).Control(1)=   "decoresultgrid(1)"
      Tab(1).Control(2)=   "decoresultgridlite(1)"
      Tab(1).Control(3)=   "lblminsg(1)"
      Tab(1).Control(4)=   "lblunitsg(1)"
      Tab(1).Control(5)=   "Label14(1)"
      Tab(1).Control(6)=   "lblzerodepg(1)"
      Tab(1).Control(7)=   "lbldepg(1)"
      Tab(1).Control(8)=   "lbltimeg(1)"
      Tab(1).Control(9)=   "lblzerotimeg(1)"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Dive 3"
      TabPicture(2)   =   "frmseqdive.frx":99F6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1(2)"
      Tab(2).Control(1)=   "decoresultgrid(2)"
      Tab(2).Control(2)=   "decoresultgridlite(2)"
      Tab(2).Control(3)=   "lblminsg(2)"
      Tab(2).Control(4)=   "lblunitsg(2)"
      Tab(2).Control(5)=   "Label14(2)"
      Tab(2).Control(6)=   "lblzerodepg(2)"
      Tab(2).Control(7)=   "lbldepg(2)"
      Tab(2).Control(8)=   "lbltimeg(2)"
      Tab(2).Control(9)=   "lblzerotimeg(2)"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Dive 4"
      TabPicture(3)   =   "frmseqdive.frx":9A12
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblzerotimeg(3)"
      Tab(3).Control(1)=   "lbltimeg(3)"
      Tab(3).Control(2)=   "lbldepg(3)"
      Tab(3).Control(3)=   "lblzerodepg(3)"
      Tab(3).Control(4)=   "Label14(3)"
      Tab(3).Control(5)=   "lblunitsg(3)"
      Tab(3).Control(6)=   "lblminsg(3)"
      Tab(3).Control(7)=   "decoresultgridlite(3)"
      Tab(3).Control(8)=   "decoresultgrid(3)"
      Tab(3).Control(9)=   "Picture1(3)"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Dive 5"
      TabPicture(4)   =   "frmseqdive.frx":9A2E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblzerotimeg(4)"
      Tab(4).Control(1)=   "lbltimeg(4)"
      Tab(4).Control(2)=   "lbldepg(4)"
      Tab(4).Control(3)=   "lblzerodepg(4)"
      Tab(4).Control(4)=   "Label14(4)"
      Tab(4).Control(5)=   "lblunitsg(4)"
      Tab(4).Control(6)=   "lblminsg(4)"
      Tab(4).Control(7)=   "decoresultgridlite(4)"
      Tab(4).Control(8)=   "decoresultgrid(4)"
      Tab(4).Control(9)=   "Picture1(4)"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "Dive 6"
      TabPicture(5)   =   "frmseqdive.frx":9A4A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblzerotimeg(5)"
      Tab(5).Control(1)=   "lbltimeg(5)"
      Tab(5).Control(2)=   "lbldepg(5)"
      Tab(5).Control(3)=   "lblzerodepg(5)"
      Tab(5).Control(4)=   "Label14(5)"
      Tab(5).Control(5)=   "lblunitsg(5)"
      Tab(5).Control(6)=   "lblminsg(5)"
      Tab(5).Control(7)=   "decoresultgridlite(5)"
      Tab(5).Control(8)=   "decoresultgrid(5)"
      Tab(5).Control(9)=   "Picture1(5)"
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Dive 7"
      TabPicture(6)   =   "frmseqdive.frx":9A66
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblzerotimeg(6)"
      Tab(6).Control(1)=   "lbltimeg(6)"
      Tab(6).Control(2)=   "lbldepg(6)"
      Tab(6).Control(3)=   "lblzerodepg(6)"
      Tab(6).Control(4)=   "Label14(6)"
      Tab(6).Control(5)=   "lblunitsg(6)"
      Tab(6).Control(6)=   "lblminsg(6)"
      Tab(6).Control(7)=   "decoresultgridlite(6)"
      Tab(6).Control(8)=   "decoresultgrid(6)"
      Tab(6).Control(9)=   "Picture1(6)"
      Tab(6).ControlCount=   10
      TabCaption(7)   =   "Dive 8"
      TabPicture(7)   =   "frmseqdive.frx":9A82
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblzerotimeg(7)"
      Tab(7).Control(1)=   "lbltimeg(7)"
      Tab(7).Control(2)=   "lbldepg(7)"
      Tab(7).Control(3)=   "lblzerodepg(7)"
      Tab(7).Control(4)=   "Label14(7)"
      Tab(7).Control(5)=   "lblunitsg(7)"
      Tab(7).Control(6)=   "lblminsg(7)"
      Tab(7).Control(7)=   "decoresultgridlite(7)"
      Tab(7).Control(8)=   "decoresultgrid(7)"
      Tab(7).Control(9)=   "Picture1(7)"
      Tab(7).ControlCount=   10
      TabCaption(8)   =   "Dive 9"
      TabPicture(8)   =   "frmseqdive.frx":9A9E
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lblzerotimeg(8)"
      Tab(8).Control(1)=   "lbltimeg(8)"
      Tab(8).Control(2)=   "lbldepg(8)"
      Tab(8).Control(3)=   "lblzerodepg(8)"
      Tab(8).Control(4)=   "Label14(8)"
      Tab(8).Control(5)=   "lblunitsg(8)"
      Tab(8).Control(6)=   "lblminsg(8)"
      Tab(8).Control(7)=   "decoresultgridlite(18)"
      Tab(8).Control(8)=   "decoresultgridlite(8)"
      Tab(8).Control(9)=   "decoresultgrid(8)"
      Tab(8).Control(10)=   "Picture1(8)"
      Tab(8).ControlCount=   11
      TabCaption(9)   =   "Dive 10"
      TabPicture(9)   =   "frmseqdive.frx":9ABA
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "lblzerotimeg(9)"
      Tab(9).Control(1)=   "lbltimeg(9)"
      Tab(9).Control(2)=   "lbldepg(9)"
      Tab(9).Control(3)=   "lblzerodepg(9)"
      Tab(9).Control(4)=   "Label14(9)"
      Tab(9).Control(5)=   "lblunitsg(9)"
      Tab(9).Control(6)=   "lblminsg(9)"
      Tab(9).Control(7)=   "decoresultgridlite(9)"
      Tab(9).Control(8)=   "decoresultgrid(9)"
      Tab(9).Control(9)=   "Picture1(9)"
      Tab(9).ControlCount=   10
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   9
         Left            =   -74640
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   224
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   480
            TabIndex        =   225
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   8
         Left            =   -74520
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   214
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   480
            TabIndex        =   215
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   7
         Left            =   -74520
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   204
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   480
            TabIndex        =   205
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   6
         Left            =   -74520
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   194
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   480
            TabIndex        =   195
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   5
         Left            =   -74520
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   184
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   480
            TabIndex        =   185
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   4
         Left            =   -74520
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   174
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   480
            TabIndex        =   175
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   3
         Left            =   -74520
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   164
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   480
            TabIndex        =   165
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   2
         Left            =   -74520
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   154
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   480
            TabIndex        =   155
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   1
         Left            =   -74520
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   144
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   480
            TabIndex        =   145
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   0
         Left            =   480
         ScaleHeight     =   2385
         ScaleWidth      =   3345
         TabIndex        =   112
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   480
            TabIndex        =   137
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   0
         Left            =   4080
         TabIndex        =   113
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   1
         Left            =   -70920
         TabIndex        =   146
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   2
         Left            =   -70920
         TabIndex        =   156
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   3
         Left            =   -70920
         TabIndex        =   166
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   4
         Left            =   -70920
         TabIndex        =   176
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   5
         Left            =   -70920
         TabIndex        =   186
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   6
         Left            =   -70920
         TabIndex        =   196
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   7
         Left            =   -70920
         TabIndex        =   206
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   8
         Left            =   -70920
         TabIndex        =   216
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   4095
         Index           =   9
         Left            =   -71160
         TabIndex        =   226
         Top             =   480
         Visible         =   0   'False
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   1
         Left            =   -70920
         TabIndex        =   234
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   2
         Left            =   -70920
         TabIndex        =   235
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   3
         Left            =   -70920
         TabIndex        =   236
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   4
         Left            =   -70920
         TabIndex        =   237
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   5
         Left            =   -70920
         TabIndex        =   238
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   6
         Left            =   -70920
         TabIndex        =   239
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   7
         Left            =   -70920
         TabIndex        =   240
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   8
         Left            =   -70920
         TabIndex        =   241
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   9
         Left            =   -71160
         TabIndex        =   242
         Top             =   480
         Visible         =   0   'False
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   0
         Left            =   4080
         TabIndex        =   243
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
         Height          =   4095
         Index           =   18
         Left            =   -70920
         TabIndex        =   244
         Top             =   480
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
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
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -73200
         TabIndex        =   233
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -75000
         TabIndex        =   232
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   9
         Left            =   -74760
         TabIndex        =   231
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   230
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -75000
         TabIndex        =   229
         Top             =   4080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -71640
         TabIndex        =   228
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -74640
         TabIndex        =   227
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -73080
         TabIndex        =   223
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   222
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   8
         Left            =   -74760
         TabIndex        =   221
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74640
         TabIndex        =   220
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   219
         Top             =   4080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -71520
         TabIndex        =   218
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74520
         TabIndex        =   217
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -73080
         TabIndex        =   213
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   212
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   7
         Left            =   -74760
         TabIndex        =   211
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74640
         TabIndex        =   210
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   209
         Top             =   4080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -71520
         TabIndex        =   208
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74520
         TabIndex        =   207
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -73080
         TabIndex        =   203
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   202
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   6
         Left            =   -74760
         TabIndex        =   201
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74640
         TabIndex        =   200
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   199
         Top             =   4080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -71520
         TabIndex        =   198
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74520
         TabIndex        =   197
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -73080
         TabIndex        =   193
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   192
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   5
         Left            =   -74760
         TabIndex        =   191
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -74640
         TabIndex        =   190
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   189
         Top             =   4080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -71520
         TabIndex        =   188
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -74520
         TabIndex        =   187
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -73080
         TabIndex        =   183
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   182
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   4
         Left            =   -74760
         TabIndex        =   181
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74640
         TabIndex        =   180
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   179
         Top             =   4080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -71520
         TabIndex        =   178
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74520
         TabIndex        =   177
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -73080
         TabIndex        =   173
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   172
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   3
         Left            =   -74760
         TabIndex        =   171
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   170
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   169
         Top             =   4080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -71520
         TabIndex        =   168
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   167
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -73080
         TabIndex        =   163
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   162
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   2
         Left            =   -74760
         TabIndex        =   161
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   160
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   159
         Top             =   4080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -71520
         TabIndex        =   158
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   157
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -73080
         TabIndex        =   153
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   152
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   1
         Left            =   -74760
         TabIndex        =   151
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   150
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   149
         Top             =   4080
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -71520
         TabIndex        =   148
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   147
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblminsg 
         BackStyle       =   0  'Transparent
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   143
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblunitsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   142
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblzerotimeg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   141
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbltimeg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   140
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbldepg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   139
         Top             =   4080
         Visible         =   0   'False
         Width           =   310
      End
      Begin VB.Label lblzerodepg 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   138
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   0
         Left            =   240
         TabIndex        =   120
         Top             =   480
         Width           =   3375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2730
      Left            =   240
      TabIndex        =   111
      Top             =   4680
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   4815
      _Version        =   393216
      Rows            =   200
      Cols            =   6
      FixedCols       =   0
      BackColor       =   16761024
      BackColorFixed  =   14737632
      BackColorBkg    =   16744576
      GridColor       =   14737632
      ScrollBars      =   2
      BorderStyle     =   0
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7800
      TabIndex        =   114
      Top             =   11160
      Width           =   2775
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   135
      Text            =   "frmseqdive.frx":9AD6
      Top             =   4680
      Visible         =   0   'False
      Width           =   8415
   End
   Begin MSComDlg.CommonDialog cmdlog 
      Left            =   0
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4lite 
      Height          =   3375
      Left            =   12600
      TabIndex        =   246
      Top             =   720
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   5953
      _Version        =   393216
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Plan Editor"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   247
      Top             =   3390
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   1
      Left            =   360
      TabIndex        =   136
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Series Construction"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   134
      Top             =   4245
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   133
      Top             =   8760
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Safety :"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   195
      TabIndex        =   132
      Top             =   8760
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Atmospheric :                      mBar"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   131
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Label lbllabel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   130
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date : "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   195
      TabIndex        =   129
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label lblseqdiveno 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   128
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dive Series No :"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   195
      TabIndex        =   127
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "New Dive"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   124
      ToolTipText     =   "Opens selected dive"
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Dive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   119
      ToolTipText     =   "Opens selected dive"
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Dive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8400
      TabIndex        =   118
      ToolTipText     =   "Deletes selected dive"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10920
      TabIndex        =   117
      ToolTipText     =   "Returns to previous screen"
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4a 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Plan No :"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " Interval:               hours"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   11655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dive List"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3360
      TabIndex        =   108
      Top             =   -120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   9840
      TabIndex        =   93
      Top             =   2715
      Width           =   525
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   9840
      TabIndex        =   92
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   9840
      TabIndex        =   91
      Top             =   2325
      Width           =   525
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   9840
      TabIndex        =   90
      Top             =   2115
      Width           =   525
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   9840
      TabIndex        =   89
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   9840
      TabIndex        =   88
      Top             =   1725
      Width           =   525
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   9840
      TabIndex        =   87
      Top             =   1515
      Width           =   525
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   9840
      TabIndex        =   86
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   9840
      TabIndex        =   85
      Top             =   1125
      Width           =   525
   End
   Begin VB.Label lbldepth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   9840
      TabIndex        =   84
      Top             =   915
      Width           =   525
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gas"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8880
      TabIndex        =   82
      Top             =   720
      Width           =   360
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gas Used"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10800
      TabIndex        =   81
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "PPO2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10365
      TabIndex        =   80
      Top             =   720
      Width           =   435
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "He"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9540
      TabIndex        =   79
      Top             =   720
      Width           =   315
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "O2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9240
      TabIndex        =   78
      Top             =   720
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   9240
      TabIndex        =   77
      Top             =   915
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   9240
      TabIndex        =   76
      Top             =   1125
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   9240
      TabIndex        =   75
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   9240
      TabIndex        =   74
      Top             =   1515
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   9240
      TabIndex        =   73
      Top             =   1725
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   9240
      TabIndex        =   72
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   9240
      TabIndex        =   71
      Top             =   2115
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   9240
      TabIndex        =   70
      Top             =   2325
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   9240
      TabIndex        =   69
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label lbl02 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   9240
      TabIndex        =   68
      Top             =   2715
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   9540
      TabIndex        =   67
      Top             =   915
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   9540
      TabIndex        =   66
      Top             =   1125
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   9540
      TabIndex        =   65
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   9540
      TabIndex        =   64
      Top             =   1515
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   9540
      TabIndex        =   63
      Top             =   1725
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   9540
      TabIndex        =   62
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   9540
      TabIndex        =   61
      Top             =   2115
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   9540
      TabIndex        =   60
      Top             =   2325
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   9540
      TabIndex        =   59
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label lblhe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   9540
      TabIndex        =   58
      Top             =   2715
      Width           =   315
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   10365
      TabIndex        =   57
      Top             =   915
      Width           =   435
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   10365
      TabIndex        =   56
      Top             =   1125
      Width           =   435
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   10365
      TabIndex        =   55
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   10365
      TabIndex        =   54
      Top             =   1515
      Width           =   435
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   10365
      TabIndex        =   53
      Top             =   1725
      Width           =   435
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   10365
      TabIndex        =   52
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   10365
      TabIndex        =   51
      Top             =   2115
      Width           =   435
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   10365
      TabIndex        =   50
      Top             =   2325
      Width           =   435
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   10365
      TabIndex        =   49
      Top             =   2520
      Width           =   435
   End
   Begin VB.Label lblppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   10365
      TabIndex        =   48
      Top             =   2715
      Width           =   435
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   10800
      TabIndex        =   47
      Top             =   915
      Width           =   1020
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   10800
      TabIndex        =   46
      Top             =   1125
      Width           =   1020
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   10800
      TabIndex        =   45
      Top             =   1320
      Width           =   1020
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   10800
      TabIndex        =   44
      Top             =   1515
      Width           =   1020
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   10800
      TabIndex        =   43
      Top             =   1725
      Width           =   1020
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   10800
      TabIndex        =   42
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   10800
      TabIndex        =   41
      Top             =   2115
      Width           =   1020
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   10800
      TabIndex        =   40
      Top             =   2325
      Width           =   1020
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   10800
      TabIndex        =   39
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Label lblgasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   10800
      TabIndex        =   38
      Top             =   2715
      Width           =   1020
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "MOD"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   37
      Top             =   720
      Width           =   525
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   8880
      TabIndex        =   36
      Top             =   2715
      Width           =   360
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   8880
      TabIndex        =   35
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   8880
      TabIndex        =   34
      Top             =   2325
      Width           =   360
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   8880
      TabIndex        =   33
      Top             =   2115
      Width           =   360
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   8880
      TabIndex        =   32
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   8880
      TabIndex        =   31
      Top             =   1725
      Width           =   360
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   8880
      TabIndex        =   30
      Top             =   1515
      Width           =   360
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   8880
      TabIndex        =   29
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   8880
      TabIndex        =   28
      Top             =   1125
      Width           =   360
   End
   Begin VB.Label gasn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   8880
      TabIndex        =   27
      Top             =   915
      Width           =   360
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   2415
      Index           =   1
      Left            =   12960
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   13800
      TabIndex        =   103
      Top             =   4320
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   13800
      TabIndex        =   102
      Top             =   5520
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   13800
      TabIndex        =   101
      Top             =   5325
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   13800
      TabIndex        =   100
      Top             =   5130
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   13800
      TabIndex        =   99
      Top             =   4920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   13800
      TabIndex        =   98
      Top             =   4725
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   13800
      TabIndex        =   97
      Top             =   4530
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   13800
      TabIndex        =   96
      Top             =   4125
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   13800
      TabIndex        =   95
      Top             =   3930
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblgasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   13680
      TabIndex        =   94
      Top             =   3720
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gas List"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   1
      Left            =   13800
      TabIndex        =   83
      Top             =   3915
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   12480
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2535
      Index           =   1
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dive Library"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   115
      Top             =   195
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   4335
      Index           =   0
      Left            =   120
      Top             =   4560
      Width           =   12135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   12135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   3
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   8640
      Width           =   12135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2535
      Index           =   0
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1935
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2535
      Index           =   2
      Left            =   9960
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   2655
      Index           =   2
      Left            =   120
      Top             =   480
      Width           =   12135
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   3
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   12135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   2
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   12135
   End
   Begin VB.Menu mnuseqdive 
      Caption         =   "&File"
      Begin VB.Menu mnunewseq 
         Caption         =   "&New Series"
      End
      Begin VB.Menu mnuseqsave 
         Caption         =   "&Save Series"
      End
      Begin VB.Menu mnuSDprint 
         Caption         =   "&Print Series"
      End
      Begin VB.Menu mnusavecsv 
         Caption         =   "Save Series as &CSV"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
      Begin VB.Menu mnuPrintCurrent 
         Caption         =   "Print &Current Dive"
      End
      Begin VB.Menu mnuPrintAll 
         Caption         =   "Print &All Dives"
      End
   End
   Begin VB.Menu mnuplan 
      Caption         =   "&Library"
      Begin VB.Menu Mnucreateplan 
         Caption         =   "&New Dive"
      End
      Begin VB.Menu mnupldelete 
         Caption         =   "&Delete Dive"
      End
      Begin VB.Menu mnuplanedit 
         Caption         =   "&Edit Dive"
      End
      Begin VB.Menu mnueditasnew 
         Caption         =   "Edit Dive as &New"
      End
      Begin VB.Menu mnuplansort 
         Caption         =   "&Sort Dive Plan"
         Begin VB.Menu mnusortbydepth 
            Caption         =   "By &Depth"
            Begin VB.Menu mnusortasec 
               Caption         =   "&Asecending"
            End
            Begin VB.Menu mnusortdesc 
               Caption         =   "&Descending"
               Checked         =   -1  'True
            End
         End
         Begin VB.Menu mnusortbyplano 
            Caption         =   "By &Plan No"
         End
      End
   End
   Begin VB.Menu mnuGenerate 
      Caption         =   "&Generate Deco"
      Begin VB.Menu mnuAutoGen 
         Caption         =   "&AutoGen"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnumanualgenerate 
         Caption         =   "&Manual Generate"
      End
   End
   Begin VB.Menu mnuVPMBdef 
      Caption         =   "&VPMB/Buhl"
      Enabled         =   0   'False
      Begin VB.Menu mnuVPMB 
         Caption         =   "&VPMB only"
         Index           =   0
      End
      Begin VB.Menu mnuVPMB 
         Caption         =   "V&PMB+Buhl"
         Index           =   1
      End
      Begin VB.Menu mnuVPMB 
         Caption         =   "&Buhlmann"
         Index           =   2
      End
   End
   Begin VB.Menu mnudecoversion 
      Caption         =   "&Schedule"
      Begin VB.Menu mnuProfessionalm 
         Caption         =   "&Professional"
      End
      Begin VB.Menu mnulitem 
         Caption         =   "&Lite"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnusetting 
      Caption         =   "&Settings"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu MNUTTIPS 
         Caption         =   "&Tool Tips"
         Begin VB.Menu mnuttipson 
            Caption         =   "&On"
         End
         Begin VB.Menu MNUTTIPSOFF 
            Caption         =   "0ff"
         End
      End
   End
   Begin VB.Menu mnufileexit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frmseqdive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nick data start here
'
 '=======================================================================
'     varying permeability model (vpm) vdecompression subroutine in fortran
'     with boyle's law compensation valgorithm (vpm-b)
'
'     author:  erik c. baker
'
'     "distribute freely - credit the authors"
'
'     this subroutine extends the 1986 vpm valgorithm (yount & hoffman) to i
'     vmixed gas, repetitive, and valtitude diving.  developments to the a
'     were made by david e. yount, eric b. maiken, and erik c. baker ove
'     period from 1999 to 2001.  this work is dedicated in remembrance o
'     professor david e. yount who passed away on april 27, 2000.
'
'     notes:
'     1.  this subroutine uses the sixteen (16) half-time compartments of t
'         buhlmann zh-l16 model.  the optional compartment 1b is used he
'         half-times of 1.88 vminutes for vhelium and 5.0 vminutes for nitr
'
'     2.  this subroutine uses various dec, ibm, and microsoft extensions w
'         may not be supported by all fortran compilers.  comments are m
'         a capital "c" in the first column or an exclamation point "!"
'         in a line after code.  an asterisk "*" in column 6 is a contin
'         of the previous line.  all code, except for line vnumbers, star
'         column 7.
'
'     3.  comments and suggestions for improvements are welcome.  please
'         respond by e-mail to:  ebaker@se.aeieng.com
'
'     acknowledgment:  thanks to kurt spaugh for recommendations on how
'     up the code.
'=======================================================================
'      implicit none
'=======================================================================
'     Global variables - main subroutine
'=======================================================================
Dim m As String * 1
Dim os_command As String * 3
Dim word As String * 7
Dim units As String * 3
'Dim Line1 As String * 70
Dim critical_volume_valgorithm As String * 3
Dim units_word1 As String * 4
Dim units_word2 As String * 7
Dim valtitude_dive_valgorithm As String * 3
'Dim i As Integer
'Dim j As Integer                                               'loop as integer
Dim vmonth As Integer
Dim vday As Integer
Dim vyear As Integer
Dim clock_hour As Integer
Dim vminute As Integer
Dim vnumber_of_vmixes As Integer
Dim vnumber_of_changes As Integer
Dim vprofile_code As Integer
Dim temptext, temptext2, temptext3, temptext4, temptext5, temptext6, temptext7, temptext8 As String
Dim vsegment_vnumber_start_of_ascent As Integer
Dim repetitive_dive_flag As Integer
Dim schedule_converged As Boolean
Dim critical_volume_valgorithm_off As Boolean
Dim valtitude_dive_valgorithm_off As Boolean
Dim ascent_vceiling_vdepth As Double
Dim vdeco_vstop_vdepth As Double
Dim vstep_size As Double
Dim sum_of_vfractions As Double
Dim sum_check As Double
Dim vdepth As Double
Dim ending_vdepth As Double
Dim starting_vdepth As Double
Dim rate As Double
Dim rounding_operation1 As Double
Dim run_vtime_end_of_vsegment As Double
Dim last_run_vtime As Double
Dim vstop_vtime As Double
Dim vdepth_start_of_vdeco_zone As Double
Dim rounding_operation2 As Double
Dim deepest_possible_vstop_vdepth As Double
Dim first_vstop_vdepth As Double
Dim critical_volume_comparison As Double
Dim next_vstop As Double
Dim run_vtime_start_of_vdeco_zone As Double
Dim critical_radius_vn2_microns As Double
Dim critical_radius_vhe_microns As Double
Dim run_vtime_start_of_ascent As Double
Dim valtitude_of_dive As Double
Dim vdeco_phase_volume_vtime As Double
Dim surface_interval_vtime As Double
Dim vpressure_other_gases_mmhg As Double
'New data 220704
Dim vdepth_change_new As Double
'=======================================================================
'     Global arrays - main subroutine
'=======================================================================
Dim vmix_change(20) As Integer
Dim vfraction_voxygen(20) As Double
Dim vdepth_change(20) As Double
Dim rate_change(20)  As Double
Dim vstep_size_change(20) As Double
Dim vsetPoint_Change(20)
Dim vhelium_half_vtime(17)  As Double
Dim vnitrogen_half_vtime(17) As Double
Dim vhe_vpressure_start_of_ascent(17) As Double
Dim vn2_vpressure_start_of_ascent(17) As Double
Dim vhe_vpressure_start_of_vdeco_zone(17) As Double
Dim vn2_vpressure_start_of_vdeco_zone(17) As Double
Dim phase_volume_vtime(17) As Double
Dim last_phase_volume_vtime(17) As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
Dim water_vapor_vpressure As Double
'common /block_8/ water_vapor_vpressure
Dim surface_tension_gamma As Double
Dim skin_compression_gammac As Double
'common /block_19/ surface_tension_gamma, skin_compression_gammac
Dim crit_volume_parameter_lambda As Double
'common /block_20/ crit_volume_parameter_lambda
Dim minimum_vdeco_vstop_vtime As Double
'common /block_21/ minimum_vdeco_vstop_vtime
Dim regeneration_vtime_constant As Double
'common /block_22/ regeneration_vtime_constant
Dim constant_vpressure_other_gases As Double
'common /block_17/ constant_vpressure_other_gases
Dim gradient_onset_of_imperm_atm As Double
'common /block_14/ gradient_onset_of_imperm_atm
'=======================================================================
'     global variables in named common blocks
'=======================================================================
Dim vsegment_vnumber As Integer
Dim run_vtime As Double
Dim vsegment_vtime As Double
'common /block_2/ run_vtime, vsegment_vnumber, vsegment_vtime
Dim starting_ambient_vpressure As Double
Dim ending_ambient_vpressure As Double
Dim ambient_vpressure As Double
'common /block_4/ ending_ambient_vpressure
Dim vmix_vnumber As Integer
'common /block_9/ vmix_vnumber
Dim barometric_vpressure As Double
'common /block_18/ barometric_vpressure
Dim units_equal_fsw As Boolean
Dim units_equal_msw As Boolean
'common /block_15/ units_equal_fsw, units_equal_msw
Dim units_factor As Double
'common /block_16/ units_factor

Dim SetPoint As Double '                             ! current setpoint in effect
Dim PO2_CCR As Double '                              ! current setpoint converted to fsw/msw
Dim FO2_CCR As Double '                               ! effective FO2 at start of waypoint
Dim Effective_FHE As Double '                          ! effective FHE at start of waypoint
Dim Effective_FN2 As Double '                          ! effective FN2 at start of waypoint
Dim InertSum_Diluent As Double '                     ! sum of Diluent fractions
Dim Is_CCR As Boolean '                            ! true if current leg is setpoint controlled
'
'     This variable is controlled by addition to VPMDECO.IN namelist ('program_settings')
'
Dim SetPoint_Is_Bar As Boolean '                   !
'      COMMON /CCR_Block/ SetPoint,PO2_CCR,Effective_FHE,Effective_FN2,
'     * FO2_CCR,InertSum_Diluent,Is_CCR,SetPoint_Is_Bar

'=======================================================================
'     global arrays in named common blocks
'=======================================================================
Dim vhelium_vtime_constant(17) As Double
'common /block_1a/ vhelium_vtime_constant
Dim vnitrogen_vtime_constant(17) As Double
'common /block_1b/ vnitrogen_vtime_constant
Dim vhelium_vpressure(17)  As Double
Dim vnitrogen_vpressure(17) As Double
'common /block_3/ vhelium_vpressure, vnitrogen_vpressure
Dim vfraction_vhelium(20)  As Double
Dim vfraction_vnitrogen(20) As Double
'common /block_5/ vfraction_vhelium, vfraction_vnitrogen
Dim initial_critical_radius_he(17) As Double
Dim initial_critical_radius_n2(17) As Double
'common /block_6/ initial_critical_radius_he,                                 initial_critical_radius_n2
Dim adjusted_critical_radius_he(17) As Double
Dim adjusted_critical_radius_n2(17) As Double
'common /block_7/ adjusted_critical_radius_he,                                adjusted_critical_radius_n2
Dim max_crushing_vpressure_he(17)  As Double
Dim max_crushing_vpressure_n2(17) As Double
'common /block_10/ max_crushing_vpressure_he,                                         max_crushing_vpressure_n2
Dim surface_phase_volume_vtime(17) As Double
'common /block_11/ surface_phase_volume_vtime
Dim max_actual_gradient(17) As Double
'common /block_12/ max_actual_gradient
Dim amb_vpressure_onset_of_imperm(17) As Double
Dim gas_tension_onset_of_imperm(17) As Double
'common /block_13/ amb_vpressure_onset_of_imperm,                               gas_tension_onset_of_imperm
Dim initial_vhelium_vpressure(17)  As Double
Dim initial_vnitrogen_vpressure(17) As Double
''common /block_23/ initial_vhelium_vpressure,                                          initial_vnitrogen_vpressure
Dim regenerated_radius_he(17)  As Double
Dim regenerated_radius_n2(17) As Double
'common /block_24/ regenerated_radius_he, regenerated_radius_n2
Dim adjusted_crushing_vpressure_he(17) As Double
Dim adjusted_crushing_vpressure_n2(17) As Double
'common /block_25/ adjusted_crushing_vpressure_he,                                    adjusted_crushing_vpressure_n2
Dim allowable_gradient_he(17)  As Double
Dim allowable_gradient_n2(17) As Double
'common /block_26/ allowable_gradient_he, allowable_gradient_n2
Dim initial_allowable_gradient_he(17) As Double
Dim initial_allowable_gradient_n2(17) As Double
'common /block_27/                                                     initial_allowable_gradient_he, initial_allowable_gradient_n2
Dim vdeco_gradient_he(17)  As Double
Dim vdeco_gradient_n2(17) As Double

Dim initial_inspired_vhe_vpressure As Double
Dim initial_inspired_vn2_vpressure As Double
Dim inspired_vhelium_vpressure As Double
Dim inspired_vnitrogen_vpressure As Double

'common /block_34/ vdeco_gradient_he, vdeco_gradient_n2
'=======================================================================
'     namelist for subroutine settings (read in from ascii text file)
'=======================================================================
'=======================================================================
'     assign half-time values to buhlmann compartment arrays
'=======================================================================
Dim Plan_Depth(1000) As Double
Dim Plan_Time(1000) As Double
Dim Plan_o2(1000) As Double
Dim Plan_he(1000) As Double
Dim Plan_OpenClosed(1000) As Integer 'String
Dim Plan_GasID(1000) As Integer
Dim Plan_PPo2(1000) As Double
Dim Plan_Gas_list_o2(1000) As Double
Dim Plan_Gas_list_he(1000) As Double
Dim Plan_Gas_list_n2(1000) As Double
Dim Plan_Gas_list_mod(1000) As Double
Dim Plan_Gas_list_used(1000) As Integer
Dim Plan_Gas_list_deco(1000) As Integer
Dim Plan_Gas_list_setpoint(1000) As Double
Dim Plan_Gas_list_numgasdeco As Integer
Dim Number_of_planpoints As Long

Dim current_vdepth As Double
Dim current_vmix_vnumber As Double

Dim Number_Dives As Integer

Dim buhlptoln2(16) As Double
Dim buhlptolhe(16) As Double
Dim an2(16) As Double
Dim ahe(16) As Double
Dim bn2(16) As Double
Dim bhe(16) As Double
    Dim nfrac As Double
    Dim hefrac As Double
    Dim atotal As Double
    Dim btotal As Double

Dim no_deco_found As Integer

'Nick data end here

Dim display_grid2 As Integer
Dim deco_update As Integer

Dim i As Integer
Dim tempplanno As String
Dim tempinterval As String
Dim checkgasusedselected, formstarted, profilerecordexist As Boolean
Dim row_count As Integer
Dim X(400) As Single
Dim Y(400) As Single
Dim xscale As Single
Dim yscale As Single
Dim runtime_graph As Single
Dim dum As Integer


Private Sub deleteseqdpmain()
 SQL = "select * FROM seqdplist "
 Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("seqdiveidmain")
   If tempdpid Like "T*" Then
      tempdpid2 = tempdpid
      RS3.Delete
   End If
   RS3.MoveNext
Wend
SQL = "select * FROM dpserialno "
Set RS3 = DB.OpenRecordset(SQL)
tempdpid1 = Right$(tempdpid2, 8)
tempdpid1 = tempdpid1 - 1
lengthsn = Len(tempdpid1)
  Select Case lengthsn
  Case 1
     tempdpid = "SM0000000" & tempdpid1
  Case 2
     tempdpid = "SM000000" & tempdpid1
  Case 3
     tempdpid = "SM00000" & tempdpid1
  Case 4
     tempdpid = "SM0000" & tempdpid1
  Case 5
     tempdpid = "SM000" & tempdpid1
  Case 6
     tempdpid = "SM00" & tempdpid1
  Case 7
     tempdpid = "SM0" & tempdpid1
  Case 8
     tempdpid = "SM" & tempdpid1
  End Select
RS3.Edit
RS3!seqdiveserialno = tempdpid
RS3.Update
RS3.Close
End Sub
Private Sub saveseqdpmain()
Screen.MousePointer = 11
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
      SQL = "select * FROM dpserialno "
      Set RS4 = DB.OpenRecordset(SQL)
      tempsediveid = RS4("seqdiveserialno")
      If tempsediveid Like "T*" Then
         tempdpid2 = Right(tempsediveid, 9)
         tempdpid2 = "S" & tempdpid2
         RS4.Edit
         RS4!seqdiveserialno = tempdpid2
         RS4.Update
      End If
    End If
   RS3.MoveNext
Wend
Screen.MousePointer = 0
End Sub
Private Sub reloadgriddata()
cleargriddata
If profilerecordexist = True Then
SQL = "SELECT * FROM seqdplist"
SQL = SQL & " where seqdiveidmain  = '" & tempseqdiveno & "' "
SQL = SQL & " order by seqdiveidmain, seqdiveidseq"
Set RS = DB.OpenRecordset(SQL)
RS.MoveFirst
MSFlexGrid1.Rows = 1
While RS.EOF = False
     MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
     K = MSFlexGrid1.Rows
     MSFlexGrid1.Row = K - 1
     MSFlexGrid1.Col = 0
     MSFlexGrid1.Text = K - 1
     MSFlexGrid1.Col = 1
     MSFlexGrid1.Text = RS("seqdiveid")
     MSFlexGrid1.Col = 2
     MSFlexGrid1.Text = RS("seqdiveidinterval")
     RS.MoveNext
Wend
Else
   cmdinsert.Visible = False
   cmdmodify.Visible = False
   cmdremove.Visible = False
   cmdSave.Enabled = False
   'cmdadd.Caption = "Create"
End If
  
End Sub
Private Sub griddataexist()
profilerecordexist = False
SQL = "SELECT COUNT(*) FROM seqdplist "
  SQL = SQL & " WHERE "
  SQL = SQL & " seqdiveidmain ='" & tempseqdiveno & "' "
  Set RS3 = DB.OpenRecordset(SQL)
  If RS3.Fields(0) <> 0 Then
     profilerecordexist = True
  End If
End Sub
Private Sub checkforchanges()
MSFlexGrid1.Row = rowindentified
For i = 1 To 2
Select Case i
   Case 1
      MSFlexGrid1.Col = i
      If Trim(txtplanno.Text) <> Trim(MSFlexGrid1.Text) Then
         DataChanged = "True"
      End If
   Case 2
      MSFlexGrid1.Col = i
      If Trim(txtinterval.Text) <> Trim(MSFlexGrid1.Text) Then
         DataChanged = "True"
      End If
End Select
Next
End Sub
Private Sub savechangerecord()
  For K = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = K
    saveseqrecord
  Next K
End Sub
Private Sub removerecord()
  SQL = "SELECT * FROM seqdplist"
  SQL = SQL & " where seqdiveidmain = '" & tempseqdiveno & "' "
  Set RS = DB.OpenRecordset(SQL)
  While RS.EOF = False
   RS.Delete
   RS.MoveNext
  Wend
End Sub
Private Sub saveprerowdata()
MSFlexGrid1.Row = i + 1
MSFlexGrid1.Col = 0
MSFlexGrid1.Text = i + 1
MSFlexGrid1.Col = 1
MSFlexGrid1.Text = tempplanno
MSFlexGrid1.Col = 2
MSFlexGrid1.Text = tempinterval
End Sub
Private Sub readprerowval()
  MSFlexGrid1.Row = i
  MSFlexGrid1.Col = 1
  tempplanno = MSFlexGrid1.Text
  MSFlexGrid1.Col = 2
  tempinterval = MSFlexGrid1.Text
End Sub
Private Sub saveseqrecord()
SQL = "SELECT * FROM seqdplist "
Set RS = DB.OpenRecordset(SQL)
   RS.AddNew
   RS!seqdiveidmain = tempseqdiveno
   MSFlexGrid1.Col = 0
   RS!seqdiveidseq = MSFlexGrid1.Text
   tempdiveidseq = MSFlexGrid1.Text
   MSFlexGrid1.Col = 1
   RS!seqdiveid = MSFlexGrid1.Text
   tempseqdiveid = MSFlexGrid1.Text
   MSFlexGrid1.Col = 2
   RS!seqdiveidinterval = MSFlexGrid1.Text
   RS!seqdiveidsafetyfac = safetytext.Text
   RS!seqdiveidatm = atmtext.Text
   RS.Update
   RS.Close
   If Len(tempseqdiveid) < 1 Then Exit Sub
   SQL = "SELECT * FROM seqdpmain"
   SQL = SQL & " where diveplanid = '" & tempseqdiveid & "' "
   Set RS5 = DB.OpenRecordset(SQL)
   RS5.MoveFirst
   While RS5.EOF = False
      tempmaxdepth = RS5("maxdepth")
      RS5.MoveNext
   Wend
   RS5.Close
   SQL = "SELECT * FROM seqdplist "
   SQL = SQL & " where seqdiveidmain = '" & tempseqdiveno & "' and "
   SQL = SQL & " seqdiveidseq = '" & tempdiveidseq & "' "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   RS!seqdiveidmaxdepth = tempmaxdepth
   RS.Update
   RS.Close
End Sub
Private Sub saveseqrecord2()
For i = 1 To MSFlexGrid1.Rows - 1

SQL = "SELECT * FROM seqdplist "
Set RS = DB.OpenRecordset(SQL)
   RS.AddNew
   RS!seqdiveidmain = lblseqdiveno.Caption
   MSFlexGrid1.Row = i
   MSFlexGrid1.Col = 0
   RS!seqdiveidseq = MSFlexGrid1.Text
   tempdiveidseq = MSFlexGrid1.Text
   MSFlexGrid1.Col = 1
   RS!seqdiveid = MSFlexGrid1.Text
   tempseqdiveid = MSFlexGrid1.Text
   MSFlexGrid1.Col = 2
   RS!seqdiveidinterval = MSFlexGrid1.Text
   RS!seqdiveidsafetyfac = safetytext.Text
   RS!seqdiveidatm = atmtext.Text
   RS.Update
   RS.Close
   If Len(tempseqdiveid) < 1 Then Exit Sub
   SQL = "SELECT * FROM seqdpmain"
   SQL = SQL & " where diveplanid = '" & tempseqdiveid & "' "
   Set RS5 = DB.OpenRecordset(SQL)
   RS5.MoveFirst
   While RS5.EOF = False
      tempmaxdepth = RS5("maxdepth")
      RS5.MoveNext
   Wend
   RS5.Close
   SQL = "SELECT * FROM seqdplist "
   SQL = SQL & " where seqdiveidmain = '" & tempseqdiveno & "' and "
   SQL = SQL & " seqdiveidseq = '" & tempdiveidseq & "' "
   Set RS5 = DB.OpenRecordset(SQL)
   RS5.Edit
   RS5!seqdiveidmaxdepth = tempmaxdepth
   RS5.Update
   RS5.Close
   
  Next i
End Sub
Private Sub updateseqserialno()
SQL = "SELECT * FROM dpserialno "
Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   
   RS!seqdiveserialno = tempseqdiveno
   RS.Update
End Sub
Private Sub loaddpprofiledata()
Dim i As Integer
On Error Resume Next
  
  For i = 0 To 9
    If i < (MSFlexGrid1.Rows - 1) Then SSTab1.TabVisible(i) = True Else SSTab1.TabVisible(i) = False
    Text2(deco_grid_display).Visible = False
  Next
  
'  Frame3.Caption = "Decompression Result for " & txtplanno.Text
  Frame6.Caption = "Gas Profile for " & txtplanno.Text
  Frame7.Caption = "Plan Profile for " & txtplanno.Text
  If display_grid2 = 1 Then
    Frame6.BackColor = &HC0FFFF    'vbGreen
    txtplanno.BackColor = &HC0FFFF
    Frame7.BackColor = &HC0FFFF    'vbGreen
    Frame6.Visible = True
   
  Else
    Frame6.BackColor = &HFFFFC0    'vbBlue
    txtplanno.BackColor = &HFFFFC0
    Frame7.BackColor = &HFFFFC0    'vbBlue
 
    view_graph_gaslist
  End If
'  If display_grid2 = 1 Then
'    Frame6.BackColor = &HC0FFFF    'vbGreen
'    txtplanno.BackColor = &HC0FFFF
'    Frame7.BackColor = &HC0FFFF    'vbGreen
'    Frame6.Visible = True
'    picture1(grid_num).Visible = False
'  Else
'    Frame6.BackColor = &HFFFFC0    'vbBlue
'    txtplanno.BackColor = &HFFFFC0
'    Frame7.BackColor = &HFFFFC0    'vbBlue
'    view_graph_gaslist
'  End If
  MSFlexGrid3.Rows = 1
  SQL = "SELECT * FROM seqdpprofile"
  SQL = SQL & " where dpprofileid = '" & txtplanno.Text & "' "
  SQL = SQL & " order by dpnumseq "
  Set RS = DB.OpenRecordset(SQL)
    RS.MoveFirst
  While RS.EOF = False
     MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
     K = MSFlexGrid3.Rows
     MSFlexGrid3.Row = K - 1
     MSFlexGrid3.RowHeight(K - 1) = 200
     MSFlexGrid3.Col = 0
     MSFlexGrid3.Text = RS("dpnumseq")
     MSFlexGrid3.Col = 1
     MSFlexGrid3.Text = RS("depth")
     MSFlexGrid3.Col = 2
     MSFlexGrid3.Text = RS("duration")
     MSFlexGrid3.Col = 3
     MSFlexGrid3.Text = RS("dpo2")
     MSFlexGrid3.Col = 4
     MSFlexGrid3.Text = RS("dphe")
     MSFlexGrid3.Col = 5
     MSFlexGrid3.Text = RS("po2")
     MSFlexGrid3.Col = 6
     MSFlexGrid3.Text = RS("dpcircuit")
     MSFlexGrid3.Col = 7
     MSFlexGrid3.Text = RS("gasid")
     MSFlexGrid3.Col = 8
     MSFlexGrid3.Text = Format(CStr(CDbl(RS("depth")) * feetormeter_factor), "###0.0")
     RS.MoveNext
   Wend
  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & " where dpmainid = '" & txtplanno.Text & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  While RS.EOF = False
     tempgasindex = RS("dpgasid")
     i = Right$(tempgasindex, 1)
     lblgasindex(i).Caption = tempgasindex
     tempnitrogen = RS("dpgasnitrogen")
     temphelium = RS("dpgashelium")
     lbl02(i).Caption = 100 - CInt(temphelium) - CInt(tempnitrogen)
     lblhe(i).Caption = temphelium
     tempdepth = RS("dpgasmaxopdepth")
     tempdepth = CInt(tempdepth) / 10
     lbldepth(i).Caption = Format(tempdepth, "###0")
     lblppo2(i).Caption = (CInt(lbl02(i).Caption) / 100) * ((CInt(lbldepth(i).Caption) / 10) + 1)
     lblppo2(i).Caption = Format(lblppo2(i).Caption, "###.00")
     lblgasused(i).Caption = RS("dpgasused")
     txtmaxdft(i).Text = Format(tempdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
     RS.MoveNext
   Wend
End Sub

Private Sub atmtext_Change()
  If IsNumeric(atmtext.Text) Then
  Else
    atmtext.Text = "1000"
  End If
End Sub

Private Sub atmtext_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    If CInt(atmtext.Text) < 400 Or CInt(atmtext.Text) > 1000 Then
     MsgBox "Value must be between 400 and 1000 mBar !"
     atmtext.Text = "1000"
    Else
         atmtext.SetFocus
         SendKeys "{HOME}+{END}"
         cmdgenerate_Click
    End If
   Else
    If KeyAscii <> 8 Then
      If KeyAscii < 46 Or KeyAscii > 57 Then
         MsgBox "Sorry, Only numeric characters allowed !"
         atmtext.SetFocus
         SendKeys "{HOME}+{END}"
      End If
    End If
   End If
End Sub

Private Sub atmtext_LostFocus()
  If CInt(atmtext.Text) < 400 Or CInt(atmtext.Text) > 1000 Then
     MsgBox "Value must be between 400 and 1000 mBar !"
     atmtext.Text = "1000"
  Else
    cmdgenerate_Click
  End If
End Sub

Private Sub cmdadd_Click()
If MSFlexGrid1.Rows > 10 Then
  MsgBox "Too many dives!!"
  Exit Sub
End If

cleargrid1
If Trim(txtplanno.Text) <> "" And CInt(txtinterval.Text) > 0 Then
   MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
   K = MSFlexGrid1.Rows
   MSFlexGrid1.Row = K - 1
   MSFlexGrid1.Col = 0
   MSFlexGrid1.Text = K - 1
   MSFlexGrid1.Col = 1
   MSFlexGrid1.Text = txtplanno.Text
   MSFlexGrid1.Col = 2
   MSFlexGrid1.Text = txtinterval.Text
   For q = 0 To 2
   MSFlexGrid1.Row = K - 1
   MSFlexGrid1.Col = q
   MSFlexGrid1.CellForeColor = vbWhite
   MSFlexGrid1.CellBackColor = vbBlue
   rowindentified = MSFlexGrid1.Row
   Next q
   MSFlexGrid1.Col = 1
   txtplanno.Text = MSFlexGrid1.Text
      MSFlexGrid1.Col = 2
   txtinterval.Text = MSFlexGrid1.Text
   loaddpprofiledata
  ' saveprorecord
  If CInt(MSFlexGrid1.Rows) > 1 Then
     cmdinsert.Visible = True
     cmdremove.Visible = True
     cmdmodify.Visible = False
     'cmdadd.Caption = "Add To end"
     cmdSave.Enabled = True
  Else
     cmdinsert.Visible = False
     'cmdadd.Caption = "Create"
  End If
  
updateseqserialno
saveseqrecord
clearhlgrid2
Else
   Title = "Error on System Validation.."
   MsgBox "Incomplete Profile Data !", 48, Title
End If
  
  cmdgenerate_Click 'nick
  display_deco_text 'nick
End Sub

Private Sub CMDCLOSE_Click()
Frame6.Visible = True
End Sub

Private Sub Cmdcreate_Click()
tempsnfound = "False"
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
   Title = "Dive Plan not Save.."
   ans = MsgBox("You have Dive plan that was not saved, " & Chr(13) & "Press No will remove all previous unsaved plans !" & Chr(13) & "Do you want to save now ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
   Select Case ans
   Case vbYes
      saveseqdpmain
   Case vbNo
      deleteseqdpmain
   End Select
End If
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
 Unload Me
 frmseqdive.Show
End Sub

Private Sub cmddelete_Click()
mnupldelete_Click
End Sub

Private Sub cmddetails_Click()
mnuplanedit_Click
End Sub

Private Sub cmdeditasnew_Click()
Mnucreateplan_Click
End Sub

Private Sub cmdgenerate_Click()
'  Frame6.Visible = False
  If mnuAutoGen.Checked = True Then
    vimportdb_data
  Else
    deco_update = 0
  End If
  'Sequence_deco
End Sub

Private Sub cmdinsert_Click()
If MSFlexGrid1.Rows > 10 Then
  MsgBox "Too many dives!!"
  Exit Sub
End If

rowchanged = rowindentified - 1
 If rowindentified <> "" Then
   totalrow = MSFlexGrid1.Rows - 1
   MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'   MsgBox totalrow
'   MsgBox rowchanged
   For i = totalrow To rowchanged Step -1
      If CInt(i) = CInt(rowchanged) Then
         MSFlexGrid1.Row = rowchanged + 1
         MSFlexGrid1.Col = 0
         MSFlexGrid1.Text = i + 1
         MSFlexGrid1.Col = 1
         MSFlexGrid1.Text = txtplanno.Text
         MSFlexGrid1.Col = 2
         MSFlexGrid1.Text = txtinterval.Text

      Else
         readprerowval ' read previous row value
         saveprerowdata
      End If
   Next i
   removerecord
   For K = 1 To MSFlexGrid1.Rows - 1
     MSFlexGrid1.Row = K
     saveseqrecord
  Next K
   For q = 0 To 2
      MSFlexGrid1.Row = rowindentified
      MSFlexGrid1.Col = q
      MSFlexGrid1.CellForeColor = vbWhite
      MSFlexGrid1.CellBackColor = vbBlue
   Next q
      MSFlexGrid1.Col = 1
      txtplanno.Text = MSFlexGrid1.Text
      MSFlexGrid1.Col = 2
      txtinterval.Text = MSFlexGrid1.Text
      rowindentified = MSFlexGrid1.Row
      loaddpprofiledata
Else
   Title = "Dive Profile"
   MsgBox "You must selected a record in the list to insert the sequence", 48, Title
End If
cmdmodify.Visible = False
clearhlgrid2
  cmdgenerate_Click 'nick
  display_deco_text 'nick
End Sub

Private Sub cmdmodify_Click()
DataChanged = "False"
checkforchanges
If DataChanged = "True" Then
   MSFlexGrid1.Row = rowindentified
   MSFlexGrid1.Col = 1
   MSFlexGrid1.Text = txtplanno.Text
   MSFlexGrid1.Col = 2
   MSFlexGrid1.Text = txtinterval.Text
   SQL = "SELECT * FROM seqdplist "
   SQL = SQL & " where seqdiveidmain = '" & tempseqdiveno & "' and seqdiveidseq = '" & rowindentified & "' "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   MSFlexGrid1.Col = 1
   RS("seqdiveid") = MSFlexGrid1.Text
   MSFlexGrid1.Col = 2
   RS("seqdiveidinterval") = MSFlexGrid1.Text
   RS.Update
   clearhlgrid2
   
   MSFlexGrid1.Col = 1
   
   txtplanno.Text = MSFlexGrid1.Text
   MSFlexGrid1.Col = 2
   txtinterval.Text = MSFlexGrid1.Text
   cmdgenerate_Click 'nick
   display_deco_text 'nick
End If

End Sub

Private Sub cmdremove_Click()
Screen.MousePointer = 11
numrow = MSFlexGrid1.Rows
Totalcount = numrow - 1
For K = 0 To Totalcount
    MSFlexGrid1.Row = K
    If MSFlexGrid1.CellBackColor = vbBlue Then
     MSFlexGrid1.Col = 0
     tempseq = MSFlexGrid1.Text
     MSFlexGrid1.Col = 1
     tempserialid = MSFlexGrid1.Text
  End If
Next K
SQL = "SELECT * FROM seqdplist "
SQL = SQL & "where seqdiveidseq = '" & tempseq & "'  and seqdiveid = '" & tempserialid & "' "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   RS3.Delete
   RS3.MoveNext
Wend
griddataexist
If profilerecordexist = False Then
  MSFlexGrid1.Rows = 1
  Screen.MousePointer = 0
  deco_update = 0
  For i = 0 To 9
     decoresultgrid(i).Rows = 0
  Next i
    Picture1(grid_num).Visible = False
    mnugaslist_Click
     cmdinsert.Visible = False
  cmdmodify.Visible = False
  cmdremove.Visible = False
  cmdSave.Enabled = False
  SSTab1.Visible = False '(0) = False
   Exit Sub
Else
  reloadgriddata
  removerecord
  savechangerecord
  MSFlexGrid1_Click
  cmdgenerate_Click 'nick
  display_deco_text 'nick
  cmdinsert.Visible = False
  cmdmodify.Visible = False
  cmdremove.Visible = True
  cmdSave.Enabled = True
End If
Screen.MousePointer = 0
End Sub

Private Sub cmdsave_Click()
saveseqdpmain
End Sub

Private Sub savemaxdepth()
  K = 1
  tempmaxdepth = 0
   MSFlexGrid3.Row = K
   MSFlexGrid3.Col = 1
   For p = K To MSFlexGrid3.Rows - 1
      MSFlexGrid3.Row = K
      MaxDepth = MSFlexGrid3.Text
      
      If CInt(tempmaxdepth) <= CInt(MaxDepth) Then
         tempmaxdepth = MaxDepth
         
      End If
      K = K + 1
   Next p
   SQL = "select * FROM seqdpmain "
   SQL = SQL & " where diveplanid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      tempmaxdepth = Format(tempmaxdepth, "###0.0")
      Select Case Len(tempmaxdepth)
      Case 3
         tempmaxdepth = "00" & tempmaxdepth
      Case 4
         tempmaxdepth = "0" & tempmaxdepth
      Case 5
         tempmaxdepth = tempmaxdepth
      End Select
      RS3.Edit
      RS3!MaxDepth = tempmaxdepth
      RS3.Update
      RS3.MoveNext
   Wend

End Sub
Private Sub saveprorecord()
SQL = "SELECT * FROM seqdpprofile"
Set RS = DB.OpenRecordset(SQL)
   RS.AddNew
   RS!Dpprofileid = tempserialno
   MSFlexGrid3.Col = 0
   RS!dpnumseq = MSFlexGrid3.Text
   MSFlexGrid3.Col = 1
   RS!depth = MSFlexGrid3.Text
   MSFlexGrid3.Col = 2
   RS!Duration = MSFlexGrid3.Text
   MSFlexGrid3.Col = 7
   RS!gasid = MSFlexGrid3.Text
   MSFlexGrid3.Col = 6
   RS!dpcircuit = MSFlexGrid3.Text
   MSFlexGrid3.Col = 5
   RS!po2 = MSFlexGrid3.Text
   MSFlexGrid3.Col = 3
   RS!dpo2 = MSFlexGrid3.Text
   MSFlexGrid3.Col = 4
   RS!dphe = MSFlexGrid3.Text
   
   RS.Update
End Sub

Private Sub reloadgrid2a()
MSFlexGrid2.Rows = 1
SQL = "SELECT * FROM seqdpmain"
SQL = SQL & " order by diveplanid "
Set RS5 = DB.OpenRecordset(SQL) 'nick changed to RS5
 If RS5.EOF = True Then
     Exit Sub
     MsgBox "No Plan detected, please create some plan first"
     Splanmain.Show
 Else
  RS5.MoveFirst
     
     While RS5.EOF = False
     MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
     K = MSFlexGrid2.Rows
     MSFlexGrid2.Row = K - 1
     MSFlexGrid2.Col = 0
     tempdpid = RS5("diveplanid")
      MSFlexGrid2.Text = tempdpid
      MSFlexGrid2.Col = 1
      tempdepth = RS5("maxdepth")
      If IsNumeric(tempdepth) = False Then
       tempdepth = "0.1"
      End If
      MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring)
    '
    ' tempdepth = RS5("maxdepth")
  '   List1.AddItem tempdpid
  '   List2.AddItem tempdepth
      RS5.MoveNext
'     End If
  Wend
End If
 End Sub
  
  
Private Sub Form_Load()
Dim i As Integer

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 30
'const float a[2][16] = {
  an2(0) = 1.2599
  an2(1) = 1#
  an2(2) = 0.8618
  an2(3) = 0.7562
  an2(4) = 0.6667
  an2(5) = 0.56
  an2(6) = 0.4947
  an2(7) = 0.45
  an2(8) = 0.4187
  an2(9) = 0.3798
  an2(10) = 0.3497
  an2(11) = 0.3223
  an2(12) = 0.285
  an2(13) = 0.2737
  an2(14) = 0.2523
  an2(15) = 0.2327
  ahe(0) = 1.7424
  ahe(1) = 1.383
  ahe(2) = 1.1919
  ahe(3) = 1.0458
  ahe(4) = 0.922
  ahe(5) = 0.8205
  ahe(6) = 0.7305
  ahe(7) = 0.6502
  ahe(8) = 0.595
  ahe(9) = 0.5545
  ahe(10) = 0.5333
  ahe(11) = 0.5189
  ahe(12) = 0.5181
  ahe(13) = 0.5176
  ahe(14) = 0.5172
  ahe(15) = 0.5119

'const float b[2][16] = {
  bn2(0) = 0.505
  bn2(1) = 0.6514
  bn2(2) = 0.7222
  bn2(3) = 0.7825
  bn2(4) = 0.8126
  bn2(5) = 0.8434
  bn2(6) = 0.8693
  bn2(7) = 0.891
  bn2(8) = 0.9092
  bn2(9) = 0.9222
  bn2(10) = 0.9319
  bn2(11) = 0.9403
  bn2(12) = 0.9477
  bn2(13) = 0.9544
  bn2(14) = 0.9602
  bn2(15) = 0.9653
  bhe(0) = 0.4245
  bhe(1) = 0.5747
  bhe(2) = 0.6527
  bhe(3) = 0.7223
  bhe(4) = 0.7582
  bhe(5) = 0.7957
  bhe(6) = 0.8279
  bhe(7) = 0.8553
  bhe(8) = 0.8757
  bhe(9) = 0.8903
  bhe(10) = 0.8997
  bhe(11) = 0.9073
  bhe(12) = 0.9122
  bhe(13) = 0.9171
  bhe(14) = 0.9217
  bhe(15) = 0.9267

  For i = 0 To 9
    Label14(i).Caption = ""
  Next i
If systemversion = "Pro" Then
   mnudecoversion.Visible = True
Else
   mnudecoversion.Visible = True
End If

buhl_mode = 1
mnuVPMB_Click (buhl_mode)
rowindentified = ""
display_grid2 = 0
deco_update = 0
lblseqdiveno = newseqdiveno
tempdate = Format$(Date, "dd/mmm/yyyy")
lbllabel.Caption = tempdate
'add plan list to list box
initialgrid
SQL = "SELECT * FROM seqdpmain"
SQL = SQL & " order by maxdepth desc"
  Set RS5 = DB.OpenRecordset(SQL) 'nick changed to RS5
 If RS5.EOF = True Then
    
     Exit Sub
    MsgBox "No Plan detected, please create a plan first"
    Splanmain.Show
 Else
reloadgrid2
  If tempchoice = "NSP" Then
  MSFlexGrid2.Col = 0
  txtinterval = "10"
  txtplanno = MSFlexGrid2.Text
  MSFlexGrid2_Click
  mnugaslist_Click
 End If
 If tempchoice = "SPP" Or tempchoice = "GSP" Then
  lblseqdiveno = tempseqdiveno
  SQL = "SELECT * FROM seqdplist "
  SQL = SQL & " where seqdiveidmain = '" & tempseqdiveno & "' "
  SQL = SQL & " order by seqdiveidseq "
  
  Set RS5 = DB.OpenRecordset(SQL)
  If RS5.EOF Then
    Unload Me
    tempchoice = "NSP"
    frmseqdive.Show
    Exit Sub
  End If
  RS5.MoveFirst
  While RS5.EOF = False
     MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
     K = MSFlexGrid1.Rows
     MSFlexGrid1.Row = K - 1
     MSFlexGrid1.Col = 0
     MSFlexGrid1.Text = K - 1
     MSFlexGrid1.Col = 1
     MSFlexGrid1.Text = RS5("seqdiveid")
     txtplanno.Text = MSFlexGrid1.Text
     MSFlexGrid1.Col = 2
     MSFlexGrid1.Text = RS5("seqdiveidinterval")
     txtinterval.Text = MSFlexGrid1.Text
     atmtext.Text = RS5("seqdiveidatm")
     safetytext.Text = RS5("seqdiveidsafetyfac")
     RS5.MoveNext
   
     loaddpprofiledata 'nick
   Wend
   For q = 0 To 2
   MSFlexGrid1.Row = K - 1
   MSFlexGrid1.Col = q
   MSFlexGrid1.CellForeColor = vbWhite
   MSFlexGrid1.CellBackColor = vbBlue
   Next q
   rowindentified = MSFlexGrid1.Row
      loaddpprofiledata
 End If
  If CInt(MSFlexGrid1.Rows) > 1 Then
    'cmdadd.Caption = "Add To End"
 Else
    'cmdadd.Caption = "Create"
 End If
 If rowindentified = "" Then
    cmdinsert.Visible = False
    cmdremove.Visible = False
    cmdmodify.Visible = False
    cmdSave.Enabled = False
 Else
    cmdinsert.Visible = True
    cmdremove.Visible = True
    cmdmodify.Visible = True
    cmdSave.Enabled = True
 End If
 If temptooltips = "Off" Then
    cmdadd.ToolTipText = ""
    cmdinsert.ToolTipText = ""
    cmdmodify.ToolTipText = ""
    cmdremove.ToolTipText = ""
    cmddetails.ToolTipText = ""
    cmdSave.ToolTipText = ""
    Cmdcreate.ToolTipText = ""
 End If
End If
  initialgrid4
  initialgrid4lite
  cmdgenerate_Click
  display_deco_text
  view_graph_gaslist
  
For j = 0 To 9
  If SSTab1.TabVisible(j) = False Then Exit For
  rowindentified = CStr(j + 1)
  display_deco_text
Next
End Sub
Private Sub initialgrid4()
MSFlexGrid4.Cols = 10
MSFlexGrid4.Col = 0
MSFlexGrid4.Rows = 1
MSFlexGrid4.Row = 0
MSFlexGrid4.Text = "No."
MSFlexGrid4.Col = 1
MSFlexGrid4.Text = "Duration"
MSFlexGrid4.Col = 2
MSFlexGrid4.Text = "RunTime"
MSFlexGrid4.Col = 3
MSFlexGrid4.Text = "Mix"
MSFlexGrid4.Col = 4
MSFlexGrid4.Text = "Depth"
MSFlexGrid4.Col = 6
MSFlexGrid4.Text = "CNS"
MSFlexGrid4.Col = 7
MSFlexGrid4.Text = "OTU"
MSFlexGrid4.Col = 8
MSFlexGrid4.Text = "Rate"
MSFlexGrid4.Col = 5
MSFlexGrid4.Text = "SPO2"

MSFlexGrid4.ColWidth(0) = 550
MSFlexGrid4.ColWidth(1) = 720
MSFlexGrid4.ColWidth(2) = 720
MSFlexGrid4.ColWidth(3) = 750
MSFlexGrid4.ColWidth(4) = 580
MSFlexGrid4.ColWidth(5) = 550
MSFlexGrid4.ColWidth(6) = 500
MSFlexGrid4.ColWidth(7) = 500
MSFlexGrid4.ColWidth(8) = 1000
End Sub
Private Sub initialgrid4lite()
MSFlexGrid4lite.Cols = 9
MSFlexGrid4lite.Col = 0
MSFlexGrid4lite.Rows = 1
MSFlexGrid4lite.Row = 0
MSFlexGrid4lite.Text = "No."
MSFlexGrid4lite.Col = 1
MSFlexGrid4lite.Text = "Duration"
MSFlexGrid4lite.Col = 2
MSFlexGrid4lite.Text = "RunTime"
MSFlexGrid4lite.Col = 3
MSFlexGrid4lite.Text = "Mix"
MSFlexGrid4lite.Col = 4
MSFlexGrid4lite.Text = "Depth"
MSFlexGrid4lite.Col = 6
MSFlexGrid4lite.Text = "CNS"
MSFlexGrid4lite.Col = 7
MSFlexGrid4lite.Text = "OTU"
MSFlexGrid4lite.Col = 8
MSFlexGrid4lite.Text = "Rate"
MSFlexGrid4lite.Col = 5
MSFlexGrid4lite.Text = "SPO2"

MSFlexGrid4lite.ColWidth(0) = 550
MSFlexGrid4lite.ColWidth(1) = 720
MSFlexGrid4lite.ColWidth(2) = 720
MSFlexGrid4lite.ColWidth(3) = 750
MSFlexGrid4lite.ColWidth(4) = 580
MSFlexGrid4lite.ColWidth(5) = 550
MSFlexGrid4lite.ColWidth(6) = 500
MSFlexGrid4lite.ColWidth(7) = 500
MSFlexGrid4lite.ColWidth(8) = 1000

End Sub
Private Sub initialgrid()
MSFlexGrid1.Cols = 3
MSFlexGrid1.Col = 0
MSFlexGrid1.Rows = 1
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "#"
MSFlexGrid1.Col = 1
MSFlexGrid1.Text = "Plan No"
MSFlexGrid1.Col = 2
MSFlexGrid1.Text = "Interval"
MSFlexGrid1.ColWidth(0) = 255
MSFlexGrid1.ColWidth(1) = 1210
MSFlexGrid1.ColWidth(2) = 750

MSFlexGrid3.Cols = 9
MSFlexGrid3.Col = 0
MSFlexGrid3.Rows = 1
MSFlexGrid3.Row = 0
MSFlexGrid3.Text = "#"
MSFlexGrid3.Col = 1
MSFlexGrid3.Text = "Depth"
MSFlexGrid3.Col = 2
MSFlexGrid3.Text = "Mins"
MSFlexGrid3.Col = 3
MSFlexGrid3.Text = "O2"
MSFlexGrid3.Col = 4
MSFlexGrid3.Text = "He"
MSFlexGrid3.Col = 5
MSFlexGrid3.Text = "Po2"
MSFlexGrid3.Col = 6
MSFlexGrid3.Text = "Circuit"
MSFlexGrid3.Col = 7
MSFlexGrid3.Text = "Gas Index"
MSFlexGrid3.Col = 8
MSFlexGrid3.Text = "Depth"
MSFlexGrid3.ColWidth(0) = 265
If feetormeter_feeton = 0 Then
  MSFlexGrid3.ColWidth(1) = 630
  MSFlexGrid3.ColWidth(8) = 0
Else
  MSFlexGrid3.ColWidth(1) = 0
  MSFlexGrid3.ColWidth(8) = 630
End If
MSFlexGrid3.ColWidth(2) = 630
MSFlexGrid3.ColWidth(3) = 280
MSFlexGrid3.ColWidth(4) = 280
MSFlexGrid3.ColWidth(5) = 530
MSFlexGrid3.ColWidth(6) = 950 '1200
MSFlexGrid3.ColWidth(7) = 0 '820
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then Me.WindowState = 0
If Me.WindowState = 0 Then
  Me.Width = 12500
  Me.Height = 10400 '11340
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
tempsnfound = "False"
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
'   MsgBox "Dive plan Saved !!"
Case vbNo
   deleteseqdpmain
End Select
End If
   Select Case previousform
   Case "SEQLIST"
      Splanmain.Show
   Case "SEQPLAN"
      If do_not_load = 1 Then Exit Sub
      Splanmain.Show
   End Select
End Sub

Private Sub List1_Click()

End Sub
Private Sub cleargrid1()
numrow = MSFlexGrid1.Rows
Totalcount = numrow - 1
rowindentified = MSFlexGrid1.Row
For K = 0 To Totalcount
  For p = 0 To 0
    MSFlexGrid1.Row = K
    MSFlexGrid1.Col = p
    If MSFlexGrid1.CellBackColor = vbBlue Then
      For H = 0 To 2
        MSFlexGrid1.Row = K
        MSFlexGrid1.Col = H
        MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
        MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor
      Next H
    End If
  Next p
Next K
End Sub
Private Sub clearhlgrid2()
numrow = MSFlexGrid2.Rows
Totalcount = numrow - 1
rowidentified2 = MSFlexGrid2.Row
For K = 0 To Totalcount
  For p = 0 To 0
    MSFlexGrid2.Row = K
    MSFlexGrid2.Col = p
    If MSFlexGrid2.CellBackColor = vbGreen Then
      For H = 0 To 7
        MSFlexGrid2.Row = K
        MSFlexGrid2.Col = H
        MSFlexGrid2.CellForeColor = MSFlexGrid2.ForeColor
        MSFlexGrid2.CellBackColor = MSFlexGrid2.BackColor
      Next H
    End If
  Next p
Next K
End Sub
Private Sub cleargriddata()
For K = 1 To MSFlexGrid1.Rows - 1
For p = 0 To 2
   MSFlexGrid1.Col = p
   MSFlexGrid1.Row = K
   MSFlexGrid1.Text = ""
Next p

Next K
End Sub

Private Sub Label11_Click()
  Unload Me
End Sub

Private Sub Label12_Click()
  mnupldelete_Click
End Sub

Private Sub Label13_Click()
  mnuplanedit_Click
End Sub

Private Sub Label15_Click()
  mnueditasnew_Click
End Sub

Private Sub mnuAutoGen_Click()
  If mnuAutoGen.Checked = True Then mnuAutoGen.Checked = False Else mnuAutoGen.Checked = True
End Sub

Private Sub Mnucreateplan_Click()
Unload Me
  previousform = "SEQLIST"
  tempchoice = "NPP"
  frmgasprofile2.Show
  Unload frmgasprofile2
End Sub

Private Sub mnueditasnew_Click()
SQL = "SELECT * FROM dpserialno"
Set RS = DB.OpenRecordset(SQL)
tempserialno = RS("lastseqdserialno")
tempserialno2 = Right(tempserialno, 8)
tempserialno = tempserialno2 + 1
lengthsn = Len(tempserialno)
  Select Case lengthsn
  Case 1
     tempserialno = "SP0000000" & tempserialno
  Case 2
     tempserialno = "SP000000" & tempserialno
  Case 3
    tempserialno = "SP00000" & tempserialno
  Case 4
    tempserialno = "SP0000" & tempserialno
  Case 5
    tempserialno = "SP000" & tempserialno
  Case 6
    tempserialno = "SP00" & tempserialno
  Case 7
    tempserialno = "SP0" & tempserialno
  Case 8
    tempserialno = "SP" & tempserialno
 End Select
  RS.Edit
  RS!lastseqdserialno = tempserialno
  RS.Update
  'lblserialno.Caption = tempserialno
SQL = "SELECT * FROM dpmaingaslist"
Set RS = DB.OpenRecordset(SQL)
For i = 0 To 9
   RS.AddNew
   RS!dpmainid = tempserialno
   RS!dpgasid = lblgasindex(i).Caption
   RS!dpgashelium = lblhe(i).Caption
   tempnitrogen = 100 - CInt(lblhe(i).Caption) - CInt(lbl02(i).Caption)
   RS!dpgasnitrogen = tempnitrogen
   RS!dpgasmaxopdepth = CInt(lbldepth(i).Caption) * 10
   RS!dpgaspo2setpoint = lblppo2(i).Caption
   RS!dpgasused = lblgasused(i).Caption
   RS.Update
Next i
SQL = "SELECT * FROM seqdpmain"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!diveplanid = tempserialno
RS.Update
For K = 1 To MSFlexGrid3.Rows - 1
    MSFlexGrid3.Row = K
    saveprorecord
  Next K
savemaxdepth
MsgBox "Saved as Dive : " & tempserialno & "  !"
previousform = "SEQPLAN"
do_not_load = 1
Unload Me
do_not_load = 0
Planprofile2.Show
End Sub

Private Sub mnufileexit_Click()
Unload Me
End Sub

Private Sub mnugaslist_Click()
End Sub

Private Sub mnugraph_Click()
End Sub

Private Sub mnulitem_Click()
systemversion = "Lite"
initialgrid4
initialgrid4lite
mnuProfessionalm.Checked = False
mnulitem.Checked = True
cmdgenerate_Click
display_deco_text
view_graph_gaslist
 For j = 0 To 9
    If systemversion = "Pro" Then
       decoresultgrid(j).Visible = True
       decoresultgridlite(j).Visible = False
    Else
       decoresultgrid(j).Visible = False
       decoresultgridlite(j).Visible = True
    End If
  Next
End Sub

Private Sub mnumanualgenerate_Click()
vimportdb_data 'cmdgenerate_Click
display_deco_text
End Sub

Private Sub mnunewseq_Click()
tempsnfound = "False"
tempchoice = "NSP"
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
   Title = "Dive Plan not Save.."
   ans = MsgBox("You have Dive plan that was not saved, " & Chr(13) & "Press No will remove all previous unsaved plans !" & Chr(13) & "Do you want to save now ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
   Select Case ans
   Case vbYes
      saveseqdpmain
   Case vbNo
      deleteseqdpmain
   End Select
End If
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
 Unload Me
 tempseqdiveno = "T" & Right(newseqdiveno, 9)
 frmseqdive.Show
End Sub

Private Sub mnuplanedit_Click()
tempseqdiveno = lblseqdiveno.Caption
tempserialno = txtplanno.Text
saveseqdpmain
previousform = "SEQPLAN"
If MSFlexGrid1.Rows > 1 And tempchoice = "NSP" Then
  tempchoice = "SPP"
End If
do_not_load = 1
Unload Me
do_not_load = 0
Planprofile2.Show
End Sub

Private Sub mnupldelete_Click()

If Trim(rowidentified2) <> "" Then
MSFlexGrid2.Row = rowidentified2
tempserialno = MSFlexGrid2.Text
tempseqduplicate = False
   If MSFlexGrid2.CellBackColor = vbGreen Then
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
            MsgBox "Dive deleted from database!"
            Screen.MousePointer = 0
         Case vbNo
            Me.MousePointer = 0
            Exit Sub
         End Select
      End If
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

Private Sub mnuPrintAll_Click()
  print_dives (10)
End Sub

Private Sub mnuPrintCurrent_Click()
  print_dives (SSTab1.Tab)
End Sub

Private Function print_dives(index As Integer)
Dim comptext As String
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim p As Integer
Dim q As Integer
On Error GoTo ErrorHandler:
  CommonDialog1.ShowPrinter

  If index > 9 Then
    q = 9
    j = 0
  Else
    q = index
    j = index
  End If
  For j = j To q
    If SSTab1.TabVisible(j) = False Then Exit For
    rowindentified = CStr(j + 1)
    Text1.Text = ""
    comptext = "DO NOT DIVE USING THESE TABLES. BETA SOFTWARE TESTING ONLY"
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = "Sequential Dive Serial No : " & lblseqdiveno.Caption
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = "  Atmospheric : " & atmtext.Text
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = "  Safety : " & safetytext.Text
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = ""
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = "Sequence Of the Sequential Dive : " & lblseqdiveno.Caption
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = ""
    For K = 0 To MSFlexGrid1.Rows - 1
      For i = 0 To MSFlexGrid1.Cols - 1
         MSFlexGrid1.Row = K
         MSFlexGrid1.Col = i
         rowtext = MSFlexGrid1.Text
         comptext = comptext + (rowtext + vbTab)
      Next i
     Text1.Text = Text1.Text + comptext + vbCrLf
     comptext = ""
    Next K
   
    comptext = ""
    Text1.Text = Text1.Text + comptext + vbCrLf
    display_deco_text
    comptext = vbCrLf + "Dive: " + CStr(j + 1)
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = Label14(j).Caption
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = "Gas #" & vbTab & "O2" & vbTab & "He" & vbTab & "Depth" & vbTab & "PPO2" & vbTab & "Gas Used"
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = ""
    For v = 0 To 9
      temptext = lblgasindex(v).Caption
      temptext2 = lbl02(v).Caption
      temptext3 = lblhe(v).Caption
      temptext4 = txtmaxdft(v).Text 'lbldepth(v).Caption
      temptext5 = lblppo2(v).Caption
      temptext6 = lblgasused(v).Caption
      comptext = temptext & vbTab & temptext2 & vbTab & temptext3 & vbTab & temptext4 & vbTab & temptext5 & vbTab & temptext6
      Text1.Text = Text1.Text + comptext + vbCrLf
      comptext = ""
    Next v
    comptext = ""
    Text1.Text = Text1.Text + comptext + vbCrLf
    'comptext = " " & Frame3.Caption
    'text1.text=text1.text +  comptext
    comptext = ""
    Text1.BackColor = &H808080
    Text1.Text = Text1.Text + comptext + vbCrLf
    comptext = ""
    If systemversion = "Pro" Then
       For K = 0 To decoresultgrid(j).Rows - 1
          For p = 0 To decoresultgrid(j).Cols - 2
             decoresultgrid(j).Row = K
             decoresultgrid(j).Col = p
             rowtext = decoresultgrid(j).Text
             rowtext = Left(rowtext, 8)
             If Len(rowtext) < 7 Then
                rowtext = rowtext + vbTab
             End If
             comptext = comptext + (rowtext + vbTab)
           Next p
           If K Mod 2 = 1 Then Text1.BackColor = vbYellow Else Text1.BackColor = vbRed
           If K Mod 2 = 1 Then Printer.FillColor = vbYellow Else Printer.FillColor = vbRed
           Text1.Text = Text1.Text + comptext + vbCrLf
           comptext = ""
        Next K
    Else
        For K = 0 To decoresultgridlite(j).Rows - 1
          For p = 0 To decoresultgridlite(j).Cols - 2
             decoresultgridlite(j).Row = K
             decoresultgridlite(j).Col = p
             rowtext = decoresultgridlite(j).Text
             rowtext = Left(rowtext, 8)
             If Len(rowtext) < 7 Then
                rowtext = rowtext + vbTab
             End If
             comptext = comptext + (rowtext + vbTab)
           Next p
           If K Mod 2 = 1 Then Text1.BackColor = vbYellow Else Text1.BackColor = vbRed
           If K Mod 2 = 1 Then Printer.FillColor = vbYellow Else Printer.FillColor = vbRed
           Text1.Text = Text1.Text + comptext + vbCrLf
           comptext = ""
        Next K
    End If
    Printer.Print Text1.Text
    Text1.Text = ""
    Printer.NewPage
      
  Next
  Printer.EndDoc
ErrorHandler:
   MsgBox "Printer error !!"

End Function


Private Sub mnuProfessionalm_Click()
systemversion = "Pro"
initialgrid4
initialgrid4lite
mnuProfessionalm.Checked = True
mnulitem.Checked = False
cmdgenerate_Click
display_deco_text
view_graph_gaslist
End Sub

Private Sub mnusavecsv_Click()
Dim comptext As String
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim p As Integer
Dim q As Integer

 On Error GoTo ErrorHandler2
    cmdlog.Action = 2
 Open cmdlog.FileName For Output As #1
  comptext = "DO NOT DIVE USING THESE TABLES. BETA SOFTWARE TESTING ONLY"
   Print #1, comptext
  comptext = "Sequential Dive Serial No : " & lblseqdiveno.Caption & "  Atmospheric : " & atmtext.Text & "  Safety : " & safetytext.Text
   Print #1, comptext
   comptext = ""
   Print #1, comptext
   comptext = "Sequence Of the Sequential Dive : " & lblseqdiveno.Caption
   Print #1, comptext
   comptext = ""
   For K = 0 To MSFlexGrid1.Rows - 1
      For j = 0 To MSFlexGrid1.Cols - 1
         MSFlexGrid1.Row = K
         MSFlexGrid1.Col = j
         rowtext = MSFlexGrid1.Text
         comptext = comptext + (rowtext + ",")
      Next j
    Print #1, comptext
    comptext = ""
   Next K
   
   comptext = ""
   Print #1, comptext
  j = rowindentified
  For j = 0 To 9
    If SSTab1.TabVisible(j) = False Then Exit For
    rowindentified = CStr(j + 1)
    display_deco_text
    comptext = vbCrLf + "Dive: " + CStr(j + 1)
    Print #1, comptext
    comptext = Label14(j).Caption
    Print #1, comptext
    comptext = "Gas Index" & "," & "O2" & "," & "He" & "," & "Depth" & "," & "PPO2" & "," & "Gas Used"
    Print #1, comptext
    comptext = ""
    For v = 0 To 9
      temptext = lblgasindex(v).Caption
      temptext2 = lbl02(v).Caption
      temptext3 = lblhe(v).Caption
      temptext4 = txtmaxdft(v).Text 'lbldepth(v).Caption
      temptext5 = lblppo2(v).Caption
      temptext6 = lblgasused(v).Caption
      comptext = temptext & "," & temptext2 & "," & temptext3 & "," & temptext4 & "," & temptext5 & "," & temptext6
      Print #1, comptext
      comptext = ""
    Next v
    comptext = ""
    Print #1, comptext
    'comptext = " " & Frame3.Caption
    'Print #1, comptext
    comptext = ""
    Print #1, comptext
    comptext = ""
    If systemversion = "Pro" Then
       For K = 0 To decoresultgrid(j).Rows - 1
          For p = 0 To decoresultgrid(j).Cols - 1
             decoresultgrid(j).Row = K
             decoresultgrid(j).Col = p
             rowtext = decoresultgrid(j).Text
             comptext = comptext + (rowtext + ",")
          Next p
          Print #1, comptext
          comptext = ""
       Next K
    Else
       For K = 0 To decoresultgridlite(j).Rows - 1
          For p = 0 To decoresultgridlite(j).Cols - 1
             decoresultgridlite(j).Row = K
             decoresultgridlite(j).Col = p
             rowtext = decoresultgridlite(j).Text
             comptext = comptext + (rowtext + ",")
          Next p
          Print #1, comptext
          comptext = ""
       Next K
    End If
  Next
  Close #1
  MsgBox "Data saved to CSV file....!!"

ErrorHandler2:
    If Err = 32755 Then
        MsgBox "Data is not saved to a file....!!"
        Exit Sub
    End If
End Sub

Private Sub mnuSDprint_Click()
  On Error GoTo ErrorHandler
 cmdlog.Action = 5
    Printer.FontSize = 15
  '  Printer.FontName = "Courier"
   Printer.FontName = "Arial"
    Printer.Print ""
    Printer.CurrentX = xPos
    Printer.CurrentY = yPos
    xPos = 100
    yPos = 100
    Printer.CurrentX = xPos
    Printer.CurrentY = yPos
    
    Printer.Print "DO NOT DIVE USING THESE TABLE. BETA TESTING ONLY  Sequential Dive Serial No : " & lblseqdiveno.Caption & "  Atmospheric : " & atmtext.Text & "  Safety : " & safetytext.Text
    Printer.FontSize = 12
    Printer.Print Spc(3);
    
    xPos = 100
    yPos = 740
    Printer.CurrentX = xPos
    Printer.CurrentY = yPos
    Printer.Print "Sequence Of the Sequential Dive :"
    Printer.Print ""
    printgrid1
    
Printer.EndDoc

ErrorHandler:
    
    If Err = 32755 Then
        MsgBox "Cancel printing ... "
        Exit Sub
    End If


End Sub

Private Sub mnuseqdelete_Click()

End Sub

Private Sub mnuseqsave_Click()
Screen.MousePointer = 11
SQL = "select * FROM seqdplist "
Set RS3 = DB.OpenRecordset(SQL)
tempseqduplicate = False
While RS3.EOF = False
   tempsediveid = RS3("seqdiveidmain")
   If tempsediveid = Trim(lblseqdiveno.Caption) Then
      tempseqduplicate = True
   End If
   RS3.MoveNext
Wend
If tempseqduplicate = True Then
   SQL = "select * FROM seqdplist "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      tempsediveid = RS3("seqdiveidmain")
      If tempsediveid = Trim(lblseqdiveno.Caption) Then
         RS3.Delete
      End If
     ' If tempsediveid = tempseqdiveno Then
     '    RS3.Delete
     ' End If
      RS3.MoveNext
   Wend
   saveseqrecord2
Else
  ' If tempchoice = "NSP" Then
  '    saveseqrecord2
  ' Else
      saveseqdpmain
  ' End If
End If
  Screen.MousePointer = 0
End Sub
Private Sub mnusortasec_Click()
lastsortplanid = txtplanno.Text
mnusortbyplano.Checked = False
mnusortasec.Checked = True
mnusortdesc.Checked = False
sortorder = "Asec"
changegrid2
End Sub

Private Sub mnusortbyplano_Click()
MSFlexGrid2.Col = 0
lastsortplanid = MSFlexGrid2.Text
mnusortasec.Checked = False
mnusortdesc.Checked = False
mnusortbyplano.Checked = True
changegrid2
End Sub
Private Sub changegrid2()

If mnusortasec.Checked = True Then
  SQL = "SELECT * FROM seqdpmain"
  SQL = SQL & " order by maxdepth "
  Set RS5 = DB.OpenRecordset(SQL)
End If
If mnusortdesc.Checked = True Then
  SQL = "SELECT * FROM seqdpmain"
  SQL = SQL & " order by maxdepth desc "
  Set RS5 = DB.OpenRecordset(SQL)
End If
If mnusortbyplano.Checked = True Then
   SQL = "SELECT * FROM seqdpmain"
   SQL = SQL & " order by diveplanid "
   Set RS5 = DB.OpenRecordset(SQL)
End If
   If RS5.EOF = True Then
      Exit Sub
      MsgBox "No Plan detected, please create a plan first"
      Splanmain.Show
   Else
      MSFlexGrid2.Rows = 1
      RS5.MoveFirst
      While RS5.EOF = False
         MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
         K = MSFlexGrid2.Rows
         MSFlexGrid2.Row = K - 1
         MSFlexGrid2.Col = 0
         tempdpid = RS5("diveplanid")
         MSFlexGrid2.Text = tempdpid
         MSFlexGrid2.Col = 1
         tempdepth = RS5("maxdepth")
         If IsNumeric(tempdepth) = False Then
            tempdepth = "0.1"
         Else
            MSFlexGrid2.Text = tempdpid
         End If
         MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring)
         SQL = "SELECT * FROM seqdpprofile"
         SQL = SQL & " where dpprofileid = '" & tempdpid & "' "
         SQL = SQL & " order by dpnumseq "
         Set RS6 = DB.OpenRecordset(SQL)
         If RS6.EOF = True Then
            Exit Sub
            MsgBox "No Plan detected, please create some plan first"
            Splanmain.Show
         Else
            RS6.MoveFirst
            While RS6.EOF = False
               MSFlexGrid2.Col = 1
               If IsNumeric(RS6("depth")) = False Then
                  tempdepth = "0.1"
               End If
               If tempdepth < RS6("depth") Then
                  tempdepth = RS6("depth")
                  If IsNumeric(tempdepth) = False Then
                     tempdepth = "0.1"
                  End If
                  MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring) + "   "
               End If
                  'MSFlexGrid2.Col = 2
                  'MSFlexGrid2.Text = CStr(CInt(MSFlexGrid2.Text) + CInt(RS6("duration")))
               MSFlexGrid2.Col = 2
               If MSFlexGrid2.Text <> "" Then
                  MSFlexGrid2.Text = CStr(CInt(MSFlexGrid2.Text) + CInt(RS6("duration")))
               Else
                  MSFlexGrid2.Text = RS6("duration")
               End If
               MSFlexGrid2.Col = 3
               MSFlexGrid2.Text = RS6("gasid")
               MSFlexGrid2.Col = 4
               MSFlexGrid2.Text = Format(RS6("po2"), "0.00") + "   "
               MSFlexGrid2.Col = 5
               MSFlexGrid2.Text = "   " + RS6("dpcircuit")
               MSFlexGrid2.Col = 6
               MSFlexGrid2.Text = RS6("dpo2") ' + "   "
               MSFlexGrid2.Col = 7
               MSFlexGrid2.Text = RS6("dphe") ' + "   "
               RS6.MoveNext
            Wend
          End If
'          MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
          RS5.MoveNext
       Wend
    End If
    For K = 1 To MSFlexGrid2.Rows - 1
       For p = 0 To 0
          MSFlexGrid2.Row = K
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
          For H = 0 To 7
             MSFlexGrid2.Row = K
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
          MSFlexGrid2.Col = 2
          MSFlexGrid2.Text = MSFlexGrid2.Text + "mins   "
     Next K
For K = 1 To MSFlexGrid2.Rows - 1
   For p = 0 To 0
      MSFlexGrid2.Row = K
      MSFlexGrid2.Col = p 'End If
      If MSFlexGrid2.Text = lastsortplanid Then
         MSFlexGrid2_Click
         Exit Sub
      End If
   Next p
Next K
'rowidentified2 = MSFlexGrid2.Rows - 1
'MSFlexGrid2_Click
         
End Sub

Private Sub mnusortdesc_Click()
MSFlexGrid2.Col = 0
lastsortplanid = MSFlexGrid2.Text
mnusortdesc.Checked = True
mnusortbyplano.Checked = False
mnusortasec.Checked = False
sortorder = "Dsec"
changegrid2
End Sub

Private Sub mnuVPMB_Click(index As Integer)
Dim i As Integer
Dim j As Integer
Dim st As String

  mnuVPMB(0).Checked = False
  mnuVPMB(1).Checked = False
  mnuVPMB(2).Checked = False
  mnuVPMB(index).Checked = True
  mnuVPMBdef.Caption = mnuVPMB(index).Caption
  If buhl_mode = index Then
  Else
    buhl_mode = index
    SQL = "SELECT * FROM dpserialno"
    Set RS = DB.OpenRecordset(SQL)
    RS.Edit
    RS!buhl = CStr(buhl_mode)
    RS.Update
    vimportdb_data
  End If
  For i = 0 To 9
    If InStr(1, Label14(i).Caption, "Deco", vbTextCompare) Then
      st = ""
      st = Left(Label14(i).Caption, InStr(1, Label14(i).Caption, "Deco", vbTextCompare) - 1)
      st = st + "Deco Algorithm: " + mnuVPMB(buhl_mode).Caption ' + vbCrLf
      st = st + Right(Label14(i).Caption, Len(Label14(i).Caption) - InStr(InStr(1, Label14(i).Caption, "Deco", vbTextCompare), Label14(i).Caption, vbCrLf, vbTextCompare) + 1)
      Label14(i).Caption = st
    End If
  Next i
End Sub

Private Sub MSFlexGrid1_Click()
clearhlgrid2
numrow = MSFlexGrid1.Rows
If numrow > 1 Then
Totalcount = numrow - 1
rowindentified = MSFlexGrid1.Row
For K = 0 To Totalcount
  For p = 0 To 0
    MSFlexGrid1.Row = K
    MSFlexGrid1.Col = p
    If MSFlexGrid1.CellBackColor = vbBlue Then
      For H = 0 To 2
        MSFlexGrid1.Row = K
        MSFlexGrid1.Col = H
        MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
        MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor
      Next H
    End If
  Next p
Next K
For q = 0 To 2
   MSFlexGrid1.Row = rowindentified
   MSFlexGrid1.Col = q
   MSFlexGrid1.CellForeColor = vbWhite
   MSFlexGrid1.CellBackColor = vbBlue
Next q
For p = 1 To 2
  MSFlexGrid1.Row = rowindentified
  MSFlexGrid1.Col = p
  Select Case p
  Case 1
     txtplanno.Text = MSFlexGrid1.Text
     clearhlgrid2
     For T = 0 To MSFlexGrid2.Rows - 1
        MSFlexGrid2.Row = T
        MSFlexGrid2.Col = 0
        If MSFlexGrid2.Text = txtplanno.Text Then
           MSFlexGrid2_Click
'           For r = 0 To 7
'              MSFlexGrid2.Col = r
'              MSFlexGrid2.CellForeColor = vbWhite
'              MSFlexGrid2.CellBackColor = vbGreen
'           Next r
        End If
    Next T
  Case 2
     txtinterval.Text = MSFlexGrid1.Text
  End Select
Next p
  loaddpprofiledata
  cmdinsert.Visible = True
  cmdremove.Visible = True
  cmdmodify.Visible = False
  cmdSave.Enabled = True
 
  'cmdgenerate_Click
  display_deco_text
   For j = 0 To 9
    If systemversion = "Pro" Then
       decoresultgrid(j).Visible = True
       decoresultgridlite(j).Visible = False
    Else
       decoresultgrid(j).Visible = False
       decoresultgridlite(j).Visible = True
    End If
  Next
Else
  MsgBox "No Series Contruction yet, Please add one !"
End If
End Sub

'Nick code start here

Private Sub Command1_Click()
  'atmtext.Text
  'safetytext.Text
  vimportdb_data
  'Sequence_deco
  display_deco_graph (0)
End Sub


'Nick code start here

Private Sub Sequence_deco()
Dim Planpoint As Integer
Dim ratelasttemp As Double

'If vimportdb_data > 0 Then Exit Sub
no_deco_found = 0
cleardecogrid

 run_vtime = 0#
 vsegment_vtime = 0
 vsegment_vnumber = 0
 run_vtime_end_of_vsegment = 0

If repetitive_dive_flag < 0 Then
 vhelium_half_vtime(1) = 1.88: vhelium_half_vtime(2) = 3.02: vhelium_half_vtime(3) = 4.72: vhelium_half_vtime(4) = 6.99: vhelium_half_vtime(5) = 10.21: vhelium_half_vtime(6) = 14.48: vhelium_half_vtime(7) = 20.53: vhelium_half_vtime(8) = 29.11: vhelium_half_vtime(9) = 41.2: vhelium_half_vtime(10) = 55.19: vhelium_half_vtime(11) = 70.69: vhelium_half_vtime(12) = 90.34: vhelium_half_vtime(13) = 115.29: vhelium_half_vtime(14) = 147.42: vhelium_half_vtime(15) = 188.24: vhelium_half_vtime(16) = 240.03
 vnitrogen_half_vtime(1) = 5#: vnitrogen_half_vtime(2) = 8#: vnitrogen_half_vtime(3) = 12.5: vnitrogen_half_vtime(4) = 18.5: vnitrogen_half_vtime(5) = 27#: vnitrogen_half_vtime(6) = 38.3: vnitrogen_half_vtime(7) = 54.3: vnitrogen_half_vtime(8) = 77#: vnitrogen_half_vtime(9) = 109#: vnitrogen_half_vtime(10) = 146#: vnitrogen_half_vtime(11) = 187#: vnitrogen_half_vtime(12) = 239#: vnitrogen_half_vtime(13) = 305#: vnitrogen_half_vtime(14) = 390#: vnitrogen_half_vtime(15) = 498#: vnitrogen_half_vtime(16) = 635#
 '=======================================================================
 '     open files for subroutine input/output
 '=======================================================================
 '       open (unit = 7, file = 'vpmvdeco.in', status = 'unknown',                   access = 'sequential', form = 'formatted')
 '       open (unit = 8, file = 'vpmvdeco.out', status = 'unknown',                  access = 'sequential', form = 'formatted')
 '       open (unit = 10, file = 'vpmvdeco.set', status = 'unknown',                 access = 'sequential', form = 'formatted')
 '=======================================================================
 '     begin subroutine execution with output message to screen
 '=======================================================================
 'Open "report.txt" For Append As #1
 cleardecogrid
 't1print (" ")                                'to ms operating sys
 't1print ("subroutine vpmvdeco")
 't1print (" ")                          'asterisk indicates t1print( t
 't1print (vbCrLf)
 '=======================================================================
 '     read in subroutine settings and check for errors
 '     if there are errors, write an error message and terminate subroutine
 '=======================================================================
 'read #1,
 units = "msw"
 valtitude_dive_valgorithm = "on" '"off"
 minimum_vdeco_vstop_vtime = 1.0000001
 critical_radius_vn2_microns = 0.6 + (CDbl(safetytext.Text) * 0.6 / 100)
 critical_radius_vhe_microns = 0.5 + (CDbl(safetytext.Text) * 0.5 / 100)
 critical_volume_valgorithm = "on"
 crit_volume_parameter_lambda = 7500#
 gradient_onset_of_imperm_atm = 8.2 ' - (CDbl(safetytext.Text) * 8.2 / 100)
 surface_tension_gamma = 0.0179
 skin_compression_gammac = 0.257
 regeneration_vtime_constant = 20160#
 vpressure_other_gases_mmhg = 102#
 If (InStr(1, units, "fsw")) Then
     units_equal_fsw = (True)
     units_equal_msw = (False)
 Else
     units_equal_fsw = (False)
     units_equal_msw = (True)
 End If
 
 
 'If ((units = "fsw") Or (units = "fsw")) Then
 '    units_equal_fsw = (True)
 '    units_equal_msw = (False)
 'ElseIf ((units = "msw") Or (units = "msw")) Then
 '    units_equal_fsw = (False)
 '    units_equal_msw = (True)
 'Else
 '    'Call systemqq(os_command)
 '    no_deco_found=3 ' MsgBox "root not in brackets"
 'End If
 If (InStr(1, valtitude_dive_valgorithm, "on")) Then
     valtitude_dive_valgorithm_off = (False)
 Else 'If ((valtitude_dive_valgorithm = "off") Or (valtitude_dive_valgorithm = "off")) Then
     valtitude_dive_valgorithm_off = (True)
 End If
 'If ((valtitude_dive_valgorithm = "on") Or (valtitude_dive_valgorithm = "on")) Then
 '    valtitude_dive_valgorithm_off = (False)
 'ElseIf ((valtitude_dive_valgorithm = "off") Or (valtitude_dive_valgorithm = "off")) Then
 '    valtitude_dive_valgorithm_off = (True)
 'Else
 '    no_deco_found=3 ' MsgBox "root not in brackets"
 'End If
 If ((critical_radius_vn2_microns < 0.2) Or (critical_radius_vn2_microns > 1.35)) Then
     no_deco_found = 3 ' MsgBox "root not in brackets"
 End If
 If ((critical_radius_vhe_microns < 0.2) Or (critical_radius_vhe_microns > 1.35)) Then
     no_deco_found = 3 ' MsgBox "root not in brackets"
 End If
 If (InStr(1, critical_volume_valgorithm, "on ")) Then
     critical_volume_valgorithm_off = (False)
 Else 'If ((critical_volume_valgorithm = "off") Or (critical_volume_valgorithm = "off")) Then
     critical_volume_valgorithm_off = (True)
 End If
 'If ((critical_volume_valgorithm = "on ") Or (critical_volume_valgorithm = "on")) Then
 '    critical_volume_valgorithm_off = (False)
 'ElseIf ((critical_volume_valgorithm = "off") Or (critical_volume_valgorithm = "off")) Then
 '    critical_volume_valgorithm_off = (True)
 'Else
 '    no_deco_found=3 ' MsgBox "root not in brackets"
 'End If
 '=======================================================================
 '     initialize constants/variables based on selection of units - fsw o
 '     fsw = feet of seawater, a unit of vpressure
 '     msw = meters of seawater, a unit of vpressure
 '=======================================================================
 If (units_equal_fsw) Then
     units_word1 = "fswg"
     units_word2 = "fsw/min"
     units_factor = 33#
     water_vapor_vpressure = 1.607     'based on respiratory quotien
     '(schreiner value)
 End If
 If (units_equal_msw) Then
     units_word1 = "mswg"
     units_word2 = "msw/min"
     units_factor = 10.1325
     water_vapor_vpressure = 0.493     'based on respiratory quotien
 End If                               '(schreiner value)
 '=======================================================================
 '     initialize constants/variables
 '=======================================================================
 constant_vpressure_other_gases = (vpressure_other_gases_mmhg / 760#) * units_factor
 run_vtime = 0#
 vsegment_vtime = 0
 vsegment_vnumber = 0
 run_vtime_end_of_vsegment = 0
 For i = 1 To 16
     vhelium_vtime_constant(i) = Log(2#) / vhelium_half_vtime(i)
     vnitrogen_vtime_constant(i) = Log(2#) / vnitrogen_half_vtime(i)
     max_crushing_vpressure_he(i) = 0#
     max_crushing_vpressure_n2(i) = 0#
     max_actual_gradient(i) = 0#
     surface_phase_volume_vtime(i) = 0#
     amb_vpressure_onset_of_imperm(i) = 0#
     gas_tension_onset_of_imperm(i) = 0#
     initial_critical_radius_n2(i) = critical_radius_vn2_microns * 0.000001
     initial_critical_radius_he(i) = critical_radius_vhe_microns * 0.000001
 Next i
 '=======================================================================
 '     initialize variables for sea level or valtitude dive
 '     see subroutines for explanation of valtitude calculations.  purpose
 '     1) to determine barometric vpressure and 2) set or adjust the vpm c
 '     radius variables and gas loadings, as applicable, based on altitud
 '     ascent to valtitude before the dive, and time at valtitude before th
 '=======================================================================

End If 'initialise

If (valtitude_dive_valgorithm_off) Then
    valtitude_of_dive = 0#
    Call calc_barometric_vpressure(valtitude_of_dive)            'su
    t1print CStr("Alt")
    t1print CStr(valtitude_of_dive)
    t1print CStr("Atmmospheric")
    t1print CStr(barometric_vpressure)
    t1print (vbCrLf)
    If repetitive_dive_flag < 0 Then
     For i = 1 To 16
        adjusted_critical_radius_n2(i) = initial_critical_radius_n2(i)
        adjusted_critical_radius_he(i) = initial_critical_radius_he(i)
        vhelium_vpressure(i) = 0#
        vnitrogen_vpressure(i) = (barometric_vpressure - water_vapor_vpressure) * 0.79
     Next i
    End If
Else
    If repetitive_dive_flag < 0 Then Call vpm_valtitude_dive_valgorithm                           'su
    t1print CStr("Alt")
    t1print CStr(valtitude_of_dive)
    t1print CStr("Atmmospheric")
    t1print CStr(barometric_vpressure)
    t1print (vbCrLf)
End If

'=======================================================================
'     start of repetitive dive loop
'     this is the largest loop in the main subroutine and operates between
'     30 and 330.  if there is one or more repetitive dives, the subroutine
'     return to this point to process each repetitive dive.
'=======================================================================
L30:
Do While (True) 'loop will run continuous


    'there is an exit stateme
    '=======================================================================
    '     input dive description and gas vmix data from ascii text input file
    '     begin writing headings/output to ascii text output file
    ' '     see separate explanation of format for input file.
    '=======================================================================
    'read (7,805) line1
    'Call clock(vyear, vmonth, vday, clock_hour, vminute, m)            'su
    'print, vmonth, vday, vyear, clock_hour, vminute, m
    'print, line1
    vnumber_of_vmixes = 10
    For i = 1 To vnumber_of_vmixes
        vfraction_vnitrogen(i) = Plan_Gas_list_n2(i)
        vfraction_vhelium(i) = Plan_Gas_list_he(i)
        vfraction_voxygen(i) = 1# - vfraction_vnitrogen(i) - vfraction_vhelium(i)
        i = i
        'If i = 1 Then
        '  vfraction_voxygen(i) = 0.15
        '  vfraction_vhelium(i) = 0.45
        '  vfraction_vnitrogen(i) = 0.4
        'End If
        'If i = 2 Then
        '  vfraction_voxygen(i) = 0.36
        '  vfraction_vhelium(i) = 0#
        '  vfraction_vnitrogen(i) = 0.64
        'End If
        'If i = 3 Then
        '  vfraction_voxygen(i) = 0.99
        '  vfraction_vhelium(i) = 0#
        '  vfraction_vnitrogen(i) = 0.01
        'End If
        sum_of_vfractions = vfraction_voxygen(i) + vfraction_vhelium(i) + vfraction_vnitrogen(i)
        sum_check = sum_of_vfractions
        If (sum_check <> 1#) Then
            no_deco_found = 3 ' MsgBox "root not in brackets"
        End If
    Next i
    For j = 1 To vnumber_of_vmixes
        'print, j, vfraction_voxygen(j), vfraction_vhelium(j), vfraction_vnitrogen(j)
    Next j
    'print, units_word1, units_word1, units_word2, units_word1
    '=======================================================================
    '     dive profile loop - input dive profile data from ascii text input
    '     and process dive as a series of ascent/descent and constant vdepth
    '     vsegments.  this allows for multi-level dives and unusual profiles.
    '     gas loadings for each vsegment.  if it is a descent vsegment, calc c
    '     vpressure on critical radii in each compartment.
    '     "instantaneous" descents are not used in the vpm.  all ascent/desc
    '     vsegments must have a realistic rate of ascent/descent.  unlike hal
    '     models, the vpm is actually more conservative when the descent rat
    '     slower becuase the effective crushing vpressure is reduced.  also,
    '     realistic actual supersaturation gradient must be calculated durin
    '     ascents as this affects critical radii adjustments for repetitive
    '     profile codes: 1 = ascent/descent, 2 = constant vdepth, 99 = vdecomp
    '=======================================================================
    vprofile_code = 0
    Planpoint = 0
    initialgrid4
    initialgrid4lite
    t1print ("   #     Dur      RT     Mix   Depth    SDep    EDep    Rate    " + vbCrLf)
    Do While Planpoint <= Number_of_planpoints                     'loop will run continuous
        'there is an exit stateme
        'vprofile_code = vprofile_code + 1
        If Planpoint = 0 Or vprofile_code = 2 Then 'get next point and start with rate change
          vprofile_code = 1
          vmix_vnumber = Plan_GasID(Planpoint)
          If (Planpoint = 0) Then
            Planpoint = Planpoint + 1
            vmix_vnumber = Plan_GasID(Planpoint)
            starting_vdepth = 0#
            ending_vdepth = Plan_Depth(Planpoint)
'            vdepth = Plan_Depth(Planpoint)
            vdepth = 0
            If Plan_OpenClosed(Planpoint) = 1 Then
              SetPoint = Plan_PPo2(Planpoint)
            Else
              SetPoint = 0#
            End If
            rate = 20
          Else
            starting_vdepth = Plan_Depth(Planpoint - 1)
            ending_vdepth = Plan_Depth(Planpoint)
            If Plan_OpenClosed(Planpoint) = 1 Then
              SetPoint = Plan_PPo2(Planpoint)
            Else
              SetPoint = 0#
            End If
            If starting_vdepth < ending_vdepth Then
              rate = 20
            Else
              If starting_vdepth > ending_vdepth Then
                rate = -10
              Else
                vprofile_code = 2 'constant depth
                vdepth = Plan_Depth(Planpoint)
                run_vtime_end_of_vsegment = run_vtime_end_of_vsegment + Plan_Time(Planpoint)
                vmix_vnumber = Plan_GasID(Planpoint)
                'rate = 0
              End If
            End If
          End If
        Else
                vprofile_code = 2 'constant depth
                vdepth = Plan_Depth(Planpoint)
                run_vtime_end_of_vsegment = run_vtime_end_of_vsegment + Plan_Time(Planpoint)
                vmix_vnumber = Plan_GasID(Planpoint)
                'rate = 0
        End If
        
        If vprofile_code > 2 Then vprofile_code = 99
        If (vprofile_code = 1) Then
            'starting_vdepth = 0#
            'ending_vdepth = 80#
            'rate = 20
            'vmix_vnumber = 1
            setUpSetpoint (units_factor)
            Call gas_loadings_ascent_descent(starting_vdepth, ending_vdepth, rate)
            If (ending_vdepth > starting_vdepth) Then
                Call calc_crushing_vpressure(starting_vdepth, ending_vdepth, rate)
            End If
            If (ending_vdepth > starting_vdepth) Then
                'word = "descent"
                ElseIf (starting_vdepth > ending_vdepth) Then
                'word = "ascent "
                Else
                'word = "error"
            End If
            t1print8 CStr(vsegment_vnumber)
            t1print8 CStr(vsegment_vtime)
            t1print8 CStr(run_vtime)
            t1print8 CStr(vmix_vnumber - 1)
            t1print8 CStr(vdepth * feetormeter_factor) & feetormeter_shortstring
            t1print8 CStr(starting_vdepth * feetormeter_factor) & feetormeter_shortstring
            t1print8 CStr(ending_vdepth * feetormeter_factor) & feetormeter_shortstring
            t1print8 CStr(rate)
            t1print (vbCrLf)
           ' If systemversion = "Lite" Then
               printtogridlite
           ' Else
               printtogrid
           ' End If
'            t1print (CStr(vsegment_vnumber) + CStr(vsegment_vtime) + CStr(run_vtime) + CStr(vmix_vnumber) + CStr(starting_vdepth) + CStr(ending_vdepth) + CStr(rate) + vbCrLf)
        ElseIf (vprofile_code = 2) Then
            'vdepth = 80#
            'run_vtime_end_of_vsegment = 30#
            'vmix_vnumber = 1
            If (run_vtime_end_of_vsegment - run_vtime) <= 0 Then
              MsgBox "Segment time too short Ascent/Descent at Segment: " + CStr(CInt(vsegment_vnumber / 2) + 1)
              'vhighlite_line (CInt(vsegment_vnumber / 2) + 1)
              Exit Sub
            End If
            setUpSetpoint (units_factor)
            Call gas_loadings_constant_vdepth(vdepth, run_vtime_end_of_vsegment)
            t1print8 CStr(vsegment_vnumber)
            t1print8 CStr(vsegment_vtime)
            t1print8 CStr(run_vtime)
            t1print8 CStr(vmix_vnumber - 1)
            t1print8 CStr(vdepth * feetormeter_factor) & feetormeter_shortstring
'            t1print8 cstr(rate)
           ' t1print (vbCrLf)
         '  If systemversion = "Lite" Then
              printtogrid2lite
         '  Else
              printtogrid2
         '  End If
            Planpoint = Planpoint + 1
        ElseIf (vprofile_code = 99) Then
                Exit Do
        Else
              no_deco_found = 3 ' MsgBox "root not in brackets"
        End If
        current_vdepth = vdepth
        current_vmix_vnumber = vmix_vnumber
        
    Loop 'Next Planpoint
    '=======================================================================
    '     begin process of ascent and vdecompression
    '     first, calculate the regeneration of critical radii that takes pla
    '     the dive time.  the regeneration time constant has a time scale of
    '     so this will have very little impact on dives of normal length, bu
    '     have major impact for saturation dives.
    '=======================================================================
    Call nuclear_regeneration(run_vtime)
    '=======================================================================
    '     calculate initial allowable gradients for ascent
    '     this is based on the maximum effective crushing vpressure on critic
    '     in each compartment achieved during the dive profile.
    '=======================================================================
    Call calc_initial_allowable_gradient
    '=======================================================================
    '     save variables at start of ascent (end of bottom time) since these
    '     be used later to compute the final ascent profile that is written
    '     output file.
    '     the vpm uses an iterative process to compute vdecompression schedul
    '     there will be more than one pass through the vdecompression loop.
    '=======================================================================
    For i = 1 To 16
        vhe_vpressure_start_of_ascent(i) = vhelium_vpressure(i)
        vn2_vpressure_start_of_ascent(i) = vnitrogen_vpressure(i)
    Next i
    run_vtime_start_of_ascent = run_vtime
    vsegment_vnumber_start_of_ascent = vsegment_vnumber
    '=======================================================================
    '     input parameters to be used for staged vdecompression and save in a
    '     assign inital parameters to be used at start of ascent
    '     the user has the ability to change vmix, ascent rate, and step size
    '     combination at any vdepth during the ascent.
    '=======================================================================
    vnumber_of_changes = Plan_Gas_list_numgasdeco
    For i = 1 To vnumber_of_changes
        If i = 1 Then
          vdepth_change(i) = current_vdepth
          vmix_change(i) = current_vmix_vnumber
          rate_change(i) = -10#
          vstep_size_change(i) = feetormeter_decostep
          vsetPoint_Change(i) = SetPoint
        Else
          vdepth_change(i) = Plan_Gas_list_mod(Plan_Gas_list_deco(i)) '33#
          vmix_change(i) = Plan_Gas_list_deco(i)
          rate_change(i) = -10#
          vstep_size_change(i) = feetormeter_decostep
          If Plan_Gas_list_used(Plan_Gas_list_deco(i)) = 5 Then 'Plan_Gas_list_numgasdeco)) = 5 Then
            vsetPoint_Change(i) = Plan_Gas_list_setpoint(Plan_Gas_list_deco(i))
          Else
            vsetPoint_Change(i) = 0
          End If
        End If
        'If i = 3 Then
        '  vdepth_change(i) = 6#
        '  vmix_change(i) = 3
        '  rate_change(i) = -3#
        '  vstep_size_change(i) = 3
        'End If
    Next i
    starting_vdepth = vdepth_change(1)
    vmix_vnumber = vmix_change(1)
    rate = rate_change(1)
    vstep_size = vstep_size_change(1)
    SetPoint = vsetPoint_Change(1)
    setUpSetpoint (units_factor)
      
    '=======================================================================
    '     calculate the vdepth where the vdecompression zone begins for this p
    '     based on the initial ascent parameters and write the deepest possi
    '     vdecompression vstop vdepth to the output file
    '     knowing where the vdecompression zone starts is very important.  be
    ' '     that vdepth there is no possibility for bubble formation because th
    '     will be no supersaturation gradients.  vdeco vstops should never sta
    '     below the vdeco zone.  the deepest possible vstop vdeco vstop vdepth is
    '     defined as the next "standard" vstop vdepth above the point where th
    '     leading compartment enters the vdeco zone.  thus, the subroutine will
    '     base this calculation on step sizes larger than 10 fsw or 3 msw.
    '     deepest possible vstop vdepth is not used in the subroutine, per se, ra
    ' '     it is information to tell the diver where to start putting on the
    '     during ascent.  this should be prominently displayed by any vdeco p
    '=======================================================================
    Call calc_start_of_vdeco_zone(starting_vdepth, rate, vdepth_start_of_vdeco_zone)
    If (units_equal_fsw) Then
        If (vstep_size < 10#) Then
            rounding_operation1 = (vdepth_start_of_vdeco_zone / vstep_size) - 0.5
            deepest_possible_vstop_vdepth = CDbl(CInt(rounding_operation1)) * vstep_size
            Else
            rounding_operation1 = (vdepth_start_of_vdeco_zone / 10#) - 0.5
            deepest_possible_vstop_vdepth = CDbl(CInt(rounding_operation1)) * 10#
        End If
    End If
    If (units_equal_msw) Then
        If (vstep_size < 3#) Then
            rounding_operation1 = (vdepth_start_of_vdeco_zone / vstep_size) - 0.5
            deepest_possible_vstop_vdepth = CDbl(CInt(rounding_operation1)) * vstep_size
            Else
            rounding_operation1 = (vdepth_start_of_vdeco_zone / 3#) - 0.5
            deepest_possible_vstop_vdepth = CDbl(CInt(rounding_operation1)) * 3#
        End If
    End If
    't1print (vdepth_start_of_vdeco_zone)
    't1print (deepest_possible_vstop_vdepth)
    't1print (units_word1)
    't1print (units_word1)
    't1print (units_word2)
    't1print (units_word1)
    't1print (vbCrLf)
    '=======================================================================
    '     temporarily ascend profile to the start of the vdecompression zone,
    '     variables at this point, and initialize variables for critical vol
    '     the iterative process of the vpm critical volume valgorithm will op
    '     only in the vdecompression zone since it deals with excess gas volu
    '     released as a result of supersaturation gradients (not possible be
    '     vdecompression zone) .
    '=======================================================================
    Call gas_loadings_ascent_descent(starting_vdepth, vdepth_start_of_vdeco_zone, rate)
    run_vtime_start_of_vdeco_zone = run_vtime
    vdeco_phase_volume_vtime = 0#
    last_run_vtime = 0#
    schedule_converged = (False)
    For i = 1 To 16
        last_phase_volume_vtime(i) = 0#
        vhe_vpressure_start_of_vdeco_zone(i) = vhelium_vpressure(i)
        vn2_vpressure_start_of_vdeco_zone(i) = vnitrogen_vpressure(i)
        max_actual_gradient(i) = 0#
    Next i
    '=======================================================================
    '     start of critical volume loop
    '     this loop operates between lines 50 and 100.  if the critical volu
    '     valgorithm is toggled "off" in the subroutine settings, there will onl
    '     one pass through this loop.  otherwise, there will be two or more
    '     through this loop until the vdeco schedule is "converged" - that is
    '     comparison between the phase volume time of the present iteration
    '     last iteration is less than or equal to one vminute.  this implies
    '     the volume of released gas in the most recent iteration differs fr
    '     "critical" volume limit by an acceptably small amount.  the critic
    '     volume limit is set by the critical volume parameter lambda in the
    '     settings (default setting is 7500 fsw-min with adjustability range
    '     from 6500 to 8300 fsw-min according to bruce wienke) .
    '=======================================================================
L50:

    Do While (True)                     'loop will run continuous
        'there is an exit stateme
        '=======================================================================
        '     calculate initial ascent vceiling based on allowable supersaturatio
        '     gradients and set first vdeco vstop.  check to make sure that select
        '     size will not round up first vstop to a vdepth that is below the dec
        '=======================================================================
        Call calc_ascent_vceiling(ascent_vceiling_vdepth)
        If (ascent_vceiling_vdepth <= 0#) Then
            vdeco_vstop_vdepth = 0#
            Else
            rounding_operation2 = (ascent_vceiling_vdepth / vstep_size) + 0.5
            vdeco_vstop_vdepth = CDbl(CInt(rounding_operation2)) * vstep_size
        End If
        If (vdeco_vstop_vdepth > vdepth_start_of_vdeco_zone) Then
               no_deco_found = 3 ' MsgBox "root not in brackets"
               Exit Do
        End If
        '=======================================================================
        '     perform a separate "projected ascent" outside of the main subroutine
        '     sure that an increase in gas loadings during ascent to the first s
        '     not cause a violation of the vdeco vceiling.  if so, adjust the firs
        '     deeper based on step size until a safe ascent can be made.
        '     note: this situation is a possibility when ascending from extremel
        '     dives or due to an unusual gas vmix selection.
        '     check again to make sure that adjusted first vstop will not be belo
        '     vdeco zone.
        '=======================================================================
        Call projected_ascent(vdepth_start_of_vdeco_zone, rate, vdeco_vstop_vdepth, vstep_size)
        If (vdeco_vstop_vdepth > vdepth_start_of_vdeco_zone) Then
            no_deco_found = 3 ' MsgBox "root not in brackets"
        End If
        '=======================================================================
        '     handle the special case when no vdeco vstops are required - ascent c
        '     made directly to the surface
        '     write ascent data to output file and exit the critical volume loop
        '=======================================================================
        If (vdeco_vstop_vdepth = 0#) Then
            For i = 1 To 16
                vhelium_vpressure(i) = vhe_vpressure_start_of_ascent(i)
                vnitrogen_vpressure(i) = vn2_vpressure_start_of_ascent(i)
            Next i
            run_vtime = run_vtime_start_of_ascent
            vsegment_vnumber = vsegment_vnumber_start_of_ascent
            starting_vdepth = vdepth_change(1)
            ending_vdepth = 0#
            Call gas_loadings_ascent_descent(starting_vdepth, ending_vdepth, rate)
           'printtogrid3
            t1print ("Decompression" + vbCr + "  #     Dur     RT      Mix     Depth   Rate    ")
            t1print8 CStr(vsegment_vnumber)
            t1print8 CStr(vsegment_vtime)
            t1print8 CStr(run_vtime)
            t1print8 CStr(vmix_vnumber - 1)
            t1print8 CStr(vdeco_vstop_vdepth)
            t1print8 CStr(rate)
            t1print (vbCrLf)
         '   If systemversion = "Lite" Then
               printtogrid4lite
         '   Else
               printtogrid4
         '   End If
               
            Exit Do '           exit                       !exit the critical volume loop at
        End If
        '=======================================================================
        '     assign variables for ascent from start of vdeco zone to first vstop.
        '     first vstop vdepth for later use when computing the final ascent pro
        '=======================================================================
        starting_vdepth = vdepth_start_of_vdeco_zone
        first_vstop_vdepth = vdeco_vstop_vdepth
        '=======================================================================
        '     vdeco vstop loop block within critical volume loop
        '     this loop computes a vdecompression schedule to the surface during
        '     iteration of the critical volume loop.  no output is written from
        '     loop, rather it computes a schedule from which the in-water portio
        '     total phase volume time (vdeco_phase_volume_vtime) can be extracted.
        '     the gas loadings computed at the end of this loop are used the sub
        '     which computes the out-of-water portion of the total phase volume
        '     (surface_phase_volume_vtime) for that schedule.
        '
        '     note that exit is made from the loop after last ascent is made to
        '     vstop vdepth that is less than or equal to zero.  a final vdeco vstop
        '     than zero can happen when the user makes an odd step size change d
        '     ascent - such as specifying a 5 msw step size change at the 3 msw
        '=======================================================================
        Do While (True)                          'loop will run continuous
            'there is an exit stateme
            Call gas_loadings_ascent_descent(starting_vdepth, vdeco_vstop_vdepth, rate)
            If vdeco_vstop_vdepth <= 0# Then Exit Do
            If (vnumber_of_changes > 1) Then
                For i = 2 To vnumber_of_changes
                    If (vdepth_change(i) >= vdeco_vstop_vdepth) Then
                        vmix_vnumber = vmix_change(i)
                        rate = rate_change(i)
                        vstep_size = vstep_size_change(i)
                        SetPoint = vsetPoint_Change(i)
                        setUpSetpoint (units_factor)
                    End If
                Next i
            End If
            If laststop_index = 2 Then
              If ((vstep_size_change(1) * 2) >= vdeco_vstop_vdepth) Then
                        'vmix_vnumber = vmix_change(i)
                        'rate = rate_change(i)
                        vstep_size = vstep_size_change(1) / 2
                        'SetPoint = vsetPoint_Change(i)
                        'setUpSetpoint (units_factor)
'                        vstep_size = vstep_size_change(1) / 2
              End If
              If (((vstep_size_change(1) * 2) - 1) >= vdeco_vstop_vdepth) Then
                        'vmix_vnumber = vmix_change(i)
                        'rate = rate_change(i)
                        vstep_size = vstep_size_change(1) * 1.5
                        'SetPoint = vsetPoint_Change(i)
                        'setUpSetpoint (units_factor)
              End If
            End If
            If laststop_index = 3 Then
              If ((vstep_size_change(1) * 2) >= vdeco_vstop_vdepth) Then
                        'vmix_vnumber = vmix_change(i)
                        'rate = rate_change(i)
                        vstep_size = vstep_size_change(1) * 2
                        'SetPoint = vsetPoint_Change(i)
                        'setUpSetpoint (units_factor)
'                        vstep_size = vstep_size_change(1) / 2
              End If
            End If
            Call boyles_law_compensation(first_vstop_vdepth, vdeco_vstop_vdepth, vstep_size)                                                   'su
            Call vdecompression_vstop(vdeco_vstop_vdepth, vstep_size)
            starting_vdepth = vdeco_vstop_vdepth
            next_vstop = vdeco_vstop_vdepth - vstep_size
            vdeco_vstop_vdepth = next_vstop
            last_run_vtime = run_vtime
L60:

        Loop
        '=======================================================================
        '     compute total phase volume time and make critical volume compariso
        '     the vdeco phase volume time is computed from the run time.  the sur
        '     phase volume time is computed in a subroutine based on the surfaci
        '     loadings from previous vdeco loop block.  next the total phase volu
        '     (in-water + surface) for each compartment is compared against the
        '     total phase volume time.  the schedule is converged when the diffe
        '     less than or equal to 1 vminute in any one of the 16 compartments.
        '
        '     note:  the "phase volume time" is somewhat of a mathematical conce
        '     it is the time divided out of a total integration of supersaturati
        '     gradient x time (in-water and surface) .  this integration is multi
        '     by the excess bubble vnumber to represent the amount of free-gas re
        '     as a result of allowing a certain vnumber of excess bubbles to form
        '=======================================================================
        vdeco_phase_volume_vtime = run_vtime - run_vtime_start_of_vdeco_zone
        Call calc_surface_phase_volume_vtime                            'su
        For i = 1 To 16
            phase_volume_vtime(i) = vdeco_phase_volume_vtime + surface_phase_volume_vtime(i)
            critical_volume_comparison = Abs(phase_volume_vtime(i) - last_phase_volume_vtime(i))
            If (critical_volume_comparison <= 1#) Then
                schedule_converged = (True)
            End If
        Next i
        '=======================================================================
        '     critical volume decision tree between lines 70 and 99
        '     there are two options here.  if the critical volume agorithm setti
        '     "on" and the schedule is converged, or the critical volume algorit
        '     setting was "off" in the first place, the subroutine will re-assign v
        '     to their values at the start of ascent (end of bottom time) and pr
        '     a complete vdecompression schedule once again using all the same as
        '     parameters and first vstop vdepth.  this vdecompression schedule will
        '     the last iteration of the critical volume loop and the subroutine wil
        '     the final vdeco schedule to the output file.
        '
        '     note: if the critical volume agorithm setting was "off", the final
        '     schedule will be based on "initial allowable supersaturation gradi
        '     if it was "on", the final schedule will be based on "adjusted allo
        '     supersaturation gradients" (gradients that are "relaxed" as a resu
        '     the critical volume valgorithm) .
        '
        '     if the critical volume agorithm setting is "on" and the schedule i
        '     converged, the subroutine will re-assign variables to their values at
        '     start of the vdeco zone and process another trial vdecompression sch
        '=======================================================================
L70:

        If ((schedule_converged) Or (critical_volume_valgorithm_off)) Then
            For i = 1 To 16
                vhelium_vpressure(i) = vhe_vpressure_start_of_ascent(i)
                vnitrogen_vpressure(i) = vn2_vpressure_start_of_ascent(i)
            Next i
            run_vtime = run_vtime_start_of_ascent
            vsegment_vnumber = vsegment_vnumber_start_of_ascent
            starting_vdepth = vdepth_change(1)
            vmix_vnumber = vmix_change(1)
            rate = rate_change(1)
            vstep_size = vstep_size_change(1)
            vdeco_vstop_vdepth = first_vstop_vdepth
            SetPoint = vsetPoint_Change(1)
            setUpSetpoint (units_factor)
            last_run_vtime = 0#
            '=======================================================================
            '     vdeco vstop loop block for final vdecompression schedule
            '=======================================================================
            t1print ("Decompression" + vbCrLf + "  " + "   #     Dur      RT     Mix   Depth    Rate   Stime" + vbCrLf)
           'printtogrid3
            Do While (True)                      'loop will run continuous
                'there is an exit stateme
                Call gas_loadings_ascent_descent(starting_vdepth, vdeco_vstop_vdepth, rate)
                '=======================================================================
                '     during final vdecompression schedule process, compute maximum actua
                '     supersaturation gradient resulting in each compartment
                '     if there is a repetitive dive, this will be used later in the vpm
                '     repetitive valgorithm to adjust the values for critical radii.
                '=======================================================================
                Call calc_max_actual_gradient(vdeco_vstop_vdepth)         'su
                t1print8 CStr(vsegment_vnumber)
                t1print8dbl (CDbl(CInt(vsegment_vtime * 10# + 0.999) / 10))
                t1print8dbl (CDbl(CLng(run_vtime * 10# + 0.999) / 10))
                t1print8 CStr(vmix_vnumber - 1)
                t1print8dbl (CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.999) / 10))
                t1print8 CStr(rate)
                t1print (vbCrLf)
             '   If systemversion = "Lite" Then
                   printtogrid5lite
             '   Else
                   printtogrid5
             '   End If
                If vdeco_vstop_vdepth <= 0# Then Exit Do   ' .le. 0.0) exit                !exit a
                If (vnumber_of_changes > 1) Then
                    vdepth_change_new = 9999
                    For i = 2 To vnumber_of_changes
                        If (vdepth_change(i) >= vdeco_vstop_vdepth) And vdepth_change(i) < vdepth_change_new Then
                            vmix_vnumber = vmix_change(i)
                            rate = rate_change(i)
                            vstep_size = vstep_size_change(i)
                            vdepth_change_new = vdepth_change(i)
                            SetPoint = vsetPoint_Change(i)
                            setUpSetpoint (units_factor)
                        End If
                    Next i
                End If
                If laststop_index = 2 Then
                  If ((vstep_size_change(1) * 2) >= vdeco_vstop_vdepth) Then
                        'vmix_vnumber = vmix_change(i)
                        'rate = rate_change(i)
                        vstep_size = vstep_size_change(1) / 2
                        'SetPoint = vsetPoint_Change(i)
                        'setUpSetpoint (units_factor)
'                        vstep_size = vstep_size_change(1) / 2
                  End If
                  If (((vstep_size_change(1) * 2) - 1) >= vdeco_vstop_vdepth) Then
                        'vmix_vnumber = vmix_change(i)
                        'rate = rate_change(i)
                        vstep_size = vstep_size_change(1) * 1.5
                        'SetPoint = vsetPoint_Change(i)
                        'setUpSetpoint (units_factor)
                  End If
                End If
                If laststop_index = 3 Then
                  If ((vstep_size_change(1) * 2) >= vdeco_vstop_vdepth) Then
                        'vmix_vnumber = vmix_change(i)
                        'rate = rate_change(i)
                        vstep_size = vstep_size_change(1) * 2
                        'SetPoint = vsetPoint_Change(i)
                        'setUpSetpoint (units_factor)
'                        vstep_size = vstep_size_change(1) / 2
                  End If
                End If
                Call boyles_law_compensation(first_vstop_vdepth, vdeco_vstop_vdepth, vstep_size)                                                   'su
                Call vdecompression_vstop(vdeco_vstop_vdepth, vstep_size)    'su
                '=======================================================================
                '     this next bit justs rounds up the vstop time at the first vstop to b
                '     whole increments of the minimum vstop time (to make for a nice vdeco
                '=======================================================================
                If (last_run_vtime = 0#) Then
                    vstop_vtime = CDbl(CInt((vsegment_vtime / minimum_vdeco_vstop_vtime) + 0.5)) * minimum_vdeco_vstop_vtime
                    Else
                    vstop_vtime = run_vtime - last_run_vtime
                End If
                '=======================================================================
                '     during final vdecompression schedule, if minimum vstop time paramete
                '     whole vnumber (i.e. 1 vminute) then write vdeco schedule using intege
                '     vnumbers (looks nicer) .  otherwise, use decimal vnumbers.
                '     note: per the request of a noted exploration diver(!) , subroutine now
                '     a minimum vstop time of less than one vminute so that total ascent t
                '     be minimized on very long dives.  in fact, with step size set at 1
                '     0.2 msw and minimum vstop time set at 0.1 vminute (6 seconds) , a nea
                '     continuous vdecompression schedule can be computed.
                '=======================================================================
                If (CDbl(CInt(minimum_vdeco_vstop_vtime)) = minimum_vdeco_vstop_vtime) Then
                    t1print CStr(vsegment_vnumber)
                    t1print CStr(vsegment_vtime)
                    t1print CStr(Left(CStr(run_vtime), 6))
                    t1print CStr(vmix_vnumber - 1)
                    t1print CStr(CInt(vdeco_vstop_vdepth))
                    t1print CStr(CInt(vstop_vtime))
                    t1print CStr(CInt(run_vtime))
                    t1print (vbCrLf)
             '       If systemversion = "Lite" Then
                       printtogrid5lite
             '       Else
                       printtogrid5
             '       End If
                Else
                    t1print8 CStr(vsegment_vnumber)
                    t1print8dbl (CDbl(CInt(vsegment_vtime * 10# + 0.999) / 10))
                    t1print8dbl (CDbl(CLng(run_vtime * 10# + 0.999) / 10))
                    t1print8 CStr(vmix_vnumber - 1)
                    t1print8dbl (CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.999) / 10))
                    t1print8 CStr(rate)
                    t1print8dbl (CDbl(CInt(vstop_vtime * 10#) / 10))
                    t1print (vbCrLf)
                    ratelasttemp = rate
                    rate = 0#
              '      If systemversion = "Lite" Then
                       printtogrid5lite
              '      Else
                       printtogrid5
              '      End If
                    
                    rate = ratelasttemp '-10#
                End If
                starting_vdepth = vdeco_vstop_vdepth
                next_vstop = vdeco_vstop_vdepth - vstep_size
                vdeco_vstop_vdepth = next_vstop
                last_run_vtime = run_vtime
L80:

            Loop
            'for final vdeco sche
            Exit Do '          exit                            !exit critical volume loop at
            'final vdeco schedule
        Else
            '=======================================================================
            '     if schedule not converged, compute relaxed allowable supersaturati
            '     gradients with vpm critical volume valgorithm and process another
            '     iteration of the critical volume loop
            '=======================================================================
            Call critical_volume(vdeco_phase_volume_vtime)               'su
            vdeco_phase_volume_vtime = 0#
            run_vtime = run_vtime_start_of_vdeco_zone
            starting_vdepth = vdepth_start_of_vdeco_zone
            vmix_vnumber = vmix_change(1)
            rate = rate_change(1)
            vstep_size = vstep_size_change(1)
            SetPoint = vsetPoint_Change(1)
            setUpSetpoint (units_factor)
            For i = 1 To 16
                last_phase_volume_vtime(i) = phase_volume_vtime(i)
                vhelium_vpressure(i) = vhe_vpressure_start_of_vdeco_zone(i)
                vnitrogen_vpressure(i) = vn2_vpressure_start_of_vdeco_zone(i)
            Next i
            '          cycle                         !return to start of critical vo
            '(line 50) to process another it
L99:

        End If                               'end of critical volume decis
    Loop             'end of critical vol
L100:

    '    continue                                      'end of critical vol
        '=======================================================================
        '     processing of dive complete.  read input file to determine if ther
        '     repetitive dive.  if none, then exit repetitive loop.
        '=======================================================================
        'repetitive_dive_flag = 0
'        If (repetitive_dive_flag = 0) Then
'            GoTo L330 'Exit Do '          exit                                        !exit repetitive
'            'at line 330
'            '=======================================================================
'            '     if there is a repetitive dive, compute gas loadings (off-gassing)
'            '     surface interval time.  adjust critical radii using vpm repetitive
'            '     valgorithm.  re-initialize selected variables and return to start o
'            '     repetitive loop at line 30.
'            '=======================================================================
'        ElseIf (repetitive_dive_flag = 1) Then
'            surface_interval_vtime = 60
            Call gas_loadings_surface_interCint(surface_interval_vtime)  'su
            Call vpm_repetitive_valgorithm(surface_interval_vtime)       'su
            For i = 1 To 16
                max_crushing_vpressure_he(i) = 0#
                max_crushing_vpressure_n2(i) = 0#
                max_actual_gradient(i) = 0#
            Next i
            run_vtime = 0#
            vsegment_vnumber = 0
            vmix_vnumber = 1 'air
            Is_CCR = False
            dum = ppo2exposuretime(0#, surface_interval_vtime)   ' cns_current = cns_current - (barometric_vpressure * surface_interval_vtime)
            'If cns_current < 0# Then cns_current = 0#
            'otu_current = otu_current - (barometric_vpressure * surface_interval_vtime)
            'If otu_current < 0# Then otu_current = 0#
            GoTo L330 'exit do
            '          cycle      !return to start of repetitive loop to process ano
            '=======================================================================
            '     write error message and terminate subroutine if there is an error in
            '     input file for the repetitive dive flag
            '=======================================================================
'        Else
'            no_deco_found=3 ' MsgBox "root not in brackets"
'        End If
Loop  '        continue                                           'end of repetit
L330:

        '=======================================================================
        '     final writes to output and close subroutine files
        '=======================================================================
        'close (unit = 7, status = "keep")
        'close (unit = 8, status = "keep")
        'close (unit = 10, status = "keep")
        '=======================================================================
        ' '     format statements - subroutine input/output
        '=======================================================================
        ' 800   format ('0units = feet of seawater (fsw) ')
        ' 801   format ('0units = meters of seawater (msw) ')
        ' 802   format ('0valtitude = ',1x,f7.1,4x,'barometric vpressure = ',       f6.3)
        ' 805   format (a70)
        ' 811   format (26x,'vdecompression calculation subroutine')
        ' 812   format (24x,'developed in fortran by erik c. baker')
        ' 814   format ('subroutine run:',4x,i2.2,'-',i2.2,'-',i4,1x,'at',1x,i2.2,           ':',i2.2,1x,a1,'m',23x,'model: vpm-b')
        ' 815   format ('description:',4x,a70)
        ' 813   format (' ')
        ' 820   format ('gasvmix summary:',24x,'fo2',4x,'fhe',4x,'fn2')
        ' 821   format (26x,'gasvmix #',i2,2x,f5.3,2x,f5.3,2x,f5.3)
        ' 830   format (36x,'dive profile')
        ' 831   format ('seg-',2x,'segm.',2x,'run',3x,'|',1x,'gasvmix',1x,'|',1x,         'ascent',4x,'from',5x,'to',6x,'rate',4x,'|',1x,'constant')
        ' 832   format ('ment',2x,'time',3x,'time',2x,'|',2x,'used',2x,'|',3x,            'or',5x,'vdepth',3x,'vdepth',4x,'+dn/-up',2x,'|',2x,'vdepth')
        ' 833   format (2x,'#',3x,'(min) ',2x,'(min) ',1x,'|',4x,'#',3x,'|',1x,             'descent',2x,'(',a4,') ',2x,'(',a4,') ',2x,'(',a7,') ',1x,           '|',2x,'(',a4,') ')
        ' 834   format ('-----',1x,'-----',2x,'-----',1x,'|',1x,'------',1x,'|',          1x,'-------',2x,'------',2x,'------',2x,'---------',1x,           '|',1x,'--------')
        ' 840   format (i3,3x,f5.1,1x,f6.1,1x,'|',3x,i2,3x,'|',1x,a7,f7.0,                    1x,f7.0,3x,f7.1,3x,'|')
        ' 845   format (i3,3x,f5.1,1x,f6.1,1x,'|',3x,i2,3x,'|',36x,'|',f7.0)
        ' 850   format (31x,'vdecompression profile')
        ' 851   format ('seg-',2x,'segm.',2x,'run',3x,'|',1x,'gasvmix',1x,'|',1x,          'ascent',3x,'ascent',3x,'col',3x,'|',2x,'vdeco',3x,'vstop',         3x,'run')
        ' 852   format ('ment',2x,'time',3x,'time',2x,'|',2x,'used',2x,'|',3x,            'to',6x,'rate',4x,'not',3x,'|',2x,'vstop',3x,'time',3x,             'time')
        ' 853   format (2x,'#',3x,'(min) ',2x,'(min) ',1x,'|',4x,'#',3x,'|',1x,             '(',a4,') ',1x,'(',a7,') ',2x,'used',2x,'|',1x,'(',a4,') ',          2x,'(min) ',2x,'(min) ')
        ' 854   format ('-----',1x,'-----',2x,'-----',1x,'|',1x,'------',1x,'|',          1x,'------',1x,'---------',1x,'------',1x,'|',1x,                 '------',2x,'-----',2x,'-----')
        ' 857   format (10x,'leading compartment enters the vdecompression zone',          1x,'at',f7.1,1x,a4)
        ' 858   format (17x,'deepest possible vdecompression vstop is',f7.1,1x,a4)
        ' 860   format (i3,3x,f5.1,1x,f6.1,1x,'|',3x,i2,3x,'|',2x,f4.0,3x,f6.1,           10x,'|')
        ' 862   format (i3,3x,f5.1,1x,f6.1,1x,'|',3x,i2,3x,'|',25x,'|',2x,i4,3x,          i4,2x,i5)
        ' 863   format (i3,3x,f5.1,1x,f6.1,1x,'|',3x,i2,3x,'|',25x,'|',2x,f5.0,1x,        f6.1,1x,f7.1)
        ' 871   format (' subroutine calculations complete')
        ' 872   format ('0output data is located in the file vpmvdeco.out')
        ' 880   format (' ')
        ' 890   format ('repetitive dive:')
        '=======================================================================
        ' '     format statements - error messages
        '=======================================================================
        ' 900   format (' ')
        ' 901   format ('0error! units must be fsw or msw')
        ' 902   format ('0error! valtitude dive valgorithm must be on or off')
        ' 903   format ('0error! radius must be between 0.2 and 1.35 microns')
        ' 904   format ('0error! critical volume valgorithm must be on or off')
        ' 905   format ('0error! step size is too large to vdecompress')
        ' 906   format ('0error in input file (gasvmix data) ')
        ' 907   format ('0error in input file (profile code) ')
        ' 908   format ('0error in input file (repetitive dive code) ')
        '=======================================================================
        '     end of main subroutine
        '=======================================================================
        'Close #1:
End Sub


'
'=======================================================================
'     function subsubroutine for gas loading calculations - ascent and desc
'=======================================================================
Public Function schreiner_equation(initial_inspired_gas_vpressure As Double, rate_change_insp_gas_vpressure As Double, interval_vtime As Double, gas_vtime_constant As Double, initial_gas_vpressure As Double) As Double
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim initial_inspired_gas_vpressure as double
' sub parameter : do not dim ! Dim rate_change_insp_gas_vpressure as double
' sub parameter : do not dim ! Dim interval_vtime as double
' sub parameter : do not dim ! Dim  gas_vtime_constant as double
' sub parameter : do not dim ! Dim initial_gas_vpressure as double
'Dim schreiner_equation As Double
'=======================================================================
'     note: the schreiner equation is applied when calculating the uptak
'     elimination of compartment gases during linear ascents or descents
'     constant rate.  for ascents, a negative vnumber for rate must be us
'=======================================================================
schreiner_equation = initial_inspired_gas_vpressure + rate_change_insp_gas_vpressure * (interval_vtime - 1# / gas_vtime_constant) - (initial_inspired_gas_vpressure - initial_gas_vpressure - rate_change_insp_gas_vpressure / gas_vtime_constant) * Exp(-gas_vtime_constant * interval_vtime)
Exit Function
Exit Function
'=======================================================================
'     function subsubroutine for gas loading calculations - constant vdepth
'=======================================================================
End Function
Function haldane_equation(initial_gas_vpressure As Double, inspired_gas_vpressure As Double, gas_vtime_constant As Double, interval_vtime As Double) As Double
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim initial_gas_vpressure as double
' sub parameter : do not dim ! Dim  inspired_gas_vpressure as double
' sub parameter : do not dim ! Dim gas_vtime_constant as double
' sub parameter : do not dim ! Dim  interval_vtime as double
'Dim haldane_equation As Double
'=======================================================================
'     note: the haldane equation is applied when calculating the uptake
'     elimination of compartment gases during intervals at constant dept
'     outside ambient vpressure does not change) .
'=======================================================================
'haldane_equation = initial_gas_vpressure + (inspired_gas_vpressure - initial_gas_vpressure) * (1# - Exp(-gas_vtime_constant * interval_vtime))
'haldane_equation = (inspired_gas_vpressure - initial_gas_vpressure)
'f = (1# - Exp(-gas_vtime_constant * interval_vtime))
'f = f * haldane_equation
'f = initial_gas_vpressure + f
'f = (1# - Exp(-gas_vtime_constant * interval_vtime))
'haldane_equation = initial_gas_vpressure + (inspired_gas_vpressure - initial_gas_vpressure) * (1# - Exp(-gas_vtime_constant * interval_vtime))
haldane_equation = initial_gas_vpressure + (inspired_gas_vpressure - initial_gas_vpressure) * (1# - Exp(-gas_vtime_constant * interval_vtime))
Exit Function
Exit Function
'=======================================================================
'     subroutine gas_loadings_ascent_descent
'     purpose: this subsubroutine applies the schreiner equation to update
'     gas loadings (partial vpressures of vhelium and vnitrogen) in the hal
'     compartments due to a linear ascent or descent vsegment at a consta
'=======================================================================
End Function
Sub gas_loadings_ascent_descent(starting_vdepth As Double, ending_vdepth As Double, rate As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim starting_vdepth as double
' sub parameter : do not dim ! Dim  ending_vdepth as double
' sub parameter : do not dim ! Dim  rate as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim last_vsegment_vnumber As Integer
Dim last_run_vtime As Double
Dim vhelium_rate As Double
Dim vnitrogen_rate As Double
'Dim starting_ambient_vpressure As Double
'Dim schreiner_equation As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim water_vapor_vpressure As Double
''common /block_8/ water_vapor_vpressure
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim vsegment_vnumber                                         'bo as integer
'Dim run_vtime As Double
'Dim vsegment_vtime                                     'an as double
''common /block_2/ run_vtime, vsegment_vnumber, vsegment_vtime
'Dim ending_ambient_vpressure As Double
''common /block_4/ ending_ambient_vpressure
'Dim vmix_vnumber As Integer
''common /block_9/ vmix_vnumber
'Dim barometric_vpressure As Double
''common /block_18/ barometric_vpressure
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim vhelium_vtime_constant(16) As Double
''common /block_1a/ vhelium_vtime_constant
'Dim vnitrogen_vtime_constant(16) As Double
''common /block_1b/ vnitrogen_vtime_constant
'Dim vhelium_vpressure(16)  As Double
'Dim vnitrogen_vpressure(16)                 'bo as double
''common /block_3/ vhelium_vpressure, vnitrogen_vpressure            'an
'Dim vfraction_vhelium(20)  As Double
'Dim vfraction_vnitrogen(20) As Double
''common /block_5/ vfraction_vhelium, vfraction_vnitrogen
'Dim initial_vhelium_vpressure(16)  As Double
'Dim initial_vnitrogen_vpressure(16) As Double
''common /block_23/ initial_vhelium_vpressure,                                          initial_vnitrogen_vpressure
'=======================================================================
'     calculations
'=======================================================================
vsegment_vtime = (ending_vdepth - starting_vdepth) / rate
last_run_vtime = run_vtime
run_vtime = last_run_vtime + vsegment_vtime
last_vsegment_vnumber = vsegment_vnumber
vsegment_vnumber = last_vsegment_vnumber + 1
ending_ambient_vpressure = ending_vdepth + barometric_vpressure
starting_ambient_vpressure = starting_vdepth + barometric_vpressure
initial_inspired_vhe_vpressure = (starting_ambient_vpressure - water_vapor_vpressure) * vfraction_vhelium(vmix_vnumber)
initial_inspired_vn2_vpressure = (starting_ambient_vpressure - water_vapor_vpressure) * vfraction_vnitrogen(vmix_vnumber)
vhelium_rate = rate * vfraction_vhelium(vmix_vnumber)
vnitrogen_rate = rate * vfraction_vnitrogen(vmix_vnumber)

'a = Set_Inspired_Inert(vmix_vnumber, starting_ambient_vpressure, water_vapor_vpressure, vfraction_vhelium(vmix_vnumber), vfraction_vnitrogen(vmix_vnumber), initial_inspired_vhe_vpressure, initial_inspired_vn2_vpressure, vhelium_rate, vnitrogen_rate)
Set_Inspired_Inert_Starting

For i = 1 To 16
    initial_vhelium_vpressure(i) = vhelium_vpressure(i)
    initial_vnitrogen_vpressure(i) = vnitrogen_vpressure(i)
    vhelium_vpressure(i) = schreiner_equation(initial_inspired_vhe_vpressure, vhelium_rate, vsegment_vtime, vhelium_vtime_constant(i), initial_vhelium_vpressure(i))
    vnitrogen_vpressure(i) = schreiner_equation(initial_inspired_vn2_vpressure, vnitrogen_rate, vsegment_vtime, vnitrogen_vtime_constant(i), initial_vnitrogen_vpressure(i))
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
'=======================================================================
'     subroutine calc_crushing_vpressure
'     purpose: compute the effective "crushing vpressure" in each compart
'     a result of descent vsegment(s) .  the crushing vpressure is the grad
'     (difference in vpressure) between the outside ambient vpressure and
'     gas tension inside a vpm nucleus (bubble seed) .  this gradient act
'     reduce (shrink) the radius smaller than its initial value at the s
'     this phenomenon has important ramifications because the smaller th
'     of a vpm nucleus, the greater the allowable supersaturation gradie
'     ascent.  gas loading (uptake) during descent, especially in the fa
'     compartments, will reduce the magnitude of the crushing vpressure.
'     crushing vpressure is not cumulative over a multi-level descent.  i
'     be the maximum value obtained in any one discrete vsegment of the o
'     descent.  thus, the subroutine must compute and store the maximum cru
'     vpressure for each compartment that was obtained across all vsegment
'     the descent profile.
'
'     the calculation of crushing vpressure will be different depending o
'     whether or not the gradient is in the vpm permeable range (gas can
'     across skin of vpm nucleus) or the vpm impermeable range (molecule
'     skin of nucleus are squeezed together so tight that gas can no lon
'     diffuse in or out of nucleus; the gas becomes trapped and further
'     the crushing vpressure) .  the solution for crushing vpressure in the
'     permeable range is a simple linear equation.  in the vpm impermeab
'     range, a cubic equation must be solved using a numerical method.
'
'     separate crushing vpressures are tracked for vhelium and vnitrogen be
'     they can have different critical radii.  the crushing vpressures wi
'     the same for vhelium and vnitrogen in the permeable range of the mod
'     they will start to diverge in the impermeable range.  this is due
'     the differences between starting radius, radius at the onset of
'     impermeability, and radial compression in the impermeable range.
'=======================================================================
End Sub
Sub calc_crushing_vpressure(starting_vdepth As Double, ending_vdepth As Double, rate As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim starting_vdepth as double
' sub parameter : do not dim ! Dim  ending_vdepth as double
' sub parameter : do not dim ! Dim  rate as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim starting_ambient_vpressure As Double
Dim ending_ambient_vpressure As Double
Dim starting_gas_tension As Double
Dim ending_gas_tension As Double
Dim crushing_vpressure_he As Double
Dim crushing_vpressure_n2 As Double
Dim gradient_onset_of_imperm As Double
Dim gradient_onset_of_imperm_pa As Double
Dim ending_ambient_vpressure_pa As Double
Dim amb_press_onset_of_imperm_pa As Double
Dim gas_tension_onset_of_imperm_pa As Double
Dim crushing_vpressure_pascals_he As Double
Dim crushing_vpressure_pascals_n2 As Double
Dim starting_gradient As Double
Dim ending_gradient As Double
Dim a_he As Double
Dim b_he As Double
Dim c_he As Double
Dim ending_radius_he As Double
Dim high_bound_he As Double
Dim low_bound_he As Double
Dim a_n2 As Double
Dim b_n2 As Double
Dim c_n2 As Double
Dim ending_radius_n2 As Double
Dim high_bound_n2 As Double
Dim low_bound_n2 As Double
Dim radius_onset_of_imperm_he As Double
Dim radius_onset_of_imperm_n2 As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim gradient_onset_of_imperm_atm As Double
''common /block_14/ gradient_onset_of_imperm_atm
'Dim constant_vpressure_other_gases As Double
''common /block_17/ constant_vpressure_other_gases
'Dim surface_tension_gamma As Double
'Dim skin_compression_gammac As Double
''common /block_19/ surface_tension_gamma, skin_compression_gammac
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim units_factor As Double
''common /block_16/ units_factor
'Dim barometric_vpressure As Double
''common /block_18/ barometric_vpressure
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim vhelium_vpressure(16)  As Double
'Dim vnitrogen_vpressure(16) As Double
''common /block_3/ vhelium_vpressure, vnitrogen_vpressure
'Dim adjusted_critical_radius_he(16) As Double
'Dim adjusted_critical_radius_n2(16) As Double
''common /block_7/ adjusted_critical_radius_he,                                adjusted_critical_radius_n2
'Dim max_crushing_vpressure_he(16)  As Double
'Dim max_crushing_vpressure_n2(16) As Double
''common /block_10/ max_crushing_vpressure_he,                                         max_crushing_vpressure_n2
'Dim amb_vpressure_onset_of_imperm(16) As Double
'Dim gas_tension_onset_of_imperm(16) As Double
''common /block_13/ amb_vpressure_onset_of_imperm,                               gas_tension_onset_of_imperm
'Dim initial_vhelium_vpressure(16)  As Double
'Dim initial_vnitrogen_vpressure(16) As Double
''common /block_23/ initial_vhelium_vpressure,                                          initial_vnitrogen_vpressure
'=======================================================================
'     calculations
'     first, convert the gradient for onset of impermeability from units
'     atmospheres to diving vpressure units (either fsw or msw) and to pa
'     (si units) .  the reason that the gradient for onset of impermeabil
'     given in the subroutine settings in units of atmospheres is because t
'     how it was reported in the original research papers by yount and
'     colleauges.
'=======================================================================
gradient_onset_of_imperm = gradient_onset_of_imperm_atm * units_factor                                                      'divi
gradient_onset_of_imperm_pa = gradient_onset_of_imperm_atm * 101325#
'=======================================================================
'     assign values of starting and ending ambient vpressures for descent
'=======================================================================
starting_ambient_vpressure = starting_vdepth + barometric_vpressure
ending_ambient_vpressure = ending_vdepth + barometric_vpressure
'=======================================================================
'     main loop with nested decision tree
'     for each compartment, the subroutine computes the starting and ending
'     gas tensions and gradients.  the vpm is different than some dissol
'     valgorithms, buhlmann for example, in that it considers the pressur
'     voxygen, carbon dioxide, and water vapor in each compartment in add
'     the inert gases vhelium and vnitrogen.  these "other gases" are incl
'     the calculation of gas tensions and gradients.
'=======================================================================
For i = 1 To 16
    starting_gas_tension = initial_vhelium_vpressure(i) + initial_vnitrogen_vpressure(i) + constant_vpressure_other_gases
    starting_gradient = starting_ambient_vpressure - starting_gas_tension
    ending_gas_tension = vhelium_vpressure(i) + vnitrogen_vpressure(i) + constant_vpressure_other_gases
    ending_gradient = ending_ambient_vpressure - ending_gas_tension
    '=======================================================================
    '     compute radius at onset of impermeability for vhelium and vnitrogen
    '     critical radii
    '=======================================================================
    radius_onset_of_imperm_he = 1# / (gradient_onset_of_imperm_pa / (2# * (skin_compression_gammac - surface_tension_gamma)) + 1# / adjusted_critical_radius_he(i))
    radius_onset_of_imperm_n2 = 1# / (gradient_onset_of_imperm_pa / (2# * (skin_compression_gammac - surface_tension_gamma)) + 1# / adjusted_critical_radius_n2(i))
    '=======================================================================
    '     first branch of decision tree - permeable range
    '     crushing vpressures will be the same for vhelium and vnitrogen
    '=======================================================================
    If (ending_gradient <= gradient_onset_of_imperm) Then
        crushing_vpressure_he = ending_ambient_vpressure - ending_gas_tension
        crushing_vpressure_n2 = ending_ambient_vpressure - ending_gas_tension
    End If
    '=======================================================================
    '     second branch of decision tree - impermeable range
    '     both the ambient vpressure and the gas tension at the onset of
    '     impermeability must be computed in order to properly solve for the
    '     radius and resultant crushing vpressure.  the first decision block
    '     addresses the special case when the starting gradient just happens
    '     equal to the gradient for onset of impermeability (not very likely
    '=======================================================================
    If (ending_gradient > gradient_onset_of_imperm) Then
        If (starting_gradient = gradient_onset_of_imperm) Then
            amb_vpressure_onset_of_imperm(i) = starting_ambient_vpressure
            gas_tension_onset_of_imperm(i) = starting_gas_tension
        End If
        '=======================================================================
        '     in most cases, a subroutine will be called to find these values us
        '     numerical method.
        '=======================================================================
        If (starting_gradient < gradient_onset_of_imperm) Then
            Call onset_of_impermeability(starting_ambient_vpressure, ending_ambient_vpressure, rate, i)
        End If
        '=======================================================================
        '     next, using the values for ambient vpressure and gas tension at the
        '     of impermeability, the equations are set up to process the calcula
        '     through the radius root finder subroutine.  this subsubroutine will f
        '     root (solution) to the cubic equation using a numerical method.  i
        '     to do this efficiently, the equations are placed in the form
        '     ar^3 - br^2 - c = 0, where r is the ending radius after impermeabl
        '     compression.  the coefficients a, b, and c for vhelium and vnitrogen
        '     computed and passed to the subroutine as arguments.  the high and
        '     bounds to be used by the numerical method of the subroutine are al
        '     computed (see separate page posted on vdeco list ftp site entitled
        '     "vpm: solving for radius in the impermeable regime") .  the subprog
        '     will return the value of the ending radius and then the crushing
        '     vpressures for vhelium and vnitrogen can be calculated.
        '=======================================================================
        ending_ambient_vpressure_pa = (ending_ambient_vpressure / units_factor) * 101325#
        amb_press_onset_of_imperm_pa = (amb_vpressure_onset_of_imperm(i) / units_factor) * 101325#
        gas_tension_onset_of_imperm_pa = (gas_tension_onset_of_imperm(i) / units_factor) * 101325#
        b_he = 2# * (skin_compression_gammac - surface_tension_gamma)
        a_he = ending_ambient_vpressure_pa - amb_press_onset_of_imperm_pa + gas_tension_onset_of_imperm_pa + (2# * (skin_compression_gammac - surface_tension_gamma)) / radius_onset_of_imperm_he
        c_he = gas_tension_onset_of_imperm_pa * radius_onset_of_imperm_he ^ 3
        high_bound_he = radius_onset_of_imperm_he
        low_bound_he = b_he / a_he
        Call radius_root_finder(a_he, b_he, c_he, low_bound_he, high_bound_he, ending_radius_he)
        crushing_vpressure_pascals_he = gradient_onset_of_imperm_pa + ending_ambient_vpressure_pa - amb_press_onset_of_imperm_pa + gas_tension_onset_of_imperm_pa * (1# - radius_onset_of_imperm_he ^ 3 / ending_radius_he ^ 3)
        crushing_vpressure_he = (crushing_vpressure_pascals_he / 101325#) * units_factor
        b_n2 = 2# * (skin_compression_gammac - surface_tension_gamma)
        a_n2 = ending_ambient_vpressure_pa - amb_press_onset_of_imperm_pa + gas_tension_onset_of_imperm_pa + (2# * (skin_compression_gammac - surface_tension_gamma)) / radius_onset_of_imperm_n2
        c_n2 = gas_tension_onset_of_imperm_pa * radius_onset_of_imperm_n2 ^ 3
        high_bound_n2 = radius_onset_of_imperm_n2
        low_bound_n2 = b_n2 / a_n2
        Call radius_root_finder(a_n2, b_n2, c_n2, low_bound_n2, high_bound_n2, ending_radius_n2)
        crushing_vpressure_pascals_n2 = gradient_onset_of_imperm_pa + ending_ambient_vpressure_pa - amb_press_onset_of_imperm_pa + gas_tension_onset_of_imperm_pa * (1# - radius_onset_of_imperm_n2 ^ 3 / ending_radius_n2 ^ 3)
        crushing_vpressure_n2 = (crushing_vpressure_pascals_n2 / 101325#) * units_factor
    End If
    '=======================================================================
    '     update values of max crushing vpressure in global arrays
    '=======================================================================
    max_crushing_vpressure_he(i) = Max(max_crushing_vpressure_he(i), crushing_vpressure_he)
    max_crushing_vpressure_n2(i) = Max(max_crushing_vpressure_n2(i), crushing_vpressure_n2)
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
End Sub

'=======================================================================
'     subroutine onset_of_impermeability
'     purpose:  this subroutine uses the bisection method to find the am
'     vpressure and gas tension at the onset of impermeability for a give
'     compartment.  source:  "numerical recipes in fortran 77",
'     cambridge university press, 1992.
'=======================================================================
Sub onset_of_impermeability(starting_ambient_vpressure As Double, ending_ambient_vpressure As Double, rate As Double, i As Integer)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! dim i as integer                       'input - array subscript for com as integer
' sub parameter : do not dim ! Dim starting_ambient_vpressure as double
' sub parameter : do not dim ! Dim  ending_ambient_vpressure as double
' sub parameter : do not dim ! Dim  rate as double
'=======================================================================
'     local variables
'=======================================================================
Dim j As Integer                                                  'loop as integer
'Dim initial_inspired_vhe_vpressure As Double
'Dim initial_inspired_vn2_vpressure As Double
Dim xtime As Double
Dim vhelium_rate As Double
Dim vnitrogen_rate As Double
Dim low_bound As Double
Dim high_bound As Double
Dim high_bound_vhelium_vpressure As Double
Dim high_bound_vnitrogen_vpressure As Double
Dim mid_range_vhelium_vpressure As Double
Dim mid_range_vnitrogen_vpressure As Double
Dim last_diff_change As Double
Dim funcion_at_high_bound As Double
Dim funcion_at_low_bound As Double
Dim mid_range_vtime As Double
Dim funcion_at_mid_range As Double
Dim differential_change As Double
Dim mid_range_ambient_vpressure As Double
Dim gas_tension_at_mid_range As Double
Dim gradient_onset_of_imperm As Double
Dim starting_gas_tension As Double
Dim ending_gas_tension As Double
'Dim schreiner_equation                               'funcion sub as double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim water_vapor_vpressure As Double
'common /block_8/ water_vapor_vpressure
'Dim gradient_onset_of_imperm_atm As Double
'common /block_14/ gradient_onset_of_imperm_atm
'Dim constant_vpressure_other_gases As Double
'common /block_17/ constant_vpressure_other_gases
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim vmix_vnumber As Integer
'common /block_9/ vmix_vnumber
'Dim units_factor As Double
'common /block_16/ units_factor
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim vhelium_vtime_constant(16) As Double
'common /block_1a/ vhelium_vtime_constant
'Dim vnitrogen_vtime_constant(16) As Double
'common /block_1b/ vnitrogen_vtime_constant
'Dim vfraction_vhelium(20)  As Double
'Dim vfraction_vnitrogen(20) As Double
'common /block_5/ vfraction_vhelium, vfraction_vnitrogen
'Dim amb_vpressure_onset_of_imperm(16) As Double
'Dim gas_tension_onset_of_imperm(16) As Double
'common /block_13/ amb_vpressure_onset_of_imperm,                               gas_tension_onset_of_imperm
'Dim initial_vhelium_vpressure(16)  As Double
'Dim initial_vnitrogen_vpressure(16) As Double
'common /block_23/ initial_vhelium_vpressure,                                          initial_vnitrogen_vpressure
''=======================================================================
'     calculations
'     first convert the gradient for onset of impermeability to the divi
'     vpressure units that are being used
'=======================================================================
gradient_onset_of_imperm = gradient_onset_of_imperm_atm * units_factor
'=======================================================================
'     establish the bounds for the root search using the bisection metho
'     in this case, we are solving for time - the time when the ambient
'     minus the gas tension will be equal to the gradient for onset of
'     impermeabliity.  the low bound for time is set at zero and the hig
'     bound is set at the elapsed time (vsegment time) it took to go from
'     starting ambient vpressure to the ending ambient vpressure.  the des
'     ambient vpressure and gas tension at the onset of impermeability wi
'     be found somewhere between these endpoints.  the valgorithm checks
'     make sure that the solution lies in between these bounds by first
'     computing the low bound and high bound funcion values.
'=======================================================================
initial_inspired_vhe_vpressure = (starting_ambient_vpressure - water_vapor_vpressure) * vfraction_vhelium(vmix_vnumber)
initial_inspired_vn2_vpressure = (starting_ambient_vpressure - water_vapor_vpressure) * vfraction_vnitrogen(vmix_vnumber)
vhelium_rate = rate * vfraction_vhelium(vmix_vnumber)
vnitrogen_rate = rate * vfraction_vnitrogen(vmix_vnumber)
Set_Inspired_Inert_Starting
low_bound = 0#
high_bound = (ending_ambient_vpressure - starting_ambient_vpressure) / rate
starting_gas_tension = initial_vhelium_vpressure(i) + initial_vnitrogen_vpressure(i) + constant_vpressure_other_gases
funcion_at_low_bound = starting_ambient_vpressure - starting_gas_tension - gradient_onset_of_imperm
high_bound_vhelium_vpressure = schreiner_equation(initial_inspired_vhe_vpressure, vhelium_rate, high_bound, vhelium_vtime_constant(i), initial_vhelium_vpressure(i))
high_bound_vnitrogen_vpressure = schreiner_equation(initial_inspired_vn2_vpressure, vnitrogen_rate, high_bound, vnitrogen_vtime_constant(i), initial_vnitrogen_vpressure(i))
ending_gas_tension = high_bound_vhelium_vpressure + high_bound_vnitrogen_vpressure + constant_vpressure_other_gases
funcion_at_high_bound = ending_ambient_vpressure - ending_gas_tension - gradient_onset_of_imperm
If ((funcion_at_high_bound * funcion_at_low_bound) >= 0#) Then
    MsgBox "error' root is not within brackets"
End If
'=======================================================================
'     apply the bisection method in several iterations until a solution
'     the desired accuracy is found
'     note: the subroutine allows for up to 100 iterations.  normally an ex
'     be made from the loop well before that vnumber.  if, for some reaso
'     subroutine exceeds 100 iterations, there will be a pause to alert the
'=======================================================================
If (funcion_at_low_bound < 0#) Then
    xtime = low_bound
    differential_change = high_bound - low_bound
    Else
    xtime = high_bound
    differential_change = low_bound - high_bound
End If
For j = 1 To 200
    last_diff_change = differential_change
    differential_change = last_diff_change * 0.5
    mid_range_vtime = xtime + differential_change
    mid_range_ambient_vpressure = (starting_ambient_vpressure + rate * mid_range_vtime)
    mid_range_vhelium_vpressure = schreiner_equation(initial_inspired_vhe_vpressure, vhelium_rate, mid_range_vtime, vhelium_vtime_constant(i), initial_vhelium_vpressure(i))
    mid_range_vnitrogen_vpressure = schreiner_equation(initial_inspired_vn2_vpressure, vnitrogen_rate, mid_range_vtime, vnitrogen_vtime_constant(i), initial_vnitrogen_vpressure(i))
    gas_tension_at_mid_range = mid_range_vhelium_vpressure + mid_range_vnitrogen_vpressure + constant_vpressure_other_gases
    funcion_at_mid_range = mid_range_ambient_vpressure - gas_tension_at_mid_range - gradient_onset_of_imperm
    If (funcion_at_mid_range <= 0#) Then xtime = mid_range_vtime
    If ((Abs(differential_change) < 0.001) Or (funcion_at_mid_range = 0#)) Then GoTo L100
Next j
MsgBox "error' root search exceeded maximum iterations"
'=======================================================================
'     when a solution with the desired accuracy is found, the subroutine ju
'     of the loop to line 100 and assigns the solution values for ambien
'     vpressure and gas tension at the onset of impermeability.
'=======================================================================
L100:
amb_vpressure_onset_of_imperm(i) = mid_range_ambient_vpressure
gas_tension_onset_of_imperm(i) = gas_tension_at_mid_range
'=======================================================================
'     end of subroutine
'=======================================================================
End Sub

'=======================================================================
'     subroutine radius_root_finder
'     purpose: this subroutine is a "fail-safe" routine that combines th
'     bisection method and the newton-raphson method to find the desired
'     this hybrid valgorithm takes a bisection step whenever newton-raphs
'     take the solution out of bounds, or whenever newton-raphson is not
'     converging fast enough.  source:  "numerical recipes in fortran 77
'     cambridge university press, 1992.
'=======================================================================
Sub radius_root_finder(a As Double, b As Double, C As Double, low_bound As Double, high_bound As Double, ending_radius As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim a as double
' sub parameter : do not dim ! Dim  b as double
' sub parameter : do not dim ! Dim  c as double
' sub parameter : do not dim ! Dim  low_bound as double
' sub parameter : do not dim ! Dim  high_bound as double
' sub parameter : do not dim ! Dim ending_radius as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim funcion As Double
Dim derivative_of_funcion As Double
Dim differential_change As Double
Dim last_diff_change As Double
Dim last_ending_radius As Double
Dim radius_at_low_bound As Double
Dim radius_at_high_bound As Double
Dim funcion_at_low_bound As Double
Dim funcion_at_high_bound As Double
'=======================================================================
'     begin calculations by making sure that the root lies within bounds
'     in this case we are solving for radius in a cubic equation of the
'     ar^3 - br^2 - c = 0.  the coefficients a, b, and c were passed to
'     subroutine as arguments.
'=======================================================================
funcion_at_low_bound = low_bound * (low_bound * (a * low_bound - b)) - C
funcion_at_high_bound = high_bound * (high_bound * (a * high_bound - b)) - C
If ((funcion_at_low_bound > 0#) And (funcion_at_high_bound > 0#)) Then
    MsgBox "error' root is not within brackets"
End If
'=======================================================================
'     next the valgorithm checks for special conditions and then prepares
'     the first bisection.
'=======================================================================
If ((funcion_at_low_bound < 0#) And (funcion_at_high_bound < 0#)) Then
    MsgBox "error' root is not within brackets"
End If
If (funcion_at_low_bound = 0#) Then
    ending_radius = low_bound
    'Close #1: Exit Sub
    ElseIf (funcion_at_high_bound = 0#) Then
    ending_radius = high_bound
    'Close #1: Exit Sub
    ElseIf (funcion_at_low_bound < 0#) Then
    radius_at_low_bound = low_bound
    radius_at_high_bound = high_bound
    Else
    radius_at_high_bound = low_bound
    radius_at_low_bound = high_bound
End If
ending_radius = 0.5 * (low_bound + high_bound)
last_diff_change = Abs(high_bound - low_bound)
differential_change = last_diff_change
'=======================================================================
'     at this point, the newton-raphson method is applied which uses a f
'     and its first derivative to rapidly converge upon a solution.
'     note: the subroutine allows for up to 100 iterations.  normally an ex
'     be made from the loop well before that vnumber.  if, for some reaso
'     subroutine exceeds 100 iterations, there will be a pause to alert the
'     when a solution with the desired accuracy is found, exit is made f
'     loop by returning to the calling subroutine.  the last value of endin
'     radius has been assigned as the solution.
'=======================================================================
funcion = ending_radius * (ending_radius * (a * ending_radius - b)) - C
derivative_of_funcion = ending_radius * (ending_radius * 3# * a - 2# * b)
For i = 1 To 200
    If ((((ending_radius - radius_at_high_bound) * derivative_of_funcion - funcion) * ((ending_radius - radius_at_low_bound) * derivative_of_funcion - funcion) >= 0#) Or (Abs(2# * funcion) > (Abs(last_diff_change * derivative_of_funcion)))) Then
        last_diff_change = differential_change
        differential_change = 0.5 * (radius_at_high_bound - radius_at_low_bound)
        ending_radius = radius_at_low_bound + differential_change
        If (radius_at_low_bound = ending_radius) Then Exit Sub
        Else
        last_diff_change = differential_change
        differential_change = funcion / derivative_of_funcion
        last_ending_radius = ending_radius
        ending_radius = ending_radius - differential_change
        If (last_ending_radius = ending_radius) Then Exit Sub
    End If
    If (Abs(differential_change) < 0.000000000001) Then Exit Sub
    funcion = ending_radius * (ending_radius * (a * ending_radius - b)) - C
    derivative_of_funcion = ending_radius * (ending_radius * 3# * a - 2# * b)
    If (funcion < 0#) Then
        radius_at_low_bound = ending_radius
        Else
        radius_at_high_bound = ending_radius
    End If
Next i
MsgBox "error' root search exceeded maximum iterations"
'=======================================================================
'     end of subroutine
'=======================================================================
End Sub

'     subroutine gas_loadings_constant_vdepth
'     purpose: this subsubroutine applies the haldane equation to update th
'     gas loadings (partial vpressures of vhelium and vnitrogen) in the hal
'     compartments for a vsegment at constant vdepth.
'=======================================================================
Sub gas_loadings_constant_vdepth(vdepth As Double, run_vtime_end_of_vsegment As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim vdepth as double
' sub parameter : do not dim ! Dim  run_vtime_end_of_vsegment as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim last_vsegment_vnumber As Integer
Dim initial_vhelium_vpressure As Double
Dim initial_vnitrogen_vpressure As Double
'Dim ambient_vpressure As Double
Dim last_run_vtime As Double
'Dim haldane_equation                                 'function su as double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim water_vapor_vpressure As Double
''common /block_8/ water_vapor_vpressure
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim vsegment_vnumber                                         'bo as integer
'Dim run_vtime As Double
'Dim vsegment_vtime                                     'an as double
''common /block_2/ run_vtime, vsegment_vnumber, vsegment_vtime
'Dim ending_ambient_vpressure As Double
''common /block_4/ ending_ambient_vpressure
'Dim vmix_vnumber As Integer
''common /block_9/ vmix_vnumber
'Dim barometric_vpressure As Double
''common /block_18/ barometric_vpressure
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim vhelium_vtime_constant(16) As Double
''common /block_1a/ vhelium_vtime_constant
'Dim vnitrogen_vtime_constant(16) As Double
''common /block_1b/ vnitrogen_vtime_constant
'Dim vhelium_vpressure(16)  As Double
'Dim vnitrogen_vpressure(16)                 'bo as double
'common /block_3/ vhelium_vpressure, vnitrogen_vpressure            'an
'Dim vfraction_vhelium(20)  As Double
'Dim vfraction_vnitrogen(20) As Double
'common /block_5/ vfraction_vhelium, vfraction_vnitrogen
'=======================================================================
'     calculations
'=======================================================================
vsegment_vtime = run_vtime_end_of_vsegment - run_vtime
last_run_vtime = run_vtime_end_of_vsegment
run_vtime = last_run_vtime
last_vsegment_vnumber = vsegment_vnumber
vsegment_vnumber = last_vsegment_vnumber + 1
ambient_vpressure = vdepth + barometric_vpressure
inspired_vhelium_vpressure = (ambient_vpressure - water_vapor_vpressure) * vfraction_vhelium(vmix_vnumber)
inspired_vnitrogen_vpressure = (ambient_vpressure - water_vapor_vpressure) * vfraction_vnitrogen(vmix_vnumber)
ending_ambient_vpressure = ambient_vpressure
Set_Inspired_Inert
For i = 1 To 16
    initial_vhelium_vpressure = vhelium_vpressure(i)
    initial_vnitrogen_vpressure = vnitrogen_vpressure(i)
    vhelium_vpressure(i) = haldane_equation(initial_vhelium_vpressure, inspired_vhelium_vpressure, vhelium_vtime_constant(i), vsegment_vtime)
    vnitrogen_vpressure(i) = haldane_equation(initial_vnitrogen_vpressure, inspired_vnitrogen_vpressure, vnitrogen_vtime_constant(i), vsegment_vtime)
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
'=======================================================================
'     subroutine nuclear_regeneration
'     purpose: this subsubroutine calculates the regeneration of vpm critic
'     radii that takes place over the dive time.  the regeneration time
'     has a time scale of weeks so this will have very little impact on
'     normal length, but will have a major impact for saturation dives.
'=======================================================================
End Sub
Sub nuclear_regeneration(dive_vtime As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim dive_vtime as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim crushing_vpressure_pascals_he As Double
Dim crushing_vpressure_pascals_n2 As Double
Dim ending_radius_he As Double
Dim ending_radius_n2 As Double
Dim crush_vpressure_adjust_ratio_he As Double
Dim crush_vpressure_adjust_ratio_n2 As Double
Dim adj_crush_vpressure_vhe_pascals As Double
Dim adj_crush_vpressure_vn2_pascals As Double
On Error Resume Next
''=======================================================================
''     global constants in named common blocks
''=======================================================================
'Dim surface_tension_gamma As Double
'Dim skin_compression_gammac As Double
'common /block_19/ surface_tension_gamma, skin_compression_gammac
'Dim regeneration_vtime_constant As Double
'common /block_22/ regeneration_vtime_constant
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim units_factor As Double
'common /block_16/ units_factor
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim adjusted_critical_radius_he(16) As Double
'Dim adjusted_critical_radius_n2(16) As Double
'common /block_7/ adjusted_critical_radius_he,                                      adjusted_critical_radius_n2
'Dim max_crushing_vpressure_he(16)  As Double
'Dim max_crushing_vpressure_n2(16) As Double
'common /block_10/ max_crushing_vpressure_he,                                         max_crushing_vpressure_n2
'Dim regenerated_radius_he(16)  As Double
'Dim regenerated_radius_n2(16) As Double
'common /block_24/ regenerated_radius_he, regenerated_radius_n2
'Dim adjusted_crushing_vpressure_he(16) As Double
'Dim adjusted_crushing_vpressure_n2(16) As Double
'common /block_25/ adjusted_crushing_vpressure_he,                                    adjusted_crushing_vpressure_n2
''=======================================================================
'     calculations
'     first convert the maximum crushing vpressure obtained for each comp
'     to pascals.  next, compute the ending radius for vhelium and nitrog
'     critical nuclei in each compartment.
'=======================================================================
For i = 1 To 16
    crushing_vpressure_pascals_he = (max_crushing_vpressure_he(i) / units_factor) * 101325#
    crushing_vpressure_pascals_n2 = (max_crushing_vpressure_n2(i) / units_factor) * 101325#
    ending_radius_he = 1# / (crushing_vpressure_pascals_he / (2# * (skin_compression_gammac - surface_tension_gamma)) + 1# / adjusted_critical_radius_he(i))
    ending_radius_n2 = 1# / (crushing_vpressure_pascals_n2 / (2# * (skin_compression_gammac - surface_tension_gamma)) + 1# / adjusted_critical_radius_n2(i))
    '=======================================================================
    '     a "regenerated" radius for each nucleus is now calculated based on
    '     regeneration time constant.  this means that after application of
    '     crushing vpressure and reduction in radius, a nucleus will slowly g
    '     back to its original initial radius over a period of time.  this
    '     phenomenon is probabilistic in nature and depends on absolute temp
    '     it is independent of crushing vpressure.
    '=======================================================================
    regenerated_radius_he(i) = adjusted_critical_radius_he(i) + (ending_radius_he - adjusted_critical_radius_he(i)) * Exp(-dive_vtime / regeneration_vtime_constant)
    regenerated_radius_n2(i) = adjusted_critical_radius_n2(i) + (ending_radius_n2 - adjusted_critical_radius_n2(i)) * Exp(-dive_vtime / regeneration_vtime_constant)
    '=======================================================================
    '     in order to preserve reference back to the initial critical radii
    '     regeneration, an "adjusted crushing vpressure" for the nuclei in ea
    '     compartment must be computed.  in other words, this is the value o
    '     crushing vpressure that would have reduced the original nucleus to
    '     to the present radius had regeneration not taken place.  the ratio
    '     for adjusting crushing vpressure is obtained from algebraic manipul
    '     of the standard vpm equations.  the adjusted crushing vpressure, in
    '     of the original crushing vpressure, is then applied in the vpm crit
    '     volume valgorithm and the vpm repetitive valgorithm.
    '=======================================================================
    crush_vpressure_adjust_ratio_he = (ending_radius_he * (adjusted_critical_radius_he(i) - regenerated_radius_he(i))) / (regenerated_radius_he(i) * (adjusted_critical_radius_he(i) - ending_radius_he))
    crush_vpressure_adjust_ratio_n2 = (ending_radius_n2 * (adjusted_critical_radius_n2(i) - regenerated_radius_n2(i))) / (regenerated_radius_n2(i) * (adjusted_critical_radius_n2(i) - ending_radius_n2))
    adj_crush_vpressure_vhe_pascals = crushing_vpressure_pascals_he * crush_vpressure_adjust_ratio_he
    adj_crush_vpressure_vn2_pascals = crushing_vpressure_pascals_n2 * crush_vpressure_adjust_ratio_n2
    adjusted_crushing_vpressure_he(i) = (adj_crush_vpressure_vhe_pascals / 101325#) * units_factor
    adjusted_crushing_vpressure_n2(i) = (adj_crush_vpressure_vn2_pascals / 101325#) * units_factor
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
End Sub
'=======================================================================
'     subroutine calc_initial_allowable_gradient
'     purpose: this subsubroutine calculates the initial allowable gradient
'     vhelium and nitrogren in each compartment.  these are the gradients
'     will be used to set the vdeco vceiling on the first pass through the
'     loop.  if the critical volume valgorithm is set to "off", then thes
'     gradients will determine the final vdeco schedule.  otherwise, if t
'     critical volume valgorithm is set to "on", these gradients will be
'     "relaxed" by the critical volume valgorithm subroutine.  the initia
'     allowable gradients are referred to as "pssmin" in the papers by y
'     and colleauges, i.e., the minimum supersaturation vpressure gradien
' '     that will probe bubble formation in the vpm nuclei that started wi
'     designated minimum initial radius (critical radius) .
'
'     the initial allowable gradients are computed directly from the
'     "regenerated" radii after the nuclear regeneration subroutine.  th
'     gradients are tracked separately for vhelium and vnitrogen.
'=======================================================================
Sub calc_initial_allowable_gradient()
'      implicit none
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim initial_allowable_grad_vhe_pa As Double
Dim initial_allowable_grad_vn2_pa As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim surface_tension_gamma As Double
'Dim skin_compression_gammac As Double
'common /block_19/ surface_tension_gamma, skin_compression_gammac
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim units_factor As Double
'common /block_16/ units_factor
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim regenerated_radius_he(16)  As Double
'Dim regenerated_radius_n2(16) As Double
'common /block_24/ regenerated_radius_he, regenerated_radius_n2
'Dim allowable_gradient_he(16)  As Double
'Dim allowable_gradient_n2(16) As Double
'common /block_26/ allowable_gradient_he, allowable_gradient_n2
'Dim initial_allowable_gradient_he(16) As Double
'Dim initial_allowable_gradient_n2(16) As Double
'common /block_27/                                                     initial_allowable_gradient_he, initial_allowable_gradient_n2
''=======================================================================
'     calculations
'     the initial allowable gradients are computed in pascals and then c
'     to the diving vpressure units.  two different sets of arrays are us
'     save the calculations - initial allowable gradients and allowable
'     gradients.  the allowable gradients are assigned the values from i
'     allowable gradients however the allowable gradients can be changed
'     by the critical volume subroutine.  the values for the initial all
'     gradients are saved in a global array for later use by both the cr
'     volume subroutine and the vpm repetitive valgorithm subroutine.
'=======================================================================
For i = 1 To 16
    initial_allowable_grad_vn2_pa = ((2# * surface_tension_gamma * (skin_compression_gammac - surface_tension_gamma)) / (regenerated_radius_n2(i) * skin_compression_gammac))
    initial_allowable_grad_vhe_pa = ((2# * surface_tension_gamma * (skin_compression_gammac - surface_tension_gamma)) / (regenerated_radius_he(i) * skin_compression_gammac))
    initial_allowable_gradient_n2(i) = (initial_allowable_grad_vn2_pa / 101325#) * units_factor
    initial_allowable_gradient_he(i) = (initial_allowable_grad_vhe_pa / 101325#) * units_factor
    allowable_gradient_he(i) = initial_allowable_gradient_he(i)
    allowable_gradient_n2(i) = initial_allowable_gradient_n2(i)
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
'=======================================================================
'     subroutine calc_ascent_vceiling
'     purpose: this subsubroutine calculates the ascent vceiling (the safe a
'     vdepth) in each compartment, based on the allowable gradients, and
'     finds the deepest ascent vceiling across all compartments.
'=======================================================================
End Sub
Sub calc_ascent_vceiling(ascent_vceiling_vdepth As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim ascent_vceiling_vdepth as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim gas_loading As Double
Dim weighted_allowable_gradient As Double
Dim tolerated_ambient_vpressure As Double
'=======================================================================
'     local arrays
'=======================================================================
Dim compartment_ascent_vceiling(16) As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim constant_vpressure_other_gases As Double
'common /block_17/ constant_vpressure_other_gases
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim barometric_vpressure As Double
'common /block_18/ barometric_vpressure
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim vhelium_vpressure(16)  As Double
'Dim vnitrogen_vpressure(16) As Double
'common /block_3/ vhelium_vpressure, vnitrogen_vpressure
'Dim allowable_gradient_he(16)  As Double
'Dim allowable_gradient_n2(16) As Double
'common /block_26/ allowable_gradient_he, allowable_gradient_n2
'=======================================================================
'     calculations
'     since there are two sets of allowable gradients being tracked, one
'     vhelium and one for vnitrogen, a "weighted allowable gradient" must
'     computed each time based on the proportions of vhelium and vnitrogen
'     each compartment.  this proportioning follows the methodology of
'     buhlmann/keller.  if there is no vhelium and vnitrogen in the compar
'     such as after extended periods of voxygen breathing, then the minim
'     across both gases will be used.  it is important to note that if a
'     compartment is empty of vhelium and vnitrogen, then the weighted all
'     gradient formula cannot be used since it will result in division b
'=======================================================================
For i = 1 To 16
    gas_loading = vhelium_vpressure(i) + vnitrogen_vpressure(i)
    If (gas_loading > 0#) Then
        weighted_allowable_gradient = (allowable_gradient_he(i) * vhelium_vpressure(i) + allowable_gradient_n2(i) * vnitrogen_vpressure(i)) / (vhelium_vpressure(i) + vnitrogen_vpressure(i))
        tolerated_ambient_vpressure = (gas_loading + constant_vpressure_other_gases) - weighted_allowable_gradient
        Else
        weighted_allowable_gradient = Min(allowable_gradient_he(i), allowable_gradient_n2(i))
        tolerated_ambient_vpressure = constant_vpressure_other_gases - weighted_allowable_gradient
    End If
    '=======================================================================
    '     the tolerated ambient vpressure cannot be less than zero absolute,
    '     the vacuum of outer space!
    '=======================================================================
    If (tolerated_ambient_vpressure < 0#) Then
        tolerated_ambient_vpressure = 0#
    End If
    '=======================================================================
    '     the ascent vceiling vdepth is computed in a loop after all of the in
    '     compartment ascent vceilings have been calculated.  it is important
    '     the ascent vceiling vdepth (max ascent vceiling across all compartmen
    '     be extracted from the compartment values and not be compared again
    '     initialization value.  for example, if max(ascent_vceiling_vdepth .
    '     compared against zero, this could cause a subroutine lockup because s
    '     the ascent vceiling vdepth needs to be negative (but not less than z
    '     absolute ambient vpressure) in order to vdecompress to the last vstop
    '     vdepth.
    '=======================================================================
    compartment_ascent_vceiling(i) = tolerated_ambient_vpressure - barometric_vpressure
Next i
ascent_vceiling_vdepth = compartment_ascent_vceiling(1)
For i = 2 To 16
    ascent_vceiling_vdepth = Max(ascent_vceiling_vdepth, compartment_ascent_vceiling(i))
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
'=======================================================================
'     subroutine calc_max_actual_gradient
'     purpose: this subsubroutine calculates the actual supersaturation gra
'     obtained in each compartment as a result of the ascent profile dur
'     vdecompression.  similar to the concept with crushing vpressure, the
'     supersaturation gradients are not cumulative over a multi-level, s
'     ascent.  rather, it will be the maximum value obtained in any one
'     step of the overall ascent.  thus, the subroutine must compute and st
'     maximum actual gradient for each compartment that was obtained acr
'     steps of the ascent profile.  this subroutine is invoked on the la
'     through the vdeco vstop loop block when the final vdeco schedule is b
'     generated.
'
'     the max actual gradients are later used by the vpm repetitive algo
'     determine if adjustments to the critical radii are required.  if t
'     actual gradient did not exceed the initial alllowable gradient, th
'     adjustment will be made.  however, if the max actual gradient did
'     the intitial allowable gradient, such as permitted by the critical
'     valgorithm, then the critical radius will be adjusted (made larger)
'     repetitive dive to compensate for the bubbling that was allowed on
'     previous dive.  the use of the max actual gradients is intended to
'     the repetitive valgorithm from being overly conservative.
'=======================================================================
End Sub
Sub calc_max_actual_gradient(vdeco_vstop_vdepth As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
''Dim vdeco_vstop_vdepth as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim compartment_gradient As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim constant_vpressure_other_gases As Double
'common /block_17/ constant_vpressure_other_gases
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim barometric_vpressure As Double
'common /block_18/ barometric_vpressure
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim vhelium_vpressure(16)  As Double
'Dim vnitrogen_vpressure(16) As Double
'common /block_3/ vhelium_vpressure, vnitrogen_vpressure
'Dim max_actual_gradient(16) As Double
'common /block_12/ max_actual_gradient
''=======================================================================
'     calculations
'     note: negative supersaturation gradients are meaningless for this
'     application, so the values must be equal to or greater than zero.
'=======================================================================
For i = 1 To 16
    compartment_gradient = (vhelium_vpressure(i) + vnitrogen_vpressure(i) + constant_vpressure_other_gases) - (vdeco_vstop_vdepth + barometric_vpressure)
    If (compartment_gradient <= 0#) Then
        compartment_gradient = 0#
    End If
    max_actual_gradient(i) = Max(max_actual_gradient(i), compartment_gradient)
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
'=======================================================================
'     subroutine calc_surface_phase_volume_vtime
'     purpose: this subsubroutine computes the surface portion of the total
'     volume time.  this is the time factored out of the integration of
'     supersaturation gradient x time over the surface interval.  the vp
'     considers the gradients that allow bubbles to form or to drive bub
'     growth both in the water and on the surface after the dive.
'
'     this subroutine is a new development to the vpm valgorithm in that
'     computes the time course of supersaturation gradients on the surfa
'     when both vhelium and vnitrogen are present.  refer to separate writ
'     for a more detailed explanation of this valgorithm.
'=======================================================================
End Sub
Sub calc_surface_phase_volume_vtime()
'      implicit none
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim integral_gradient_x_vtime As Double
Dim decay_vtime_to_zero_gradient As Double
Dim surface_inspired_vn2_vpressure As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
''Dim water_vapor_vpressure As Double
'common /block_8/ water_vapor_vpressure
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim barometric_vpressure As Double
'common /block_18/ barometric_vpressure
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim vhelium_vtime_constant(16) As Double
'common /block_1a/ vhelium_vtime_constant
'Dim vnitrogen_vtime_constant(16) As Double
'common /block_1b/ vnitrogen_vtime_constant
'Dim vhelium_vpressure(16)  As Double
'Dim vnitrogen_vpressure(16) As Double
'common /block_3/ vhelium_vpressure, vnitrogen_vpressure
'Dim surface_phase_volume_vtime(16) As Double
'common /block_11/ surface_phase_volume_vtime
''=======================================================================
'     calculations
'=======================================================================
surface_inspired_vn2_vpressure = (barometric_vpressure - water_vapor_vpressure) * 0.79
For i = 1 To 16
    If (vnitrogen_vpressure(i) > surface_inspired_vn2_vpressure) Then
        surface_phase_volume_vtime(i) = (vhelium_vpressure(i) / vhelium_vtime_constant(i) + (vnitrogen_vpressure(i) - surface_inspired_vn2_vpressure) / vnitrogen_vtime_constant(i)) / (vhelium_vpressure(i) + vnitrogen_vpressure(i) - surface_inspired_vn2_vpressure)
        ElseIf ((vnitrogen_vpressure(i) <= surface_inspired_vn2_vpressure) And (vhelium_vpressure(i) + vnitrogen_vpressure(i) >= surface_inspired_vn2_vpressure)) Then
        decay_vtime_to_zero_gradient = 1# / (vnitrogen_vtime_constant(i) - vhelium_vtime_constant(i)) * Log((surface_inspired_vn2_vpressure - vnitrogen_vpressure(i)) / vhelium_vpressure(i))
        integral_gradient_x_vtime = vhelium_vpressure(i) / vhelium_vtime_constant(i) * (1# - Exp(-vhelium_vtime_constant(i) * decay_vtime_to_zero_gradient)) + (vnitrogen_vpressure(i) - surface_inspired_vn2_vpressure) / vnitrogen_vtime_constant(i) * (1# - Exp(-vnitrogen_vtime_constant(i) * decay_vtime_to_zero_gradient))
        surface_phase_volume_vtime(i) = integral_gradient_x_vtime / (vhelium_vpressure(i) + vnitrogen_vpressure(i) - surface_inspired_vn2_vpressure)
        Else
        surface_phase_volume_vtime(i) = 0#
    End If
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
'=======================================================================
'     subroutine critical_volume
'     purpose: this subsubroutine applies the vpm critical volume valgorithm
'     valgorithm will compute "relaxed" gradients for vhelium and vnitrogen
'     on the setting of the critical volume parameter lambda.
'=======================================================================
End Sub
Sub critical_volume(vdeco_phase_volume_vtime As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim vdeco_phase_volume_vtime as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim parameter_lambda_pascals As Double
Dim adj_crush_vpressure_vhe_pascals As Double
Dim adj_crush_vpressure_vn2_pascals As Double
Dim initial_allowable_grad_vhe_pa As Double
Dim initial_allowable_grad_vn2_pa As Double
Dim new_allowable_grad_vhe_pascals As Double
Dim new_allowable_grad_vn2_pascals As Double
Dim b As Double
Dim C As Double
'=======================================================================
'     local arrays
'=======================================================================
Dim phase_volume_vtime(16) As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim surface_tension_gamma As Double
'Dim skin_compression_gammac As Double
'common /block_19/ surface_tension_gamma, skin_compression_gammac
'Dim crit_volume_parameter_lambda As Double
'common /block_20/ crit_volume_parameter_lambda
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim units_factor As Double
'common /block_16/ units_factor
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim adjusted_critical_radius_he(16) As Double
'Dim adjusted_critical_radius_n2(16) As Double
'common /block_7/ adjusted_critical_radius_he,                                adjusted_critical_radius_n2
'Dim surface_phase_volume_vtime(16) As Double
'common /block_11/ surface_phase_volume_vtime
'Dim adjusted_crushing_vpressure_he(16) As Double
'Dim adjusted_crushing_vpressure_n2(16) As Double
'common /block_25/ adjusted_crushing_vpressure_he,                                    adjusted_crushing_vpressure_n2
'Dim allowable_gradient_he(16)  As Double
'Dim allowable_gradient_n2(16) As Double
'common /block_26/ allowable_gradient_he, allowable_gradient_n2
'Dim initial_allowable_gradient_he(16) As Double
'Dim initial_allowable_gradient_n2(16) As Double
'common /block_27/                                                     initial_allowable_gradient_he, initial_allowable_gradient_n2
''=======================================================================
'     calculations
'     note:  since the critical volume parameter lambda was defined in u
'     fsw-min in the original papers by yount and colleauges, the same
'     convention is retained here.  although lambda is adjustable only i
'     of fsw-min in the subroutine settings (range from 6500 to 8300 with d
'     7500) , it will convert to the proper value in pascals-min in this
'     subroutine regardless of which diving vpressure units are being use
'     the main subroutine - feet of seawater (fsw) or meters of seawater (m
'     the allowable gradient is computed using the quadratic formula (re
'     separate write-up posted on the vdeco list web site) .
'=======================================================================
parameter_lambda_pascals = (crit_volume_parameter_lambda / 33#) * 101325#
For i = 1 To 16
    phase_volume_vtime(i) = vdeco_phase_volume_vtime + surface_phase_volume_vtime(i)
Next i
For i = 1 To 16
    adj_crush_vpressure_vhe_pascals = (adjusted_crushing_vpressure_he(i) / units_factor) * 101325#
    initial_allowable_grad_vhe_pa = (initial_allowable_gradient_he(i) / units_factor) * 101325#
    b = initial_allowable_grad_vhe_pa + (parameter_lambda_pascals * surface_tension_gamma) / (skin_compression_gammac * phase_volume_vtime(i))
    C = (surface_tension_gamma * (surface_tension_gamma * (parameter_lambda_pascals * adj_crush_vpressure_vhe_pascals))) / (skin_compression_gammac * (skin_compression_gammac * phase_volume_vtime(i)))
    new_allowable_grad_vhe_pascals = (b + Sqr(b ^ 2 - 4# * C)) / 2#
    allowable_gradient_he(i) = (new_allowable_grad_vhe_pascals / 101325#) * units_factor
Next i
For i = 1 To 16
    adj_crush_vpressure_vn2_pascals = (adjusted_crushing_vpressure_n2(i) / units_factor) * 101325#
    initial_allowable_grad_vn2_pa = (initial_allowable_gradient_n2(i) / units_factor) * 101325#
    b = initial_allowable_grad_vn2_pa + (parameter_lambda_pascals * surface_tension_gamma) / (skin_compression_gammac * phase_volume_vtime(i))
    C = (surface_tension_gamma * (surface_tension_gamma * (parameter_lambda_pascals * adj_crush_vpressure_vn2_pascals))) / (skin_compression_gammac * (skin_compression_gammac * phase_volume_vtime(i)))
    new_allowable_grad_vn2_pascals = (b + Sqr(b ^ 2 - 4# * C)) / 2#
    allowable_gradient_n2(i) = (new_allowable_grad_vn2_pascals / 101325#) * units_factor
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
End Sub

'=======================================================================
'     subroutine calc_start_of_vdeco_zone
'     purpose: this subroutine uses the bisection method to find the dep
'     which the leading compartment just enters the vdecompression zone.
'     source:  "numerical recipes in fortran 77", cambridge university p
'     1992.
'=======================================================================
Sub calc_start_of_vdeco_zone(starting_vdepth As Double, rate As Double, vdepth_start_of_vdeco_zone As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim starting_vdepth as double
' sub parameter : do not dim ! Dim  rate as double
' sub parameter : do not dim ! Dim  vdepth_start_of_vdeco_zone as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer
Dim j As Integer                                               'loop as integer
Dim initial_vhelium_vpressure As Double
Dim initial_vnitrogen_vpressure As Double
'Dim initial_inspired_vhe_vpressure As Double
'Dim initial_inspired_vn2_vpressure As Double
Dim time_to_start_of_vdeco_zone As Double
Dim vhelium_rate As Double
Dim vnitrogen_rate As Double
'Dim starting_ambient_vpressure As Double
Dim cpt_vdepth_start_of_vdeco_zone As Double
Dim low_bound As Double
Dim high_bound As Double
Dim high_bound_vhelium_vpressure As Double
Dim high_bound_vnitrogen_vpressure As Double
Dim mid_range_vhelium_vpressure As Double
Dim mid_range_vnitrogen_vpressure As Double
Dim funcion_at_high_bound As Double
Dim funcion_at_low_bound As Double
Dim mid_range_vtime As Double
Dim funcion_at_mid_range As Double
Dim differential_change As Double
Dim last_diff_change As Double
'Dim schreiner_equation                               'funcion sub as double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim water_vapor_vpressure As Double
'common /block_8/ water_vapor_vpressure
'Dim constant_vpressure_other_gases As Double
'common /block_17/ constant_vpressure_other_gases
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim vmix_vnumber As Integer
'common /block_9/ vmix_vnumber
'Dim barometric_vpressure As Double
'common /block_18/ barometric_vpressure
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim vhelium_vtime_constant(16) As Double
'common /block_1a/ vhelium_vtime_constant
'Dim vnitrogen_vtime_constant(16) As Double
'common /block_1b/ vnitrogen_vtime_constant
'Dim vhelium_vpressure(16)  As Double
'Dim vnitrogen_vpressure(16) As Double
'common /block_3/ vhelium_vpressure, vnitrogen_vpressure
'Dim vfraction_vhelium(20)  As Double
'Dim vfraction_vnitrogen(20) As Double
'common /block_5/ vfraction_vhelium, vfraction_vnitrogen
''=======================================================================
'     calculations
'     first initialize some variables
'=======================================================================
vdepth_start_of_vdeco_zone = 0#
starting_ambient_vpressure = starting_vdepth + barometric_vpressure
initial_inspired_vhe_vpressure = (starting_ambient_vpressure - water_vapor_vpressure) * vfraction_vhelium(vmix_vnumber)
initial_inspired_vn2_vpressure = (starting_ambient_vpressure - water_vapor_vpressure) * vfraction_vnitrogen(vmix_vnumber)
vhelium_rate = rate * vfraction_vhelium(vmix_vnumber)
vnitrogen_rate = rate * vfraction_vnitrogen(vmix_vnumber)
Set_Inspired_Inert_Starting
'=======================================================================
'     establish the bounds for the root search using the bisection metho
'     and check to make sure that the root will be within bounds.  proce
'     each compartment individually and find the maximum vdepth across al
'     compartments (leading compartment)
'     in this case, we are solving for time - the time when the gas tens
'     the compartment will be equal to ambient vpressure.  the low bound
'     is set at zero and the high bound is set at the time it would take
'     ascend to zero ambient vpressure (absolute) .  since the ascent rate
'     negative, a multiplier of -1.0 is used to make the time positive.
'     desired point when gas tension equals ambient vpressure is found at
'     somewhere between these endpoints.  the valgorithm checks to make s
'     the solution lies in between these bounds by first computing the l
'     and high bound funcion values.
'=======================================================================
low_bound = 0#
high_bound = -1# * (starting_ambient_vpressure / rate)
For i = 1 To 16
    initial_vhelium_vpressure = vhelium_vpressure(i)
    initial_vnitrogen_vpressure = vnitrogen_vpressure(i)
    funcion_at_low_bound = initial_vhelium_vpressure + initial_vnitrogen_vpressure + constant_vpressure_other_gases - starting_ambient_vpressure
    high_bound_vhelium_vpressure = schreiner_equation(initial_inspired_vhe_vpressure, vhelium_rate, high_bound, vhelium_vtime_constant(i), initial_vhelium_vpressure)
    high_bound_vnitrogen_vpressure = schreiner_equation(initial_inspired_vn2_vpressure, vnitrogen_rate, high_bound, vnitrogen_vtime_constant(i), initial_vnitrogen_vpressure)
    funcion_at_high_bound = high_bound_vhelium_vpressure + high_bound_vnitrogen_vpressure + constant_vpressure_other_gases
    If ((funcion_at_high_bound * funcion_at_low_bound) >= 0#) Then
        'Open "report.txt" For Append As #1
        'Write #1, "error' root is not within brackets"
        no_deco_found = 1 ' MsgBox "root not in brackets"
    End If
    '=======================================================================
    '     apply the bisection method in several iterations until a solution
    '     the desired accuracy is found
    '     note: the subroutine allows for up to 100 iterations.  normally an ex
    '     be made from the loop well before that vnumber.  if, for some reaso
    '     subroutine exceeds 100 iterations, there will be a pause to alert the
    '=======================================================================
    If (funcion_at_low_bound < 0#) Then
        time_to_start_of_vdeco_zone = low_bound
        differential_change = high_bound - low_bound
        Else
        time_to_start_of_vdeco_zone = high_bound
        differential_change = low_bound - high_bound
    End If
    For j = 1 To 200
        last_diff_change = differential_change
        differential_change = last_diff_change * 0.5
        mid_range_vtime = time_to_start_of_vdeco_zone + differential_change
        mid_range_vhelium_vpressure = schreiner_equation(initial_inspired_vhe_vpressure, vhelium_rate, mid_range_vtime, vhelium_vtime_constant(i), initial_vhelium_vpressure)
        mid_range_vnitrogen_vpressure = schreiner_equation(initial_inspired_vn2_vpressure, vnitrogen_rate, mid_range_vtime, vnitrogen_vtime_constant(i), initial_vnitrogen_vpressure)
        funcion_at_mid_range = mid_range_vhelium_vpressure + mid_range_vnitrogen_vpressure + constant_vpressure_other_gases - (starting_ambient_vpressure + rate * mid_range_vtime)
        If (funcion_at_mid_range <= 0#) Then time_to_start_of_vdeco_zone = mid_range_vtime
        If ((Abs(differential_change) < 0.001) Or (funcion_at_mid_range = 0#)) Then GoTo L170
    Next j
L150:
        'Write #1, "error' root search exceeded maximum iterations"
    no_deco_found = 2 ' MsgBox "root not in brackets"
        '=======================================================================
        '     when a solution with the desired accuracy is found, the subroutine ju
        '     of the loop to line 170 and assigns the solution value for the ind
        '     compartment.
        '=======================================================================
L170:
    cpt_vdepth_start_of_vdeco_zone = (starting_ambient_vpressure + rate * time_to_start_of_vdeco_zone) - barometric_vpressure
        '=======================================================================
        '     the overall solution will be the compartment with the maximum dept
        '     gas tension equals ambient vpressure (leading compartment) .
        '=======================================================================
    vdepth_start_of_vdeco_zone = Max(vdepth_start_of_vdeco_zone, cpt_vdepth_start_of_vdeco_zone)
Next i
L200:
        '=======================================================================
        '     end of subroutine
        '=======================================================================
        'Close #1: Exit Sub
        'Close #1: Exit Sub
        'Close #1
End Sub

'=======================================================================
'     subroutine projected_ascent
'     purpose: this subsubroutine performs a simulated ascent outside of th
'     subroutine to ensure that a vdeco vceiling will not be violated due to
'     gas loading during ascent (on-gassing) .  if the vdeco vceiling is vi
'     the vstop vdepth will be adjusted deeper by the step size until a sa
'     ascent can be made.
'=======================================================================
Sub projected_ascent(starting_vdepth As Double, rate As Double, vdeco_vstop_vdepth As Double, vstep_size As Double)
        '      implicit none
        '=======================================================================
        '     arguments
        '=======================================================================
        ' sub parameter : do not dim ! Dim starting_vdepth as double
        ' sub parameter : do not dim ! Dim  rate as double
        ' sub parameter : do not dim ! Dim  vstep_size as double
        'Dim vdeco_vstop_vdepth                                     'input an as double
        '=======================================================================
        '     local variables
        '=======================================================================
        Dim i As Integer                                                  'loop as integer
        'Dim initial_inspired_vhe_vpressure As Double
        'Dim initial_inspired_vn2_vpressure As Double
        Dim vhelium_rate As Double
        Dim vnitrogen_rate As Double
        'Dim starting_ambient_vpressure As Double
        Dim ending_ambient_vpressure As Double
        Dim new_ambient_vpressure As Double
        Dim vsegment_vtime As Double
        Dim temp_vhelium_vpressure As Double
        Dim temp_vnitrogen_vpressure As Double
        Dim weighted_allowable_gradient As Double
        'Dim schreiner_equation                               'funcion sub as double
        '=======================================================================
        '     local arrays
        '=======================================================================
        Dim initial_vhelium_vpressure(16)  As Double
        Dim initial_vnitrogen_vpressure(16) As Double
        Dim temp_gas_loading(16)  As Double
        Dim allowable_gas_loading(16) As Double
        '=======================================================================
        '     global constants in named common blocks
        '=======================================================================
        'Dim water_vapor_vpressure As Double
        'common /block_8/ water_vapor_vpressure
        'Dim constant_vpressure_other_gases As Double
        'common /block_17/ constant_vpressure_other_gases
        ''=======================================================================
        ''     global variables in named common blocks
        ''=======================================================================
        'Dim vmix_vnumber As Integer
        'common /block_9/ vmix_vnumber
        'Dim barometric_vpressure As Double
        'common /block_18/ barometric_vpressure
        ''=======================================================================
        ''     global arrays in named common blocks
        ''=======================================================================
        'Dim vhelium_vtime_constant(16) As Double
        'common /block_1a/ vhelium_vtime_constant
        'Dim vnitrogen_vtime_constant(16) As Double
        'common /block_1b/ vnitrogen_vtime_constant
        'Dim vhelium_vpressure(16)  As Double
        'Dim vnitrogen_vpressure(16) As Double
        'common /block_3/ vhelium_vpressure, vnitrogen_vpressure
        'Dim vfraction_vhelium(20)  As Double
        'Dim vfraction_vnitrogen(20) As Double
        'common /block_5/ vfraction_vhelium, vfraction_vnitrogen
        'Dim allowable_gradient_he(16)  As Double
        'Dim allowable_gradient_n2(16) As Double
        'common /block_26/ allowable_gradient_he, allowable_gradient_n2
        ''=======================================================================
        '     calculations
        '=======================================================================
        new_ambient_vpressure = vdeco_vstop_vdepth + barometric_vpressure
        starting_ambient_vpressure = starting_vdepth + barometric_vpressure
        initial_inspired_vhe_vpressure = (starting_ambient_vpressure - water_vapor_vpressure) * vfraction_vhelium(vmix_vnumber)
        initial_inspired_vn2_vpressure = (starting_ambient_vpressure - water_vapor_vpressure) * vfraction_vnitrogen(vmix_vnumber)
        vhelium_rate = rate * vfraction_vhelium(vmix_vnumber)
        vnitrogen_rate = rate * vfraction_vnitrogen(vmix_vnumber)
        Set_Inspired_Inert_Starting
        For i = 1 To 16
            initial_vhelium_vpressure(i) = vhelium_vpressure(i)
            initial_vnitrogen_vpressure(i) = vnitrogen_vpressure(i)
        Next i
L665:
        ending_ambient_vpressure = new_ambient_vpressure
        vsegment_vtime = (ending_ambient_vpressure - starting_ambient_vpressure) / rate
        For i = 1 To 16
            temp_vhelium_vpressure = schreiner_equation(initial_inspired_vhe_vpressure, vhelium_rate, vsegment_vtime, vhelium_vtime_constant(i), initial_vhelium_vpressure(i))
            temp_vnitrogen_vpressure = schreiner_equation(initial_inspired_vn2_vpressure, vnitrogen_rate, vsegment_vtime, vnitrogen_vtime_constant(i), initial_vnitrogen_vpressure(i))
            temp_gas_loading(i) = temp_vhelium_vpressure + temp_vnitrogen_vpressure
            If (temp_gas_loading(i) > 0#) Then
                weighted_allowable_gradient = (allowable_gradient_he(i) * temp_vhelium_vpressure + allowable_gradient_n2(i) * temp_vnitrogen_vpressure) / temp_gas_loading(i)
                Else
                weighted_allowable_gradient = Min(allowable_gradient_he(i), allowable_gradient_n2(i))
            End If
            allowable_gas_loading(i) = ending_ambient_vpressure + weighted_allowable_gradient - constant_vpressure_other_gases
            
            nfrac = temp_vnitrogen_vpressure / (temp_vnitrogen_vpressure + temp_vhelium_vpressure)
            hefrac = temp_vhelium_vpressure / (temp_vnitrogen_vpressure + temp_vhelium_vpressure)
            If (hefrac = 0) Then hefrac = 0.005
            If (nfrac = 0) Then nfrac = 0.005
            atotal = (nfrac * an2(i - 1) + hefrac * ahe(i - 1)) / (nfrac + hefrac)
            btotal = (nfrac * bn2(i - 1) + hefrac * bhe(i - 1)) / (nfrac + hefrac)
        
            atotal = atotal * 0.9 * (1# - (CDbl(safetytext.Text) / 300#))
            btotal = btotal * 1.1 * (1# + CDbl(safetytext.Text) / 300#)
        
            buhlptoln2(i - 1) = new_ambient_vpressure / btotal + (atotal * units_factor)
            If buhlptoln2(i - 1) < allowable_gas_loading(i) Then
              If buhl_mode > 0 Then allowable_gas_loading(i) = buhlptoln2(i - 1)
            End If
            If buhl_mode = 2 Then allowable_gas_loading(i) = buhlptoln2(i - 1)
        
            'If (tolerated_ambient_vpressure < 0#) Then
            '    tolerated_ambient_vpressure = 0#
            'End If

        Next i
L670:
        For i = 1 To 16
                If (temp_gas_loading(i) > allowable_gas_loading(i)) Then
                    new_ambient_vpressure = ending_ambient_vpressure + vstep_size
                    vdeco_vstop_vdepth = vdeco_vstop_vdepth + vstep_size
                    GoTo L665
                End If
        Next i
L671:
                '=======================================================================
                '     end of subroutine
                '=======================================================================
        Exit Sub
        Exit Sub
End Sub

'=======================================================================
'     subroutine boyles_law_compensation
'     purpose: this subsubroutine calculates the reduction in allowable gra
'     with decreasing ambient vpressure during the vdecompression profile
'     on boyle's law considerations.
'=======================================================================

Sub boyles_law_compensation(first_vstop_vdepth As Double, vdeco_vstop_vdepth As Double, vstep_size As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim first_vstop_vdepth as double
' sub parameter : do not dim ! Dim  vdeco_vstop_vdepth as double
' sub parameter : do not dim ! Dim  vstep_size as double
'=======================================================================
'     local variables
'=======================================================================
Dim i As Integer                                                  'loop as integer
Dim next_vstop As Double
Dim ambient_vpressure_first_vstop As Double
Dim ambient_vpressure_next_vstop As Double
Dim amb_press_first_vstop_pascals As Double
Dim amb_press_next_vstop_pascals As Double
Dim a As Double
Dim b As Double
Dim C As Double
Dim low_bound As Double
Dim high_bound As Double
Dim ending_radius As Double
Dim vdeco_gradient_pascals As Double
Dim allow_grad_first_vstop_vhe_pa As Double
Dim radius_first_vstop_he As Double
Dim allow_grad_first_vstop_vn2_pa As Double
Dim radius_first_vstop_n2 As Double
'=======================================================================
'     local arrays
'=======================================================================
Dim radius1_he(16)  As Double
Dim radius2_he(16) As Double
Dim radius1_n2(16)  As Double
Dim radius2_n2(16) As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim surface_tension_gamma As Double
'Dim skin_compression_gammac As Double
'common /block_19/ surface_tension_gamma, skin_compression_gammac
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim barometric_vpressure As Double
'common /block_18/ barometric_vpressure
'Dim units_factor As Double
'common /block_16/ units_factor
''=======================================================================
''     global arrays in named common blocks
''=======================================================================
'Dim allowable_gradient_he(16)  As Double
'Dim allowable_gradient_n2(16) As Double
'common /block_26/ allowable_gradient_he, allowable_gradient_n2
'Dim vdeco_gradient_he(16)  As Double
'Dim vdeco_gradient_n2(16) As Double
'common /block_34/ vdeco_gradient_he, vdeco_gradient_n2
'=======================================================================
'     calculations
'=======================================================================
next_vstop = vdeco_vstop_vdepth - vstep_size
ambient_vpressure_first_vstop = first_vstop_vdepth + barometric_vpressure
ambient_vpressure_next_vstop = next_vstop + barometric_vpressure
amb_press_first_vstop_pascals = (ambient_vpressure_first_vstop / units_factor) * 101325#
amb_press_next_vstop_pascals = (ambient_vpressure_next_vstop / units_factor) * 101325#
For i = 1 To 16
    allow_grad_first_vstop_vhe_pa = (allowable_gradient_he(i) / units_factor) * 101325#
    radius_first_vstop_he = (2# * surface_tension_gamma) / allow_grad_first_vstop_vhe_pa
    radius1_he(i) = radius_first_vstop_he
    a = amb_press_next_vstop_pascals
    b = -2# * surface_tension_gamma
    C = (amb_press_first_vstop_pascals + (2# * surface_tension_gamma) / radius_first_vstop_he) * radius_first_vstop_he * (radius_first_vstop_he * (radius_first_vstop_he))
    low_bound = radius_first_vstop_he
'    f = (amb_press_first_vstop_pascals / amb_press_next_vstop_pascals)
'    If (f > 0#) Then
      high_bound = radius_first_vstop_he * ((amb_press_first_vstop_pascals / amb_press_next_vstop_pascals) ^ (1# / 3#))
'    Else
'      f = 1#
'    End If
    Call radius_root_finder(a, b, C, low_bound, high_bound, ending_radius)
    radius2_he(i) = ending_radius
    vdeco_gradient_pascals = (2# * surface_tension_gamma) / ending_radius
    vdeco_gradient_he(i) = (vdeco_gradient_pascals / 101325#) * units_factor
Next i
For i = 1 To 16
    allow_grad_first_vstop_vn2_pa = (allowable_gradient_n2(i) / units_factor) * 101325#
    radius_first_vstop_n2 = (2# * surface_tension_gamma) / allow_grad_first_vstop_vn2_pa
    radius1_n2(i) = radius_first_vstop_n2
    a = amb_press_next_vstop_pascals
    b = -2# * surface_tension_gamma
    C = (amb_press_first_vstop_pascals + (2# * surface_tension_gamma) / radius_first_vstop_n2) * radius_first_vstop_n2 * (radius_first_vstop_n2 * (radius_first_vstop_n2))
    low_bound = radius_first_vstop_n2
    'f = (amb_press_first_vstop_pascals / amb_press_next_vstop_pascals)
    'If f > 0# Then
    high_bound = radius_first_vstop_n2 * (amb_press_first_vstop_pascals / amb_press_next_vstop_pascals) ^ (1# / 3#)
    Call radius_root_finder(a, b, C, low_bound, high_bound, ending_radius)
    radius2_n2(i) = ending_radius
    vdeco_gradient_pascals = (2# * surface_tension_gamma) / ending_radius
    vdeco_gradient_n2(i) = (vdeco_gradient_pascals / 101325#) * units_factor
Next i
'=======================================================================
'     end of subroutine
'=======================================================================
Exit Sub
Exit Sub
End Sub

'=======================================================================
'     subroutine vdecompression_vstop
'     purpose: this subsubroutine calculates the required time at each
'     vdecompression vstop.
'=======================================================================
Sub vdecompression_vstop(vdeco_vstop_vdepth As Double, vstep_size As Double)
'      implicit none
'=======================================================================
'     arguments
'=======================================================================
' sub parameter : do not dim ! Dim vdeco_vstop_vdepth as double
' sub parameter : do not dim ! Dim  vstep_size as double
'=======================================================================
'     local variables
'=======================================================================
Dim os_command As String * 3
Dim i As Integer                                                  'loop as integer
Dim last_vsegment_vnumber As Integer
'Dim ambient_vpressure As Double
'Dim inspired_vhelium_vpressure As Double
'Dim inspired_vnitrogen_vpressure As Double
Dim last_run_vtime As Double
Dim vdeco_vceiling_vdepth As Double
Dim next_vstop As Double
Dim round_up_operation As Double
Dim temp_vsegment_vtime As Double
Dim time_counter As Double
Dim weighted_allowable_gradient As Double
'Dim haldane_equation                                 'function su as double
'=======================================================================
'     local arrays
'=======================================================================
Dim initial_vhelium_vpressure(16) As Double
Dim initial_vnitrogen_vpressure(16) As Double
'=======================================================================
'     global constants in named common blocks
'=======================================================================
'Dim water_vapor_vpressure As Double
'common /block_8/ water_vapor_vpressure
'Dim constant_vpressure_other_gases As Double
'common /block_17/ constant_vpressure_other_gases
'Dim minimum_vdeco_vstop_vtime As Double
'common /block_21/ minimum_vdeco_vstop_vtime
''=======================================================================
''     global variables in named common blocks
''=======================================================================
'Dim vsegment_vnumber As Integer
'Dim run_vtime As Double
'Dim vsegment_vtime As Double
'common /block_2/ run_vtime, vsegment_vnumber, vsegment_vtime
'Dim ending_ambient_vpressure As Double
'common /block_4/ ending_ambient_vpressure
'Dim vmix_vnumber As Integer
'common /block_9/ vmix_vnumber
'Dim barometric_vpressure As Double
'common /block_18/ barometric_vpressure
''=======================================================================
'''     global arrays in named common blocks
''=======================================================================
'Dim vhelium_vtime_constant(16) As Double
'common /block_1a/ vhelium_vtime_constant
'Dim vnitrogen_vtime_constant(16) As Double
'common /block_1b/ vnitrogen_vtime_constant
'Dim vhelium_vpressure(16)  As Double
'Dim vnitrogen_vpressure(16)                 'bo as double
'common /block_3/ vhelium_vpressure, vnitrogen_vpressure            'an
'Dim vfraction_vhelium(20)  As Double
'Dim vfraction_vnitrogen(20) As Double
'common /block_5/ vfraction_vhelium, vfraction_vnitrogen
'Dim vdeco_gradient_he(16)  As Double
'Dim vdeco_gradient_n2(16) As Double
'common /block_34/ vdeco_gradient_he, vdeco_gradient_n2
''=======================================================================
'     calculations
'=======================================================================
'os_command = "cls"
last_run_vtime = run_vtime
round_up_operation = CDbl(CLng((last_run_vtime / minimum_vdeco_vstop_vtime) + 0.49999)) * minimum_vdeco_vstop_vtime
vsegment_vtime = round_up_operation - run_vtime
run_vtime = round_up_operation
temp_vsegment_vtime = vsegment_vtime
last_vsegment_vnumber = vsegment_vnumber
vsegment_vnumber = last_vsegment_vnumber + 1
ambient_vpressure = vdeco_vstop_vdepth + barometric_vpressure
ending_ambient_vpressure = ambient_vpressure
next_vstop = vdeco_vstop_vdepth - vstep_size
inspired_vhelium_vpressure = (ambient_vpressure - water_vapor_vpressure) * vfraction_vhelium(vmix_vnumber)
inspired_vnitrogen_vpressure = (ambient_vpressure - water_vapor_vpressure) * vfraction_vnitrogen(vmix_vnumber)
Set_Inspired_Inert
'=======================================================================
'     check to make sure that subroutine won't lock up if unable to vdecompr
'     to the next vstop.  if so, write error message and terminate progra
'=======================================================================
For i = 1 To 16
    If ((inspired_vhelium_vpressure + inspired_vnitrogen_vpressure) > 0#) Then
        weighted_allowable_gradient = (vdeco_gradient_he(i) * inspired_vhelium_vpressure + vdeco_gradient_n2(i) * inspired_vnitrogen_vpressure) / (inspired_vhelium_vpressure + inspired_vnitrogen_vpressure)
        If ((inspired_vhelium_vpressure + inspired_vnitrogen_vpressure + constant_vpressure_other_gases - weighted_allowable_gradient) > (next_vstop + barometric_vpressure)) Then
            no_deco_found = 3 ' MsgBox "root not in brackets"
        End If
    End If
Next i
L700:
For i = 1 To 16
    initial_vhelium_vpressure(i) = vhelium_vpressure(i)
    initial_vnitrogen_vpressure(i) = vnitrogen_vpressure(i)
'    f = haldane_equation(initial_vhelium_vpressure(i), inspired_vhelium_vpressure, vhelium_vtime_constant(i), vsegment_vtime)
'    If i = 1 Then
'      f = f
'    End If
    vhelium_vpressure(i) = haldane_equation(initial_vhelium_vpressure(i), inspired_vhelium_vpressure, vhelium_vtime_constant(i), vsegment_vtime)
    vnitrogen_vpressure(i) = haldane_equation(initial_vnitrogen_vpressure(i), inspired_vnitrogen_vpressure, vnitrogen_vtime_constant(i), vsegment_vtime)
Next i
L720:
    Call calc_vdeco_vceiling(vdeco_vceiling_vdepth)
    If temp_vsegment_vtime >= 10000 Then
      i = 1
    End If
    If (vdeco_vceiling_vdepth > next_vstop) And temp_vsegment_vtime < 10000 Then
        vsegment_vtime = minimum_vdeco_vstop_vtime
        time_counter = temp_vsegment_vtime
        temp_vsegment_vtime = time_counter + minimum_vdeco_vstop_vtime
        last_run_vtime = run_vtime
        run_vtime = last_run_vtime + minimum_vdeco_vstop_vtime
        GoTo L700
    End If
    vsegment_vtime = temp_vsegment_vtime
Exit Sub
    '=======================================================================
    ' '     format statements - error messages
    '=======================================================================
    ' 905   format ('0error! off-gassing gradient is too small to vdecompress' 1x,'at the',f6.1,1x,'vstop')
    ' 906   format ('0reduce step size or increase voxygen vfraction')
    ' 907   format (' ')
    '=======================================================================
    '     end of subroutine
    '=======================================================================
Exit Sub
    'Close #1
End Sub

    '=======================================================================
    '     subroutine calc_vdeco_vceiling
    '     purpose: this subsubroutine calculates the vdeco vceiling (the safe asc
    '     vdepth) in each compartment, based on the allowable "vdeco gradients
    '     computed in the boyle's law compensation subroutine, and then find
    '     deepest vdeco vceiling across all compartments.  this deepest value
    '     (vdeco vceiling vdepth) is then used by the vdecompression vstop subrou
    '     to determine the actual vdeco schedule.
    '=======================================================================
Sub calc_vdeco_vceiling(vdeco_vceiling_vdepth As Double)
    '      implicit none
    '=======================================================================
    '     arguments
    '=======================================================================
    ' sub parameter : do not dim ! Dim vdeco_vceiling_vdepth as double
    '=======================================================================
    '     local variables
    '=======================================================================
    Dim i As Integer                                                  'loop as integer
    Dim gas_loading As Double
    Dim weighted_allowable_gradient As Double
    Dim tolerated_ambient_vpressure As Double
    '=======================================================================
    '     local arrays
    '=======================================================================
    Dim compartment_vdeco_vceiling(16) As Double
    '=======================================================================
    '     global constants in named common blocks
    '=======================================================================
    'Dim constant_vpressure_other_gases As Double
    'common /block_17/ constant_vpressure_other_gases
    ''=======================================================================
    ''     global variables in named common blocks
    ''=======================================================================
    'Dim barometric_vpressure As Double
    'common /block_18/ barometric_vpressure
    ''=======================================================================
    ''     global arrays in named common blocks
    ''=======================================================================
    'Dim vhelium_vpressure(16)  As Double
    'Dim vnitrogen_vpressure(16) As Double
    'common /block_3/ vhelium_vpressure, vnitrogen_vpressure
    'Dim vdeco_gradient_he(16)  As Double
    'Dim vdeco_gradient_n2(16) As Double
    'common /block_34/ vdeco_gradient_he, vdeco_gradient_n2
    ''=======================================================================
    ''     calculations
    ''     since there are two sets of vdeco gradients being tracked, one for
    '     vhelium and one for vnitrogen, a "weighted allowable gradient" must
    '     computed each time based on the proportions of vhelium and vnitrogen
    '     each compartment.  this proportioning follows the methodology of
    '     buhlmann/keller.  if there is no vhelium and vnitrogen in the compar
    '     such as after extended periods of voxygen breathing, then the minim
    '     across both gases will be used.  it is important to note that if a
    '     compartment is empty of vhelium and vnitrogen, then the weighted all
    '     gradient formula cannot be used since it will result in division b
    '=======================================================================
    
    For i = 1 To 16
        gas_loading = vhelium_vpressure(i) + vnitrogen_vpressure(i)
        If (gas_loading > 0#) Then
            weighted_allowable_gradient = (vdeco_gradient_he(i) * vhelium_vpressure(i) + vdeco_gradient_n2(i) * vnitrogen_vpressure(i)) / (vhelium_vpressure(i) + vnitrogen_vpressure(i))
            tolerated_ambient_vpressure = (gas_loading + constant_vpressure_other_gases) - weighted_allowable_gradient
            Else
            weighted_allowable_gradient = Min(vdeco_gradient_he(i), vdeco_gradient_n2(i))
            tolerated_ambient_vpressure = constant_vpressure_other_gases - weighted_allowable_gradient
        End If
        '=======================================================================
        '     the tolerated ambient vpressure cannot be less than zero absolute,
        '     the vacuum of outer space!
        '=======================================================================

        nfrac = vnitrogen_vpressure(i) / (vnitrogen_vpressure(i) + vhelium_vpressure(i))
        hefrac = vhelium_vpressure(i) / (vnitrogen_vpressure(i) + vhelium_vpressure(i))
        If (hefrac = 0) Then hefrac = 0.005
        If (nfrac = 0) Then nfrac = 0.005
        atotal = (nfrac * an2(i - 1) + hefrac * ahe(i - 1)) / (nfrac + hefrac)
        btotal = (nfrac * bn2(i - 1) + hefrac * bhe(i - 1)) / (nfrac + hefrac)
        
        atotal = atotal * 0.9 * (1# - (CDbl(safetytext.Text) / 300#))
        btotal = btotal * 1.1 * (1# + CDbl(safetytext.Text) / 300#)
        
        buhlptoln2(i - 1) = btotal * (vnitrogen_vpressure(i) + vhelium_vpressure(i) - atotal * units_factor)
        
        If buhlptoln2(i - 1) > tolerated_ambient_vpressure Then
          If buhl_mode > 0 Then tolerated_ambient_vpressure = buhlptoln2(i - 1)
        End If
        If buhl_mode = 2 Then tolerated_ambient_vpressure = buhlptoln2(i - 1)
        
        If (tolerated_ambient_vpressure < 0#) Then
            tolerated_ambient_vpressure = 0#
        End If
        '=======================================================================
        '     the vdeco vceiling vdepth is computed in a loop after all of the indi
        '     compartment vdeco vceilings have been calculated.  it is important t
        '     vdeco vceiling vdepth (max vdeco vceiling across all compartments) only
        '     extracted from the compartment values and not be compared against
        '     initialization value.  for example, if max(vdeco_vceiling_vdepth . .)
        '     compared against zero, this could cause a subroutine lockup because s
        '     the vdeco vceiling vdepth needs to be negative (but not less than abs
        '     zero) in order to vdecompress to the last vstop at zero vdepth.
        '=======================================================================
        compartment_vdeco_vceiling(i) = tolerated_ambient_vpressure - barometric_vpressure
    Next i
    vdeco_vceiling_vdepth = compartment_vdeco_vceiling(1)
    For i = 2 To 16
        vdeco_vceiling_vdepth = Max(vdeco_vceiling_vdepth, compartment_vdeco_vceiling(i))
    Next i
    '=======================================================================
    '     end of subroutine
    '=======================================================================
    Exit Sub
    Exit Sub
    '=======================================================================
    '     subroutine gas_loadings_surface_interval
    '     purpose: this subsubroutine calculates the gas loading (off-gassing)
    '     a surface interval.
    '=======================================================================
End Sub
Sub gas_loadings_surface_interCint(surface_interval_vtime As Double)
    '      implicit none
    '=======================================================================
    '     arguments
    '=======================================================================
    ' sub parameter : do not dim ! Dim surface_interval_vtime as double
    '=======================================================================
    '     local variables
    '=======================================================================
    Dim i As Integer                                                  'loop as integer
    'Dim inspired_vhelium_vpressure As Double
    'Dim inspired_vnitrogen_vpressure As Double
    Dim initial_vhelium_vpressure As Double
    Dim initial_vnitrogen_vpressure As Double
    'Dim haldane_equation                                 'function su as double
    '=======================================================================
    '     global constants in named common blocks
    '=======================================================================
    'Dim water_vapor_vpressure As Double
    'common /block_8/ water_vapor_vpressure
    ''=======================================================================
    ''     global variables in named common blocks
    ''=======================================================================
    'Dim barometric_vpressure As Double
    'common /block_18/ barometric_vpressure
    ''=======================================================================
    ''     global arrays in named common blocks
    ''=======================================================================
    'Dim vhelium_vtime_constant(16) As Double
    'common /block_1a/ vhelium_vtime_constant
    'Dim vnitrogen_vtime_constant(16) As Double
    'common /block_1b/ vnitrogen_vtime_constant
    'Dim vhelium_vpressure(16)  As Double
    'Dim vnitrogen_vpressure(16)                 'bo as double
    'common /block_3/ vhelium_vpressure, vnitrogen_vpressure            'an
    '=======================================================================
    '     calculations
    '=======================================================================
    inspired_vhelium_vpressure = 0#
    inspired_vnitrogen_vpressure = (barometric_vpressure - water_vapor_vpressure) * 0.79
    For i = 1 To 16
        initial_vhelium_vpressure = vhelium_vpressure(i)
        initial_vnitrogen_vpressure = vnitrogen_vpressure(i)
        vhelium_vpressure(i) = haldane_equation(initial_vhelium_vpressure, inspired_vhelium_vpressure, vhelium_vtime_constant(i), surface_interval_vtime)
        vnitrogen_vpressure(i) = haldane_equation(initial_vnitrogen_vpressure, inspired_vnitrogen_vpressure, vnitrogen_vtime_constant(i), surface_interval_vtime)
    Next i
    '=======================================================================
    '     end of subroutine
    '=======================================================================
    '=======================================================================
    '     subroutine vpm_repetitive_valgorithm
    '     purpose: this subsubroutine implements the vpm repetitive valgorithm t
    '     envisioned by professor david e. yount only vmonths before his pass
    '=======================================================================
End Sub
Sub vpm_repetitive_valgorithm(surface_interval_vtime As Double)
    '      implicit none
    '=======================================================================
    '     arguments
    '=======================================================================
    ' sub parameter : do not dim ! Dim surface_interval_vtime as double
    '=======================================================================
    '     local variables
    '=======================================================================
    Dim i As Integer                                                  'loop as integer
    Dim max_actual_gradient_pascals As Double
    Dim adj_crush_vpressure_vhe_pascals As Double
    Dim adj_crush_vpressure_vn2_pascals As Double
    Dim initial_allowable_grad_vhe_pa As Double
    Dim initial_allowable_grad_vn2_pa As Double
    Dim new_critical_radius_he As Double
    Dim new_critical_radius_n2 As Double
    '=======================================================================
    '     global constants in named common blocks
    '=======================================================================
    'Dim surface_tension_gamma As Double
    'Dim skin_compression_gammac As Double
    'common /block_19/ surface_tension_gamma, skin_compression_gammac
    'Dim regeneration_vtime_constant As Double
    'common /block_22/ regeneration_vtime_constant
    ''=======================================================================
    ''     global variables in named common blocks
    ''=======================================================================
    'Dim units_factor As Double
    'common /block_16/ units_factor
    ''=======================================================================
    ''     global arrays in named common blocks
    ''=======================================================================
    'Dim initial_critical_radius_he(16) As Double
    'Dim initial_critical_radius_n2(16) As Double
    'common /block_6/ initial_critical_radius_he,                                 initial_critical_radius_n2
    'Dim adjusted_critical_radius_he(16) As Double
    'Dim adjusted_critical_radius_n2(16) As Double
    'common /block_7/ adjusted_critical_radius_he,                                      adjusted_critical_radius_n2
    'Dim max_actual_gradient(16) As Double
    'common /block_12/ max_actual_gradient
    'Dim adjusted_crushing_vpressure_he(16) As Double
    'Dim adjusted_crushing_vpressure_n2(16) As Double
    'common /block_25/ adjusted_crushing_vpressure_he,                                    adjusted_crushing_vpressure_n2
    'Dim initial_allowable_gradient_he(16) As Double
    'Dim initial_allowable_gradient_n2(16) As Double
    'common /block_27/                                                     initial_allowable_gradient_he, initial_allowable_gradient_n2
    ''=======================================================================
    '     calculations
    '=======================================================================
    For i = 1 To 16
        max_actual_gradient_pascals = (max_actual_gradient(i) / units_factor) * 101325#
        adj_crush_vpressure_vhe_pascals = (adjusted_crushing_vpressure_he(i) / units_factor) * 101325#
        adj_crush_vpressure_vn2_pascals = (adjusted_crushing_vpressure_n2(i) / units_factor) * 101325#
        initial_allowable_grad_vhe_pa = (initial_allowable_gradient_he(i) / units_factor) * 101325#
        initial_allowable_grad_vn2_pa = (initial_allowable_gradient_n2(i) / units_factor) * 101325#
        If (max_actual_gradient(i) > initial_allowable_gradient_n2(i)) Then
            new_critical_radius_n2 = ((2# * surface_tension_gamma * (skin_compression_gammac - surface_tension_gamma))) / (max_actual_gradient_pascals * skin_compression_gammac - surface_tension_gamma * adj_crush_vpressure_vn2_pascals)
            adjusted_critical_radius_n2(i) = initial_critical_radius_n2(i) + (initial_critical_radius_n2(i) - new_critical_radius_n2) * Exp(-surface_interval_vtime / regeneration_vtime_constant)
            Else
            adjusted_critical_radius_n2(i) = initial_critical_radius_n2(i)
        End If
        If (max_actual_gradient(i) > initial_allowable_gradient_he(i)) Then
            new_critical_radius_he = ((2# * surface_tension_gamma * (skin_compression_gammac - surface_tension_gamma))) / (max_actual_gradient_pascals * skin_compression_gammac - surface_tension_gamma * adj_crush_vpressure_vhe_pascals)
            adjusted_critical_radius_he(i) = initial_critical_radius_he(i) + (initial_critical_radius_he(i) - new_critical_radius_he) * Exp(-surface_interval_vtime / regeneration_vtime_constant)
            Else
            adjusted_critical_radius_he(i) = initial_critical_radius_he(i)
        End If
    Next i
    '=======================================================================
    '     end of subroutine
    '=======================================================================
End Sub

    '=======================================================================
    '     subroutine calc_barometric_vpressure
    '     purpose: this sub calculates barometric vpressure at valtitude based
    '     publication "u.s. standard atmosphere, 1976", u.s. government prin
    '     office, washington, d.c. the source for this code is a fortran 90
    '     written by ralph l. carmichael (retired nasa researcher) and endor
    '     the national geophysical data center of the national oceanic and
    '     atmospheric administration.  it is available for download free fro
    '     public domain aeronautical software at:  http://www.pdas.com/atmos
    '=======================================================================

Sub calc_barometric_vpressure(valtitude As Double)
    '      implicit none
    '=======================================================================
    '     arguments
    '=======================================================================
    ' sub parameter : do not dim ! Dim valtitude as double
    '=======================================================================
    '     local constants
    '=======================================================================
    Dim radius_of_earth As Double
    Dim acceleration_of_gravity As Double
    Dim molecular_weight_of_air As Double
    Dim gas_constant_r As Double
    Dim temp_at_sea_level As Double
    Dim temp_gradient As Double
    Dim vpressure_at_sea_level_fsw As Double
    Dim vpressure_at_sea_level_msw As Double
    '=======================================================================
    '     local variables
    '=======================================================================
    Dim vpressure_at_sea_level As Double
    Dim gmr_factor As Double
    Dim valtitude_feet As Double
    Dim valtitude_meters As Double
    Dim valtitude_kilometers As Double
    Dim geopotential_valtitude As Double
    Dim temp_at_geopotential_valtitude As Double
    '=======================================================================
    '     global variables in named common blocks
    '=======================================================================
    'Dim units_equal_fsw As Boolean
    'Dim units_equal_msw As Boolean
    ''common /block_15/ units_equal_fsw, units_equal_msw
    'Dim barometric_vpressure As Double
    ''common /block_18/ barometric_vpressure
    '=======================================================================
    '     calculations
    '=======================================================================
    Exit Sub
    radius_of_earth = 6369#                                        'ki
    acceleration_of_gravity = 9.80665                         'meters/
    molecular_weight_of_air = 28.9644
    gas_constant_r = 8.31432                            'joules/mol*de
    temp_at_sea_level = 288.15                                 'degree
    vpressure_at_sea_level_fsw = 33#       'feet of seawater based on 1
    'at sea level (standard atm
    vpressure_at_sea_level_msw = 10#     'meters of seawater based on 1
    'at sea level (european
    temp_gradient = -6.5                       'change in temp deg kel
    'change in geopotential a
    'valid for first layer of at
    'up to 11 kilometers or 36,
    gmr_factor = acceleration_of_gravity * molecular_weight_of_air / gas_constant_r
    If (units_equal_fsw) Then
        valtitude_feet = valtitude
        valtitude_kilometers = valtitude_feet / 3280.839895
        vpressure_at_sea_level = vpressure_at_sea_level_fsw
    End If
    If (units_equal_msw) Then
        valtitude_meters = valtitude
        valtitude_kilometers = valtitude_meters / 1000#
        vpressure_at_sea_level = vpressure_at_sea_level_msw
    End If
    geopotential_valtitude = (valtitude_kilometers * radius_of_earth) / (valtitude_kilometers + radius_of_earth)
    temp_at_geopotential_valtitude = temp_at_sea_level + temp_gradient * geopotential_valtitude
    barometric_vpressure = vpressure_at_sea_level * Exp(Log(temp_at_sea_level / temp_at_geopotential_valtitude) * gmr_factor / temp_gradient)
    '=======================================================================
    '     end of subroutine
    '=======================================================================
End Sub
    
    
    '=======================================================================
    '     subroutine vpm_valtitude_dive_valgorithm
    '     purpose:  this subsubroutine updates gas loadings and adjusts critica
    '     (as required) based on whether or not diver is acclimatized at alt
    '     makes an ascent to valtitude before the dive.
    '=======================================================================
Sub vpm_valtitude_dive_valgorithm()
    '      implicit none
    '=======================================================================
    '     local variables
    '=======================================================================
    Dim diver_acclimatized_at_valtitude As String * 3
    Dim os_command As String * 3
    Dim i As Integer                                                  'loop as integer
    Dim diver_acclimatized As Boolean
    Dim valtitude_of_dive As Double
    Dim starting_acclimatized_valtitude As Double
    Dim ascent_to_valtitude_hours As Double
    Dim hours_at_valtitude_before_dive As Double
    Dim ascent_to_valtitude_vtime As Double
    Dim time_at_valtitude_before_dive As Double
'    Dim starting_ambient_vpressure As Double
    Dim ending_ambient_vpressure As Double
    'Dim initial_inspired_vn2_vpressure As Double
    Dim rate As Double
    Dim vnitrogen_rate As Double
    'Dim inspired_vnitrogen_vpressure As Double
    Dim initial_vnitrogen_vpressure As Double
    Dim compartment_gradient As Double
    Dim compartment_gradient_pascals As Double
    Dim gradient_vhe_bubble_formation As Double
    Dim gradient_vn2_bubble_formation As Double
    Dim new_critical_radius_he As Double
    Dim new_critical_radius_n2 As Double
    Dim ending_radius_he As Double
    Dim ending_radius_n2 As Double
    Dim regenerated_radius_he As Double
    Dim regenerated_radius_n2 As Double
    'Dim haldane_equation                                 'function su as double
'    Dim schreiner_equation As Double
    '=======================================================================
    '     global constants in named common blocks
    '=======================================================================
    'Dim water_vapor_vpressure As Double
    'common /block_8/ water_vapor_vpressure
    'Dim constant_vpressure_other_gases As Double
    'common /block_17/ constant_vpressure_other_gases
    'Dim surface_tension_gamma As Double
    'Dim skin_compression_gammac As Double
    'common /block_19/ surface_tension_gamma, skin_compression_gammac
    'Dim regeneration_vtime_constant As Double
    'common /block_22/ regeneration_vtime_constant
    ''=======================================================================
    ''     global variables in named common blocks
    ''=======================================================================
    'Dim units_equal_fsw As Boolean
    'Dim units_equal_msw As Boolean
    'common /block_15/ units_equal_fsw, units_equal_msw
    'Dim units_factor As Double
    'common /block_16/ units_factor
    'Dim barometric_vpressure As Double
    'common /block_18/ barometric_vpressure
    ''=======================================================================
    ''     global arrays in named common blocks
    ''=======================================================================
    'Dim vnitrogen_vtime_constant(16) As Double
    'common /block_1b/ vnitrogen_vtime_constant
    'Dim vhelium_vpressure(16)  As Double
    'Dim vnitrogen_vpressure(16)                 'bo as double
    'common /block_3/ vhelium_vpressure, vnitrogen_vpressure            'an
    'Dim initial_critical_radius_he(16)                            'bo as double
    'Dim initial_critical_radius_n2(16)                            'an as double
    'common /block_6/ initial_critical_radius_he,                                 initial_critical_radius_n2
    'Dim adjusted_critical_radius_he(16) As Double
    'Dim adjusted_critical_radius_n2(16) As Double
    'common /block_7/ adjusted_critical_radius_he,                                      adjusted_critical_radius_n2
    ''=======================================================================
    '     namelist for subroutine settings (read in from ascii text file)
    '=======================================================================
    '=======================================================================
    '     calculations
    '=======================================================================
    os_command = "cls"
    '       open (unit = 12, file = 'valtitude.set', status = 'unknown',                access = 'sequential', form = 'formatted')
    valtitude_of_dive = 0
    diver_acclimatized_at_valtitude = "yes"
    starting_acclimatized_valtitude = 0
    ascent_to_valtitude_hours = 1
    hours_at_valtitude_before_dive = 30
    If ((units_equal_fsw) And (valtitude_of_dive > 30000#)) Then
        no_deco_found = 3 ' MsgBox "root not in brackets"
    End If
    If ((units_equal_msw) And (valtitude_of_dive > 9144#)) Then
        no_deco_found = 3 ' MsgBox "root not in brackets"
    End If
    If (InStr(1, diver_acclimatized_at_valtitude, "yes")) Then ' Or (diver_acclimatized_at_valtitude = "yes")) Then
        diver_acclimatized = (True)
    Else 'If ((diver_acclimatized_at_valtitude = "no") Or (diver_acclimatized_at_valtitude = "no")) Then
        diver_acclimatized = (False)
    End If
'    If ((diver_acclimatized_at_valtitude = "yes") Or (diver_acclimatized_at_valtitude = "yes")) Then
'        diver_acclimatized = (True)
'        ElseIf ((diver_acclimatized_at_valtitude = "no") Or (diver_acclimatized_at_valtitude = "no")) Then
'        diver_acclimatized = (False)
'        Else
'        no_deco_found=3 ' MsgBox "root not in brackets"
'    End If
    ascent_to_valtitude_vtime = ascent_to_valtitude_hours * 60#
    time_at_valtitude_before_dive = hours_at_valtitude_before_dive * 60#
    If (diver_acclimatized) Then
        Call calc_barometric_vpressure(valtitude_of_dive)            'su
        'Write #1, valtitude_of_dive, barometric_vpressure
        For i = 1 To 16
            adjusted_critical_radius_n2(i) = initial_critical_radius_n2(i)
            adjusted_critical_radius_he(i) = initial_critical_radius_he(i)
            vhelium_vpressure(i) = 0#
            vnitrogen_vpressure(i) = (barometric_vpressure - water_vapor_vpressure) * 0.79
        Next i
    Else
        If ((starting_acclimatized_valtitude >= valtitude_of_dive) Or (starting_acclimatized_valtitude < 0#)) Then
            no_deco_found = 3 ' MsgBox "root not in brackets"
        End If
        Call calc_barometric_vpressure(starting_acclimatized_valtitude)
        starting_ambient_vpressure = barometric_vpressure
        For i = 1 To 16
            vhelium_vpressure(i) = 0#
            vnitrogen_vpressure(i) = (barometric_vpressure - water_vapor_vpressure) * 0.79
        Next i
        Call calc_barometric_vpressure(valtitude_of_dive)            'su
        'Write #1, valtitude_of_dive, barometric_vpressure
        ending_ambient_vpressure = barometric_vpressure
        initial_inspired_vn2_vpressure = (starting_ambient_vpressure - water_vapor_vpressure) * 0.79
        rate = (ending_ambient_vpressure - starting_ambient_vpressure) / ascent_to_valtitude_vtime
        vnitrogen_rate = rate * 0.79
        For i = 1 To 16
            initial_vnitrogen_vpressure = vnitrogen_vpressure(i)
            vnitrogen_vpressure(i) = schreiner_equation(initial_inspired_vn2_vpressure, vnitrogen_rate, ascent_to_valtitude_vtime, vnitrogen_vtime_constant(i), initial_vnitrogen_vpressure)
            compartment_gradient = (vnitrogen_vpressure(i) + constant_vpressure_other_gases) - ending_ambient_vpressure
            compartment_gradient_pascals = (compartment_gradient / units_factor) * 101325#
            gradient_vhe_bubble_formation = ((2# * surface_tension_gamma * (skin_compression_gammac - surface_tension_gamma)) / (initial_critical_radius_he(i) * skin_compression_gammac))
            If (compartment_gradient_pascals > gradient_vhe_bubble_formation) Then
              new_critical_radius_he = ((2# * surface_tension_gamma * (skin_compression_gammac - surface_tension_gamma))) / (compartment_gradient_pascals * skin_compression_gammac)
              adjusted_critical_radius_he(i) = initial_critical_radius_he(i) + (initial_critical_radius_he(i) - new_critical_radius_he) * Exp(-time_at_valtitude_before_dive / regeneration_vtime_constant)
              initial_critical_radius_he(i) = adjusted_critical_radius_he(i)
            Else
              ending_radius_he = 1# / (compartment_gradient_pascals / (2# * (surface_tension_gamma - skin_compression_gammac)) + 1# / initial_critical_radius_he(i))
              regenerated_radius_he = initial_critical_radius_he(i) + (ending_radius_he - initial_critical_radius_he(i)) * Exp(-time_at_valtitude_before_dive / regeneration_vtime_constant)
              initial_critical_radius_he(i) = regenerated_radius_he
              adjusted_critical_radius_he(i) = initial_critical_radius_he(i)
            End If
            gradient_vn2_bubble_formation = ((2# * surface_tension_gamma * (skin_compression_gammac - surface_tension_gamma)) / (initial_critical_radius_n2(i) * skin_compression_gammac))
            If (compartment_gradient_pascals > gradient_vn2_bubble_formation) Then
              new_critical_radius_n2 = ((2# * surface_tension_gamma * (skin_compression_gammac - surface_tension_gamma))) / (compartment_gradient_pascals * skin_compression_gammac)
              adjusted_critical_radius_n2(i) = initial_critical_radius_n2(i) + (initial_critical_radius_n2(i) - new_critical_radius_n2) * Exp(-time_at_valtitude_before_dive / regeneration_vtime_constant)
              initial_critical_radius_n2(i) = adjusted_critical_radius_n2(i)
            Else
              ending_radius_n2 = 1# / (compartment_gradient_pascals / (2# * (surface_tension_gamma - skin_compression_gammac)) + 1# / initial_critical_radius_n2(i))
              regenerated_radius_n2 = initial_critical_radius_n2(i) + (ending_radius_n2 - initial_critical_radius_n2(i)) * Exp(-time_at_valtitude_before_dive / regeneration_vtime_constant)
              initial_critical_radius_n2(i) = regenerated_radius_n2
              adjusted_critical_radius_n2(i) = initial_critical_radius_n2(i)
            End If
        Next i
        inspired_vnitrogen_vpressure = (barometric_vpressure - water_vapor_vpressure) * 0.79
        For i = 1 To 16
            initial_vnitrogen_vpressure = vnitrogen_vpressure(i)
            vnitrogen_vpressure(i) = haldane_equation(initial_vnitrogen_vpressure, inspired_vnitrogen_vpressure, vnitrogen_vtime_constant(i), time_at_valtitude_before_dive)
        Next i
    End If
''close (unit = 12, status = "keep")
''Close #1: Exit Sub
'=======================================================================
' '     format statements - subroutine output
'=======================================================================
' 802   format ('0valtitude = ',1x,f7.1,4x,'barometric vpressure = ',       f6.3)
'=======================================================================
' '     format statements - error messages
'=======================================================================
' 900   format ('0error! valtitude of dive higher than mount everest')
' 901   format (' ')
' 902   format ('0error! diver acclimatized at valtitude',                 1x,'must be yes or no')
' 903   format ('0error! starting acclimatized valtitude must be less',    1x,'than valtitude of dive')
' 904   format (' and greater than or equal to zero')
'=======================================================================
'     end of subroutine
'=======================================================================
End Sub

'=======================================================================
'     subroutine clock
' '     purpose:  this subsubroutine retrieves clock information from the mic
'     operating system so that date and time stamp can be included on pr
'     output.
'=======================================================================

'Sub clock(vyear, vmonth, vday, clock_hour, vminute, m)
''      implicit none
''=======================================================================
''     arguments
''=======================================================================
'' sub parameter : do not dim ! Dim m as string * 1
'' sub parameter : do not dim ! Dim vmonth as integer
'' sub parameter : do not dim ! Dim  vday as integer
'' sub parameter : do not dim ! Dim  vyear as integer
'' sub parameter : do not dim ! Dim vminute as integer
'' sub parameter : do not dim ! Dim  clock_hour as integer
''=======================================================================
''     local variables
''=======================================================================
'Dim hour As Integer
'Dim second As Integer
'Dim hundredth As Integer
''=======================================================================
''     calculations
''=======================================================================
''Call getdat(vyear, vmonth, vday)                          'microsoft
''Call gettim(hour, vminute, second, hundredth)                  'sub
'If (hour > 12) Then
'    clock_hour = hour - 12
'    m = "p"
'    Else
'    clock_hour = hour
'    m = "a"
'End If
''=======================================================================
''     end of subroutine
''=======================================================================
'Exit Sub
'Exit Sub
'End Sub

Private Function Max(m1 As Double, m2 As Double) As Double
  If m1 > m2 Then Max = m1 Else Max = m2
  
End Function

Private Function Min(m1 As Double, m2 As Double) As Double
  If m1 < m2 Then Min = m1 Else Min = m2
  
End Function

Private Sub t1print(t1 As String)
  'Text10.Text = Text10.Text + "  " + CStr(t1)
  
End Sub

Private Sub t1print8(t1 As String)
Dim S As String
'MsgBox CStr(t1)
  S = CStr(t1)
  Do While (Len(S) < 6)
    S = " " + S
  Loop
  If (Len(S) > 6) Then
    S = Left(S, 6)
  End If
  '  Text10.Text = Text10.Text + S + "  "
End Sub

Private Sub t1print8dbl(t1 As Double)
Dim S As String
 'MsgBox CStr(t1)
  S = CStr(t1)
  If InStr(1, S, ".", vbTextCompare) < 1 Then
  S = S + ".0"
  End If
  Do While (Len(S) < 6)
    S = " " + S
  Loop
  If (Len(S) > 6) Then
    S = Left(S, 6)
  End If
  
  'Text10.Text = Text10.Text + S + "  "
End Sub

Private Function vimportdb_data() As Integer
Dim i As Integer
Dim j As Integer
Dim diveplan_num As String
cns_current = 0
otu_current = 0
 deco_update = 1
 deco_grid_display_last = -1
 For i = 0 To 9
   decoresultgrid(i).Rows = 0
   decoresultgridlite(i).Rows = 0
 Next i
 Number_Dives = MSFlexGrid1.Rows - 1
 If Number_Dives = 0 Then
   SSTab1.Visible = False
   Screen.MousePointer = 0
   Exit Function
 End If
 SSTab1.Visible = True
 For j = 1 To Number_Dives
  MSFlexGrid1.Row = j
  MSFlexGrid1.Col = 1
  diveplan_num = MSFlexGrid1.Text
  MSFlexGrid1.Col = 2
  surface_interval_vtime = CDbl(MSFlexGrid1.Text) * 60
  
  SQL = "SELECT * FROM seqdpprofile"
  SQL = SQL & " where dpprofileid = '" & diveplan_num & "' "
  SQL = SQL & " order by dpnumseq "
  Set RS = DB.OpenRecordset(SQL)
  If RS.EOF Then
    MsgBox "Add Profile Plan Points before calculating decompression"
    vimportdb_data = 99
    Exit Function
  End If
   
  laststop_index = 1
   
   SQL = "select * FROM seqdpmain "
   SQL = SQL & " where diveplanid = '" & diveplan_num & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
     If RS3("divecategories") = "2" Then laststop_index = 2
     If RS3("divecategories") = "3" Then laststop_index = 3
     RS3.MoveNext
   Wend
  
  vimportdb_data = 0
Screen.MousePointer = 11
  
  RS.MoveFirst
  i = 1
  While RS.EOF = False
'     MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
'     K = MSFlexGrid3.Rows
'     MSFlexGrid3.Row = K - 1
'     MSFlexGrid3.Col = 0
'     MSFlexGrid3.Text = RS("dpnumseq")
'     MSFlexGrid3.Col = 1
     Plan_Depth(i) = CDbl(RS("depth"))
'     MSFlexGrid3.Col = 2
     Plan_Time(i) = RS("duration")
'     MSFlexGrid3.Col = 3
     Plan_o2(i) = RS("dpo2")
'     MSFlexGrid3.Col = 4
     Plan_he(i) = RS("dphe")
'     MSFlexGrid3.Col = 5
     If InStr(RS("dpcircuit"), "Closed") Then
       Plan_OpenClosed(i) = 1
     Else
       Plan_OpenClosed(i) = 0
     End If
'     MSFlexGrid3.Col = 6
     Plan_PPo2(i) = RS("po2") 'MSFlexGrid3.Text = RS("po2")
     Plan_GasID(i) = CInt(Right(CStr(RS("gasid")), 1)) + 1
     RS.MoveNext
     i = i + 1
   Wend
  i = i - 1
  If i > 0 Then
    Number_of_planpoints = i
  Else
    MsgBox "Add Profile Plan Points before calculating decompression"
  End If

'  SQL = "SELECT * FROM dpmaingaslist "
'  SQL = SQL & " where dpmainid = '" & tempserialno & "' "
'  SQL = SQL & " order by dpgasid "
'  Set RS = DB.OpenRecordset(SQL)
  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & " where dpmainid = '" & diveplan_num & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  Plan_Gas_list_numgasdeco = 1 ' make first gas the last bottom gas when main deco calc done. the extra deco gases are then added below
  For i = 0 To 9
     Plan_Gas_list_n2(i + 1) = CDbl(RS("dpgasnitrogen")) / 100#
     Plan_Gas_list_he(i + 1) = CDbl(RS("dpgashelium")) / 100#
     Plan_Gas_list_mod(i + 1) = CDbl(RS("dpgasmaxopdepth")) / 10#
     Plan_Gas_list_used(i + 1) = CInt(Left(CStr(RS("dpgasused")), 1))
     Plan_Gas_list_setpoint(i + 1) = CDbl(RS("dpgaspo2setpoint"))
     If Plan_Gas_list_used(i + 1) > 3 Then
       Plan_Gas_list_numgasdeco = Plan_Gas_list_numgasdeco + 1
       Plan_Gas_list_deco(Plan_Gas_list_numgasdeco) = i + 1
     End If
     RS.MoveNext
  Next i
  If j < Number_Dives Then repetitive_dive_flag = 1 Else repetitive_dive_flag = 0
  If j = 1 Then repetitive_dive_flag = -1 'initialise data
  barometric_vpressure = CDbl(atmtext.Text) / 100# 'this gives a value of 10 for sea level..
  Sequence_deco
  ' Decoresult(j - 1).Text = Text10.Text
  T = j
  If systemversion = "Lite" Then
     For p = 1 To MSFlexGrid4lite.Rows - 1
        MSFlexGrid4lite.Row = p
        MSFlexGrid4lite.Col = 0
        MSFlexGrid4lite.Text = p
        
     Next p
  End If
  
  duplicategrid
  If no_deco_found > 0 Then
    decoresultgrid(j - 1).Rows = 0
    decoresultgridlite(j - 1).Rows = 0
    Picture1(j - 1).Cls
    MsgBox "Bad depth points!! Re-enter sensible values!"
    Screen.MousePointer = 0
    Exit Function
  End If
 Next j
Screen.MousePointer = 0
End Function



Private Sub display_deco_text()
    
    Dim j As Integer

   If rowindentified = "" Then
     Exit Sub
   Else
     If rowindentified = 0 Then Exit Sub
   End If
  For j = 0 To 9
    If systemversion = "Pro" Then
       decoresultgrid(j).Enabled = True
    Else
       decoresultgridlite(j).Enabled = True
    End If
    If rowindentified = j Then
      Frame3.Caption = "Decompression Result for " & "Dive " & CStr(j) & " - " & txtplanno.Text

  '    changed '12:28pm
  
      If systemversion = "Pro" Then
         decoresultgrid(j - 1).Visible = True
         decoresultgridlite(j - 1).Visible = False
      Else
         decoresultgridlite(j - 1).Visible = True
         decoresultgrid(j - 1).Visible = False
      End If
      SSTab1.TabVisible(j - 1) = True
      SSTab1.TabEnabled(j - 1) = True
      SSTab1.Tab = j - 1
      SSTab1.Caption = "Dive " + CStr(j)
      Label14(j - 1).Caption = "Dive Plan: " + txtplanno.Text + vbCrLf + "Interval to next dive: " + txtinterval.Text + " hours" + vbCrLf + "Deco Algorithm: " + mnuVPMB(buhl_mode).Caption
      display_deco_graph (j - 1)
     ' decoresultgrid(j - 2).Visible = True
    End If
  Next j
 
End Sub

Private Sub setUpSetpoint(units_factor As Double)
      If (SetPoint > 0) Then
            Is_CCR = 1
'            If (SetPoint_Is_Bar) Then
'               SetPoint = SetPoint / 1.01325
'            End If
            PO2_CCR = SetPoint * units_factor
       Else
          Is_CCR = 0
      End If

End Sub

'C===============================================================================
'C Subroutine Set_Inspired_Inert
'C
'C This subroutine calculates and sets partial pressure for inert gases and
'C passes the calculated values back in the arguments piHe and piN2.
'C
'C When the calculation is for a CCR profile, the pre-calculated rates supplied
'C in the arguments (rateHe and rateN2) are modified to corrected values for
'C Schreiner computations.
'C
'C Call this routine before any use of the subroutines Schreiner_equation() or
'C Haldane_Equation().
'C
'C Requirements:
'C   The arguments 'rateN2' and 'rateHe' must be calculated by normal means
'C   prior to calling this method. e.g. rateN2=rate*vfraction_vnitrogen(vmix_vnumber).
'C
'C Warnings:
'C   The rate variables though not necessary for Haldane_equation must still
'C   be supplied. allocate a variable 'Real DUMMY' and supply it as both arguments.
'C   See gas_loadings_constant_depth in this file for example.
'C   FORTRAN passes arguments by reference. Whatever is supplied to these args
'C   will be the ADDRESS of the argument. You might be able to get away with
'C   passing a constant i.e. '0.0', but whether or not it works will be dependent
'C   on the compiler, linker and optimization. 'Zero' is a popular constant, and
'C   the compiler/linker could fold all instances into one location. In this case
'C   '0.0' would be trashed for every location in the program.
'C
'C   Be safe. Pass a dummy REAL.
'C
'C <CCR>
'C===============================================================================
Private Sub Set_Inspired_Inert_Starting() '(vmix_vnumber As Double, starting_ambient_vpressure As Double, water_vapor_vpressure As Double, vfraction_vhelium(vmix_vnumber) As Double, vfraction_vnitrogen(vmix_vnumber) As Double, piHe As Double, piN As double2, rateHe As Double, rateN2 As Double)
'a = Set_Inspired_Inert(vmix_vnumber, starting_ambient_vpressure, water_vapor_vpressure, vfraction_vhelium(vmix_vnumber), vfraction_vnitrogen(vmix_vnumber), initial_inspired_vhe_vpressure, initial_inspired_vn2_vpressure, vhelium_rate, vnitrogen_rate)


'      IMPLICIT NONE
'C===============================================================================
'C ARGUMENTS
'C===============================================================================
'      INTEGER mix       !input  current mix number
'      REAL pAmb         !input  current ambient pressure
'      REAL wvp          !input  Water Vapor Pressure
'      REAL fhe(10)      !input  helium mix array
'      REAL fn2(10)      !input  nitrogen mix array
'      REAL piHe         !OUTPUT inspired helium partial pressure at start of leg
'      REAL piN2         !OUTPUT inspired nitrogen partial pressure at start of leg
'      REAL rateHe       !I/O    rate of change of nitrogen pp/fsw
'      REAL rateN2       !I/O    rate of change of helium pp/fsw
'C===============================================================================
'C LOCAL VARIABLES
'C===============================================================================
'      Real adjAmb
'C===============================================================================
'C GLOBAL VARIABLES IN NAMED COMMON BLOCKS
'C===============================================================================
'      Real Barometric_Pressure
'      COMMON /Block_18/ Barometric_Pressure'

'      Logical units_equal_fsw, units_equal_msw
'      COMMON /Block_15/ Units_Equal_Fsw, Units_Equal_Msw

'      Real units_factor
'      COMMON /Block_16/ Units_Factor
'
'      REAL    SetPoint              ! setpoint currently in effect, if any
'      REAL    PO2_CCR               ! setpoint converted to fsw/msw
'      REAL    Effective_FO2           ! calculated here - effective fo2 at start of leg
'      REAL    Effective_FHE           ! calculated here - effective fHe at start of leg
'      REAL    Effective_FN2           ! calculated here - effective fN2 at start of leg
'      REAL    InertSum_Diluent      ! sum of ACTUAL inert gas fractions in diluent
'      LOGICAL Is_CCR                ! modified here - true if not OC
'      LOGICAL SetPoint_Is_Bar       ! true if setpoint is in BAR
'      COMMON /CCR_Block/ SetPoint,PO2_CCR,Effective_FO2,Effective_FHE,
'     * Effective_FN2,InertSum_Diluent,Is_CCR,SetPoint_Is_Bar
'C===============================================================================
'C calculations
'C===============================================================================
'      !
'      ! all calculations need this figured
'      !
      adjAmb = starting_ambient_vpressure - water_vapor_vpressure

      If (Is_CCR) Then

         adjAmb = adjAmb - PO2_CCR

         If (adjAmb < 0) Then
'            !
'            ! Don't disturb setpoint. Use adjAmb for adjustment and
'            ! message printing. That's OK, because we' going to change it
'            ! to a new value anyway.
'            !
            adjAmb = SetPoint
            'If (SetPoint_Is_Bar) Then
            '   adjAmb = adjAmb * 1.01325
            'End If

'            WRITE(*,
'     *            '("WARNING: SetPoint ",F4.2," unreachable at depth: ",
'     *            F5.1,".")') adjAmb,(pAmb-Barometric_Pressure)

'            !
'            ! Use adjAmb as a scratch variable again.
'            !
            adjAmb = ((starting_ambient_vpressure - water_vapor_vpressure) / units_factor)
            'If (SetPoint_Is_Bar) Then
            '   adjAmb = adjAmb * 1.01325
            'End If

'            WRITE(*,'("         Reducing SetPoint to: ",F4.2,".")')
'     *              adjAmb
            PO2_CCR = ((starting_ambient_vpressure - water_vapor_vpressure))
'            !
'            ! set adjAmb to new value
'            !
            adjAmb = starting_ambient_vpressure - water_vapor_vpressure - PO2_CCR
         End If

'         ! AdjAmb pressure now has all the pressure remaining
'         ! Calculate the partial pressures which will be
'         ! proportional to fInert/fInertSum
'         !
         InertSum_Diluent = vfraction_vhelium(vmix_vnumber) + vfraction_vnitrogen(vmix_vnumber)
         initial_inspired_vhe_vpressure = adjAmb * vfraction_vhelium(vmix_vnumber) / InertSum_Diluent
         initial_inspired_vn2_vpressure = adjAmb * vfraction_vnitrogen(vmix_vnumber) / InertSum_Diluent

'         !
'         ! Now adjust the OC rates (fInert*rate) so that their sum is 1.0
'         !
         If (vfraction_vhelium(vmix_vnumber) > 0) Then
           vhelium_rate = vhelium_rate * (vfraction_vhelium(vmix_vnumber) / InertSum_Diluent) / vfraction_vhelium(vmix_vnumber)
         Else
           vhelium_rate = 0
         End If

         If (vfraction_vnitrogen(vmix_vnumber) > 0) Then
           vnitrogen_rate = vnitrogen_rate * (vfraction_vnitrogen(vmix_vnumber) / InertSum_Diluent) / vfraction_vnitrogen(vmix_vnumber)
         Else
           vnitrogen_rate = 0
         End If
'         !
'         ! These values are not used by any component of the program
'         ! They are comuted for demonstrative purposes
'         !
         Effective_FO2 = PO2_CCR / starting_ambient_vpressure
         adjAmb = starting_ambient_vpressure - water_vapor_vpressure
         Effective_FHE = initial_inspired_vhe_vpressure / adjAmb
         Effective_FN2 = initial_inspired_vn2_vpressure / adjAmb

      Else
         initial_inspired_vhe_vpressure = adjAmb * vfraction_vhelium(vmix_vnumber)
         initial_inspired_vn2_vpressure = adjAmb * vfraction_vnitrogen(vmix_vnumber)
'         !
'         !just for completeness
'         !
         Effective_FO2 = 1# - (vfraction_vhelium(vmix_vnumber) + vfraction_vnitrogen(vmix_vnumber))
         Effective_FHE = vfraction_vhelium(vmix_vnumber)
         Effective_FN2 = vfraction_vnitrogen(vmix_vnumber)
      End If


End Sub

Private Sub Set_Inspired_Inert() '(vmix_vnumber As Double, starting_ambient_vpressure As Double, water_vapor_vpressure As Double, vfraction_vhelium(vmix_vnumber) As Double, vfraction_vnitrogen(vmix_vnumber) As Double, piHe As Double, piN As double2, rateHe As Double, rateN2 As Double)
'a = Set_Inspired_Inert(vmix_vnumber, starting_ambient_vpressure, water_vapor_vpressure, vfraction_vhelium(vmix_vnumber), vfraction_vnitrogen(vmix_vnumber), initial_inspired_vhe_vpressure, initial_inspired_vn2_vpressure, vhelium_rate, vnitrogen_rate)


'      IMPLICIT NONE
'C===============================================================================
'C ARGUMENTS
'C===============================================================================
'      INTEGER mix       !input  current mix number
'      REAL pAmb         !input  current ambient pressure
'      REAL wvp          !input  Water Vapor Pressure
'      REAL fhe(10)      !input  helium mix array
'      REAL fn2(10)      !input  nitrogen mix array
'      REAL piHe         !OUTPUT inspired helium partial pressure at start of leg
'      REAL piN2         !OUTPUT inspired nitrogen partial pressure at start of leg
'      REAL rateHe       !I/O    rate of change of nitrogen pp/fsw
'      REAL rateN2       !I/O    rate of change of helium pp/fsw
'C===============================================================================
'C LOCAL VARIABLES
'C===============================================================================
'      Real adjAmb
'C===============================================================================
'C GLOBAL VARIABLES IN NAMED COMMON BLOCKS
'C===============================================================================
'      Real Barometric_Pressure
'      COMMON /Block_18/ Barometric_Pressure'

'      Logical units_equal_fsw, units_equal_msw
'      COMMON /Block_15/ Units_Equal_Fsw, Units_Equal_Msw

'      Real units_factor
'      COMMON /Block_16/ Units_Factor
'
'      REAL    SetPoint              ! setpoint currently in effect, if any
'      REAL    PO2_CCR               ! setpoint converted to fsw/msw
'      REAL    Effective_FO2           ! calculated here - effective fo2 at start of leg
'      REAL    Effective_FHE           ! calculated here - effective fHe at start of leg
'      REAL    Effective_FN2           ! calculated here - effective fN2 at start of leg
'      REAL    InertSum_Diluent      ! sum of ACTUAL inert gas fractions in diluent
'      LOGICAL Is_CCR                ! modified here - true if not OC
'      LOGICAL SetPoint_Is_Bar       ! true if setpoint is in BAR
'      COMMON /CCR_Block/ SetPoint,PO2_CCR,Effective_FO2,Effective_FHE,
'     * Effective_FN2,InertSum_Diluent,Is_CCR,SetPoint_Is_Bar
'C===============================================================================
'C calculations
'C===============================================================================
'      !
'      ! all calculations need this figured
'      !
      adjAmb = ambient_vpressure - water_vapor_vpressure

      If (Is_CCR) Then

         adjAmb = adjAmb - PO2_CCR

         If (adjAmb < 0) Then
'            !
'            ! Don't disturb setpoint. Use adjAmb for adjustment and
'            ! message printing. That's OK, because we' going to change it
'            ! to a new value anyway.
'            !
            adjAmb = SetPoint
            'If (SetPoint_Is_Bar) Then
            '   adjAmb = adjAmb * 1.01325
            'End If

'            WRITE(*,
'     *            '("WARNING: SetPoint ",F4.2," unreachable at depth: ",
'     *            F5.1,".")') adjAmb,(pAmb-Barometric_Pressure)

'            !
'            ! Use adjAmb as a scratch variable again.
'            !
            adjAmb = ((ambient_vpressure - water_vapor_vpressure) / units_factor)
            'If (SetPoint_Is_Bar) Then
            '   adjAmb = adjAmb * 1.01325
            'End If

'            WRITE(*,'("         Reducing SetPoint to: ",F4.2,".")')
'     *              adjAmb
            PO2_CCR = ((ambient_vpressure - water_vapor_vpressure))
'            !
'            ! set adjAmb to new value
'            !
            adjAmb = ambient_vpressure - water_vapor_vpressure - PO2_CCR
         End If

'         ! AdjAmb pressure now has all the pressure remaining
'         ! Calculate the partial pressures which will be
'         ! proportional to fInert/fInertSum
'         !
         InertSum_Diluent = vfraction_vhelium(vmix_vnumber) + vfraction_vnitrogen(vmix_vnumber)
         inspired_vhelium_vpressure = adjAmb * vfraction_vhelium(vmix_vnumber) / InertSum_Diluent
         inspired_vnitrogen_vpressure = adjAmb * vfraction_vnitrogen(vmix_vnumber) / InertSum_Diluent

'         !
'         ! Now adjust the OC rates (fInert*rate) so that their sum is 1.0
'         !
         If (vfraction_vhelium(vmix_vnumber) > 0) Then
           vhelium_rate = vhelium_rate * (vfraction_vhelium(vmix_vnumber) / InertSum_Diluent) / vfraction_vhelium(vmix_vnumber)
         Else
           vhelium_rate = 0
         End If

         If (vfraction_vnitrogen(vmix_vnumber) > 0) Then
           vnitrogen_rate = vnitrogen_rate * (vfraction_vnitrogen(vmix_vnumber) / InertSum_Diluent) / vfraction_vnitrogen(vmix_vnumber)
         Else
           vnitrogen_rate = 0
         End If
'         !
'         ! These values are not used by any component of the program
'         ! They are comuted for demonstrative purposes
'         !
         Effective_FO2 = PO2_CCR / ambient_vpressure
         adjAmb = starting_ambient_vpressure - water_vapor_vpressure
         Effective_FHE = inspired_vhelium_vpressure / adjAmb
         Effective_FN2 = inspired_vnitrogen_vpressure / adjAmb

      Else
         inspired_vhelium_vpressure = adjAmb * vfraction_vhelium(vmix_vnumber)
         inspired_vnitrogen_vpressure = adjAmb * vfraction_vnitrogen(vmix_vnumber)
'         !
'         !just for completeness
'         !
         Effective_FO2 = 1# - (vfraction_vhelium(vmix_vnumber) + vfraction_vnitrogen(vmix_vnumber))
         Effective_FHE = vfraction_vhelium(vmix_vnumber)
         Effective_FN2 = vfraction_vnitrogen(vmix_vnumber)
      End If


End Sub



'nick code end here
Private Sub printtogrid3()
MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
MSFlexGrid4.Row = MSFlexGrid4.Rows - 1

MSFlexGrid4.Cols = 10
MSFlexGrid4.Col = 0
MSFlexGrid4.Text = "No."
MSFlexGrid4.Col = 1
MSFlexGrid4.Text = "Duration"
MSFlexGrid4.Col = 2
MSFlexGrid4.Text = "RunTime"
MSFlexGrid4.Col = 3
MSFlexGrid4.Text = "Mix"
MSFlexGrid4.Col = 4
MSFlexGrid4.Text = "Depth"
MSFlexGrid4.Col = 6
MSFlexGrid4.Text = "CNS"
MSFlexGrid4.Col = 7
MSFlexGrid4.Text = "OTU"
MSFlexGrid4.Col = 8
MSFlexGrid4.Text = "Rate"
MSFlexGrid4.Col = 5
MSFlexGrid4.Text = "Set Point"

For K = 0 To 8
  MSFlexGrid4.Col = K
  MSFlexGrid4.CellBackColor = &H8000000F
Next K
End Sub
Private Sub printtogrid4()
MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
MSFlexGrid4.Row = MSFlexGrid4.Rows - 1
MSFlexGrid4.Col = 0
MSFlexGrid4.Text = CStr(vsegment_vnumber)
MSFlexGrid4.Col = 1
MSFlexGrid4.Text = Format(vsegment_vtime, "###0.0")
MSFlexGrid4.Col = 2
MSFlexGrid4.Text = Format(run_vtime, "###0.0")
MSFlexGrid4.Col = 3
If lblhe(vmix_vnumber - 1).Caption = "0" Then
    If lbl02(vmix_vnumber - 1).Caption = "21" Then
      MSFlexGrid4.Text = "  Air"
    Else
      MSFlexGrid4.Text = " Nx" + CStr(lbl02(vmix_vnumber - 1).Caption)
    End If
Else
    MSFlexGrid4.Text = "TX" + CStr(lbl02(vmix_vnumber - 1).Caption + "/" + lblhe(vmix_vnumber - 1).Caption)
End If
MSFlexGrid4.Col = 9
MSFlexGrid4.Text = CStr(vmix_vnumber)
MSFlexGrid4.Col = 4
MSFlexGrid4.Text = Format(vdeco_vstop_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then MSFlexGrid4.Text = Format(vdeco_vstop_vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring)
'xxdum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
MSFlexGrid4.Col = 6
MSFlexGrid4.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
MSFlexGrid4.Col = 7
MSFlexGrid4.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
MSFlexGrid4.Col = 8
MSFlexGrid4.Text = Format(rate * feetormeter_factor, "###")
If rate > 0.01 Then
  MSFlexGrid4.Text = "Descent        " + MSFlexGrid4.Text
Else
  If rate < -0.01 Then
    MSFlexGrid4.Text = "Ascent        " + MSFlexGrid4.Text
  Else
    MSFlexGrid4.Text = " ---- "
  End If
End If
MSFlexGrid4.Col = 5
If SetPoint < 0.21 Then MSFlexGrid4.Text = " " Else MSFlexGrid4.Text = CStr(SetPoint) + " "
For K = 0 To 8
  MSFlexGrid4.Col = K
  MSFlexGrid4.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid4lite()
If rate > 0.01 Then
   MSFlexGrid4lite.Rows = MSFlexGrid4lite.Rows + 1
   MSFlexGrid4lite.Row = MSFlexGrid4lite.Rows - 1
   MSFlexGrid4lite.Col = 8
   MSFlexGrid4lite.Text = Format(rate * feetormeter_factor, "###")
   MSFlexGrid4lite.Text = "Descent        " + MSFlexGrid4lite.Text
Else
   If rate < -0.01 Then
      MSFlexGrid4lite.Rows = MSFlexGrid4lite.Rows + 1
      MSFlexGrid4lite.Row = MSFlexGrid4lite.Rows - 1
      MSFlexGrid4lite.Col = 8
      MSFlexGrid4lite.Text = Format(rate * feetormeter_factor, "###")
      MSFlexGrid4lite.Text = "Ascent        " + MSFlexGrid4lite.Text
   End If
End If
MSFlexGrid4lite.Col = 0
MSFlexGrid4lite.Text = CStr(vsegment_vnumber)
MSFlexGrid4lite.Col = 1
tempdurationlite = MSFlexGrid4lite.Text
MSFlexGrid4lite.Text = Format(vsegment_vtime, "###0.0")
'MsFlexgrid4lite.Text = Left(MsFlexgrid4lite.Text, Len(MsFlexgrid4lite.Text) - 2) + ":" + Format((CDbl(Right(MsFlexgrid4lite.Text, 2)) * 60#), "00")
tempdurationlitesec = Right(tempdurationlite, 2)
If Len(tempdurationlite) > 3 Then tempdurationlitemins = Left(tempdurationlite, Len(tempdurationlite) - 3)
tempgridlitesec = Right(MSFlexGrid4lite.Text, 2) * 60

If Len(MSFlexGrid4lite.Text) > 2 Then tempgridlitemins = Left(MSFlexGrid4lite.Text, Len(MSFlexGrid4lite.Text) - 2)
If Val(tempgridlitesec) + Val(tempdurationlitesec) > 59 Then
   templeftsec = (CInt(tempgridlitesec) + CInt(tempdurationlitesec)) - 60
   templeftsec = Format(CDbl(templeftsec), "00")
   tempgridlitemins = CInt(tempgridlitemins) + 1
Else
   templeftsec = Format(CDbl(Val(tempgridlitesec) + Val(tempdurationlitesec)), "00")
End If
totalmins = CInt(tempgridlitemins) + CInt(tempdurationlitemins)
MSFlexGrid4lite.Text = CDbl(totalmins) & ":" & templeftsec
MSFlexGrid4lite.Col = 2
MSFlexGrid4lite.Text = Format(run_vtime, "###0.0")
'MsFlexgrid4lite.Text = Left(MsFlexgrid4lite.Text, Len(MsFlexgrid4lite.Text) - 2) + ":" + Format((CDbl(Right(MsFlexgrid4lite.Text.Text, 2)) * 60#), "00")
MSFlexGrid4lite.Col = 3
If lblhe(vmix_vnumber - 1).Caption = "0" Then
    If lbl02(vmix_vnumber - 1).Caption = "21" Then
      MSFlexGrid4lite.Text = "  Air"
    Else
      MSFlexGrid4lite.Text = " Nx" + CStr(lbl02(vmix_vnumber - 1).Caption)
    End If
Else
    MSFlexGrid4lite.Text = "TX" + CStr(lbl02(vmix_vnumber - 1).Caption + "/" + lblhe(vmix_vnumber - 1).Caption)
End If
'MsFlexgrid4lite.Text = CStr(vmix_vnumber - 1)
MSFlexGrid4lite.Col = 4
MSFlexGrid4lite.Text = Format(vdeco_vstop_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then MSFlexGrid4lite.Text = Format(vdeco_vstop_vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring)
dum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
MSFlexGrid4lite.Col = 6
MSFlexGrid4lite.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
MSFlexGrid4lite.Col = 7
MSFlexGrid4lite.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)

MSFlexGrid4lite.Col = 5
If SetPoint < 0.21 Then MSFlexGrid4lite.Text = " " Else MSFlexGrid4lite.Text = CStr(SetPoint) + " "
For K = 0 To 8
  MSFlexGrid4lite.Col = K
  MSFlexGrid4lite.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid5()
MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
MSFlexGrid4.Row = MSFlexGrid4.Rows - 1
MSFlexGrid4.Col = 0
MSFlexGrid4.Text = CStr(vsegment_vnumber)
MSFlexGrid4.Col = 1
MSFlexGrid4.Text = Format(vsegment_vtime, "###0.0") 'CStr((CDbl(CInt(vsegment_vtime * 10# + 0.4999) / 10)))
MSFlexGrid4.Col = 2
MSFlexGrid4.Text = Format(run_vtime, "###0.0") 'CStr((CDbl(CInt(run_vtime * 10# + 0.4999) / 10)))
MSFlexGrid4.Col = 3
If lblhe(vmix_vnumber - 1).Caption = "0" Then
    If lbl02(vmix_vnumber - 1).Caption = "21" Then
      MSFlexGrid4.Text = "  Air"
    Else
      MSFlexGrid4.Text = " Nx" + CStr(lbl02(vmix_vnumber - 1).Caption)
    End If
Else
    MSFlexGrid4.Text = "TX" + CStr(lbl02(vmix_vnumber - 1).Caption + "/" + lblhe(vmix_vnumber - 1).Caption)
End If
MSFlexGrid4.Col = 9
MSFlexGrid4.Text = CStr(vmix_vnumber)
MSFlexGrid4.Col = 4
MSFlexGrid4.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.4999) / 10) * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then MSFlexGrid4.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.4999) / 10) * feetormeter_factor, "###0.0" & feetormeter_shortstring)
'xxdum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
MSFlexGrid4.Col = 6
MSFlexGrid4.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
MSFlexGrid4.Col = 7
MSFlexGrid4.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
MSFlexGrid4.Col = 8
MSFlexGrid4.Text = Format(rate * feetormeter_factor, "###")
If rate > 0.01 Then
  MSFlexGrid4.Text = "Descent        " + MSFlexGrid4.Text
Else
  If rate < -0.01 Then
    MSFlexGrid4.Text = "Ascent        " + MSFlexGrid4.Text
  Else
    MSFlexGrid4.Text = " ---- "
  End If
End If
MSFlexGrid4.Col = 5
If SetPoint < 0.21 Then MSFlexGrid4.Text = " " Else MSFlexGrid4.Text = CStr(SetPoint) + " "
For K = 0 To 8
  MSFlexGrid4.Col = K
  MSFlexGrid4.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid5lite()
If rate > 0.01 Then
   MSFlexGrid4lite.Rows = MSFlexGrid4lite.Rows + 1
   MSFlexGrid4lite.Row = MSFlexGrid4lite.Rows - 1
   MSFlexGrid4lite.Col = 8
   MSFlexGrid4lite.Text = Format(rate * feetormeter_factor, "###")
   MSFlexGrid4lite.Text = "Descent        " + MSFlexGrid4lite.Text
Else
   If rate < -0.01 Then
      MSFlexGrid4lite.Rows = MSFlexGrid4lite.Rows + 1
      MSFlexGrid4lite.Row = MSFlexGrid4lite.Rows - 1
      MSFlexGrid4lite.Col = 8
      MSFlexGrid4lite.Text = Format(rate * feetormeter_factor, "###")
      MSFlexGrid4lite.Text = "Ascent        " + MSFlexGrid4lite.Text
   End If
End If
MSFlexGrid4lite.Col = 0
MSFlexGrid4lite.Text = CStr(vsegment_vnumber)
MSFlexGrid4lite.Col = 1
tempdurationlite = MSFlexGrid4lite.Text
MSFlexGrid4lite.Text = Format(vsegment_vtime, "###0.0")
'MsFlexgrid4lite.Text = Left(MsFlexgrid4lite.Text, Len(MsFlexgrid4lite.Text) - 2) + ":" + Format((CDbl(Right(MsFlexgrid4lite.Text, 2)) * 60#), "00")
tempdurationlitesec = Right(tempdurationlite, 2)
If Len(tempdurationlite) > 3 Then tempdurationlitemins = Left(tempdurationlite, Len(tempdurationlite) - 3)
tempgridlitesec = Right(MSFlexGrid4lite.Text, 2) * 60#

If Len(MSFlexGrid4lite.Text) > 2 Then tempgridlitemins = Left(MSFlexGrid4lite.Text, Len(MSFlexGrid4lite.Text) - 2)
If Val(tempgridlitesec) + Val(tempdurationlitesec) > 59 Then
   templeftsec = (CInt(tempgridlitesec) + CInt(tempdurationlitesec)) - 60
   templeftsec = Format(CDbl(templeftsec), "00")
   tempgridlitemins = CInt(tempgridlitemins) + 1
Else
   templeftsec = Format(CDbl(Val(tempgridlitesec) + Val(tempdurationlitesec)), "00")
End If
totalmins = CInt(tempgridlitemins) + CInt(tempdurationlitemins)
MSFlexGrid4lite.Text = CDbl(totalmins) & ":" & templeftsec
MSFlexGrid4lite.Col = 2
MSFlexGrid4lite.Text = Format(run_vtime, "###0.0") 'CStr((CDbl(CInt(run_vtime * 10# + 0.4999) / 10)))
'MsFlexgrid4lite.Text = Left(MsFlexgrid4lite.Text, Len(MsFlexgrid4lite.Text) - 2) + ":" + Format((CDbl(Right(MsFlexgrid4lite.Text, 2)) * 60#), "00")
MSFlexGrid4lite.Col = 3
If lblhe(vmix_vnumber - 1).Caption = "0" Then
    If lbl02(vmix_vnumber - 1).Caption = "21" Then
      MSFlexGrid4lite.Text = "  Air"
    Else
      MSFlexGrid4lite.Text = " Nx" + CStr(lbl02(vmix_vnumber - 1).Caption)
    End If
Else
    MSFlexGrid4lite.Text = "TX" + CStr(lbl02(vmix_vnumber - 1).Caption + "/" + lblhe(vmix_vnumber - 1).Caption)
End If
'MsFlexgrid4lite.Text = CStr(vmix_vnumber - 1)
MSFlexGrid4lite.Col = 4
MSFlexGrid4lite.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.4999) / 10) * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then MSFlexGrid4lite.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.4999) / 10) * feetormeter_factor, "###0.0" & feetormeter_shortstring)
dum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
MSFlexGrid4lite.Col = 6
MSFlexGrid4lite.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
MSFlexGrid4lite.Col = 7
MSFlexGrid4lite.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)

MSFlexGrid4lite.Col = 5
If SetPoint < 0.21 Then MSFlexGrid4lite.Text = " " Else MSFlexGrid4lite.Text = CStr(SetPoint) + " "
For K = 0 To 8
  MSFlexGrid4lite.Col = K
  MSFlexGrid4lite.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid6()
MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
MSFlexGrid4.Row = MSFlexGrid4.Rows - 1
MSFlexGrid4.Col = 0
MSFlexGrid4.Text = CStr(vsegment_vnumber)
MSFlexGrid4.Col = 1
MSFlexGrid4.Text = Format(vsegment_vtime, "###0.0") 'CStr((CDbl(CInt(vsegment_vtime * 10# + 0.999) / 10)))
MSFlexGrid4.Col = 2
MSFlexGrid4.Text = Format(run_vtime, "###0.0") 'CStr((CDbl(CInt(run_vtime * 10# + 0.999) / 10)))
MSFlexGrid4.Col = 3
If lblhe(vmix_vnumber - 1).Caption = "0" Then
    If lbl02(vmix_vnumber - 1).Caption = "21" Then
      MSFlexGrid4.Text = "  Air"
    Else
      MSFlexGrid4.Text = " Nx" + CStr(lbl02(vmix_vnumber - 1).Caption)
    End If
Else
    MSFlexGrid4.Text = "TX" + CStr(lbl02(vmix_vnumber - 1).Caption + "/" + lblhe(vmix_vnumber - 1).Caption)
End If
MSFlexGrid4.Col = 9
MSFlexGrid4.Text = CStr(vmix_vnumber)
MSFlexGrid4.Col = 4
MSFlexGrid4.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.999) / 10) * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then MSFlexGrid4.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.999) / 10) * feetormeter_factor, "###0.0" & feetormeter_shortstring)
'xxdum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) '
MSFlexGrid4.Col = 6
MSFlexGrid4.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
MSFlexGrid4.Col = 7
MSFlexGrid4.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
MSFlexGrid4.Col = 8
MSFlexGrid4.Text = Format(rate * feetormeter_factor, "###")
If rate > 0.01 Then
  MSFlexGrid4.Text = "Descent        " + MSFlexGrid4.Text
Else
  If rate < -0.01 Then
    MSFlexGrid4.Text = "Ascent        " + MSFlexGrid4.Text
  Else
    MSFlexGrid4.Text = " ---- "
  End If
End If
MSFlexGrid4.Col = 5
If SetPoint < 0.21 Then MSFlexGrid4.Text = " " Else MSFlexGrid4.Text = CStr(SetPoint) + " "
For K = 0 To 8
  MSFlexGrid4.Col = K
  MSFlexGrid4.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid2()
MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
MSFlexGrid4.Row = MSFlexGrid4.Rows - 1
MSFlexGrid4.Col = 0
MSFlexGrid4.Text = CStr(vsegment_vnumber)
MSFlexGrid4.Col = 1
MSFlexGrid4.Text = Format(vsegment_vtime, "###0.0")
MSFlexGrid4.Col = 2
MSFlexGrid4.Text = Format(run_vtime, "###0.0")
MSFlexGrid4.Col = 3
If lblhe(vmix_vnumber - 1).Caption = "0" Then
    If lbl02(vmix_vnumber - 1).Caption = "21" Then
      MSFlexGrid4.Text = "  Air"
    Else
      MSFlexGrid4.Text = " Nx" + CStr(lbl02(vmix_vnumber - 1).Caption)
    End If
Else
    MSFlexGrid4.Text = "TX" + CStr(lbl02(vmix_vnumber - 1).Caption + "/" + lblhe(vmix_vnumber - 1).Caption)
End If
MSFlexGrid4.Col = 9
MSFlexGrid4.Text = CStr(vmix_vnumber)
MSFlexGrid4.Col = 4
MSFlexGrid4.Text = Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdepth = 4.5 Then MSFlexGrid4.Text = Format(vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring)
'xxdum = ppo2exposuretime(vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
MSFlexGrid4.Col = 6
MSFlexGrid4.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
MSFlexGrid4.Col = 7
MSFlexGrid4.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
MSFlexGrid4.Col = 5
If SetPoint < 0.21 Then MSFlexGrid4.Text = " " Else MSFlexGrid4.Text = CStr(SetPoint) + " "
MSFlexGrid4.Col = 8
MSFlexGrid4.Text = " ---- "
End Sub
Private Sub printtogrid2lite()

MSFlexGrid4lite.Col = 0
MSFlexGrid4lite.Text = CStr(vsegment_vnumber)
MSFlexGrid4lite.Col = 1
tempdurationlite = MSFlexGrid4lite.Text
MSFlexGrid4lite.Text = Format(vsegment_vtime, "###0.0")
'MSFlexGrid4lite.Text = Left(MSFlexGrid4lite.Text, Len(MSFlexGrid4lite.Text) - 2) + ":" + Format((CDbl(Right(MSFlexGrid4lite.Text, 2)) * 60#), "00")
tempdurationlitesec = Right(tempdurationlite, 2)
If Len(tempdurationlite) > 3 Then tempdurationlitemins = Left(tempdurationlite, Len(tempdurationlite) - 3)
tempgridlitesec = Right(MSFlexGrid4lite.Text, 2) * 60#

If Len(MSFlexGrid4lite.Text) > 2 Then tempgridlitemins = Left(MSFlexGrid4lite.Text, Len(MSFlexGrid4lite.Text) - 2)
If CInt(tempgridlitesec) + CInt(tempdurationlitesec) > 59 Then
   templeftsec = (CInt(tempgridlitesec) + CInt(tempdurationlitesec)) - 60
   templeftsec = Format(CDbl(templeftsec), "00")
   tempgridlitemins = CInt(tempgridlitemins) + 1
Else
   templeftsec = Format(CDbl(CInt(tempgridlitesec) + CInt(tempdurationlitesec)), "00")
End If
totalmins = CInt(tempgridlitemins) + CInt(tempdurationlitemins)
MSFlexGrid4lite.Text = CDbl(totalmins) & ":" & templeftsec
MSFlexGrid4lite.Col = 2
MSFlexGrid4lite.Text = Format(run_vtime, "###0.0")
'MSFlexGrid4.Text = Left(MSFlexGrid4.Text, Len(MSFlexGrid4.Text) - 2) + ":" + Format((CDbl(Right(MSFlexGrid4.Text, 2)) * 60#), "00")
MSFlexGrid4lite.Col = 3
If lblhe(vmix_vnumber - 1).Caption = "0" Then
    If lbl02(vmix_vnumber - 1).Caption = "21" Then
      MSFlexGrid4lite.Text = "  Air"
    Else
      MSFlexGrid4lite.Text = " Nx" + CStr(lbl02(vmix_vnumber - 1).Caption)
    End If
Else
    MSFlexGrid4lite.Text = "TX" + CStr(lbl02(vmix_vnumber - 1).Caption + "/" + lblhe(vmix_vnumber - 1).Caption)
End If
MSFlexGrid4lite.Col = 4
MSFlexGrid4lite.Text = Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdepth = 4.5 Then MSFlexGrid4lite.Text = Format(vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring)
dum = ppo2exposuretime(vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
MSFlexGrid4lite.Col = 6
MSFlexGrid4lite.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
MSFlexGrid4lite.Col = 7
MSFlexGrid4lite.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
MSFlexGrid4lite.Col = 5
If SetPoint < 0.21 Then MSFlexGrid4lite.Text = " " Else MSFlexGrid4lite.Text = CStr(SetPoint) + " "
'MSFlexGrid4.Col = 8
'MSFlexGrid4.Text = " ---- "
End Sub

Private Sub duplicategrid()
Dim r(11) As Integer
Dim a As Integer
Dim b As Integer

r(0) = 8
r(1) = 1
r(2) = 2
r(3) = 3
r(4) = 0
r(5) = 4
r(6) = 5
r(7) = 6
r(8) = 7
r(9) = 9
If MSFlexGrid1.Rows > 11 Then
  MsgBox "Too many dives!!"
  Exit Sub
End If
   decoresultgrid(T - 1).Rows = MSFlexGrid4.Rows
   decoresultgrid(T - 1).Cols = 10
   decoresultgrid(T - 1).FixedRows = 1
   For a = 0 To MSFlexGrid4.Rows - 1
      For b = 0 To 9
        decoresultgrid(T - 1).Row = a
        decoresultgrid(T - 1).Col = r(b)
        MSFlexGrid4.Row = a
        MSFlexGrid4.Col = b
        decoresultgrid(T - 1).Text = MSFlexGrid4.Text
        decoresultgrid(T - 1).CellBackColor = MSFlexGrid4.CellBackColor
        decoresultgrid(T - 1).ColWidth(r(b)) = MSFlexGrid4.ColWidth(b)
      Next b
   Next a
   b = decoresultgrid(T - 1).Rows
   For a = 1 To b - 1
      decoresultgrid(T - 1).Row = a
      decoresultgrid(T - 1).Col = 1
      S = decoresultgrid(T - 1).Text
      decoresultgrid(T - 1).Text = Left(decoresultgrid(T - 1).Text, Len(decoresultgrid(T - 1).Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid(T - 1).Text, 2)) * 60#), "00")
      S = decoresultgrid(T - 1).Text
      decoresultgrid(T - 1).Col = 2
      S = decoresultgrid(T - 1).Text
      decoresultgrid(T - 1).Text = Left(decoresultgrid(T - 1).Text, Len(decoresultgrid(T - 1).Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid(T - 1).Text, 2)) * 60#), "00")
      S = decoresultgrid(T - 1).Text
    Next a
    
    
    
    decoresultgridlite(T - 1).Rows = MSFlexGrid4lite.Rows
    decoresultgridlite(T - 1).Cols = 10
    decoresultgridlite(T - 1).FixedRows = 1
    For a = 0 To MSFlexGrid4lite.Rows - 1
      For b = 0 To 8
        decoresultgridlite(T - 1).Row = a
        decoresultgridlite(T - 1).Col = r(b)
        MSFlexGrid4lite.Row = a
        MSFlexGrid4lite.Col = b
        If b = 2 Then
           decoresultgridlite(T - 1).Text = MSFlexGrid4lite.Text
        Else
           decoresultgridlite(T - 1).Text = MSFlexGrid4lite.Text
        End If
        
        decoresultgridlite(T - 1).CellBackColor = MSFlexGrid4lite.CellBackColor
        decoresultgridlite(T - 1).ColWidth(r(b)) = MSFlexGrid4lite.ColWidth(b)
      Next b
    Next a
    b = decoresultgridlite(T - 1).Rows
    
    For a = 1 To b - 1
       decoresultgridlite(T - 1).Row = a
       decoresultgridlite(T - 1).Col = 1
       S = decoresultgridlite(T - 1).Text
'       decoresultgridlite(T - 1).Text = Left(S, Len(S) - 3) + ":" + Format((CDbl(Right(S, 2))), "00")
       S = decoresultgridlite(T - 1).Text
       decoresultgridlite(T - 1).Col = 2
       S = decoresultgridlite(T - 1).Text
       decoresultgridlite(T - 1).Text = Left(decoresultgridlite(T - 1).Text, Len(decoresultgridlite(T - 1).Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite(T - 1).Text, 2))), "00")
       S = decoresultgridlite(T - 1).Text
    Next a
 'Else

End Sub
Private Sub cleardecogrid()
   MSFlexGrid4.Rows = 1
End Sub
Private Sub printtogrid()
MSFlexGrid4.Rows = MSFlexGrid4.Rows + 1
MSFlexGrid4.Row = MSFlexGrid4.Rows - 1
MSFlexGrid4.Col = 0
MSFlexGrid4.Text = CStr(vsegment_vnumber)
MSFlexGrid4.Col = 1
MSFlexGrid4.Text = Format(vsegment_vtime, "###0.0")
MSFlexGrid4.Col = 2
MSFlexGrid4.Text = Format(run_vtime, "###0.0")
MSFlexGrid4.Col = 3
If lblhe(vmix_vnumber - 1).Caption = "0" Then
    If lbl02(vmix_vnumber - 1).Caption = "21" Then
      MSFlexGrid4.Text = "  Air"
    Else
      MSFlexGrid4.Text = " Nx" + CStr(lbl02(vmix_vnumber - 1).Caption)
    End If
Else
    MSFlexGrid4.Text = "TX" + CStr(lbl02(vmix_vnumber - 1).Caption + "/" + lblhe(vmix_vnumber - 1).Caption)
End If
MSFlexGrid4.Col = 9
MSFlexGrid4.Text = CStr(vmix_vnumber)
MSFlexGrid4.Col = 4
MSFlexGrid4.Text = Format(ending_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If ending_vdepth = 4.5 Then MSFlexGrid4.Text = Format(ending_vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'xxdum = ppo2exposuretime(ending_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
MSFlexGrid4.Col = 6
MSFlexGrid4.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
MSFlexGrid4.Col = 7
MSFlexGrid4.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
MSFlexGrid4.Col = 8
MSFlexGrid4.Text = Format(rate * feetormeter_factor, "###")
If rate > 0.01 Then
  MSFlexGrid4.Text = "Descent        " + MSFlexGrid4.Text
Else
  If rate < -0.01 Then
    MSFlexGrid4.Text = "Ascent        " + MSFlexGrid4.Text
  Else
    MSFlexGrid4.Text = " ---- "
  End If
End If
MSFlexGrid4.Col = 5
If SetPoint < 0.21 Then MSFlexGrid4.Text = " " Else MSFlexGrid4.Text = CStr(SetPoint) + " "
End Sub
Private Sub printtogridlite()
If rate > 0.01 Then
   MSFlexGrid4lite.Rows = MSFlexGrid4lite.Rows + 1
   MSFlexGrid4lite.Row = MSFlexGrid4lite.Rows - 1
   MSFlexGrid4lite.Col = 8
   MSFlexGrid4lite.Text = Format(rate * feetormeter_factor, "###")
   MSFlexGrid4lite.Text = "Descent        " + MSFlexGrid4lite.Text
Else
   MSFlexGrid4lite.Rows = MSFlexGrid4lite.Rows + 1
   MSFlexGrid4lite.Row = MSFlexGrid4lite.Rows - 1
   MSFlexGrid4lite.Col = 8
   MSFlexGrid4lite.Text = Format(rate * feetormeter_factor, "###")
   MSFlexGrid4lite.Text = "Ascent        " + MSFlexGrid4lite.Text
End If

MSFlexGrid4lite.Col = 0
MSFlexGrid4lite.Text = CStr(vsegment_vnumber)
MSFlexGrid4lite.Col = 1
If rate > 0.01 Then
  MSFlexGrid4lite.Text = Format(vsegment_vtime, "###0.0")
  MSFlexGrid4lite.Text = Left(MSFlexGrid4lite.Text, Len(MSFlexGrid4lite.Text) - 2) + ":" + Format((CDbl(Right(MSFlexGrid4lite.Text, 2)) * 60#), "00")
Else
  If rate < -0.01 Then
    MSFlexGrid4lite.Text = Format(vsegment_vtime, "###0.0")
    MSFlexGrid4lite.Text = Left(MSFlexGrid4lite.Text, Len(MSFlexGrid4lite.Text) - 2) + ":" + Format((CDbl(Right(MSFlexGrid4lite.Text, 2)) * 60#), "00")
  Else
  End If
End If
MSFlexGrid4lite.Col = 2
MSFlexGrid4lite.Text = Format(run_vtime, "###0.0")
'MsFlexgrid4lite.Text = Left(MsFlexgrid4lite.Text, Len(MsFlexgrid4lite.Text) - 2) + ":" + Format((CDbl(Right(MsFlexgrid4lite.Text, 2)) * 60#), "00")
MSFlexGrid4lite.Col = 3
If lblhe(vmix_vnumber - 1).Caption = "0" Then
    If lbl02(vmix_vnumber - 1).Caption = "21" Then
      MSFlexGrid4lite.Text = "  Air"
    Else
      MSFlexGrid4lite.Text = " Nx" + CStr(lbl02(vmix_vnumber - 1).Caption)
    End If
Else
    MSFlexGrid4lite.Text = "TX" + CStr(lbl02(vmix_vnumber - 1).Caption + "/" + lblhe(vmix_vnumber - 1).Caption)
End If
'MsFlexgrid4lite.Text = CStr(vmix_vnumber - 1)
MSFlexGrid4lite.Col = 4
MSFlexGrid4lite.Text = Format(ending_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If ending_vdepth = 4.5 Then MSFlexGrid4lite.Text = Format(ending_vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
dum = ppo2exposuretime(ending_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
MSFlexGrid4lite.Col = 6
MSFlexGrid4lite.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
MSFlexGrid4lite.Col = 7
MSFlexGrid4lite.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'MsFlexgrid4lite.Col = 8
'MsFlexgrid4lite.Text = Format(rate * feetormeter_factor, "###")
'If rate > 0.01 Then
'  MsFlexgrid4lite.Text = "Descent        " + MsFlexgrid4lite.Text
'Else
'  If rate < -0.01 Then
'    MsFlexgrid4lite.Text = "Ascent        " + MsFlexgrid4lite.Text
'  Else
'    MsFlexgrid4lite.Text = " ---- "
'  End If
'End If
MSFlexGrid4lite.Col = 5
If SetPoint < 0.21 Then MSFlexGrid4lite.Text = " " Else MSFlexGrid4lite.Text = CStr(SetPoint) + " "
End Sub
Private Sub MSFlexGrid2_Click()
'cleargrid1
If MSFlexGrid2.Rows > 1 Then
checkslected = False
rowidentified2 = MSFlexGrid2.Row
For K = 0 To MSFlexGrid2.Rows - 1
  For p = 0 To 0
    MSFlexGrid2.Row = K
    MSFlexGrid2.Col = p
    If MSFlexGrid2.CellBackColor = vbGreen Then
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
         For H = 0 To 7
            MSFlexGrid2.Row = K
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
         For H = 0 To 7
            MSFlexGrid2.Row = K
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
Next K
MSFlexGrid2.Col = 0
MSFlexGrid2.Row = rowidentified2
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
For p = 0 To 7
    MSFlexGrid2.Col = p
    MSFlexGrid2.Row = rowidentified2
    MSFlexGrid2.CellForeColor = vbWhite
    MSFlexGrid2.CellBackColor = vbGreen
Next
  MSFlexGrid2.Col = 0
  SQL = "SELECT * FROM seqdpprofile"
  SQL = SQL & " where dpprofileid = '" & MSFlexGrid2.Text & "' "
  SQL = SQL & " order by dpnumseq "
  Set RS = DB.OpenRecordset(SQL)
    RS.MoveFirst
    Label5(1).Caption = "Dive " + MSFlexGrid2.Text + ":" + vbCrLf
    Label5(1).Alignment = 2
  While RS.EOF = False
'     Label5(1).Caption = Label5(1).Caption + RS("dpnumseq") + " "
     Label5(1).Caption = Label5(1).Caption + Format(CStr(CDbl(RS("depth")) * feetormeter_factor), "###") + feetormeter_shortstring + " "
     Label5(1).Caption = Label5(1).Caption + RS("duration") + "mins "
     Label5(1).Caption = Label5(1).Caption + RS("dpcircuit") + vbCrLf
'     Label5(1).Caption = Label5(1).Caption + RS("dpo2")
'     Label5(1).Caption = Label5(1).Caption + RS("dphe")
'     Label5(1).Caption = Label5(1).Caption + RS("po2")
'     Label5(1).Caption = Label5(1).Caption + RS("gasid")
     RS.MoveNext
   Wend
Else
   MsgBox "No Dive found, please insert some Dive(s) !"
End If
rowidentified2 = MSFlexGrid2.Row
'clearhlgrid2
For q = 0 To 7
   MSFlexGrid2.Row = rowidentified2
   MSFlexGrid2.Col = q
   MSFlexGrid2.CellForeColor = vbWhite
   MSFlexGrid2.CellBackColor = vbGreen
Next q
For p = 0 To 0
  MSFlexGrid2.Row = rowidentified2
  MSFlexGrid2.Col = p
  txtplanno.Text = MSFlexGrid2.Text
Next p
If MSFlexGrid3.Rows > 1 Then
   'cleargriddata
   MSFlexGrid3.Rows = 1
End If
If txtplanno.Text <> "" Then
   display_grid2 = 1
   loaddpprofiledata
   display_grid2 = 0
End If
If MSFlexGrid1.Rows > 1 Then cmdmodify.Visible = True

If rowidentified2 < 9 Then MSFlexGrid2.TopRow = 1 Else MSFlexGrid2.TopRow = rowidentified2 - 8
End Sub

Private Sub display_deco_graph(grid_num As Integer)

Dim i As Integer
Dim K As Integer
Dim maxd As Integer
Dim dive_time As Integer
Dim deco_section As Integer
Dim C As Long

'Add by Goh on 08/11/2004 12:30pm
 
If systemversion = "Pro" Then
   If decoresultgrid(grid_num).Rows = 0 Then
      deco_update = 0
   Else
      deco_update = 1
   End If
Else
   If decoresultgridlite(grid_num).Rows = 0 Then
      deco_update = 0
   Else
      deco_update = 1
   End If
End If
'Add by Goh on 08/11/2004 12:30pm
  If deco_update = 0 Then
    For i = 0 To 9
      decoresultgrid(i).Rows = 0
      decoresultgridlite(i).Rows = 0
    Next i
    Picture1(grid_num).Visible = False
    mnugaslist_Click
    Exit Sub
  Else
    Picture1(grid_num).Visible = True
    deco_grid_display = grid_num
    mnugraph_Click
  End If
  row_count = 1
  deco_section = 0 '4 '6
  Picture1(grid_num).Cls
  'picture1(grid_num).Line (0, 0)-(0, 0), vbBlue
  maxd = 0
  dive_time = 0
'  If systemversion = "Pro" Then
     For i = 1 To decoresultgrid(grid_num).Rows - 1
       decoresultgrid(grid_num).Row = i
       decoresultgrid(grid_num).Col = 8
      '  MsgBox decoresultgrid(grid_num).Text
       If IsNumeric(decoresultgrid(grid_num).Text) = True Then
         If CInt(decoresultgrid(grid_num).Text) = row_count Then
           decoresultgrid(grid_num).Col = 2
           S = decoresultgrid(grid_num).Text
         '  MsgBox S
           X(row_count) = CSng(Left(decoresultgrid(grid_num).Text, Len(decoresultgrid(grid_num).Text) - 3))
           decoresultgrid(grid_num).Col = deco_section
           Y(row_count) = CSng(CDbl(Left(decoresultgrid(grid_num).Text, Len(decoresultgrid(grid_num).Text) - Len(feetormeter_shortstring))) / feetormeter_factor)
           If Y(row_count) > maxd Then maxd = Y(row_count)
           row_count = row_count + 1
         End If
       Else
         If CStr(decoresultgrid(grid_num).Text) = "" Then Exit For
         If InStr(1, (decoresultgrid(grid_num).Text), "No") Then deco_section = 4
       End If
     Next i
  
  If row_count < 2 Then Exit Sub
  row_count = row_count - 1
  yscale = ((Picture1(grid_num).Height - 150) / maxd)
  xscale = ((Picture1(grid_num).Width - 150) / X(row_count))
  runtime_graph = X(row_count)
  'ysacle = picture1(grid_num).ScaleHeight
  'xscale = picture1(grid_num).ScaleWidth
  lbldepg(grid_num).Caption = Format((maxd * feetormeter_factor), "###")
  lbltimeg(grid_num).Caption = Format(X(row_count), "###")
  lblzerodepg(grid_num).Visible = True
  lblzerotimeg(grid_num).Visible = True
  lblunitsg(grid_num).Visible = True
  lblminsg(grid_num).Visible = True
  lbldepg(grid_num).Visible = True
  lbltimeg(grid_num).Visible = True
  X(0) = 0
  Y(0) = 0
 
  decoresultgrid(grid_num).Col = 9
  For i = 1 To row_count
     decoresultgrid(grid_num).Row = i
     If InStr(1, decoresultgrid(grid_num).Text, "0") Then C = vbWhite
     If InStr(1, decoresultgrid(grid_num).Text, "1") Then C = vbYellow
     If InStr(1, decoresultgrid(grid_num).Text, "2") Then C = vbCyan
     If InStr(1, decoresultgrid(grid_num).Text, "3") Then C = &H808080
     If InStr(1, decoresultgrid(grid_num).Text, "4") Then C = vbGreen
     If InStr(1, decoresultgrid(grid_num).Text, "5") Then C = vbMagenta
     If InStr(1, decoresultgrid(grid_num).Text, "6") Then C = vbBlue
     If InStr(1, decoresultgrid(grid_num).Text, "7") Then C = &HFF8080
     If InStr(1, decoresultgrid(grid_num).Text, "8") Then C = &H8080FF
     If InStr(1, decoresultgrid(grid_num).Text, "9") Then C = vbRed
     Picture1(grid_num).Line (X(i - 1) * xscale, Y(i - 1) * yscale)-(X(i) * xscale, Y(i) * yscale), C
  Next i
  Label14(grid_num).Caption = Label14(grid_num).Caption + vbCrLf + "Max Depth: " + CStr(Fix(maxd * feetormeter_factor)) + feetormeter_shortstring
  
End Sub
Private Sub printgrid1()
    xPos = 550
    yPos = 1250
   Printer.CurrentX = xPos
   Printer.CurrentY = yPos
    For i = 0 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = i
        For j = 0 To MSFlexGrid1.Cols - 1
            MSFlexGrid1.Col = j
            Printer.CurrentX = xPos
            Printer.CurrentY = yPos
            Printer.Print MSFlexGrid1.Text

            If j = 0 Or j = 1 Or j = 2 Or j = 3 Or j = 4 Then
                xPos = xPos + 2000
            End If
        
        Next j
        xPos = 550
        yPos = yPos + 320
    Next i
    xPos = 550
    yPos = yPos + 320
    Printer.CurrentX = xPos
    Printer.CurrentY = yPos
    Printer.Print "" & Frame6.Caption
    Printer.Print ""
    printgaslist
    Printer.Print ""
    Printer.Print ""
    xPos = 550
    yPos = yPos + 320
    Printer.CurrentX = xPos
    Printer.CurrentY = yPos
    Printer.Print "" & Frame3.Caption
    Printer.Print ""
    printdecoresult

'   Printer.Print Spc(3);
'   Printer.Print temptext & temptext20 & temptext3
End Sub
Private Sub printgaslist()
yPos = yPos + 320
xPos = 550
Printer.CurrentX = xPos
Printer.CurrentY = yPos
Printer.Print "Gas Index"
xPos = 1850
Printer.CurrentX = xPos
Printer.CurrentY = yPos
 Printer.Print "O2"
 xPos = 3150
Printer.CurrentX = xPos
Printer.CurrentY = yPos
 Printer.Print "He"
 xPos = 4450
Printer.CurrentX = xPos
Printer.CurrentY = yPos
 Printer.Print "Depth"
 xPos = 5750
Printer.CurrentX = xPos
Printer.CurrentY = yPos
 Printer.Print "PPO2"
 xPos = 7050
Printer.CurrentX = xPos
Printer.CurrentY = yPos
 Printer.Print "Gas Used"
 
yPos = yPos + 320
 For v = 0 To 9
   temptext = lblgasindex(v).Caption
   temptext2 = lbl02(v).Caption
   temptext3 = lblhe(v).Caption
   temptext4 = txtmaxdft(v).Text 'lbldepth(v).Caption
   temptext5 = lblppo2(v).Caption
   temptext6 = lblgasused(v).Caption
   xPos = 550
   Printer.CurrentX = xPos
   Printer.CurrentY = yPos
   Printer.Print temptext
   xPos = 1850
    Printer.CurrentX = xPos
   Printer.CurrentY = yPos
   Printer.Print temptext2
   xPos = 3150
    Printer.CurrentX = xPos
   Printer.CurrentY = yPos
   Printer.Print temptext3
    xPos = 4450
    Printer.CurrentX = xPos
   Printer.CurrentY = yPos
   Printer.Print temptext4
    xPos = 5750
    Printer.CurrentX = xPos
   Printer.CurrentY = yPos
   Printer.Print temptext5
      xPos = 7050
    Printer.CurrentX = xPos
   Printer.CurrentY = yPos
   Printer.Print temptext6
    yPos = yPos + 320
    Next v

End Sub
Private Sub printdecoresult()
If systemversion = "Pro" Then
   j = rowindentified
   For v = 0 To decoresultgrid(j - 1).Rows - 1
      decoresultgrid(j - 1).Row = v
      decoresultgrid(j - 1).Col = 0
      temptext = decoresultgrid(j - 1).Text
      xPos = 550
      yPos = yPos + 320
      Printer.CurrentX = xPos
      Printer.CurrentY = yPos
      For i = 0 To decoresultgrid(j - 1).Rows - 1
         decoresultgrid(j - 1).Row = i
         For p = 0 To decoresultgrid(j - 1).Cols - 1
            decoresultgrid(j - 1).Col = p
            Printer.CurrentX = xPos
            Printer.CurrentY = yPos
            Printer.Print decoresultgrid(j - 1).Text
            If p < 9 Then
                xPos = xPos + 1300
            End If
         Next p
         xPos = 550
         yPos = yPos + 320
      Next i
   Next v
Else
  j = rowindentified
   For v = 0 To decoresultgridlite(j - 1).Rows - 1
      decoresultgridlite(j - 1).Row = v
      decoresultgridlite(j - 1).Col = 0
      temptext = decoresultgridlite(j - 1).Text
      xPos = 550
      yPos = yPos + 320
      Printer.CurrentX = xPos
      Printer.CurrentY = yPos
      For i = 0 To decoresultgridlite(j - 1).Rows - 1
         decoresultgridlite(j - 1).Row = i
         For p = 0 To decoresultgridlite(j - 1).Cols - 1
            decoresultgridlite(j - 1).Col = p
            Printer.CurrentX = xPos
            Printer.CurrentY = yPos
            Printer.Print decoresultgridlite(j - 1).Text
            If p < 9 Then
                xPos = xPos + 1300
            End If
         Next p
         xPos = 550
         yPos = yPos + 320
      Next i
   Next v
End If
End Sub


Private Sub Picture1_Click(index As Integer)
'  mnuplanedit_Click
End Sub

Private Sub Picture2_Click(index As Integer)
  mnupldelete_Click 'cmdadd_Click
End Sub

Private Sub Picture3_Click()
  cmdadd_Click
End Sub

Private Sub picture1_MouseDown(index As Integer, Button As Integer, Shift As Integer, Xmouse As Single, Ymouse As Single)
Dim i As Integer
Dim xc As Single
Dim timet As Single

  xc = Xmouse
  xc = xc / Picture1(grid_num).Width
  timet = xc * runtime_graph
  For i = 1 To row_count
    If X(i) > timet Then Exit For
  Next i
  If systemversion = "Pro" Then
     If decoresultgrid(index).Rows < 2 Then Exit Sub
  Else
     If decoresultgrid(index).Rows < 2 Then Exit Sub
  End If
  
  If deco_grid_display_last >= 0 Then
     If systemversion = "Pro" Then
        decoresultgrid(deco_grid_display_last).Row = deco_grid_display_rowlast
        decoresultgrid(deco_grid_display_last).Col = 0
        decoresultgrid(deco_grid_display_last).CellBackColor = deco_grid_display_celllast
     Else
      '   deco_grid_display_rowlast = (deco_grid_display_rowlast / 2) + 1
        decoresultgridlite(deco_grid_display_last).Row = deco_grid_display_rowlast
        decoresultgridlite(deco_grid_display_last).Col = 0
        decoresultgridlite(deco_grid_display_last).CellBackColor = deco_grid_display_celllast
     End If
  End If
  
  If i < 4 Then
     If systemversion = "Pro" Then
        decoresultgrid(deco_grid_display).TopRow = 1
     Else
        decoresultgridlite(deco_grid_display).TopRow = 1
     End If
  Else
     If systemversion = "Pro" Then
        decoresultgrid(deco_grid_display).TopRow = i - 2
        decoresultgrid(deco_grid_display).Col = 0
        decoresultgrid(deco_grid_display).Row = i
        deco_grid_display_last = (deco_grid_display)
        deco_grid_display_rowlast = (i)
        deco_grid_display_celllast = decoresultgrid(deco_grid_display).CellBackColor
        decoresultgrid(deco_grid_display).CellBackColor = vbBlue
     Else
        i = (i / 2) + 1
        decoresultgridlite(deco_grid_display).TopRow = i - 2
        decoresultgridlite(deco_grid_display).Col = 0
        decoresultgridlite(deco_grid_display).Row = i
        deco_grid_display_last = (deco_grid_display)
        deco_grid_display_rowlast = (i)
        deco_grid_display_celllast = decoresultgridlite(deco_grid_display).CellBackColor
        decoresultgridlite(deco_grid_display).CellBackColor = vbBlue
     End If
  End If
  xc = xc
  Text2(deco_grid_display).Text = ""
  Text2(deco_grid_display).Visible = True
  If systemversion = "Pro" Then
     Text2(deco_grid_display).Text = decoresultgrid(deco_grid_display).Text + " "
     decoresultgrid(deco_grid_display).Col = 1
     Text2(deco_grid_display).Text = Text2(deco_grid_display).Text + decoresultgrid(deco_grid_display).Text
  Else
     Text2(deco_grid_display).Text = decoresultgridlite(deco_grid_display).Text + " "
     decoresultgridlite(deco_grid_display).Col = 1
     Text2(deco_grid_display).Text = Text2(deco_grid_display).Text + decoresultgridlite(deco_grid_display).Text
  End If
  Text2(deco_grid_display).Top = Ymouse - Text2(deco_grid_display).Height
  Text2(deco_grid_display).Left = Xmouse
  Text2(deco_grid_display).Width = 780

End Sub

Private Sub Picture4_Click(index As Integer)
  Unload Me
End Sub

Private Sub Picture5_Click()
  mnueditasnew_Click
End Sub

Private Sub Picture6_Click()
  mnuplanedit_Click
End Sub

Private Sub Picture8_Click()
  mnupldelete_Click
End Sub

Private Sub safetytext_Change()
  If IsNumeric(safetytext.Text) Then
  Else
    safetytext.Text = "0"
  End If
End Sub

Private Sub safetytext_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    If CInt(safetytext.Text) < 0 Or CInt(safetytext.Text) > 50 Then
     MsgBox "Value must be between 0 and 50 !"
     safetytext.Text = "50"
    Else
         safetytext.SetFocus
         SendKeys "{HOME}+{END}"
         cmdgenerate_Click
    End If
     
   Else
    If KeyAscii <> 8 Then
      If KeyAscii < 46 Or KeyAscii > 57 Then
         MsgBox "Sorry, Only numeric characters allowed !"
         safetytext.SetFocus
         SendKeys "{HOME}+{END}"
      End If
    End If
   End If
End Sub

Private Sub safetytext_LostFocus()
  If CInt(safetytext.Text) < 0 Or CInt(safetytext.Text) > 50 Then
     MsgBox "Value must be between 0 and 50 !"
     safetytext.Text = "50"
  Else
    cmdgenerate_Click
  End If
End Sub

Public Function ppo2exposuretime(depth As Double, exposuretime As Double)
Dim i As Integer
Dim cnslooktable As Variant
Dim cnslooktable_1 As Variant
Dim otuadd As Double
Dim ppo2_now As Double
Dim absolutedepthpure As Double
'Dim exposuretime As Double
 
  cnslooktable = Array(0.3, -0.5, 0.595, 0#, 0.635, 0.14, 0.645, 0.15, 0.665, 0.16, 0.695, 0.17, 0.725, 0.18, 0.765, 0.2, 0.785, 0.21, 0.805, 0.22, 0.855, 0.24, 0.865, 0.25, 0.885, 0.26, 0.915, 0.28, 0.935, 0.29, 0.975, 0.31, 1.005, 0.33, 1.015, 0.34, 1.045, 0.35, 1.085, 0.4, 1.105, 0.42, 1.135, 0.43, 1.165, 0.45, 1.195, 0.47, 1.235, 0.5, 1.255, 0.51, 1.295, 0.55, 1.315, 0.56, 1.345, 0.6, 1.365, 0.61, 1.375, 0.62, 1.395, 0.64, 1.4, 0.65, 1.425, 0.68, 1.44, 0.71, 1.46, 0.74, 1.465, 0.76, 1.48, 0.78, 1.495, 0.81, 1.5, 0.83, 1.52, 0.93, 1.54, 1.04, 1.56, 1.19, 1.585, 1.47, 1.605, 2.22, 1.62, 5#, 1.65, 6.25, 1.67, 7.69, 1.7, 10#, 1.72, 12.5, 1.74, 20#, 1.77, 25#, 1.78, 31.25, 1.805, 50#, 2.5, 100#, 3, 999#)
  absolutedepthpure = (depth / 10#) + (barometric_vpressure / 10#)
  
  If (Is_CCR) Then
    ppo2_now = PO2_CCR / 10#
    
  Else
    ppo2_now = absolutedepthpure * (1# - vfraction_vnitrogen(vmix_vnumber) - vfraction_vhelium(vmix_vnumber))
  End If
  For i = 0 To 54
   If ppo2_now < cnslooktable(i * 2) Then Exit For
  Next i
  If i = 0 Then
    cns_current = cns_current / Exp((0.6935 / 90#) * exposuretime)
  Else
    cns_current = cns_current + (cnslooktable((i * 2) + 1) * exposuretime)
  End If
  If i > 54 Or cns_current > 999# Then cns_current = 999#
  If cns_current < 0# Then cns_current = 0#
  'If ppo2cnsmax < cns_current Then ppo2cnsmax = cns_current

  If (2# * absolutedepthpure * ((1# - vfraction_vnitrogen(vmix_vnumber) - vfraction_vhelium(vmix_vnumber))) - 1#) > 0 Then
'    otuadd = exposuretime * poww((2# * absolutedepth * (1# - nitrogenfraction - heliumfraction) - 1#), 0.833)
    otuadd = exposuretime * ((2# * absolutedepthpure * ((1# - vfraction_vnitrogen(vmix_vnumber) - vfraction_vhelium(vmix_vnumber))) - 1#))
    otu_current = otu_current + otuadd
  End If
  ppo2exposuretime = 0

End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
  If (SSTab1.Tab + 2) > MSFlexGrid1.Rows Then Exit Sub
  MSFlexGrid1.Row = SSTab1.Tab + 1
  MSFlexGrid1_Click
End Sub

Private Sub txtinterval_Change()
  cmdmodify.Visible = True
End Sub

Private Sub reloadgrid2()
Dim divelast As String

Label5(1).Caption = "These are singles dives. They assume no previous dive history. To add these dives into a mission sequence, use the Dive Series Planning features below."
divelast = "last"
MSFlexGrid2.Rows = 1
MSFlexGrid2.Cols = 8
MSFlexGrid2.Col = 0
MSFlexGrid2.Row = 0
MSFlexGrid2.Text = "Plan #"
MSFlexGrid2.Col = 1
MSFlexGrid2.Text = "MaxD"
MSFlexGrid2.Col = 2
MSFlexGrid2.Text = "Bottom Time"
MSFlexGrid2.Col = 3
MSFlexGrid2.Text = "Gas ID"
MSFlexGrid2.Col = 4
MSFlexGrid2.Text = "PPO2"
MSFlexGrid2.Col = 5
MSFlexGrid2.Text = "Open/Closed"
MSFlexGrid2.Col = 6
MSFlexGrid2.Text = "O2"
MSFlexGrid2.Col = 7
MSFlexGrid2.Text = "He"
MSFlexGrid2.ColWidth(0) = 1080
MSFlexGrid2.ColWidth(1) = 650
MSFlexGrid2.ColWidth(2) = 1250
MSFlexGrid2.ColWidth(3) = 0 '1040
MSFlexGrid2.ColWidth(4) = 750
MSFlexGrid2.ColWidth(5) = 1140
MSFlexGrid2.ColWidth(6) = 350
MSFlexGrid2.ColWidth(7) = 350
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
        MSFlexGrid2.Col = 1
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
        MSFlexGrid2.Col = 2
        MSFlexGrid2.Text = CStr(CInt(MSFlexGrid2.Text) + CInt(RS5("duration")))
     Else
        MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
        K = MSFlexGrid2.Rows
        MSFlexGrid2.Row = K - 1
        MSFlexGrid2.Col = 0
        If IsNull(tempdpid) Then Else MSFlexGrid2.Text = tempdpid
        MSFlexGrid2.Col = 1
        tempdepth = RS5("depth")
        If IsNumeric(tempdepth) = False Then
          tempdepth = "0.1"
        End If
        MSFlexGrid2.Text = Format(CDbl(tempdepth) * feetormeter_factor, "###0" & feetormeter_shortstring) + "   "
        MSFlexGrid2.Col = 2
        MSFlexGrid2.Text = RS5("duration")
        MSFlexGrid2.Col = 3
        MSFlexGrid2.Text = RS5("gasid")
        MSFlexGrid2.Col = 4
        MSFlexGrid2.Text = Format(RS5("po2"), "0.00") + "   "
        MSFlexGrid2.Col = 5
        MSFlexGrid2.Text = "   " + RS5("dpcircuit")
        MSFlexGrid2.Col = 6
        MSFlexGrid2.Text = RS5("dpo2") ' + "   "
        MSFlexGrid2.Col = 7
        MSFlexGrid2.Text = RS5("dphe") ' + "   "
        divelast = tempdpid
     End If
     RS5.MoveNext
  Wend
For K = 1 To MSFlexGrid2.Rows - 1
  For p = 0 To 0
    MSFlexGrid2.Row = K
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
         For H = 0 To 7
            MSFlexGrid2.Row = K
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
  MSFlexGrid2.Col = 2
  MSFlexGrid2.Text = MSFlexGrid2.Text + "mins   "
Next K
End If
rowidentified2 = MSFlexGrid2.Rows - 1
MSFlexGrid2_Click
End Sub
Private Sub view_graph_gaslist()
Exit Sub
  Picture1(grid_num).Visible = True
  Frame6.Visible = False
End Sub


