VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Planprofile2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "VGM ProPlanner"
   ClientHeight    =   10635
   ClientLeft      =   165
   ClientTop       =   -1080
   ClientWidth     =   13035
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Planprofile2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Palette         =   "Planprofile2.frx":2CFA
   PaletteMode     =   2  'Custom
   ScaleHeight     =   10635
   ScaleWidth      =   13035
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   4215
      Left            =   240
      TabIndex        =   294
      Top             =   4920
      Width           =   11895
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   0
         Picture         =   "Planprofile2.frx":5C1C
         ScaleHeight     =   4185
         ScaleWidth      =   11865
         TabIndex        =   295
         Top             =   0
         Width           =   11895
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   2280
      Top             =   4440
   End
   Begin VB.TextBox Textl 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4800
      TabIndex        =   261
      TabStop         =   0   'False
      Text            =   "%"
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Textl 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   260
      TabStop         =   0   'False
      Text            =   "Safety :"
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox safetytext 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4515
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Textl 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   10920
      TabIndex        =   259
      TabStop         =   0   'False
      Text            =   "mBar"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox atmtext 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   10215
      MaxLength       =   4
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   360
      Width           =   600
   End
   Begin VB.TextBox Textl 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   9240
      TabIndex        =   143
      TabStop         =   0   'False
      Text            =   "Atmospheric:"
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      TabIndex        =   275
      Top             =   3480
      Width           =   4575
      Begin VB.CommandButton vhmx_up 
         BackColor       =   &H00FFFFFF&
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
         Index           =   4
         Left            =   3120
         Picture         =   "Planprofile2.frx":16C3FE
         Style           =   1  'Graphical
         TabIndex        =   292
         Top             =   660
         Width           =   220
      End
      Begin VB.CommandButton vhmx_down 
         BackColor       =   &H00FFFFFF&
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
         Index           =   4
         Left            =   3360
         Picture         =   "Planprofile2.frx":16C5F8
         Style           =   1  'Graphical
         TabIndex        =   291
         Top             =   660
         Width           =   220
      End
      Begin VB.TextBox vhmx_text 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   4
         TabIndex        =   287
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Fast Tissue Safety"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox vhmx_text 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   1
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   286
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Mid Tissue Safety"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox vhmx_text 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   2
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   285
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Slow Tissue Safety"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox vhmx_text 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Index           =   3
         Left            =   0
         MaxLength       =   4
         TabIndex        =   284
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton vhmx_down 
         BackColor       =   &H00FFFFFF&
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
         Index           =   0
         Left            =   1230
         Picture         =   "Planprofile2.frx":16C7A2
         Style           =   1  'Graphical
         TabIndex        =   283
         Top             =   660
         Width           =   220
      End
      Begin VB.CommandButton vhmx_up 
         BackColor       =   &H00FFFFFF&
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
         Index           =   0
         Left            =   990
         Picture         =   "Planprofile2.frx":16C94C
         Style           =   1  'Graphical
         TabIndex        =   282
         Top             =   660
         Width           =   220
      End
      Begin VB.CommandButton vhmx_down 
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   1950
         Picture         =   "Planprofile2.frx":16CB46
         Style           =   1  'Graphical
         TabIndex        =   281
         Top             =   660
         Width           =   220
      End
      Begin VB.CommandButton vhmx_up 
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   1710
         Picture         =   "Planprofile2.frx":16CCF0
         Style           =   1  'Graphical
         TabIndex        =   280
         Top             =   660
         Width           =   220
      End
      Begin VB.CommandButton vhmx_down 
         BackColor       =   &H00FFFFFF&
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
         Index           =   2
         Left            =   2670
         Picture         =   "Planprofile2.frx":16CEEA
         Style           =   1  'Graphical
         TabIndex        =   279
         Top             =   660
         Width           =   220
      End
      Begin VB.CommandButton vhmx_up 
         BackColor       =   &H00FFFFFF&
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
         Index           =   2
         Left            =   2430
         Picture         =   "Planprofile2.frx":16D094
         Style           =   1  'Graphical
         TabIndex        =   278
         Top             =   660
         Width           =   220
      End
      Begin VB.CommandButton vhmx_down 
         BackColor       =   &H00FFFFFF&
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
         Index           =   3
         Left            =   0
         Picture         =   "Planprofile2.frx":16D28E
         Style           =   1  'Graphical
         TabIndex        =   277
         Top             =   840
         Visible         =   0   'False
         Width           =   220
      End
      Begin VB.CommandButton vhmx_up 
         BackColor       =   &H00FFFFFF&
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
         Index           =   3
         Left            =   0
         Picture         =   "Planprofile2.frx":16D438
         Style           =   1  'Graphical
         TabIndex        =   276
         Top             =   840
         Visible         =   0   'False
         Width           =   220
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " ALL"
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
         Height          =   360
         Index           =   2
         Left            =   3120
         TabIndex        =   293
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EGF"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   3840
         TabIndex        =   290
         ToolTipText     =   "Equivalent Gradient Factor"
         Top             =   120
         Width           =   615
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   855
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   3600
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " VGM Bubble Control"
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
         Height          =   720
         Index           =   1
         Left            =   240
         TabIndex        =   289
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " Fast          Mid          Slow"
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
         Height          =   240
         Index           =   0
         Left            =   960
         TabIndex        =   288
         Top             =   135
         Width           =   2415
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   975
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   4575
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   600
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Gas Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      TabIndex        =   46
      Top             =   360
      Width           =   6240
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   305
         Top             =   3050
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   304
         Top             =   2765
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   303
         Top             =   2485
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   302
         Top             =   2215
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   301
         Top             =   1905
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   300
         Top             =   1615
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   299
         Top             =   1340
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   298
         Top             =   1020
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   297
         Top             =   740
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   296
         Top             =   455
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.CommandButton cmdminus 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   5760
         Picture         =   "Planprofile2.frx":16D632
         Style           =   1  'Graphical
         TabIndex        =   223
         ToolTipText     =   "Decrease the gas O2 level"
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdplus 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Planprofile2.frx":16D974
         Style           =   1  'Graphical
         TabIndex        =   222
         ToolTipText     =   "Increase the gas O2 level"
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdminus 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   5760
         Picture         =   "Planprofile2.frx":16DCB6
         Style           =   1  'Graphical
         TabIndex        =   221
         ToolTipText     =   "Decrease the gas He level"
         Top             =   3120
         Width           =   255
      End
      Begin VB.CommandButton cmdplus 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   5520
         Picture         =   "Planprofile2.frx":16DFF8
         Style           =   1  'Graphical
         TabIndex        =   220
         ToolTipText     =   "Increase the gas He level"
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   9
         Left            =   1750
         TabIndex        =   98
         Text            =   "Text1"
         Top             =   3050
         Width           =   550
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   8
         Left            =   1750
         TabIndex        =   99
         Text            =   "Text1"
         Top             =   2765
         Width           =   550
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   9
         Left            =   4450
         TabIndex        =   219
         Top             =   3050
         Width           =   380
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   8
         Left            =   4450
         TabIndex        =   218
         Top             =   2765
         Width           =   380
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   7
         Left            =   4450
         TabIndex        =   217
         Top             =   2485
         Width           =   380
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   6
         Left            =   4450
         TabIndex        =   216
         Top             =   2215
         Width           =   380
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   5
         Left            =   4450
         TabIndex        =   215
         Top             =   1905
         Width           =   380
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   4
         Left            =   4450
         TabIndex        =   214
         Top             =   1615
         Width           =   380
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   3
         Left            =   4450
         TabIndex        =   213
         Top             =   1340
         Width           =   380
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   2
         Left            =   4450
         TabIndex        =   212
         Top             =   1020
         Width           =   380
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   1
         Left            =   4450
         TabIndex        =   211
         Top             =   740
         Width           =   380
      End
      Begin VB.TextBox txtbreathrate2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   0
         Left            =   4450
         TabIndex        =   210
         Top             =   450
         Width           =   380
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   9
         Left            =   4150
         TabIndex        =   208
         Top             =   3050
         Width           =   320
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   8
         Left            =   4150
         TabIndex        =   207
         Top             =   2765
         Width           =   320
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   7
         Left            =   4150
         TabIndex        =   206
         Top             =   2485
         Width           =   320
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   6
         Left            =   4150
         TabIndex        =   205
         Top             =   2215
         Width           =   320
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   5
         Left            =   4150
         TabIndex        =   204
         Top             =   1905
         Width           =   320
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   4
         Left            =   4150
         TabIndex        =   203
         Top             =   1615
         Width           =   320
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   3
         Left            =   4150
         TabIndex        =   202
         Top             =   1340
         Width           =   320
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   2
         Left            =   4150
         TabIndex        =   201
         Top             =   1020
         Width           =   320
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   1
         Left            =   4150
         TabIndex        =   200
         Top             =   740
         Width           =   320
      End
      Begin VB.TextBox txtcylcap2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   0
         Left            =   4150
         TabIndex        =   199
         Top             =   450
         Width           =   320
      End
      Begin VB.TextBox txtcylcap 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   5040
         TabIndex        =   198
         Text            =   "10"
         Top             =   3720
         Width           =   615
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   196
         Top             =   3050
         Width           =   255
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   195
         Top             =   2765
         Width           =   255
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   194
         Top             =   2485
         Width           =   255
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   193
         Top             =   2215
         Width           =   255
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   192
         Top             =   1905
         Width           =   255
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   191
         Top             =   1615
         Width           =   255
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   190
         Top             =   1340
         Width           =   255
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   189
         Top             =   1020
         Width           =   255
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   188
         Top             =   740
         Width           =   255
      End
      Begin VB.CheckBox Decochk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   187
         Top             =   455
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   3480
         TabIndex        =   184
         Top             =   3050
         Width           =   340
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   3480
         TabIndex        =   183
         Top             =   2765
         Width           =   340
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   182
         Top             =   2485
         Width           =   340
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   181
         Top             =   2215
         Width           =   340
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   180
         Top             =   1905
         Width           =   340
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   179
         Top             =   1615
         Width           =   340
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   178
         Top             =   1340
         Width           =   340
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   177
         Top             =   1020
         Width           =   340
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   176
         Top             =   740
         Width           =   340
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   174
         Top             =   455
         Width           =   225
      End
      Begin VB.TextBox txtbreathrate 
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
         Height          =   255
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   141
         Text            =   "2"
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txtbreathratecuft 
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
         Height          =   255
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   140
         Text            =   "2"
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   1
         Left            =   1750
         TabIndex        =   106
         Text            =   "Text1"
         Top             =   740
         Width           =   550
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   2
         Left            =   1750
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   1020
         Width           =   550
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   3
         Left            =   1750
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   1340
         Width           =   550
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   4
         Left            =   1750
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   1615
         Width           =   550
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   5
         Left            =   1750
         TabIndex        =   102
         Text            =   "Text1"
         Top             =   1905
         Width           =   550
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   6
         Left            =   1750
         TabIndex        =   101
         Text            =   "Text1"
         Top             =   2215
         Width           =   550
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   7
         Left            =   1750
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   2485
         Width           =   550
      End
      Begin VB.TextBox txtmaxdft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   0
         Left            =   1750
         TabIndex        =   97
         Text            =   "Text1"
         Top             =   455
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
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
         Height          =   300
         Index           =   9
         Left            =   2810
         TabIndex        =   96
         Text            =   " "
         Top             =   3050
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
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
         Height          =   300
         Index           =   8
         Left            =   2810
         TabIndex        =   95
         Top             =   2765
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
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
         Height          =   300
         Index           =   7
         Left            =   2810
         TabIndex        =   94
         Top             =   2485
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
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
         Height          =   300
         Index           =   6
         Left            =   2810
         TabIndex        =   93
         Top             =   2215
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
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
         Height          =   300
         Index           =   5
         Left            =   2810
         TabIndex        =   92
         Top             =   1905
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
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
         Height          =   300
         Index           =   4
         Left            =   2810
         TabIndex        =   91
         Top             =   1615
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
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
         Height          =   300
         Index           =   3
         Left            =   2810
         TabIndex        =   90
         Top             =   1340
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
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
         Height          =   300
         Index           =   2
         Left            =   2810
         TabIndex        =   89
         Top             =   1020
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
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
         Height          =   300
         Index           =   1
         Left            =   2810
         TabIndex        =   88
         Top             =   740
         Width           =   550
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   0
         Left            =   2810
         TabIndex        =   87
         Top             =   455
         Width           =   550
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   9
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   3050
         Width           =   420
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   9
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   3050
         Width           =   400
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   9
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   6375
         Width           =   2025
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   8
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   2765
         Width           =   420
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   8
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   2765
         Width           =   400
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   8
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   6060
         Width           =   2025
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   7
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   2485
         Width           =   420
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   7
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   2485
         Width           =   400
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   7
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   5745
         Width           =   2025
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   6
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   2215
         Width           =   420
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   6
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   2215
         Width           =   400
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   6
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   5400
         Width           =   2025
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   5
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   1905
         Width           =   420
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   5
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   1905
         Width           =   400
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   5
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   5115
         Width           =   2025
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   4
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   4800
         Width           =   2025
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   3
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   4485
         Width           =   2025
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   4
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   1615
         Width           =   420
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   4
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1615
         Width           =   400
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   2
         ItemData        =   "Planprofile2.frx":16E33A
         Left            =   3975
         List            =   "Planprofile2.frx":16E33C
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   4170
         Width           =   2025
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   3
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1340
         Width           =   420
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   3
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1340
         Width           =   400
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   1
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   3855
         Width           =   2025
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   2
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1020
         Width           =   420
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   2
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1020
         Width           =   400
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   1
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   740
         Width           =   420
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   1
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   740
         Width           =   400
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   0
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "21"
         Top             =   455
         Width           =   400
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   0
         Left            =   1330
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   455
         Width           =   420
      End
      Begin VB.ComboBox Cbogasused 
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
         Height          =   315
         Index           =   0
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   3600
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   290
         Index           =   0
         Left            =   1750
         TabIndex        =   56
         Top             =   455
         Width           =   550
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   9
         Left            =   1750
         TabIndex        =   55
         Top             =   3050
         Width           =   550
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   8
         Left            =   1750
         TabIndex        =   54
         Top             =   2765
         Width           =   550
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   7
         Left            =   1750
         TabIndex        =   53
         Top             =   2485
         Width           =   550
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   6
         Left            =   1750
         TabIndex        =   52
         Top             =   2215
         Width           =   550
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   5
         Left            =   1750
         TabIndex        =   51
         Top             =   1905
         Width           =   550
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   4
         Left            =   1750
         TabIndex        =   50
         Top             =   1615
         Width           =   550
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   290
         Index           =   3
         Left            =   1750
         TabIndex        =   49
         Top             =   1340
         Width           =   550
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Index           =   2
         Left            =   1750
         TabIndex        =   48
         Top             =   1020
         Width           =   550
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   290
         Index           =   1
         Left            =   1750
         TabIndex        =   47
         Top             =   740
         Width           =   550
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   140
         TabIndex        =   171
         Top             =   3050
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   140
         TabIndex        =   170
         Top             =   2765
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   140
         TabIndex        =   169
         Top             =   2485
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   140
         TabIndex        =   168
         Top             =   2215
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   140
         TabIndex        =   167
         Top             =   1905
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   140
         TabIndex        =   166
         Top             =   1615
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   140
         TabIndex        =   165
         Top             =   1340
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   140
         TabIndex        =   164
         Top             =   1020
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   140
         TabIndex        =   163
         Top             =   740
         Width           =   300
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   140
         TabIndex        =   162
         Top             =   455
         Width           =   300
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   3960
         TabIndex        =   245
         Top             =   1020
         Width           =   180
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   3960
         TabIndex        =   244
         Top             =   740
         Width           =   180
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   9
         Left            =   420
         TabIndex        =   120
         Top             =   3050
         Width           =   540
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   8
         Left            =   420
         TabIndex        =   119
         Top             =   2765
         Width           =   540
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   7
         Left            =   420
         TabIndex        =   118
         Top             =   2485
         Width           =   540
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   6
         Left            =   420
         TabIndex        =   117
         Top             =   2215
         Width           =   540
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   5
         Left            =   420
         TabIndex        =   116
         Top             =   1905
         Width           =   540
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   4
         Left            =   420
         TabIndex        =   115
         Top             =   1615
         Width           =   540
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   3
         Left            =   420
         TabIndex        =   114
         Top             =   1340
         Width           =   540
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   2
         Left            =   420
         TabIndex        =   113
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   1
         Left            =   420
         TabIndex        =   112
         Top             =   740
         Width           =   540
      End
      Begin VB.Label lblgasindex 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   0
         Left            =   420
         TabIndex        =   111
         Top             =   455
         Width           =   540
      End
      Begin VB.Label Lblhel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   5520
         TabIndex        =   224
         Top             =   3375
         Width           =   495
      End
      Begin VB.Label Lblo2g 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   5520
         TabIndex        =   225
         Top             =   165
         Width           =   495
      End
      Begin VB.Shape Shape10 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   1
         Left            =   5400
         Shape           =   4  'Rounded Rectangle
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   0
         Left            =   5610
         Top             =   675
         Width           =   300
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   2
         Left            =   5610
         Top             =   2355
         Width           =   300
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   2400
         Index           =   1
         Left            =   5610
         Top             =   675
         Width           =   300
      End
      Begin VB.Shape Shape10 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   0
         Left            =   5400
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   300
         Index           =   0
         Left            =   2280
         TabIndex        =   272
         Top             =   455
         Width           =   540
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   300
         Index           =   1
         Left            =   2280
         TabIndex        =   271
         Top             =   740
         Width           =   540
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   290
         Index           =   2
         Left            =   2280
         TabIndex        =   270
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   300
         Index           =   3
         Left            =   2280
         TabIndex        =   269
         Top             =   1340
         Width           =   540
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   300
         Index           =   4
         Left            =   2280
         TabIndex        =   268
         Top             =   1615
         Width           =   540
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   290
         Index           =   5
         Left            =   2280
         TabIndex        =   267
         Top             =   1905
         Width           =   540
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   300
         Index           =   6
         Left            =   2280
         TabIndex        =   266
         Top             =   2215
         Width           =   540
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   300
         Index           =   7
         Left            =   2280
         TabIndex        =   265
         Top             =   2485
         Width           =   540
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   300
         Index           =   8
         Left            =   2280
         TabIndex        =   264
         Top             =   2765
         Width           =   540
      End
      Begin VB.Label lblEan 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "EAN"
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
         Height          =   290
         Index           =   9
         Left            =   2280
         TabIndex        =   263
         Top             =   3050
         Width           =   540
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "EAND"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   2320
         TabIndex        =   262
         ToolTipText     =   "Equivalent Air Narcosis Depth"
         Top             =   180
         Width           =   460
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   0
         Left            =   3960
         TabIndex        =   255
         Top             =   450
         Width           =   180
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   9
         Left            =   3960
         TabIndex        =   252
         Top             =   3045
         Width           =   180
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   8
         Left            =   3960
         TabIndex        =   251
         Top             =   2765
         Width           =   180
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   7
         Left            =   3960
         TabIndex        =   250
         Top             =   2490
         Width           =   180
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   6
         Left            =   3960
         TabIndex        =   249
         Top             =   2220
         Width           =   180
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   3960
         TabIndex        =   248
         Top             =   1905
         Width           =   180
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   4
         Left            =   3960
         TabIndex        =   247
         Top             =   1620
         Width           =   180
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   3960
         TabIndex        =   246
         Top             =   1335
         Width           =   180
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   9
         Left            =   3360
         TabIndex        =   243
         Top             =   3045
         Width           =   600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   8
         Left            =   3360
         TabIndex        =   242
         Top             =   2765
         Width           =   600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   7
         Left            =   3360
         TabIndex        =   241
         Top             =   2490
         Width           =   600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   6
         Left            =   3360
         TabIndex        =   240
         Top             =   2220
         Width           =   600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   3360
         TabIndex        =   239
         Top             =   1905
         Width           =   600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   4
         Left            =   3360
         TabIndex        =   238
         Top             =   1620
         Width           =   600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   3360
         TabIndex        =   237
         Top             =   1335
         Width           =   600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   3360
         TabIndex        =   236
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   3360
         TabIndex        =   235
         Top             =   735
         Width           =   600
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   9
         Left            =   60
         TabIndex        =   234
         Top             =   3050
         Width           =   360
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   8
         Left            =   60
         TabIndex        =   233
         Top             =   2765
         Width           =   360
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   7
         Left            =   60
         TabIndex        =   232
         Top             =   2485
         Width           =   360
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   6
         Left            =   60
         TabIndex        =   231
         Top             =   2215
         Width           =   360
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   60
         TabIndex        =   230
         Top             =   1905
         Width           =   360
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   4
         Left            =   60
         TabIndex        =   229
         Top             =   1615
         Width           =   360
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   60
         TabIndex        =   228
         Top             =   1340
         Width           =   360
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   60
         TabIndex        =   227
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   60
         TabIndex        =   226
         Top             =   740
         Width           =   360
      End
      Begin VB.Shape Shape10 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   2
         Left            =   5400
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "SAC"
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
         Height          =   225
         Left            =   4450
         TabIndex        =   209
         ToolTipText     =   "Surface Air Comsumption"
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "WC"
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
         Height          =   225
         Left            =   4150
         TabIndex        =   197
         ToolTipText     =   "Water Capacity of cylinder"
         Top             =   180
         Width           =   280
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Deco"
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
         Height          =   225
         Left            =   3720
         TabIndex        =   186
         ToolTipText     =   "Check boxes for gases that are only used for decompression"
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "CC"
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
         Height          =   225
         Left            =   3390
         TabIndex        =   185
         ToolTipText     =   "Check boxes for diluent gasses used in a closed circuit rebreather"
         Top             =   180
         Width           =   315
      End
      Begin VB.Line Line4 
         X1              =   2880
         X2              =   4680
         Y1              =   2200
         Y2              =   2200
      End
      Begin VB.Line Line3 
         X1              =   2880
         X2              =   4680
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   480
         Y1              =   2200
         Y2              =   2200
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   480
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   0
         Left            =   3360
         TabIndex        =   175
         Top             =   450
         Width           =   600
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bottom"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   60
         TabIndex        =   173
         ToolTipText     =   "Set gasses that are active for this dive"
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   172
         Top             =   450
         Width           =   360
      End
      Begin VB.Label lblcylsize 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "l/min"
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
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   142
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblcylsize 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Used"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   4830
         TabIndex        =   139
         ToolTipText     =   "Gas Bar/PSI used during dive"
         Top             =   180
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   4830
         TabIndex        =   138
         Top             =   450
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   4830
         TabIndex        =   137
         Top             =   735
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   4830
         TabIndex        =   136
         Top             =   1020
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   4830
         TabIndex        =   135
         Top             =   1335
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   4
         Left            =   4830
         TabIndex        =   134
         Top             =   1615
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   5
         Left            =   4830
         TabIndex        =   133
         Top             =   1905
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   6
         Left            =   4830
         TabIndex        =   132
         Top             =   2220
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   7
         Left            =   4830
         TabIndex        =   131
         Top             =   2490
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   8
         Left            =   4830
         TabIndex        =   130
         Top             =   2765
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label gasusage 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   9
         Left            =   4830
         TabIndex        =   129
         Top             =   3045
         Visible         =   0   'False
         Width           =   405
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
         Height          =   225
         Left            =   2810
         TabIndex        =   121
         ToolTipText     =   "Partial Pressure of Oxygen"
         Top             =   180
         Width           =   565
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "O2"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   945
         TabIndex        =   110
         ToolTipText     =   "Oxygen Level"
         Top             =   180
         Width           =   390
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "He"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1350
         TabIndex        =   109
         ToolTipText     =   "Helium level"
         Top             =   180
         Width           =   390
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "MOD"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   1760
         TabIndex        =   108
         ToolTipText     =   "Maximum Operating Depth"
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gas #"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   495
         TabIndex        =   107
         ToolTipText     =   "Gas reference nu,ber"
         Top             =   180
         Width           =   435
      End
   End
   Begin MSFlexGridLib.MSFlexGrid decoresultgridlite 
      Height          =   3615
      Left            =   6480
      TabIndex        =   158
      ToolTipText     =   "Decompression Reuslt"
      Top             =   5280
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   6376
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16761024
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
   Begin VB.ComboBox cbogasindex 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   157
      Text            =   "Gas Index"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtdecoalg 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   145
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtserialno 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00808000&
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   4440
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   120
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   240
      TabIndex        =   122
      Top             =   4920
      Visible         =   0   'False
      Width           =   11880
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   360
         TabIndex        =   146
         Top             =   360
         Width           =   5415
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808000&
            FillColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   3135
            Left            =   360
            ScaleHeight     =   3105
            ScaleWidth      =   4785
            TabIndex        =   147
            Top             =   240
            Width           =   4815
            Begin VB.TextBox Text1 
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
               Height          =   255
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   148
               Text            =   "Text1"
               Top             =   1920
               Visible         =   0   'False
               Width           =   495
            End
            Begin MSComDlg.CommonDialog cmdlog 
               Left            =   5040
               Top             =   2880
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
               Filter          =   "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
            End
         End
         Begin VB.Label lblx 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Index           =   3
            Left            =   5040
            TabIndex        =   156
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label lblx 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   155
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label lblx 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   154
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label lblx 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   153
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label lbly 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   152
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label lbly 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   151
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label lbly 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   150
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lbly 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   149
            Top             =   120
            Width           =   375
         End
      End
      Begin MSFlexGridLib.MSFlexGrid decoresultgrid 
         Height          =   3615
         Left            =   6240
         TabIndex        =   124
         Top             =   360
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   16761024
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
      Begin VB.TextBox Text10 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   6360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   123
         Text            =   "Planprofile2.frx":16E33E
         Top             =   1560
         Visible         =   0   'False
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4560
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8265
      Left            =   3000
      TabIndex        =   16
      Top             =   10440
      Width           =   12015
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Data Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4635
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   5745
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00C0C0C0&
            Height          =   2895
            Left            =   1080
            Picture         =   "Planprofile2.frx":16E347
            ScaleHeight     =   2835
            ScaleWidth      =   3195
            TabIndex        =   22
            Top             =   960
            Width           =   3255
            Begin VB.CommandButton cmdsaveas 
               BackColor       =   &H00FFFF00&
               Caption         =   "Copy"
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
               Left            =   4440
               MaskColor       =   &H00000000&
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   3720
               Width           =   975
            End
            Begin VB.Label lblhelium 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   2760
               TabIndex        =   24
               Top             =   2880
               Width           =   1215
            End
            Begin VB.Label lblo2 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   2760
               TabIndex        =   23
               Top             =   3240
               Width           =   1095
            End
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Time :"
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
            Left            =   360
            TabIndex        =   21
            Top             =   4830
            Width           =   615
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0FFFF&
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
            Left            =   2160
            TabIndex        =   20
            Top             =   4920
            Width           =   615
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0FFFF&
            Caption         =   "PPO2 :"
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
            Left            =   270
            TabIndex        =   19
            Top             =   5460
            Width           =   615
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0FFFF&
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
            Left            =   2160
            TabIndex        =   18
            Top             =   5520
            Width           =   420
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   6960
   End
   Begin MSComDlg.CommonDialog dlgchart 
      Left            =   5400
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1560
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   10
      RTSEnable       =   -1  'True
   End
   Begin VB.CommandButton cmdaddtoseq 
      Caption         =   "Add To Sequential"
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
      Left            =   8160
      TabIndex        =   15
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
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
      Left            =   7560
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
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
      Left            =   9120
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
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
      Left            =   3600
      TabIndex        =   10
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   9840
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   6240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   125
      Text            =   "Planprofile2.frx":187601
      Top             =   5520
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   6600
      ScaleHeight     =   4215
      ScaleWidth      =   5775
      TabIndex        =   35
      Top             =   240
      Width           =   5775
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         Picture         =   "Planprofile2.frx":187607
         Style           =   1  'Graphical
         TabIndex        =   274
         ToolTipText     =   "Decrease the depth point"
         Top             =   2085
         Width           =   220
      End
      Begin VB.CommandButton CMDPPO2PLUS 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2280
         Picture         =   "Planprofile2.frx":1877B1
         Style           =   1  'Graphical
         TabIndex        =   273
         ToolTipText     =   "Increase the depth point"
         Top             =   2085
         Width           =   220
      End
      Begin VB.CommandButton cmdgeneratem 
         BackColor       =   &H0080FF80&
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   515
         Left            =   1620
         MaskColor       =   &H00C0C000&
         Style           =   1  'Graphical
         TabIndex        =   254
         ToolTipText     =   "Generate decompression result"
         Top             =   2800
         Width           =   885
      End
      Begin VB.CommandButton cmdgenerate 
         BackColor       =   &H0000FF00&
         Caption         =   "Calculate Deco"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Generate decompression result"
         Top             =   1320
         Width           =   1050
      End
      Begin VB.CommandButton cmdtimedown 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3840
         Picture         =   "Planprofile2.frx":1879AB
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Decrease the duration"
         Top             =   1080
         Width           =   220
      End
      Begin VB.CommandButton cmdtimeup 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3600
         Picture         =   "Planprofile2.frx":187B55
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Increase the duration"
         Top             =   1080
         Width           =   220
      End
      Begin VB.CommandButton cmddepthdown 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         Picture         =   "Planprofile2.frx":187D4F
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Decrease the depth point"
         Top             =   1080
         Width           =   220
      End
      Begin VB.CommandButton cmddepthup 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1320
         Picture         =   "Planprofile2.frx":187EF9
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Increase the depth point"
         Top             =   1080
         Width           =   220
      End
      Begin VB.TextBox txtdepth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   960
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "m"
         Top             =   1320
         Width           =   1275
      End
      Begin VB.TextBox txttime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3360
         TabIndex        =   1
         Text            =   "10"
         Top             =   1320
         Width           =   1275
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Closed"
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
         Left            =   3240
         MaskColor       =   &H00FFFF80&
         TabIndex        =   39
         Top             =   2560
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Open"
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
         Left            =   3240
         MaskColor       =   &H00FFFF80&
         TabIndex        =   38
         Top             =   2280
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtppo2v 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2160
         TabIndex        =   4
         Top             =   2325
         Width           =   1275
      End
      Begin VB.CommandButton cmdremove 
         BackColor       =   &H000000FF&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   2640
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Delete selected level profile"
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdclearall 
         BackColor       =   &H000000FF&
         Caption         =   "Clear All"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   3400
         MouseIcon       =   "Planprofile2.frx":1880F3
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Clear all level of the profile"
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Cmdadd 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1620
         Picture         =   "Planprofile2.frx":1883FD
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add new depth to end of list"
         Top             =   600
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdmodify 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1620
         Picture         =   "Planprofile2.frx":189BDF
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "modified the selected level"
         Top             =   2000
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmdinsert 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1620
         Picture         =   "Planprofile2.frx":18B3C1
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Insert depth before selected level"
         Top             =   1360
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtdepthft 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   960
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "ft"
         Top             =   1320
         Width           =   1275
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   2235
         Left            =   2760
         TabIndex        =   44
         Top             =   2040
         Visible         =   0   'False
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   3942
         _Version        =   393216
         Rows            =   50
         Cols            =   7
         FixedCols       =   0
         BackColor       =   12632256
         GridColor       =   14737632
         GridLines       =   0
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
      Begin VB.Label Label38 
         BackColor       =   &H00400000&
         Caption         =   "ppo2"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   960
         TabIndex        =   258
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label37 
         BackColor       =   &H00400000&
         Caption         =   "mins"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   960
         TabIndex        =   257
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label36 
         BackColor       =   &H00400000&
         Caption         =   "meter"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   325
         Left            =   960
         TabIndex        =   256
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Dive Plan : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   253
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label singlelevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Single level ?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   4320
         MouseIcon       =   "Planprofile2.frx":18CBA3
         MousePointer    =   99  'Custom
         TabIndex        =   160
         ToolTipText     =   "Return to the Single level windows"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbllevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Multi level? Travel Gas?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   4800
         MouseIcon       =   "Planprofile2.frx":18CEAD
         MousePointer    =   99  'Custom
         TabIndex        =   159
         ToolTipText     =   "Go to multi level construction screen"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblminutes 
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
         Left            =   4200
         TabIndex        =   128
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   2805
         TabIndex        =   127
         Top             =   2080
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "meter"
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
         Left            =   1920
         TabIndex        =   126
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblgasvr 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "TX10/45"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   45
         Top             =   735
         Width           =   1215
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   4215
         Left            =   0
         Picture         =   "Planprofile2.frx":18D1B7
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5775
      End
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   2655
      Index           =   3
      Left            =   200
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   6075
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   5775
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5775
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   2295
      Index           =   4
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFC0&
      Height          =   3855
      Index           =   1
      Left            =   5760
      Top             =   360
      Width           =   6375
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sequence: "
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
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   10
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   8400
      Width           =   1875
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   9
      Left            =   10560
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   1875
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   8
      Left            =   10560
      Shape           =   4  'Rounded Rectangle
      Top             =   8400
      Width           =   1875
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   7
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   1875
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   4455
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   12315
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "O2 :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   31
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "He :"
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
      Left            =   5760
      TabIndex        =   30
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblseqdiveno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Left            =   960
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
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
      Left            =   14400
      TabIndex        =   14
      Top             =   6240
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
      Left            =   14400
      TabIndex        =   13
      Top             =   5760
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
      Left            =   14880
      TabIndex        =   12
      Top             =   4920
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
      Left            =   7080
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gas Index :"
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
      Left            =   5880
      TabIndex        =   33
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Depth :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   32
      Top             =   3720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Index           =   2
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   6300
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Index           =   0
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   4455
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   12315
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   5
      Left            =   10440
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1995
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   11
      Left            =   10440
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   1995
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   2
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1995
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   6
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   1995
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilesave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnufilesaveas 
         Caption         =   "Save as &CSV"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnugas 
      Caption         =   "G&as"
      Begin VB.Menu mnugassetdefault 
         Caption         =   "&Set As Default"
      End
      Begin VB.Menu mnuloaddesetting 
         Caption         =   "&Load Default Setting"
      End
      Begin VB.Menu mnugasloadefault 
         Caption         =   "Load &Factory Setting"
      End
   End
   Begin VB.Menu mnuGenerate 
      Caption         =   "&Generate"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnugraph 
         Caption         =   "&Graph"
      End
      Begin VB.Menu mnugaslist 
         Caption         =   "Gas &List"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuVPMBdef 
      Caption         =   "&VPMB/Buhl"
      Enabled         =   0   'False
      Begin VB.Menu mnuVPMB 
         Caption         =   "VPMB only"
         Index           =   0
      End
      Begin VB.Menu mnuVPMB 
         Caption         =   "VPMB+Buhl"
         Index           =   1
      End
      Begin VB.Menu mnuVPMB 
         Caption         =   "VGM Alg"
         Index           =   2
      End
   End
   Begin VB.Menu mnudecoversion 
      Caption         =   "&Schedule"
      Begin VB.Menu mnuprofessional 
         Caption         =   "&Professional"
      End
      Begin VB.Menu mnulite 
         Caption         =   "&Lite"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuLastStop 
      Caption         =   "LastStop"
      Begin VB.Menu mnuStep 
         Caption         =   "Step"
      End
      Begin VB.Menu mnuStep15 
         Caption         =   "Step x 1.5"
      End
      Begin VB.Menu mnuStep2 
         Caption         =   "Step x 2"
      End
   End
   Begin VB.Menu mnufileexit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Planprofile2"
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

Dim barused(11) As Double

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

Dim checkgasselected As Boolean

'Dim DB As Database
Dim fp As UserDocument
Dim previouspoint
Dim profilefound As Integer
Dim maxdprofile As Integer
Dim CB As Byte
Dim C As String
Dim c1 As Byte
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
Dim T As String
Dim T2(4) As String
Dim S As String
Dim TS As String
Dim K As Integer
Dim H As Integer
Dim w(4) As Boolean
Dim zxt As Double
Dim CT As Integer
Dim Auto_run As Integer
Dim DataChanged, ppo2changed, tempdepth, tempdepthft, temptime, tempo2, temppo2, temphe, tempgasindex, tempcircuit As String
Dim zoom As Integer
Dim pan As Integer

Dim txtdepthft_focus As Integer
Dim txtdepth_focus As Integer
Dim txtmaxd_focus As Integer
Dim txtmaxdft_focus As Integer
Dim depthcount As Integer
Dim timer2buffer As Integer
Dim timecount As Integer
Dim ppo2count As Integer
Dim gascount As Integer
Dim vhmxcount As Integer

Dim Xstart As Single
Dim Ystart As Single

Dim ans_save As Integer

Dim conversion_factor As Double

Dim amultfreal As Double
Dim amult(20) As Double
Dim maxdstart30 As Double
Dim maxdnorm80 As Double
Dim timedstart30 As Double
Dim timednorm80 As Double
Dim vhmx_maxd_factor As Double
Dim vhmx_stop_factor As Double
Dim vhmx_safe_factor As Double
Dim vhmx_mid_factor As Double
Dim vdeco_vceiling_vi As Integer
Dim vhmx_tol_pressure As Double
Dim vhmx_tol_pressure_first As Double
Dim vhmx_tol_pressure_last As Double
Dim vhmx_compartment_vdeco_vceiling(20) As Double
Dim vhmx_tol_pressure_bnorm As Double
Dim vhmx_compartment_vdeco_vceiling_bnorm(20) As Double
Dim vhmx_tol_pressure_bvhmx As Double
Dim vhmx_compartment_vdeco_vceiling_bvhmx(20) As Double
Dim vhmx_tol_pressure_bnorm_first As Double
Dim vhmx_tol_pressure_bvhmx_first As Double
Dim vhmx_tol_pressure_bnorm_last As Double
Dim vhmx_tol_pressure_bvhmx_last As Double
Dim vhmx_ptissue As Double
Dim vhmx_ptissue_first As Double
Dim vhmx_ptissue_last As Double
Dim vhmx_gastotal_pressure(20) As Double
Dim gf_1 As Double
Dim gf_2 As Double






Private Sub cmd02up_Click()
txto2 = txto2 + 1
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
     MsgBox "Value must be between 400 and 1000 !"
     atmtext.Text = "1000"
    Else
         atmtext.SetFocus
         SendKeys "{HOME}+{END}"
         Command1_Click
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
     MsgBox "Value must be between 400 and 1000 !"
     atmtext.Text = "1000"
 ' Else
 '   Command1_Click
  End If
End Sub

Private Sub cbogasindex_Change()
  If IsNumeric(Right(cbogasindex.Text, 1)) = False Then Exit Sub
 
  If cmdgenerate.Visible = False Then cmdmodify.Visible = True
  tempgasused = cbogasindex.Text
  If tempgasused <> "Gas Index" Then
   For i = 1 To 9
      If lblgasindex(i).Caption = tempgasused Then
        If Option1(i).Value = True Then
          lblo2.Caption = txtoxygen(i).Text
          lblhelium.Caption = txthelium(i).Text
        End If
      End If
   Next i
  End If
End Sub

Private Sub cbogasindex_LostFocus()
tempgasused = cbogasindex.Text
If tempgasused <> "Gas Index" Then
   For i = 1 To 9
      If lblgasindex(i).Caption = tempgasused Then
         lblo2.Caption = txtoxygen(i).Text
         lblhelium.Caption = txthelium(i).Text
      End If
   Next i
End If
End Sub

Private Sub Cbogasused_Click(Index As Integer)

If formstarted = False Then
p = Cbogasused(Index).Index
'MsgBox Cbogasused(p).Text & "KKK"
checkgasused

If Cbogasused(p).Text = "5 - Deco Closed Circuit" Then
  txtppo2(p).Enabled = True
Else
txtppo2(p).Enabled = False
End If
validategasused

'If checkgasusedselected = True Then
'   ans = MsgBox("Not Used - Remove all Dive sequence as well with the same gas index ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
'
'   Select Case ans
'      Case vbYes
'        removerecordgasindex
'
'        validategasused
'
'      Case Else
'        restoregasindex
'   End Select
'   formstarted = False
'Else
'  validategasused
   
'End If
End If

st = Cbogasused(Index).Text
   If Left(Cbogasused(Index).Text, 1) = "0" Then Exit Sub
   lblgasindex_Click (Index)
End Sub

Private Sub Cbogasused_GotFocus(Index As Integer)
'   If Left(Cbogasused(Index).Text, 1) = "0" Then exit sub 'Cbogasused(Index).Text
   If Left(Cbogasused(Index).Text, 1) = "0" Then Cbogasused(Index).Text = "1 - Open Circuit"
   lblgasindex_Click (Index)
End Sub

Private Sub Cbogasused_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
p = txtoxygen(Index).Index
checkgasused
If checkgasusedselected = True Then
   ans = MsgBox("Not Used - Remove all Dive sequence as well with the same gas index ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
   Select Case ans
      Case vbYes
        removerecordgasindex
        validategasused
      Case Else
        restoregasindex
        
   End Select
Else
   validategasused
End If
End If
End Sub

Private Sub Cbogasused_LostFocus(Index As Integer)
  p = Cbogasused(Index).Index

  checkgasused
  If checkgasusedselected = True Then
     ans = MsgBox("Not Used - Remove all Dive sequence as well with the same gas index ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
     Select Case ans
        Case vbYes
           removerecordgasindex
           validategasused
        Case Else
           restoregasindex
     End Select
  Else
     validategasused
  End If
End Sub

Private Sub saveprorecord()
SQL = "SELECT * FROM seqdpprofile"
Set RS = DB.OpenRecordset(SQL)
   RS.AddNew
   RS!Dpprofileid = tempserialno
   MSFlexGrid3.Col = 0
   RS!dpnumseq = MSFlexGrid3.Text
   MSFlexGrid3.Col = 1
   RS!depth = Format(MSFlexGrid3.Text, "###0.0")
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
   
'   RS!depth = "10"
'   RS!Duration = "10"
'   RS!gasid = "8"
   RS.Update
   
   MSFlexGrid3.RowHeight(MSFlexGrid3.Row) = 200
   
End Sub



Private Sub setdatadefault()
   lblhelium.Caption = ""
   lblo2.Caption = ""
   txtdepth.Text = ""
   txttime.Text = ""
   Option3.Value = False
   Option4.Value = True
   cbogasindex.Text = "Gas Index"
   txtppo2v.Enabled = True
   txtppo2v.Text = ""
   txtppo2v.Enabled = False
   CMDPPO2PLUS.Enabled = False
   Command4.Enabled = False
   If CInt(MSFlexGrid3.Rows) > 1 Then
     cmdinsert.Visible = True
     
     cmdSave.Enabled = True
  Else
     cmdinsert.Visible = False
     cmdSave.Enabled = False
  End If
End Sub

Private Sub backcolortogreen()
'txtdepth.BackColor = &HFFFF80
'txttime.BackColor = &HFFFF80
'txtppo2v.BackColor = &HFFFF80
'atmtext.BackColor = &HFFFF80
'safetytext.BackColor = &HFFFF80
'Frame4.BackColor = &HFFFF80
End Sub
Private Sub backcolortored()
'Exit Sub
'txtdepth.BackColor = &HC0C0FF
'txttime.BackColor = &HC0C0FF
'txtppo2v.BackColor = &HC0C0FF
'atmtext.BackColor = &HC0C0FF
'safetytext.BackColor = &HC0C0FF
'Frame4.BackColor = &HC0C0FF
End Sub

Private Sub Check1_Click(Index As Integer)
If Check1(Index).Value = 0 Then
   txthelium(Index).Enabled = False
   txtoxygen(Index).Enabled = False
   txtmaxd(Index).Enabled = False
   txtppo2(Index).Enabled = False
   Check2(Index).Enabled = False
   Decochk(Index).Enabled = True 'False
   txtcylcap2(Index).Enabled = False
   txtbreathrate2(Index).Enabled = False
   gasusage(Index).Enabled = False
   Cbogasused(Index).Text = "0 - Not Used"
   If MSFlexGrid3.Rows > 1 Then
      MSFlexGrid3.Row = 1
      MSFlexGrid3.Col = 7
      lblgasindex_Click (CInt(Right(MSFlexGrid3.Text, 1)))
   Else
     For v = 0 To 9
      If Check1(v).Value = 1 Then
         lblgasindex_Click (v)
         Exit Sub
      End If
     Next
   End If

Else
   txthelium(Index).Enabled = True
   txtoxygen(Index).Enabled = True
   txtmaxd(Index).Enabled = True
   txtppo2(Index).Enabled = True
   Check2(Index).Enabled = True
   Decochk(Index).Enabled = True
   txtcylcap2(Index).Enabled = True
   txtbreathrate2(Index).Enabled = True
   gasusage(Index).Enabled = True
   lblgasindex_Click (Index)
   If Check2(Index).Value = 1 Then
      If Decochk(Index).Value = 1 Then
         Cbogasused(Index).Text = "5 - Deco Closed Circuit"
      Else
         Cbogasused(Index).Text = "2 - Closed Circuit"
      End If
   Else
      If Decochk(Index).Value = 1 Then
         Cbogasused(Index).Text = "4 - Deco Open Circuit"
      Else
         Cbogasused(Index).Text = "1 - Open Circuit"
      End If
   End If
End If

End Sub

Private Sub Check2_Click(Index As Integer)
If Check2(Index).Value = 1 Then
   If Decochk(Index).Value = 1 Then
      Cbogasused(Index).Text = "5 - Deco Closed Circuit"
   Else
      Cbogasused(Index).Text = "2 - Closed Circuit"
'      txtppo2v.SetFocus
   End If
   txtcylcap2(Index).Visible = False
   txtbreathrate2(Index).Visible = False
   gasusage(Index).Visible = False
   
Else
   If Decochk(Index).Value = 1 Then
      Cbogasused(Index).Text = "4 - Deco Open Circuit"
   Else
      Cbogasused(Index).Text = "1 - Open Circuit"
   End If
   txtcylcap2(Index).Visible = True
   txtbreathrate2(Index).Visible = True
   gasusage(Index).Visible = True
   
End If
If InStr(1, Cbogasused(Index), "Closed") Then txtppo2(Index).BackColor = vbYellow Else txtppo2(Index).BackColor = &HE0E0E0
End Sub

Private Sub cmdadd_Click()
If CDbl(txtdepth.Text) > CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) Then
  ans = MsgBox("Depth deeper than the maximum depth value allowed..... " & Chr(13) & "Do you want to reset the depth point to maximum depth value ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
  Select Case ans
    Case vbYes
      txtdepth.Text = txtmaxd(CInt(Right(cbogasindex.Text, 1))) '"10"
      txtdepthft.Text = Format(CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) * feetormeter_factor, "###") '"30"
    Case vbNo
      If feetormeter_feeton = 1 Then
         txtdepthft.SetFocus
         SendKeys "{HOME}+{END}"
      Else
         txtdepth.SetFocus
         SendKeys "{HOME}+{END}"
      End If
      Exit Sub
  End Select
End If
datachangedstatus = True
'backcolortogreen

checkgasindex
'IF Cint(txtdepth) <
If checkgasselected = True And CInt(txtdepth) > 0 And CInt(txttime) > 0 Then
validate_data
 MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
 K = MSFlexGrid3.Rows
 MSFlexGrid3.Row = K - 1
 MSFlexGrid3.Col = 0
 MSFlexGrid3.Text = K - 1
 MSFlexGrid3.Col = 1
 MSFlexGrid3.Text = Format(txtdepth.Text, "0.0")
 MSFlexGrid3.Col = 2
 MSFlexGrid3.Text = txttime
 MSFlexGrid3.Col = 3
 MSFlexGrid3.Text = lblo2.Caption
 MSFlexGrid3.Col = 4
 MSFlexGrid3.Text = lblhelium
 MSFlexGrid3.Col = 6
 If Option3.Value = True Then
    MSFlexGrid3.Text = "Closed Circuit"
    MSFlexGrid3.Col = 5
    If (CDbl(txtppo2v.Text) < (CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1)) Then
      MSFlexGrid3.Text = Format((CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1), "##0.00")
      MsgBox "PPO2 to low - changing to default diluent ppo2 at this depth"
      txtppo2v.Text = MSFlexGrid3.Text
    Else
      MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
    End If
 End If
 If Option4.Value = True Then
    MSFlexGrid3.Text = "Open Circuit"
    MSFlexGrid3.Col = 5
    MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
 End If
 MSFlexGrid3.Col = 7
 MSFlexGrid3.Text = cbogasindex.Text
 Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
 Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
 MSFlexGrid3.Col = 8
 MSFlexGrid3.Text = Format(txtdepthft.Text, "#")
 saveprorecord
 savemaxdepth
 rowindentified = MSFlexGrid3.Row
 If MSFlexGrid3.Rows > 1 Then
   For K = 0 To MSFlexGrid3.Rows - 1
      For p = 0 To 0
         MSFlexGrid3.Row = K
         'MSFlexGrid3.RowHeight(K) = 200
         MSFlexGrid3.Col = p
         If MSFlexGrid3.CellBackColor = vbBlue Then
            For H = 0 To 8
               MSFlexGrid3.Row = K
               MSFlexGrid3.Col = H
               MSFlexGrid3.CellForeColor = MSFlexGrid3.ForeColor
               MSFlexGrid3.CellBackColor = MSFlexGrid3.BackColor
            Next H
        End If
      Next p
    Next K
 End If
 For q = 0 To 8
   MSFlexGrid3.Row = rowindentified
   MSFlexGrid3.Col = q
   MSFlexGrid3.CellForeColor = vbWhite
   MSFlexGrid3.CellBackColor = vbBlue
 Next q
 rowindentified = MSFlexGrid3.Row
 If CInt(MSFlexGrid3.Rows) > 1 Then
     cmdinsert.Visible = True
     cmdmodify.Visible = False
     cmdSave.Enabled = True
 Else
     cmdinsert.Visible = False
     cmdmodify.Visible = False
     cmdSave.Enabled = False
 End If
 cmdinsert.Visible = True
 cmdsaveas.Enabled = False
 Command1_Click
 If MSFlexGrid3.Rows < 3 Then
   singlelevel.Visible = True
 End If
Else
   Title = "Error on System Validation.."
   MsgBox "Incomplete Profile Data !", 48, Title
   If MSFlexGrid3.Rows > 1 Then
      MSFlexGrid3.Rows = MSFlexGrid3.Rows - 1
   End If
   If feetormeter_feeton = 1 Then
     txtdepthft.SetFocus
   Else
     txtdepth.SetFocus
   End If
   SendKeys "{END}"
End If
End Sub

Private Sub cmdclearall_Click()
ans = MsgBox("Do you really want to remove all levels from this profile ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
   MSFlexGrid3.Col = 7
   MSFlexGrid3.Row = 1
   tempgasused = MSFlexGrid3.Text
   
   removerecord
   cleargriddata2
   setdatadefault
   datachangedstatus = True
   cmdsaveas.Enabled = False
   cmdSave.Enabled = False
   cleardecogrid
   mnugaslist_Click
   lblgasindex_Click (CInt(Right(tempgasused, 1)))
   txtdepth = 10
   txttime = 10
  ' backcolortogreen
   display_deco_graph (0)
   singlelevel_Click
   cmdgenerate_Click
   
'   MsgBox "Depth point(s) removed"
Case Else
'   MsgBox "Request cancelled. "
End Select

End Sub

Private Sub cmddepthdown_Click()
checkgasselected = False
backcolortored
checkgasindex
     If checkgasselected = True Then
        If CInt(txtdepth) > 0 And CInt(txtdepth) < 2001 Then
           txtdepth.Text = CStr(CDbl(txtdepth.Text) - (inc_depth / feetormeter_factor))  'txtdepth = txtdepth - 1
           If Option4.Value = True Then
              txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1#) * (CDbl(lblo2.Caption) / 100))
              txtppo2v = Format(txtppo2v, "###.00")
           End If
        Else
           If CInt(txtdepth) <= 0 Then
              txtdepth.Text = "10"
           Else
              txtdepth.Text = txtmaxd(p).Text
           End If
        End If
     Else
        Title = "Error on System Validation.."
        MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
     End If
End Sub

Private Sub cmddepthdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 2
timecount = 0
ppo2count = 0
gascount = 0
Timer2.Enabled = True
timer2buffer = 0
vhmxcount = 0
End Sub

Private Sub cmddepthdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer2.Enabled = False
End Sub

Private Sub cmddepthup_Click()
  checkgasselected = False
  backcolortored
  checkgasindex
  'MsgBox Cint(txtdepth)
     If checkgasselected = True Then
        If CInt(txtdepth) >= 0 And CInt(txtdepth) < 2000 Then
           txtdepth.Text = CStr(CDbl(txtdepth.Text) + (inc_depth / feetormeter_factor))
           If Option4.Value = True Then
              txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1#) * (CDbl(lblo2.Caption) / 100))
              txtppo2v = Format(txtppo2v, "###.00")
           End If
        Else
           If CInt(txtdepth) <= 0 Then
              txtdepth.Text = "10"
           Else
              txtdepth.Text = txtmaxd(p).Text
           End If
        End If
  Else
    Title = "Error on System Validation.."
    MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
  End If
End Sub
Private Sub checkgasindex()
Dim i As Integer

   checkgasselected = False
 '  If cbogasindex.Text = "Gas Index" Then
     For i = 0 To 9
        If txtoxygen(i).BackColor = vbBlue Then
          lblgasindex_Click (i)
          checkgasselected = True
    '      Exit Sub
        End If
     Next i
  '   checkgasselected = False
   'Else
   '   checkgasselected = True
  ' End If
End Sub
Private Sub checkgasused()
   checkgasusedselected = False
   If Cbogasused(p).Text = "0 - Not Used" Then
      SQL = "SELECT COUNT(*) FROM seqdpprofile "
      SQL = SQL & " WHERE gasid = '" & Trim(lblgasindex(p).Caption) & "' and "
      SQL = SQL & " dpprofileid ='" & Trim(tempserialno) & "'"
      Set RS3 = DB.OpenRecordset(SQL)
      If RS3.Fields(0) <> 0 Then
        checkgasusedselected = True
      Else
        checkgasusedselected = False
      End If
   End If
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

Private Sub cmddepthup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 1
timecount = 0
ppo2count = 0
gascount = 0
Timer2.Enabled = True
timer2buffer = 0
vhmxcount = 0
End Sub

Private Sub cmddepthup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer2.Enabled = False
End Sub

Private Sub cmdgenerate_Click()
 If CInt(MSFlexGrid3.Rows) > 2 Then
    MsgBox "too many rows"
 Else
    If CInt(MSFlexGrid3.Rows) = 1 Then
       tempseqnumber = "0"
       MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
    End If
'If CDbl(txtdepth.Text) > CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) Then
'  ans = MsgBox("Depth point more than the maximum depth value allow..... " & Chr(13) & "Do you want to reset the depth point to maximum depth value ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
'  Select Case ans
'    Case vbYes
'      txtdepth.Text = txtmaxd(CInt(Right(cbogasindex.Text, 1))) '"10"
'      txtdepthft.Text = Format(CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) * feetormeter_factor, "###") '"30"
'    Case vbNo
'      If feetormeter_feeton = 1 Then
'         txtdepthft.SetFocus
'         SendKeys "{HOME}+{END}"
'      Else
'         txtdepth.SetFocus
'         SendKeys "{HOME}+{END}"
'      End If
'      Exit Sub
'  End Select
'End If
datachangedstatus = True
'backcolortogreen

checkgasindex
'IF Cint(txtdepth) <
If checkgasselected = True And CInt(txtdepth) > 0 And CInt(txttime) > 0 Then
validate_data
 
 K = MSFlexGrid3.Rows
 MSFlexGrid3.Row = K - 1
 MSFlexGrid3.Col = 0
 MSFlexGrid3.Text = K - 1
 MSFlexGrid3.Col = 1
 MSFlexGrid3.Text = Format(txtdepth.Text, "0.0")
 MSFlexGrid3.Col = 2
 MSFlexGrid3.Text = txttime
 
 MSFlexGrid3.Col = 7
' If MSFlexGrid3.Text = cbogasindex.Text Or Len(MSFlexGrid3.Text) < 3 Then
'   ans = vbYes
' Else
'   ans = MsgBox("Change gas?", vbYesNo, "Change gas")
' End If
 ans = vbYes
If ans = vbYes Then
 MSFlexGrid3.Col = 3
 MSFlexGrid3.Text = lblo2.Caption
 MSFlexGrid3.Col = 4
 MSFlexGrid3.Text = lblhelium
 MSFlexGrid3.Col = 6
 If Option3.Value = True Then
    MSFlexGrid3.Text = "Closed Circuit"
    MSFlexGrid3.Col = 5
    If (CDbl(txtppo2v.Text) < (CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1)) Then
      MSFlexGrid3.Text = Format((CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1), "##0.00")
      MsgBox "PPO2 to low - changing to default diluent ppo2 at this depth"
      txtppo2v.Text = MSFlexGrid3.Text
    Else
      MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
    End If
 End If
 If Option4.Value = True Then
    MSFlexGrid3.Text = "Open Circuit"
    MSFlexGrid3.Col = 5
    MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
 End If
 MSFlexGrid3.Col = 7
 If MSFlexGrid3.Rows > 1 Then
    If MSFlexGrid3.Text <> cbogasindex.Text Then
      If Len(MSFlexGrid3.Text) > 3 Then
         Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = True
         Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = True
      End If
    End If
 End If
 MSFlexGrid3.Text = cbogasindex.Text
 Check1(CInt(Right(cbogasindex.Text, 1))).Enabled = False
 Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
 MSFlexGrid3.Col = 8
 MSFlexGrid3.Text = Format(txtdepthft.Text, "#")
  ' MSFlexGrid3.Row = rowindentified
   MSFlexGrid3.Col = 1
   MSFlexGrid3.Text = Format(txtdepth.Text, "0.0")
   MSFlexGrid3.Col = 2
   MSFlexGrid3.Text = txttime
   MSFlexGrid3.Col = 3
   MSFlexGrid3.Text = lblo2.Caption
   MSFlexGrid3.Col = 4
   MSFlexGrid3.Text = lblhelium.Caption
   MSFlexGrid3.Col = 6
   If Option3.Value = True Then
      MSFlexGrid3.Text = "Closed Circuit"
      MSFlexGrid3.Col = 5
      If (CDbl(txtppo2v.Text) < (CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1)) Then
        MSFlexGrid3.Text = Format((CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1), "##0.00")
        MsgBox "PPO2 to low - changing to default diluent ppo2 at this depth"
        txtppo2v.Text = MSFlexGrid3.Text
      Else
        MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00")
      End If
   End If
   If Option4.Value = True Then
      MSFlexGrid3.Text = "Open Circuit"
      MSFlexGrid3.Col = 5
      MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
   End If
   MSFlexGrid3.Col = 7
   MSFlexGrid3.Text = cbogasindex.Text
End If
   
   MSFlexGrid3.Col = 8
   MSFlexGrid3.Text = Format(txtdepthft.Text, "0")
   MSFlexGrid3.Col = 1
   
   If tempseqnumber = "0" Then
      saveprorecord
   Else
         tempseqnumber = "1"
         SQL = "SELECT * FROM seqdpprofile"
         SQL = SQL & " where dpprofileid = '" & tempserialno & "' and dpnumseq = '" & tempseqnumber & "' "
         Set RS = DB.OpenRecordset(SQL)
         RS.Edit
         MSFlexGrid3.Col = 1
         RS("depth") = Format(MSFlexGrid3.Text, "#.0")
         MSFlexGrid3.Col = 2
         RS("duration") = MSFlexGrid3.Text
         MSFlexGrid3.Col = 3
         RS("dpo2") = MSFlexGrid3.Text
         MSFlexGrid3.Col = 4
         RS("dphe") = MSFlexGrid3.Text
         MSFlexGrid3.Col = 5
         RS("po2") = MSFlexGrid3.Text
         MSFlexGrid3.Col = 6
         If Option3.Value = True Then
            RS("dpcircuit") = MSFlexGrid3.Text
         End If
         If Option4.Value = True Then
            RS("dpcircuit") = MSFlexGrid3.Text
         End If
         MSFlexGrid3.Col = 7
         RS("gasid") = MSFlexGrid3.Text
         RS("po2") = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
         RS.Update
   End If






 savemaxdepth
 rowindentified = MSFlexGrid3.Row
 If MSFlexGrid3.Rows > 1 Then
   For K = 0 To MSFlexGrid3.Rows - 1
      For p = 0 To 0
         MSFlexGrid3.Row = K
         'MSFlexGrid3.RowHeight(K) = 200
         MSFlexGrid3.Col = p
         If MSFlexGrid3.CellBackColor = vbBlue Then
            For H = 0 To 8
               MSFlexGrid3.Row = K
               MSFlexGrid3.Col = H
               MSFlexGrid3.CellForeColor = MSFlexGrid3.ForeColor
               MSFlexGrid3.CellBackColor = MSFlexGrid3.BackColor
            Next H
        End If
      Next p
    Next K
 End If
 For q = 0 To 8
   MSFlexGrid3.Row = rowindentified
   MSFlexGrid3.Col = q
   MSFlexGrid3.CellForeColor = vbWhite
   MSFlexGrid3.CellBackColor = vbBlue
 Next q
 rowindentified = MSFlexGrid3.Row
 If cmdgenerate.Visible = False Then
     cmdinsert.Visible = True
     cmdmodify.Visible = False
     cmdSave.Enabled = True
 Else
     cmdinsert.Visible = False
    
     cmdSave.Enabled = False
 End If
 'cmdinsert.Visible = True
 cmdsaveas.Enabled = False
 
 Command1_Click
 If systemversion = "Pro" Then
   decoresultgrid.Visible = True
   decoresultgridlite.Visible = False
   decographversion = "Pro"
 Else
   decoresultgrid.Visible = False
   decoresultgridlite.Visible = True
   mnudecoversion.Visible = True
   decographversion = "Lite"
 End If
Else
   Title = "Error on System Validation.."
   MsgBox "Incomplete Profile Data !", 48, Title
   If MSFlexGrid3.Rows > 1 Then
      MSFlexGrid3.Rows = MSFlexGrid3.Rows - 1
   End If
   If feetormeter_feeton = 1 Then
     txtdepthft.SetFocus
   Else
     txtdepth.SetFocus
   End If
   SendKeys "{END}"
End If
If CInt(MSFlexGrid3.Rows) = 2 Then
   If MSFlexGrid3.Visible = True Then
      singlelevel.Visible = True
      lbllevel.Visible = False
   Else
      singlelevel.Visible = False
      If buhl_mode = 2 Then Else lbllevel.Visible = True
   End If
End If
End If
mnufile.Enabled = True

End Sub




'Private Sub cmdheplus_Click(index As Integer)
'  If index = 0 Then
'    sptext.Text = Format(CDbl(sptext.Text) + 0.05, "0.00")
'    If CDbl(sptext.Text) > 1.6 Then sptext.Text = "1.60"
'  End If
'
'  If index = 1 Then
'    ScreenSaveText.Text = CStr(CInt(ScreenSaveText.Text) + 30)
'    If CInt(ScreenSaveText.Text) > 300 Then ScreenSaveText.Text = "300"
'  End If
'
'  If index = 2 Then
'    HomeText.Text = CStr(CInt(HomeText.Text) + 30)
'    If CInt(HomeText.Text) > 300 Then HomeText.Text = "300"
'  End If
'
'  If index = 3 Then
'    If current_index > 0 Then Text1(current_index).Text = CStr(CInt(Text1(current_index).Text) + 1)
'  End If
'
'  If index = 4 Then
'    If current_index > 0 Then Text2(current_index).Text = CStr(CInt(Text2(current_index).Text) + 1)
'  End If
'
'End Sub

Private Sub cmdgeneratem_Click()
For q = 1 To MSFlexGrid3.Rows - 1
MSFlexGrid3.Col = 7
MSFlexGrid3.Row = q
p = CInt(Right(MSFlexGrid3.Text, 1))

validateoxygen
validatehelium
validatemaxdepth
Next q
Command1_Click
End Sub

Private Sub cmdinsert_Click()
If CDbl(txtdepth.Text) > CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) Then
  ans = MsgBox("Depth deeper than the maximum depth value allow..... " & Chr(13) & "Do you want to reset the depth point to maximum depth value ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
  Select Case ans
    Case vbYes
      txtdepth.Text = txtmaxd(CInt(Right(cbogasindex.Text, 1))) '"10"
      txtdepthft.Text = Format(CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) * feetormeter_factor, "###") '"30"
    Case vbNo
      If feetormeter_feeton = 1 Then
         txtdepthft.SetFocus
         SendKeys "{HOME}+{END}"
      Else
         txtdepth.SetFocus
         SendKeys "{HOME}+{END}"
      End If
      Exit Sub
  End Select
End If
datachangedstatus = True
'backcolortogreen
validate_data
rowchanged = rowindentified - 1
 If rowindentified <> "0" Then
   totalrow = MSFlexGrid3.Rows - 1
   MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
   For i = totalrow To rowchanged Step -1
      If CInt(i) = CInt(rowchanged) Then
         MSFlexGrid3.Row = rowchanged + 1
         MSFlexGrid3.Col = 0
         MSFlexGrid3.Text = i + 1
         MSFlexGrid3.Col = 1
         MSFlexGrid3.Text = Format(txtdepth.Text, "0.0")
         MSFlexGrid3.Col = 2
         MSFlexGrid3.Text = txttime
         MSFlexGrid3.Col = 3
         MSFlexGrid3.Text = lblo2.Caption
         MSFlexGrid3.Col = 4
         MSFlexGrid3.Text = lblhelium.Caption
         MSFlexGrid3.Col = 6
         If Option3.Value = True Then
            MSFlexGrid3.Text = "Closed Circuit"
            MSFlexGrid3.Col = 5
            If (CDbl(txtppo2v.Text) < (CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1)) Then
              MSFlexGrid3.Text = Format((CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1), "##0.00")
              MsgBox "PPO2 to low - changing to default diluent ppo2 at this depth"
              txtppo2v.Text = MSFlexGrid3.Text
            Else
              MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
            End If
         End If
         If Option4.Value = True Then
            MSFlexGrid3.Text = "Open Circuit"
            MSFlexGrid3.Col = 5
            MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
         End If
         MSFlexGrid3.Col = 7
         MSFlexGrid3.Text = cbogasindex.Text
         Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
         Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
         MSFlexGrid3.Col = 8
         MSFlexGrid3.Text = Format(txtdepthft.Text, "0")
      Else
         readprerowval ' read previous row value
         saveprerowdata
      End If
   Next i
   removerecord
   savechangerecord
   savemaxdepth
   cmdsaveas.Enabled = False
   cmdSave.Enabled = True
   Command1_Click
   If MSFlexGrid3.Rows < 3 Then
      singlelevel.Visible = True
   End If
Else
   Title = "Dive Profile"
   MsgBox "You must selected a record in the list to insert the sequence", 48, Title
End If
cmdmodify.Visible = False
End Sub
Private Sub saveprerowdata()
MSFlexGrid3.Row = i + 1
MSFlexGrid3.Col = 0
MSFlexGrid3.Text = i + 1
MSFlexGrid3.Col = 1
MSFlexGrid3.Text = tempdepth
MSFlexGrid3.Col = 2
MSFlexGrid3.Text = temptime
MSFlexGrid3.Col = 3
MSFlexGrid3.Text = tempo2
MSFlexGrid3.Col = 4
MSFlexGrid3.Text = temphe
MSFlexGrid3.Col = 5
MSFlexGrid3.Text = temppo2
MSFlexGrid3.Col = 6
MSFlexGrid3.Text = tempcircuit
MSFlexGrid3.Col = 7
MSFlexGrid3.Text = tempgasindex
Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
MSFlexGrid3.Col = 8
MSFlexGrid3.Text = Format(tempdepthft, "0")
End Sub
Private Sub readprerowval()
  MSFlexGrid3.Row = i
  MSFlexGrid3.Col = 1
  tempdepth = MSFlexGrid3.Text
  MSFlexGrid3.Col = 2
  temptime = MSFlexGrid3.Text
  MSFlexGrid3.Col = 3
  tempo2 = MSFlexGrid3.Text
  
  MSFlexGrid3.Col = 4
  temphe = MSFlexGrid3.Text
  MSFlexGrid3.Col = 5
  temppo2 = MSFlexGrid3.Text
  MSFlexGrid3.Col = 6
  tempcircuit = MSFlexGrid3.Text
  MSFlexGrid3.Col = 7
  tempgasindex = MSFlexGrid3.Text
   MSFlexGrid3.Col = 8
  tempdepthft = MSFlexGrid3.Text
End Sub


Private Sub saveseqnewrecord()
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
  txtserialno.Text = tempserialno
SQL = "SELECT * FROM dpmaingaslist"
Set RS = DB.OpenRecordset(SQL)
For i = 0 To 9
   RS.AddNew
   RS!dpmainid = tempserialno
   RS!dpgasid = lblgasindex(i).Caption
   RS!dpgashelium = txthelium(i).Text
   tempnitrogen = 100 - CInt(txthelium(i).Text) - CInt(txtoxygen(i).Text)
   RS!dpgasnitrogen = tempnitrogen
   RS!dpgasmaxopdepth = CInt(txtmaxd(i).Text) * 10
   RS!dpgaspo2setpoint = txtppo2(i).Text
   RS!dpgasused = Cbogasused(i).Text
   RS.Update
Next i
SQL = "SELECT * FROM seqdpmain"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!diveplanid = tempserialno
RS!Plandate = Now
RS3!PlanBy = vhmx_text(0).Text + ",a" + vhmx_text(1).Text + ",b" + vhmx_text(2).Text + ",c"
RS!divecategories = "1"
If mnuStep15.Checked Then RS!divecategories = "2"
If mnuStep2.Checked Then RS!divecategories = "3"
RS.Update
For K = 1 To MSFlexGrid3.Rows - 1
    MSFlexGrid3.Row = K
    saveprorecord
  Next K
savemaxdepth
Unload Me
frmseqdive.Show
End Sub
Private Sub checkforchanges()
MSFlexGrid3.Row = rowindentified
For i = 1 To 6
Select Case i
Case 1
   MSFlexGrid3.Col = i
   If Trim(txtdepth) <> Trim(MSFlexGrid3.Text) Then
      DataChanged = "True"
   End If
Case 2
   MSFlexGrid3.Col = i
   If Trim(txttime) <> Trim(MSFlexGrid3.Text) Then
      DataChanged = "True"
   End If
Case 3
   MSFlexGrid3.Col = i
   If Trim(lblo2.Caption) <> Trim(MSFlexGrid3.Text) Then
      DataChanged = "True"
   End If
Case 4
   MSFlexGrid3.Col = i
   If Trim(lblhelium) <> Trim(MSFlexGrid3.Text) Then
      DataChanged = "True"
   End If
Case 5
   MSFlexGrid3.Col = i
   If Trim(txtppo2v) <> Trim(MSFlexGrid3.Text) Then
      ppo2changed = "True"
      DataChanged = "True"
   End If
Case 6
   MSFlexGrid3.Col = i
   If Option3.Value = True Then
      tempcircuit = "Closed Circuit"
   End If
   If Option4.Value = True Then
      tempcircuit = "Open Circuit"
   End If
   If Trim(tempcircuit) <> Trim(MSFlexGrid3.Text) Then
      DataChanged = "True"
   End If
Case 7
   MSFlexGrid3.Col = i
   If Trim(cbogasindex.Text) <> Trim(MSFlexGrid3.Text) Then
      DataChanged = "True"
   End If
End Select
Next i

End Sub

Private Sub Cmdminus_Click(Index As Integer)
  p = current_index
  If Index = 3 Then
    If (current_index > 0) And CInt(txtoxygen(current_index).Text) > 1 Then txtoxygen(current_index).Text = CStr(CInt(txtoxygen(current_index).Text) - 1)
  End If
  
  If Index = 4 Then
    If (current_index > 0) And CInt(txthelium(current_index).Text) >= 0 Then txthelium(current_index).Text = CStr(CInt(txthelium(current_index).Text) - 1)
  End If
   If CInt(txtoxygen(current_index).Text) >= 1 And CInt(txtoxygen(current_index).Text) <= 100 And ((CInt(txtoxygen(current_index).Text) + CInt(txthelium(current_index).Text)) < 101) Then
   Else
     Cmdplus_Click (Index)
   End If
 
 update_gas_graph (current_index)
 
End Sub

Private Sub cmdminus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 0
timecount = 0
ppo2count = 0
gascount = Index - 2
Timer2.Enabled = True
timer2buffer = 0
vhmxcount = 0

End Sub

Private Sub cmdminus_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer2.Enabled = False
End Sub

Private Sub cmdmodify_Click()
If CDbl(txtdepth.Text) > CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) Then
  ans = MsgBox("Depth deeper than the maximum depth value allow..... " & Chr(13) & "Do you want to reset the depth point to maximum depth value ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
  Select Case ans
    Case vbYes
      txtdepth.Text = txtmaxd(CInt(Right(cbogasindex.Text, 1))) '"10"
      txtdepthft.Text = Format(CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) * feetormeter_factor, "###") '"30"
    Case vbNo
      If feetormeter_feeton = 1 Then
         txtdepthft.SetFocus
         SendKeys "{HOME}+{END}"
      Else
         txtdepth.SetFocus
         SendKeys "{HOME}+{END}"
      End If
      Exit Sub
  End Select
End If
datachangedstatus = True
DataChanged = "False"
ppo2changed = "False"
'MsgBox MSFlexGrid3.Row
'backcolortogreen
validate_data
checkforchanges

If ppo2changed = "True" Then
   MSFlexGrid3.Row = rowindentified
   MSFlexGrid3.Col = 5
   tempppo2 = MSFlexGrid3.Text
End If
If DataChanged = "True" Then
   MSFlexGrid3.Row = rowindentified
   MSFlexGrid3.Col = 1
   MSFlexGrid3.Text = Format(txtdepth.Text, "0.0")
   MSFlexGrid3.Col = 2
   MSFlexGrid3.Text = txttime
   MSFlexGrid3.Col = 3
   MSFlexGrid3.Text = lblo2.Caption
   MSFlexGrid3.Col = 4
   MSFlexGrid3.Text = lblhelium.Caption
   MSFlexGrid3.Col = 6
   If Option3.Value = True Then
      MSFlexGrid3.Text = "Closed Circuit"
      MSFlexGrid3.Col = 5
      If (CDbl(txtppo2v.Text) < (CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1)) Then
        MSFlexGrid3.Text = Format((CInt(lblo2.Caption) / 100) * ((CInt(txtdepth) / 10) + 1), "##0.00")
        MsgBox "PPO2 to low - changing to default diluent ppo2 at this depth"
        txtppo2v.Text = MSFlexGrid3.Text
      Else
        MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00")
      End If
   End If
   If Option4.Value = True Then
      MSFlexGrid3.Text = "Open Circuit"
      MSFlexGrid3.Col = 5
      MSFlexGrid3.Text = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
   End If
   MSFlexGrid3.Col = 7
   Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = True
   Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = True
   MSFlexGrid3.Text = cbogasindex.Text
   Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
   Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
   MSFlexGrid3.Col = 8
   MSFlexGrid3.Text = Format(txtdepthft.Text, "0")
   SQL = "SELECT * FROM seqdpprofile"
   SQL = SQL & " where dpprofileid = '" & tempserialno & "' and dpnumseq = '" & rowindentified & "' "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   MSFlexGrid3.Col = 1
   RS("depth") = Format(MSFlexGrid3.Text, "#.0")
   MSFlexGrid3.Col = 2
   RS("duration") = MSFlexGrid3.Text
   MSFlexGrid3.Col = 3
   RS("dpo2") = MSFlexGrid3.Text
   MSFlexGrid3.Col = 4
   RS("dphe") = MSFlexGrid3.Text
   MSFlexGrid3.Col = 5
   RS("po2") = MSFlexGrid3.Text
   MSFlexGrid3.Col = 6
   If Option3.Value = True Then
      RS("dpcircuit") = MSFlexGrid3.Text
   End If
   If Option4.Value = True Then
      RS("dpcircuit") = MSFlexGrid3.Text
   End If
   MSFlexGrid3.Col = 7
   RS("gasid") = MSFlexGrid3.Text
   RS("po2") = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
   RS.Update
   savemaxdepth
   For K = 1 To MSFlexGrid3.Rows - 1
      MSFlexGrid3.Col = 7
      MSFlexGrid3.Row = K
      Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
      Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
   Next
   cmdsaveas.Enabled = False
   cmdSave.Enabled = True
   Command1_Click
   If MSFlexGrid3.Rows < 3 Then
      If cmdmodify.Visible = False Then
         If buhl_mode = 2 Then Else lbllevel.Visible = True
         singlelevel.Visible = False
      Else
         singlelevel.Visible = True
      End If
   End If
   cmdmodify.Visible = False
   MSFlexGrid3.Row = rowindentified
End If
   
End Sub

Private Sub cmdplan_Click()
Unload Me
Splanmain.Show
End Sub

Private Sub Cmdplus_Click(Index As Integer)
  p = current_index
'  Timer2.Enabled = False
  If Index = 3 Then
    If current_index > 0 Then txtoxygen(current_index).Text = CStr(CInt(txtoxygen(current_index).Text) + 1)
  End If
  
  If Index = 4 Then
    If current_index > 0 Then txthelium(current_index).Text = CStr(CInt(txthelium(current_index).Text) + 1)
  End If
  If CInt(txtoxygen(current_index).Text) >= 1 And CInt(txtoxygen(current_index).Text) <= 100 And ((CInt(txtoxygen(current_index).Text) + CInt(txthelium(current_index).Text)) < 101) Then
  Else
    Cmdminus_Click (Index)
  End If
  update_gas_graph (current_index)
End Sub

Private Sub cmdplus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 0
timecount = 0
ppo2count = 0
gascount = Index
Timer2.Enabled = True
timer2buffer = 0
vhmxcount = 0
End Sub

Private Sub cmdplus_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Enabled = False
End Sub

Private Sub CMDPPO2PLUS_Click()
If CDbl(txtppo2v) >= 0.15 And CDbl(txtppo2v) < 2.01 Then
   txtppo2v = CDbl(txtppo2v.Text) + 0.01
   backcolortored
Else
   txtppo2v.Text = txtppo2(p).Text
End If
txtppo2v.Text = Format(txtppo2v.Text, "0.00")
End Sub

Private Sub CMDPPO2PLUS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 0
timecount = 0
ppo2count = 1
gascount = 0
Timer2.Enabled = True
timer2buffer = 0
vhmxcount = 0
End Sub

Private Sub CMDPPO2PLUS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer2.Enabled = False
End Sub

Private Sub cmdremove_Click()
If MSFlexGrid3.Rows <= 1 Then Exit Sub
datachangedstatus = True
backcolortogreen
MSFlexGrid3.Col = 0
MSFlexGrid3.Row = rowindentified
tempseq = MSFlexGrid3.Text
SQL = "SELECT * FROM seqdpprofile "
SQL = SQL & "where dpnumseq = '" & tempseq & "'  and dpprofileid = '" & tempserialno & "' "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   RS3.Delete
   RS3.MoveNext
Wend
griddataexist
numrow = MSFlexGrid3.Rows - 2

If rowindentified = 1 Then
  If numrow = 0 Then
      cmdclearall_Click
      Exit Sub
  Else
      rowindentified = rowindentified
  End If
Else
  If rowindentified = numrow + 1 Then
     rowindentified = rowindentified - 1
  Else
     rowindentified = rowindentified
  End If
End If
reloadgriddata
removerecord
savechangerecord
savemaxdepth
cmdsaveas.Enabled = False
cmdSave.Enabled = True
cmdinsert.Visible = True
cmdmodify.Visible = False
cmdremove.Enabled = True
Command1_Click
If MSFlexGrid3.Rows < 3 Then
   singlelevel.Visible = True
End If
End Sub
Private Sub savechangerecord()
  For K = 1 To MSFlexGrid3.Rows - 1
    MSFlexGrid3.Row = K
    saveprorecord
  Next K
End Sub
Private Sub removerecord()
  SQL = "SELECT * FROM seqdpprofile"
  SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
  Set RS = DB.OpenRecordset(SQL)
  While RS.EOF = False
   RS.Delete
   RS.MoveNext
  Wend
End Sub
Private Sub removerecordgasindex()
  SQL = "SELECT * FROM seqdpprofile "
  SQL = SQL & " WHERE gasid = '" & Trim(lblgasindex(p).Caption) & "' and "
  SQL = SQL & " dpprofileid = '" & tempserialno & "' "
    Set RS = DB.OpenRecordset(SQL)
  While RS.EOF = False
   RS.Delete
   RS.MoveNext
  Wend
End Sub
Private Sub restoregasindex()
  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & " WHERE dpgasid = '" & Trim(lblgasindex(p).Caption) & "' and "
  SQL = SQL & " dpmainid = '" & tempserialno & "' "
  Set RS = DB.OpenRecordset(SQL)
  formstarted = True
    While RS.EOF = False
     Cbogasused(p).Text = RS("dpgasused")
     RS.MoveNext
  Wend
End Sub
Private Sub cleargriddata()
For K = 1 To MSFlexGrid3.Rows - 1
For q = 0 To 8
   MSFlexGrid3.Col = q
   If q = 7 Then
     Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = True
     Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = True
   End If
   MSFlexGrid3.Row = K
   MSFlexGrid3.Text = ""
Next q
Next K
MSFlexGrid3.Rows = 1
End Sub
Private Sub cleargriddata2()
 i = MSFlexGrid3.Rows
 i = i - 1
For K = i To 1 Step -1
For q = 0 To 8
   MSFlexGrid3.Col = q
    If q = 7 Then
     Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = True
     Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = True
    End If

   MSFlexGrid3.Row = K
   MSFlexGrid3.Text = ""
Next q
MSFlexGrid3.Rows = MSFlexGrid3.Rows - 1
Next K
End Sub
Private Sub griddataexist()
profilerecordexist = False
SQL = "SELECT COUNT(*) FROM seqdpprofile "
  SQL = SQL & " WHERE "
  SQL = SQL & " dpprofileid ='" & Trim(tempserialno) & "'"
  Set RS3 = DB.OpenRecordset(SQL)
  If RS3.Fields(0) <> 0 Then
     profilerecordexist = True
  End If
End Sub
Private Sub reloadgriddata()
cleargriddata
griddataexist
If profilerecordexist = True Then
cmdSave.Enabled = True
SQL = "SELECT * FROM seqdpprofile"
SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
SQL = SQL & " order by dpnumseq "
Set RS7 = DB.OpenRecordset(SQL)
RS7.MoveFirst
MSFlexGrid3.Rows = 1
While RS7.EOF = False
     MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
     K = MSFlexGrid3.Rows
     MSFlexGrid3.Row = K - 1
     MSFlexGrid3.Col = 0
     MSFlexGrid3.Text = K - 1
     MSFlexGrid3.Col = 1
     MSFlexGrid3.Text = RS7("depth")
     MSFlexGrid3.Col = 2
     MSFlexGrid3.Text = RS7("duration")
     MSFlexGrid3.Col = 3
     MSFlexGrid3.Text = RS7("dpo2")
     MSFlexGrid3.Col = 4
     MSFlexGrid3.Text = RS7("dphe")
     MSFlexGrid3.Col = 5
     MSFlexGrid3.Text = RS7("po2")
     MSFlexGrid3.Col = 6
     MSFlexGrid3.Text = RS7("dpcircuit")
     MSFlexGrid3.Col = 7
     MSFlexGrid3.Text = RS7("gasid")
     Check1(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
     Decochk(CInt(Right(MSFlexGrid3.Text, 1))).Enabled = False
     MSFlexGrid3.Col = 8
     MSFlexGrid3.Text = Format(CStr(CDbl(RS7("depth")) * feetormeter_factor + 0.1), "###0.0")
     RS7.MoveNext
     MSFlexGrid3.RowHeight(MSFlexGrid3.Row) = 200
Wend
If IsNull(rowindentified) Then rowindentified = MSFlexGrid3.Rows - 1
If IsEmpty(rowindentified) Then rowindentified = MSFlexGrid3.Rows - 1
If rowindentified < 1 Then rowindentified = MSFlexGrid3.Rows - 1
If Val(rowindentified) > MSFlexGrid3.Rows - 1 Then
   rowindentified = MSFlexGrid3.Rows - 1
Else
   MSFlexGrid3.Row = rowindentified
End If
For K = 0 To 8
  MSFlexGrid3.Col = K
  MSFlexGrid3.CellBackColor = vbBlue
  MSFlexGrid3.CellForeColor = vbWhite
Next K
Else
  ' cmdinsert.Visible = False
  If MSFlexGrid3.Rows > 1 Then
     cmdmodify.Visible = True
  Else
     cmdmodify.Visible = False
  End If
   cmdremove.Enabled = True
   cmdSave.Enabled = True
End If
'decoresultgrid.Rows = 1
backcolortored
End Sub
'Private Sub reloadgriddata2()
'cleargriddata
'If profilerecordexist = True Then
'cmdSave.Enabled = True
'SQL = "SELECT * FROM seqdpprofile"
'SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
'SQL = SQL & " order by dpnumseq "
'Set RS = DB.OpenRecordset(SQL)
'RS.MoveFirst
'MSFlexGrid3.Rows = 1
'While RS.EOF = False
'     MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
'     K = MSFlexGrid3.Rows
'     MSFlexGrid3.Row = K - 1
'     MSFlexGrid3.Col = 0
'     MSFlexGrid3.Text = K - 1
'     MSFlexGrid3.Col = 1
'     MSFlexGrid3.Text = RS("depth")
'     MSFlexGrid3.Col = 2
'     MSFlexGrid3.Text = RS("duration")
'     MSFlexGrid3.Col = 3
'     MSFlexGrid3.Text = RS("dpo2")
'     MSFlexGrid3.Col = 4
'     MSFlexGrid3.Text = RS("dphe")
'     MSFlexGrid3.Col = 5
'     MSFlexGrid3.Text = RS("po2")
'     MSFlexGrid3.Col = 6
'     MSFlexGrid3.Text = RS("dpcircuit")
'     MSFlexGrid3.Col = 7
'     MSFlexGrid3.Text = RS("gasid")
'     MSFlexGrid3.Col = 8
'     MSFlexGrid3.Text = Format(CStr(CDbl(RS("depth")) * feetormeter_factor), "###0.0.0")
'     RS.MoveNext
'Wend
'Else
'   cmdinsert.Visible = False
'   cmdmodify.Visible = False
'   cmdremove.Enabled = True
'   cmdSave.Enabled = True
'End If
  
'End Sub
Private Sub cmdsave_Click()
datachangedstatus = False
If tempchoice = "NSP" Then
   SQL = "select * FROM seqdpmain "
   SQL = SQL & " where diveplanid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!diveplanid = txtserialno.Text
      RS3!Plandate = Now
      RS3!divecategories = "1"
      RS3!PlanBy = vhmx_text(0).Text + ",a" + vhmx_text(1).Text + ",b" + vhmx_text(2).Text + ",c"
      If mnuStep15.Checked Then RS3!divecategories = "2"
      If mnuStep2.Checked Then RS3!divecategories = "3"
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpmaingaslist "
   SQL = SQL & " where dpmainid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!dpmainid = txtserialno.Text
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM seqdpprofile "
   SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!Dpprofileid = txtserialno.Text
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpserialno "
   Set RS3 = DB.OpenRecordset(SQL)
   RS3.Edit
   RS3!lastseqdserialno = tempserialno
  ' RS3!seqdiveserialno = newseqdiveno
   RS3.Update
     
   MsgBox "Record Saved!"
   
   Unload Me
   Select Case previousform
   Case "SEQLIST"
      Splanmain.Show
   Case "SEQPLAN"
      frmseqdive.Show
   End Select
End If

If tempchoice = "NPP" Then
   SQL = "select * FROM seqdpmain "
   SQL = SQL & " where diveplanid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!diveplanid = txtserialno.Text
      RS3!Plandate = Now
      RS3!divecategories = "1"
      RS3!PlanBy = vhmx_text(0).Text + ",a" + vhmx_text(1).Text + ",b" + vhmx_text(2).Text + ",c"
      If mnuStep15.Checked Then RS3!divecategories = "2"
      If mnuStep2.Checked Then RS3!divecategories = "3"
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpmaingaslist "
   SQL = SQL & " where dpmainid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!dpmainid = txtserialno.Text
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM seqdpprofile "
   SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!Dpprofileid = txtserialno.Text
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpserialno "
   Set RS3 = DB.OpenRecordset(SQL)
   RS3.Edit
   RS3!lastseqdserialno = txtserialno.Text
   RS3.Update
   If buhl_mode = 2 Then
     ans = vbNo
   Else
     Title = "Create New Dive Series"
     ans = MsgBox("Do you want to create a new Dive Series now? ", vbYesNo + vbCritical + vbDefaultButton2, "Create New Dive Series")
   End If
   Select Case ans
   Case vbYes
     
      SQL = "select * FROM dpserialno "
      Set RS = DB.OpenRecordset(SQL)
      tempseqdiveno2 = RS("seqdiveserialno")
      If IsNull(tempseqdiveno2) Then
        tempseqdiveno = "0"
      Else
        tempseqdiveno = Right(tempseqdiveno2, 8)
      End If
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
      do_not_load = 1
      Unload Me
      do_not_load = 0
      frmseqdive.Show
   Case vbNo
     Unload Me
     Splanmain.Show
   End Select
End If
If tempchoice = "SPP" And tempserialno Like "S*" Then
     
   Unload Me
   Select Case previousform
   Case "SEQLIST"
      Splanmain.Show
   Case "SEQPLAN"
      frmseqdive.Show
   End Select
End If

If tempchoice = "SPP" And tempserialno Like "T*" Then
   tempserialno2 = Right(tempserialno, 9)
   newserialno = "S" & tempserialno2
   SQL = "select * FROM seqdpmain "
   SQL = SQL & " where diveplanid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!diveplanid = txtserialno.Text
      RS3!Plandate = Now
      RS3!divecategories = "1"
      RS3!PlanBy = vhmx_text(0).Text + ",a" + vhmx_text(1).Text + ",b" + vhmx_text(2).Text + ",c"
      If mnuStep15.Checked Then RS3!divecategories = "2"
      If mnuStep2.Checked Then RS3!divecategories = "3"
      RS3.Update
      RS3.MoveNext
      
   Wend
   SQL = "select * FROM dpmaingaslist "
   SQL = SQL & " where dpmainid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!dpmainid = txtserialno.Text
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM seqdpprofile "
   SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!Dpprofileid = txtserialno.Text
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM seqdplist "
   SQL = SQL & " where seqdiveidmain = '" & tempseqdiveno & "' and "
   SQL = SQL & " seqdiveid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!seqdiveid = txtserialno.Text
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpserialno "
   Set RS3 = DB.OpenRecordset(SQL)
   RS3.Edit
   RS3!lastseqdserialno = txtserialno.Text
   RS3!seqdiveserialno = newseqdiveno
   RS3.Update
   
   MsgBox "Record Saved!"
   
      Unload Me
   Select Case previousform
   Case "SEQLIST"
      Splanmain.Show
   Case "SEQPLAN"
      frmseqdive.Show
   End Select
Else
   If tempchoice = "SPP" Then
      tempserialno2 = Right(tempserialno, 9)
      newserialno = "S" & tempserialno2
      SQL = "select * FROM seqdpmain "
   SQL = SQL & " where diveplanid = '" & txtserialno.Text & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
     ' RS3!diveplanid = txtserialno.Text
      RS3!Plandate = Now
      RS3!divecategories = "1"
      RS3!PlanBy = vhmx_text(0).Text + ",a" + vhmx_text(1).Text + ",b" + vhmx_text(2).Text + ",c"
      If mnuStep15.Checked Then RS3!divecategories = "2"
      If mnuStep2.Checked Then RS3!divecategories = "3"
      RS3.Update
      RS3.MoveNext
    Wend
    End If
End If

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
      RS3.Edit
      tempmaxdepth = Format(tempmaxdepth, "###0.0")
      Select Case Len(tempmaxdepth)
      Case 3
         tempmaxdepth = "00" & tempmaxdepth
      Case 4
         tempmaxdepth = "0" & tempmaxdepth
      Case 5
         tempmaxdepth = tempmaxdepth
      End Select
      RS3!MaxDepth = tempmaxdepth
      RS3!divecategories = "1"
      If mnuStep15.Checked Then RS3!divecategories = "2"
      If mnuStep2.Checked Then RS3!divecategories = "3"
      RS3.Update
      RS3.MoveNext
      
   Wend
   SQL = "select * FROM seqdplist "
   SQL = SQL & " where seqdiveid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   If RS3.EOF = False Then
      While RS3.EOF = False
      RS3.Edit
      RS3!seqdiveidmaxdepth = Fix(tempmaxdepth)
      RS3.Update
      RS3.MoveNext
      Wend
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

Private Sub cmdsaveas_Click()
   saveseqnewrecord
   MsgBox "New Dive : " & tempserialno & " created !"
End Sub

Private Sub cmdtimedown_Click()
checkgasselected = False
backcolortored
checkgasindex
   If checkgasselected = True Then
      If CInt(txttime) > 0 And CInt(xttime) < 9999 Then
         txttime = txttime - inc_time
      'Else
      '   txttime = "10"
      End If
   Else
      Title = "Error on System Validation.."
      MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
   End If
End Sub

Private Sub cmdtimedown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 0
timecount = 2
ppo2count = 0
gascount = 0
Timer2.Enabled = True
timer2buffer = 0
vhmxcount = 0
End Sub

Private Sub cmdtimedown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer2.Enabled = False
End Sub

Private Sub cmdtimeup_Click()
checkgasselected = False
backcolortored
checkgasindex
 If checkgasselected = True Then
     If CInt(txttime) >= 0 And CInt(txttime) < 9999 Then
        txttime = txttime + inc_time
     Else
        txttime = "10"
     End If
 Else
    Title = "Error on System Validation.."
    MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
 End If
End Sub

Private Sub Cmdtissue_Click()
Unload Me
rbtissue.Show
End Sub

Private Sub cmdup_Click()

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






Private Sub cmdtimeup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 0
timecount = 1
ppo2count = 0
gascount = 0
Timer2.Enabled = True
timer2buffer = 0
vhmxcount = 0
End Sub

Private Sub cmdtimeup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer2.Enabled = False
End Sub

Private Sub Command3_Click(Index As Integer)
If MSFlexGrid3.Rows < 2 Then
  MsgBox "Add Profile Plan Points before add to Sequential Plan "
  Exit Sub
Else
frmseqdive.Show
End If
End Sub

Private Sub Command4_Click()
If CDbl(txtppo2v) > 0.15 And CDbl(txtppo2v) < 2.01 Then
   txtppo2v = CDbl(txtppo2v) - 0.01
   backcolortored
Else
   txtppo2v.Text = txtppo2(p).Text
End If
txtppo2v.Text = Format(txtppo2v.Text, "0.00")
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 0
timecount = 0
ppo2count = 2
gascount = 0
Timer2.Enabled = True
timer2buffer = 0
vhmxcount = 0
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer2.Enabled = False
End Sub



Private Sub Decochk_Click(Index As Integer)
If Decochk(Index).Value = 1 Then
   Check1(Index).Value = 1
   If Check2(Index).Value = 1 Then
      Cbogasused(Index).Text = "5 - Deco Closed Circuit"
   Else
      Cbogasused(Index).Text = "4 - Deco Open Circuit"
   End If
   If MSFlexGrid3.Rows > 1 Then
      MSFlexGrid3.Row = 1
      MSFlexGrid3.Col = 7
      lblgasindex_Click (CInt(Right(MSFlexGrid3.Text, 1)))
   Else
     For v = 0 To 9
      If Check1(v).Value = 1 Then
'         lblgasindex_Click (v)
         Exit Sub
      End If
     Next
   End If
Else
   Check1(Index).Value = 0
   Cbogasused(Index).Text = "0 - Not Used"
For i = 0 To 9
  If Option1(i).Value = True Then
    lblgasindex_Click (i)
  End If
Next i
'   If Check2(Index).Value = 1 Then
'      Cbogasused(Index).Text = "2 - Closed Circuit"
'   Else
'      Cbogasused(Index).Text = "1 - Open Circuit"
'   End If
End If
End Sub

Private Sub Form_Activate()

  If feetormeter_feeton = 1 Then
    txtdepth.Visible = False
    txtdepthft.SetFocus
 '   lblcylsize(2).Caption = "PsiUse"
 '   lblcylsize(0).Caption = "Cylinder VCF"
 '   lblcylsize(3).Caption = "cuft/min"
  '  txtbreathratecuft.Visible = True
  '  txtbreathrate.Visible = False
  Else
    txtdepth.SetFocus
    txtdepthft.Visible = False
 '   lblcylsize(2).Caption = "BarUse"
 '   lblcylsize(0).Caption = "Cylinder WC"
 '   lblcylsize(3).Caption = "l/min"
  '  txtbreathratecuft.Visible = False
  '  txtbreathrate.Visible = True
  End If
    SendKeys "{END}"

  'If tempchoice = "SPP" Or tempchoice = "GSP" Then
   If MSFlexGrid3.Rows > 1 Then
 '   MSFlexGrid3.Row = 1
 '   MSFlexGrid3_Click
  End If
  

End Sub
Private Sub update_gas_graph(Index As Integer)
      If IsNumeric(txthelium(Index)) And IsNumeric(txtoxygen(Index)) And IsNumeric(txtmaxd(Index)) Then
        lblEan(Index).Caption = Fix(((100# - CDbl(txthelium(Index).Text) - CDbl(txtoxygen(Index).Text)) * (CDbl(txtmaxd(Index).Text) + 10#) / 79# - 10#) * feetormeter_factor + 0.499)
        Shape7(0).Height = CInt(txtoxygen(Index).Text) * 24
        Shape7(2).Height = CInt(txthelium(Index).Text) * 24
        Shape7(2).Top = Shape7(1).Top + Shape7(1).Height - Shape7(2).Height
     End If
     Lblo2g.Caption = "O2 " + txtoxygen(Index).Text
     Lblhel.Caption = "HE " + txthelium(Index).Text
End Sub
Private Sub Form_Deactivate()
  T = T
End Sub

Private Sub Form_Load()
Dim i As Integer
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
    
amult(0) = 1.98 ', //0 fast
amult(1) = 1.68  ', //1 fast
amult(2) = 1.48  ', //2 fast
amult(3) = 1.28  ', //3 fast
amult(4) = 1.16   ', //4 fast
amult(5) = 1.1   ', //5 medium //1.02,
amult(6) = 1.08  ', //6 medium //1.00, //0.98,
amult(7) = 1.06  '1#    ', //7 medium //0.97, //0.95,
amult(8) = 1.04  '1#  '0.98  ', //8 medium //0.96, //0.94,
amult(9) = 1.03  '1#  '0.96  ', //9 medium //0.94,
amult(10) = 1.02  '1#  '0.94  ', //10 medium //0.92,
amult(11) = 1.02  '1#  '0.9   ', //11 slow //0.88,
amult(12) = 1.02  '1#  '0.9   ', //12 slow //0.88,
amult(13) = 1.02  '1#  '0.88  ', //13 slow
amult(14) = 1.02  '1#  '0.88  ', //14 slow
amult(15) = 1.02  '1#  '0.88 '  //15 slow

For j = 0 To 15
  'amult(j) = amult(j) * 1.02
Next j

safetytext = 0
atmtext = 1000
do_not_load = 0

Xstart = -1
Ystart = -1

ans_save = 0

mnuVPMB_Click (buhl_mode)
  If systemversion = "Pro" Then
   decoresultgrid.Visible = True
   decoresultgridlite.Visible = False
   decographversion = "Pro"
Else
   decoresultgrid.Visible = False
   decoresultgridlite.Visible = True
   mnudecoversion.Visible = True
   decographversion = "Lite"
End If
'cmdplan.Visible = True
For i = 0 To 9
  If feetormeter_feeton = 1 Then
    txtmaxdft(i).Visible = True
    txtmaxd(i).Visible = False
    txtdepthft.Visible = True
    txtdepth.Visible = True
    conversion_factor = 32.098
  Else
    txtmaxdft(i).Visible = False
    txtmaxd(i).Visible = True
    txtdepthft.Visible = True
    txtdepth.Visible = True
    conversion_factor = 10#
  End If
Next i
Label21(0).Caption = "MOD " & feetormeter_shortstring
Label11.Caption = feetormeter_shortstring
Label36.Caption = feetormeter_shortstring 'feetormeter_string
formstarted = True
datachangedstatus = False
txtdepthft_focus = 0
txtmaxd_focus = 0
xtmaxd_focusft = 0
Top = 30
Me.Left = (Screen.Width - Me.Width) / 2
fgactivate = "0"
colselected = "false"
cols0activated = "false"
initialgrid
txtdepth.Text = "0"
txttime.Text = "10"

For i = 0 To 9
  Cbogasused(i).AddItem "0 - Not Used"
  Cbogasused(i).AddItem "1 - Open Circuit"
  Cbogasused(i).AddItem "2 - Closed Circuit"
  Cbogasused(i).AddItem "3 - Open & Closed"
  Cbogasused(i).AddItem "4 - Deco Open Circuit"
  Cbogasused(i).AddItem "5 - Deco Closed Circuit"
  txtbreathrate2(i).Text = "2"
  txtcylcap2(i).Text = "10"
'  Cbogasused(i).Height = 195
Next i
 If tempchoice = "NSP" Then
  cmdplan.Visible = False
  mnuloaddesetting.Enabled = True
  mnugasloadefault.Enabled = True
  If Trim(newserialno) = "" Then
     newserialno = tempserialno
  End If
  
    txtserialno.Text = tempserialno
    lblseqdiveno.Caption = newseqdiveno
  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & " where dpmainid = '" & tempserialno & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  For i = 0 To 9
     lblgasindex(i).Caption = RS("dpgasid")
     tempnitrogen = RS("dpgasnitrogen")
     txthelium(i) = RS("dpgashelium")
     txtoxygen(i).Text = 100 - CInt(txthelium(i).Text) - CInt(tempnitrogen)
     txtmaxd(i) = RS("dpgasmaxopdepth")
     txtmaxd(i) = txtmaxd(i) / 10
     temptxthe1 = txthelium(i)
     temptxtmaxd1 = txtmaxd(i)
     Cbogasused(i) = RS("dpgasused")
     txtppo2(i).Enabled = True
     txtppo2(i).Text = RS("dpgaspo2setpoint")
     'txtppo2(i).Text = (Cint(txtoxygen(i).Text) / 100) * ((Cint(txtmaxd(i).Text) / 10) + 1)
     txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     If Cbogasused(i).Text <> "5 - Deco Closed Circuit" Then
        txtppo2(i).Enabled = False
     Else
        txtppo2(i).Enabled = True
     End If
     If Cbogasused(i).Text <> "0 - Not Used" Then
        cbogasindex.AddItem lblgasindex(i).Caption
     End If
     checkgasuse (i)
     RS.MoveNext
  Next i
    SQL = "SELECT COUNT(*) FROM seqdpprofile "
  SQL = SQL & " WHERE "
  SQL = SQL & " dpprofileid ='" & Trim(tempserialno) & "'"
  Set RS3 = DB.OpenRecordset(SQL)
  If RS3.Fields(0) <> 0 Then
     loaddpprofiledata
  End If
   cmdsaveas.Enabled = True
    cmdSave.Enabled = False
 End If
 
 If tempchoice = "NPP" Then
  cmdplan.Visible = False
  mnuloaddesetting.Enabled = True
  mnugasloadefault.Enabled = True
  If Trim(newserialno) = "" Then
     newserialno = tempserialno
  End If
    txtserialno.Text = newserialno
    'lblseqdiveno.Caption = newseqdiveno
  SQL = "SELECT * FROM dpmaingaslist "
  'tempserialno = "TP00000007" 'tempserialno + 1
  SQL = SQL & " where dpmainid = '" & tempserialno & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  For i = 0 To 9
     lblgasindex(i).Caption = RS("dpgasid")
     tempnitrogen = RS("dpgasnitrogen")
     txthelium(i) = RS("dpgashelium")
     txtoxygen(i).Text = 100 - CInt(txthelium(i).Text) - CInt(tempnitrogen)
     txtmaxd(i) = RS("dpgasmaxopdepth")
     txtmaxd(i) = txtmaxd(i) / 10
     temptxthe1 = txthelium(i)
     temptxtmaxd1 = txtmaxd(i)
     Cbogasused(i) = RS("dpgasused")
     txtppo2(i).Enabled = True
     txtppo2(i).Text = RS("dpgaspo2setpoint")
     'txtppo2(i).Text = (Cint(txtoxygen(i).Text) / 100) * ((Cint(txtmaxd(i).Text) / 10) + 1)
     txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     If Cbogasused(i).Text <> "5 - Deco Closed Circuit" Then
        txtppo2(i).Enabled = False
     Else
        txtppo2(i).Enabled = True
     End If
     If Cbogasused(i).Text <> "0 - Not Used" Then
        cbogasindex.AddItem lblgasindex(i).Caption
     End If
     checkgasuse (i)
     RS.MoveNext
  Next i
  lblgasindex_Click (0)
    SQL = "SELECT COUNT(*) FROM seqdpprofile "
  SQL = SQL & " WHERE "
  SQL = SQL & " dpprofileid ='" & Trim(tempserialno) & "'"
   Set RS3 = DB.OpenRecordset(SQL)
  If RS3.Fields(0) <> 0 Then
     loaddpprofiledata
  
  End If
  txtdepth.Text = "10"
  'datachangedstatus = True
  decoresultgridlite.Visible = False
  cmdsaveas.Enabled = False
  cmdSave.Enabled = False
 End If
If tempchoice = "SPP" Or tempchoice = "GSP" Then
'  cmdplan.Visible = True
  mnuloaddesetting.Enabled = False
  mnugasloadefault.Enabled = False
  oldserialno = tempserialno
  If tempchoice = "GSP" Then
     tempseqdiveno = tempdiveserialno
  End If
  txtserialno.Text = oldserialno
  lblseqdiveno.Caption = tempseqdiveno
  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & " where dpmainid = '" & tempserialno & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  For i = 0 To 9
     p = i
     lblgasindex(i).Caption = RS("dpgasid")
     tempnitrogen = RS("dpgasnitrogen")
     txthelium(i) = RS("dpgashelium")
     txtoxygen(i).Text = 100 - CInt(txthelium(i).Text) - CInt(tempnitrogen)
     txtmaxd(i) = RS("dpgasmaxopdepth")
     txtmaxd(i) = txtmaxd(i) / 10
     temptxthe1 = txthelium(i)
     temptxtmaxd1 = txtmaxd(i)
     Cbogasused(i) = RS("dpgasused")
     txtppo2(i).Enabled = True
     
     txtppo2(i).Text = RS("dpgaspo2setpoint")
     'txtppo2(i).Text = (Cint(txtoxygen(i).Text) / 100) * ((Cint(txtmaxd(i).Text) / 10) + 1)
     txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     If Cbogasused(i).Text <> "5 - Deco Closed Circuit" Then
        txtppo2(i).Enabled = False
     Else
        txtppo2(i).Enabled = True
     End If
     If Cbogasused(i).Text <> "0 - Not Used" Then
        cbogasindex.AddItem lblgasindex(i).Caption
     End If
     checkgasuse (i)
     If InStr(1, Cbogasused(i).Text, "1") Or InStr(1, Cbogasused(i).Text, "2") Then
       Option1(i).Value = True
     End If
     RS.MoveNext
  Next i
  SQL = "SELECT COUNT(*) FROM seqdpprofile "
  SQL = SQL & " WHERE "
  SQL = SQL & " dpprofileid ='" & Trim(tempserialno) & "'"
  Set RS3 = DB.OpenRecordset(SQL)
  If RS3.Fields(0) <> 0 Then
     loaddpprofiledata
  End If
  If CInt(MSFlexGrid3.Rows) > 2 Then
     MSFlexGrid3.Row = 1
      Shape4.Visible = True
     Shape5.Visible = True
    Picture3.Height = 3600
     lbllevel_Click
  Else
     Label36.Visible = False
     Label37.Visible = False
     Label38.Visible = False
     Shape4.Visible = False
     Shape5.Visible = False
    
  End If
  
  cmdSave.Enabled = False
  cmdsaveas.Enabled = True
  rowindentified = 0
  cmdaddtoseq.Enabled = False
End If

  mnuStep.Checked = True
  mnuStep15.Checked = False
  mnuStep2.Checked = False
  laststop_index = 1
   
   SQL = "select * FROM seqdpmain "
   SQL = SQL & " where diveplanid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
     If RS3("divecategories") = "2" Then
       mnuStep.Checked = False
       mnuStep15.Checked = True
       mnuStep2.Checked = False
       laststop_index = 2
     End If
     If RS3("divecategories") = "3" Then
       mnuStep.Checked = False
       mnuStep15.Checked = False
       mnuStep2.Checked = True
       laststop_index = 3
     End If
     If IsNull(RS3("PlanBy")) Then
     Else
       stt = Mid(RS3("PlanBy"), 1, InStr(1, RS3("PlanBy"), ",a") - 1)
       vhmx_text(0).Text = stt
       stt = Mid(RS3("PlanBy"), InStr(1, RS3("PlanBy"), ",a") + 2, InStr(1, RS3("PlanBy"), ",b") - InStr(1, RS3("PlanBy"), ",a") - 2)
       vhmx_text(1).Text = stt
       stt = Mid(RS3("PlanBy"), InStr(1, RS3("PlanBy"), ",b") + 2, InStr(1, RS3("PlanBy"), ",c") - InStr(1, RS3("PlanBy"), ",b") - 2)
       vhmx_text(2).Text = stt
     End If
     RS3.MoveNext
   Wend

 If CInt(MSFlexGrid3.Rows) > 1 Then
  cmdgeneratem_Click 'Command1_Click
     MSFlexGrid3.Row = 1
    MSFlexGrid3_Click
  End If
 'Cbogasused(1).Style = 2
 If MSFlexGrid3.Rows > 3 Then
    singlelevel.Visible = False
    lbllevel.Visible = False
    cmdgeneratem.Visible = True
    
 Else
    If MSFlexGrid3.Rows = 3 Then
      singlelevel.Visible = False
      lbllevel.Visible = False
      cmdgeneratem.Visible = True
      
    Else
      If MSFlexGrid3.Rows = 2 Then
         singlelevel.Visible = False
         If buhl_mode = 2 Then Else lbllevel.Visible = True
         cmdgeneratem.Visible = False
        
      Else
         singlelevel.Visible = False
         lbllevel.Visible = False
         cmdgeneratem.Visible = False
         Label36.Visible = False
         Label37.Visible = False
         Label38.Visible = False
         Shape4.Visible = False
         Shape5.Visible = False
    
      End If
    End If
 End If
 For i = 0 To 9
    ' Cbogasused(i).Style = 2
 Next i
' cmdinsert.Visible = False
 cmdremove.Enabled = True
 If MSFlexGrid3.Rows > 2 Then
    cmdmodify.Visible = True
 Else
    cmdmodify.Visible = False
 End If
 formstarted = False
 mnugaslist_Click
 
 txtdecoalg.Text = "Deco Algorithm: " + mnuVPMB(buhl_mode).Caption
' txtdepthft.Text = "30"

  SQL = "SELECT * FROM dpserialno"
  Set RS = DB.OpenRecordset(SQL)
  If IsNull(RS!clyindersize) Then
    txtcylcap.Text = "10"
    RS.Edit
    RS!clyindersize = txtcylcap.Text
    RS.Update
  Else
    txtcylcap.Text = RS!clyindersize
  End If
  If IsNull(RS!breathingrate) Then
    txtbreathrate.Text = "10"
    RS.Edit
    RS!breathingrate = txtbreathrate.Text
    RS.Update
  Else
    txtbreathrate.Text = RS!breathingrate
  End If

  If feetormeter_feeton = 1 Then txtbreathratecuft.Text = CStr(CDbl(txtbreathrate.Text) * 0.0353)
  
  For i = 0 To 9
    update_gas_graph (i)
  Next i
  
mnuStep.Caption = Format((feetormeter_decostep * feetormeter_factor), "0") & feetormeter_shortstring
If feetormeter_factor < 2# Then
  mnuStep15.Caption = Format((feetormeter_decostep * 1.5 * feetormeter_factor), "0.0") & feetormeter_shortstring
Else
  mnuStep15.Caption = Format((feetormeter_decostep * 1.5 * feetormeter_factor), "0") & feetormeter_shortstring
End If
mnuStep2.Caption = Format((feetormeter_decostep * 2 * feetormeter_factor), "0") & feetormeter_shortstring

If IsNull(rowindentified) Then rowindentified = MSFlexGrid3.Rows - 1
If IsEmpty(rowindentified) Then rowindentified = MSFlexGrid3.Rows - 1
If rowindentified < 1 Then rowindentified = MSFlexGrid3.Rows - 1

End Sub
Private Sub loaddpprofiledata()
  SQL = "SELECT * FROM seqdpprofile"
  SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
  SQL = SQL & " order by dpnumseq "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  While RS.EOF = False
     MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
     K = MSFlexGrid3.Rows
     MSFlexGrid3.Row = K - 1
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
     MSFlexGrid3.Text = Format(CStr(CDbl(RS("depth")) * feetormeter_factor + 0.1), "###0.0")
     RS.MoveNext
   Wend
End Sub
Private Sub initialgrid()
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
MSFlexGrid3.Text = "O/C" '"Circuit"
MSFlexGrid3.Col = 7
MSFlexGrid3.Text = "Gas Index"
MSFlexGrid3.Col = 8
MSFlexGrid3.Text = "Depth"
MSFlexGrid3.ColWidth(0) = 0 '220
If feetormeter_feeton = 0 Then
  MSFlexGrid3.ColWidth(1) = 520
  MSFlexGrid3.ColWidth(8) = 0
Else
  MSFlexGrid3.ColWidth(1) = 0
  MSFlexGrid3.ColWidth(8) = 520
End If
MSFlexGrid3.ColWidth(2) = 340
MSFlexGrid3.ColWidth(3) = 300
MSFlexGrid3.ColWidth(4) = 300
MSFlexGrid3.ColWidth(5) = 410
MSFlexGrid3.ColWidth(6) = 490
MSFlexGrid3.ColWidth(7) = 620
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




Function determinexaxis()
 If zoom = 10 Then
  totalseconds = CInt(Totalcount) * CInt(txtinterval)
  totalseconds = totalseconds / zoom ' nick
  totalseconds = ((pan - 1) * totalseconds)
  totalbreak = totalseconds ' / 4
  genbreak = Format$(totalbreak, "#0")
  minutesbreak = genbreak / 60
  minutesbreak = minutesbreak - 0.499
  minutesbreak = Format$(minutesbreak, "#0")
  secondremainder = CInt(genbreak) - CInt(minutesbreak * 60)
  If CInt(minutesbreak) > 60 Then
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
    
    totalseconds = CInt(Totalcount) * CInt(txtinterval)
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
  secondremainder = CInt(genbreak) - CInt(minutesbreak * 60)
  If CInt(minutesbreak) > 60 Then
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
  
  'text3
     totalseconds = CInt(Totalcount) * CInt(txtinterval)
     If zoom = 10 Then
       totalseconds = totalseconds / zoom ' nick
       totalseconds = totalseconds / 2 + ((pan - 1) * totalseconds)
       'rtotalseconds = Cint(totalseconds) - 0.499
       totalbreak = totalseconds
     Else
       rtotalseconds = CInt(totalseconds) - 0.499
       totalbreak = rtotalseconds / 2
     End If
    genbreak = Format$(totalbreak, "#0")
    minutesbreak = genbreak / 60
    minutesbreak = minutesbreak - 0.499
    minutesbreak = Format$(minutesbreak, "#0")
    secondremainder = genbreak - (minutesbreak * 60)
    If CInt(minutesbreak) > 60 Then
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
    totalseconds = CInt(Totalcount) * CInt(txtinterval)
    If zoom = 10 Then
      totalseconds = totalseconds / zoom ' nick
      totalseconds = (totalseconds * 3 / 4) + ((pan - 1) * totalseconds)
      'rtotalseconds = Cint(totalseconds) - 0.499
      'totalbreak = rtotalseconds / 4
      totalbreak = totalseconds
    Else
      rtotalseconds = CInt(totalseconds) - 0.499
      totalbreak = rtotalseconds / 4
      totalbreak = totalbreak * 3
    End If
    genbreak = Format$(totalbreak, "#0")
    minutesbreak = genbreak / 60
    minutesbreak = minutesbreak - 0.499
    minutesbreak = Format$(minutesbreak, "#0")
     secondremainder = genbreak - (minutesbreak * 60)
    If CInt(minutesbreak) > 60 Then
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
    totalseconds = CInt(Totalcount) * CInt(txtinterval)
    If zoom = 10 Then
      totalseconds = totalseconds / zoom ' nick
      totalseconds = totalseconds + ((pan - 1) * totalseconds)
    End If
    minutesbreak = totalseconds / 60
    minutesbreak = minutesbreak - 0.499
    minutesbreak = Format$(minutesbreak, "#0")
    secondremainder = totalseconds - (minutesbreak * 60)
    If CInt(minutesbreak) > 60 Then
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
Private Sub view_graph_gaslist()
  If mnugraph.Checked = True Then
'    Picture1.Visible = True
'    Frame2.Visible = False
  Else
'    Picture1.Visible = False
'    Frame2.Visible = True
  End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then Me.WindowState = 0
If Me.WindowState = 0 Then
  Me.Width = 12660
  Me.Height = 10140
End If
End Sub

Private Sub lbllevel_Click()
Label35.Top = 240
txtserialno.Top = 630
Shape4.Visible = True
Shape5.Visible = True
txtserialno.BackColor = &H400000
txtserialno.ForeColor = vbWhite
Image1.Visible = False
Picture3.Top = 410
Picture3.Height = 3600
lblgasvr.Top = 170
lblgasvr.Left = 2590
lblgasvr.ForeColor = vbWhite
lblgasvr.BackColor = &H400000
Label36.Visible = True
Label37.Visible = True
Label38.Visible = True
Label11.Visible = False
cmddepthup.Left = 400
cmddepthdown.Left = 650
cmddepthup.Top = 650
cmddepthdown.Top = 650
txtdepth.Left = 280
txtdepthft.Left = 280
txtdepth.Top = 850
txtdepthft.Top = 850
txtdepth.Width = 1100
txtdepthft.Width = 1100
cmdtimeup.Top = 1590
cmdtimedown.Top = 1590
cmdtimeup.Left = 400
cmdtimedown.Left = 650
txttime.Left = 280
txttime.Width = 1100
txttime.Top = 1790
txtppo2v.Left = 280
txtppo2v.Width = 1100
txtppo2v.Top = 2750
txtppo2v.FontSize = 22
'txtppo2v.Height = 315
CMDPPO2PLUS.Left = 400
CMDPPO2PLUS.Top = 2550
Command4.Left = 650
Command4.Top = 2550

Cmdadd.Visible = True
cmdinsert.Visible = True
cmdmodify.Visible = True

cmdremove.Visible = True
cmdclearall.Visible = True
MSFlexGrid3.Visible = True
MSFlexGrid3.Left = 2570
MSFlexGrid3.Top = 700
Label1.Left = 1440
lblminutes.Left = 2280
'Label29.Left = 1320
'Label29.Top = 2220
'Option4.Left = 840
'Option4.Top = 2640
''Option3.Left = 1080
'Option3.Top = 2880
'lblgasvr.Left = 1440
cmdgenerate.Visible = False
If MSFlexGrid3.Rows < 3 Then
   singlelevel.Visible = True
End If
cmdgeneratem.Visible = True
lbllevel.Visible = False
End Sub

Private Sub mnufilesave_Click()
  ans_save = 1
  Unload Me
End Sub

Private Sub mnufilesaveas_Click()
Dim comptext As String
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim p As Integer
Dim q As Integer

 On Error GoTo ErrorHandler2
    CommonDialog3.Action = 2
 Open CommonDialog3.FileName For Output As #1
  comptext = "DO NOT DIVE USING THESE TABLES. BETA SOFTWARE TESTING ONLY"
   Print #1, comptext

   comptext = "Dive Plan No : " & txtserialno.Text
   Print #1, comptext

   comptext = "  Atmospheric : " & atmtext.Text
   Print #1, comptext
   comptext = "  Safety : " & safetytext.Text
   Print #1, comptext
   comptext = txtdecoalg.Text
   Print #1, comptext
   comptext = ""
   Print #1, comptext

   comptext = "Sequence Of the Dive Plan : "
   Print #1, comptext
   comptext = ""
   If feetormeter_feeton = 1 Then
      For K = 0 To MSFlexGrid3.Rows - 1
        For j = 1 To MSFlexGrid3.Cols - 1
           MSFlexGrid3.Row = K
           MSFlexGrid3.Col = j
           rowtext = MSFlexGrid3.Text
           comptext = comptext + "   " + (rowtext)
        Next j
        Print #1, comptext
        comptext = ""
      Next K
    Else
       For K = 0 To MSFlexGrid3.Rows - 1
         For j = 0 To MSFlexGrid3.Cols - 2
            MSFlexGrid3.Row = K
            MSFlexGrid3.Col = j
            rowtext = MSFlexGrid3.Text
            comptext = comptext + "   " + (rowtext)
          Next j
          Print #1, comptext
          comptext = ""
        Next K
    End If
   Print #1, comptext

   comptext = ""
   Print #1, comptext
    
    comptext = "Active" & vbTab & "Gas #" & vbTab & "O2" & vbTab & "He" & vbTab & "Depth" & vbTab & "PPO2" & vbTab & "CC" & vbTab & "Deco" & vbTab & "WC" & vbTab & "SAC" & "BarUse"
    Print #1, comptext
    comptext = ""
    For v = 0 To 9
      If Check1(v).Value = 1 Then
         temptext = "On"
      Else
         temptext = "Off"
      End If
      
      temptext2 = "Gas " + CStr(v)
      temptext3 = txtoxygen(v).Text
      temptext4 = txthelium(v).Text
      temptext5 = txtmaxdft(v).Text 'lbldepth(v).Caption
      temptext6 = txtppo2(v).Text
      If Check2(v).Value = 1 Then
         temptext7 = "CC"
      Else
         temptext7 = "OC"
      End If
      If Decochk(v).Value = 1 Then
         temptext8 = "Deco"
      Else
         temptext8 = ""
      End If
      If txtcylcap2(v).Visible = True Then
         temptext9 = txtcylcap2(v).Text
      Else
         temptext9 = ""
      End If
      If txtbreathrate2(v).Visible = True Then
         temptext10 = txtbreathrate2(v).Text
      Else
         temptext10 = ""
      End If
      If gasusage(v).Visible = True Then
         temptext11 = gasusage(v).Caption
      Else
           temptext11 = ""
      End If
      comptext = temptext & vbTab & temptext2 & vbTab & temptext3 & vbTab & temptext4 & vbTab & temptext5 & vbTab & temptext6
      comptext = comptext & vbTab & temptext7 & vbTab & temptext8 & vbTab & temptext9 & vbTab & temptext10 & vbTab & temptext11
      Print #1, comptext
      comptext = ""
    Next v
    comptext = ""
    Print #1, comptext
    'comptext = " " & Frame3.Caption
    'text2.text=text2.text +  comptext
    comptext = ""
    Print #1, comptext
    comptext = ""
    If systemversion = "Pro" Then
    For K = 0 To decoresultgrid.Rows - 1
      For p = 0 To decoresultgrid.Cols - 1
        decoresultgrid.Row = K
        decoresultgrid.Col = p
        rowtext = CStr(decoresultgrid.Text) 'Format(decoresultgrid.Text, "")
        rowtext = Left(rowtext, 8)
        If Len(rowtext) < 7 Then
          rowtext = rowtext + vbTab
        End If
        comptext = comptext + (rowtext + vbTab)
      Next p
     
      Print #1, comptext
      comptext = ""
    Next K
    Else
    For K = 0 To decoresultgridlite.Rows - 1
      For p = 0 To decoresultgridlite.Cols - 1
        decoresultgridlite.Row = K
        decoresultgridlite.Col = p
        rowtext = CStr(decoresultgridlite.Text) 'Format(decoresultgrid.Text, "")
        rowtext = Left(rowtext, 8)
        If Len(rowtext) < 7 Then
          rowtext = rowtext + vbTab
        End If
        comptext = comptext + (rowtext + vbTab)
      Next p
     
      Print #1, comptext
      comptext = ""
    Next K
    End If
    
  Close #1
  MsgBox "Data saved to CSV file....!!"

ErrorHandler2:
    If Err = 32755 Then
        MsgBox "Data is not saved to a file....!!"
        Exit Sub
    End If
End Sub

Private Sub mnugaslist_Click()
  mnugraph.Checked = False
  mnugaslist.Checked = True
  view_graph_gaslist
End Sub

Private Sub mnugraph_Click()
 mnugraph.Checked = True
  mnugaslist.Checked = False
  view_graph_gaslist
End Sub




Private Sub Textchange_GotFocus()
MSFlexGrid3.Text = textchange.Text
  If cols0activated <> "true" Then
   ChangeCellText
  End If
End Sub
Public Sub ChangeCellText() ' Move Textbox to active cell.
End Sub

Private Sub Form_Unload(Cancel As Integer)

  SQL = "SELECT * FROM dpserialno"
  Set RS = DB.OpenRecordset(SQL)
    RS.Edit
    RS!clyindersize = txtcylcap.Text
    RS.Update
    RS.Edit
    RS!breathingrate = txtbreathrate.Text
    RS.Update

If datachangedstatus = True And no_deco_found = 0 Then
ans_save = 1 'added to force save without ambiguity
If ans_save = 0 Then
  Title = "Error on System Validation.."
  ans = MsgBox("Data changed, but not yet saved, " & Chr(13) & "Press No will not save changes !" & Chr(13) & "Do you want to save now ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Else
  ans = vbYes
End If
Select Case ans
Case vbYes
   cmdsave_Click
Case vbNo
   deleteseqdpmain
  Unload Me
   Select Case previousform
   Case "SEQLIST"
      Splanmain.Show
   Case "SEQPLAN"
      frmseqdive.Show
   End Select
End Select

Else
   deleteseqdpmain2
   Select Case previousform
   Case "SEQLIST"
      If do_not_load = 1 Then Exit Sub
      Splanmain.Show
   Case "SEQPLAN"
      frmseqdive.Show
   End Select
'Unload Planprofile2

End If

End Sub
Private Sub deleteseqdpmain2()
SQL = "select * FROM seqdpmain "
SQL = SQL & "order by DIVEPLANID "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("diveplanid")
   If tempdpid Like "T*" Then
      RS3.Delete
   End If
    RS3.MoveNext
Wend
RS3.Close

SQL = "select * FROM dpmaingaslist "
SQL = SQL & "order by dpmainid "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("dpmainid")
   If tempdpid Like "T*" Then
      RS3.Delete
   End If
    RS3.MoveNext
Wend
RS3.Close
SQL = "select * FROM seqdpprofile "
SQL = SQL & "order by dpprofileid "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("dpprofileid")
   If tempdpid Like "T*" Then
      RS3.Delete
   End If
    RS3.MoveNext
Wend
RS3.Close
End Sub
Private Sub deleteseqdpmain()
SQL = "select * FROM seqdpmain "
SQL = SQL & "order by DIVEPLANID "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("diveplanid")
   If tempdpid Like "T*" Then
      RS3.Delete
   End If
    RS3.MoveNext
Wend
RS3.Close

SQL = "select * FROM dpmaingaslist "
SQL = SQL & "order by dpmainid "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("dpmainid")
   If tempdpid Like "T*" Then
      RS3.Delete
   End If
    RS3.MoveNext
Wend
RS3.Close
SQL = "select * FROM seqdpprofile "
SQL = SQL & "order by dpprofileid "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("dpprofileid")
   If tempdpid Like "T*" Then
      RS3.Delete
   End If
    RS3.MoveNext
Wend
RS3.Close
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
  barused(vmix_vnumber) = barused(vmix_vnumber) + absolutedepthpure * exposuretime

End Function
Private Sub printtogrid()
decoresultgrid.Rows = decoresultgrid.Rows + 1
decoresultgrid.Row = decoresultgrid.Rows - 1
'MSFlexGrid3.RowHeight = 100
decoresultgrid.Col = 7
decoresultgrid.Text = CStr(vsegment_vnumber)
decoresultgrid.Col = 1
decoresultgrid.Text = Format(vsegment_vtime, "###0.0")
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 2
decoresultgrid.Text = Format(run_vtime, "###0.0")
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 3
'decoresultgrid.Text = CStr(vmix_vnumber - 1)
If IsEmpty(vmix_vnumber) Then
Else
  If txthelium(vmix_vnumber - 1).Text = "0" Then
    If txtoxygen(vmix_vnumber - 1).Text = "21" Then
      decoresultgrid.Text = "  Air"
    Else
      decoresultgrid.Text = " Nx" + CStr(txtoxygen(vmix_vnumber - 1).Text)
    End If
  Else
    decoresultgrid.Text = "TX" + CStr(txtoxygen(vmix_vnumber - 1).Text + "/" + txthelium(vmix_vnumber - 1).Text)
  End If
End If
decoresultgrid.Col = 9
decoresultgrid.Text = CStr(vmix_vnumber)
decoresultgrid.Col = 0
decoresultgrid.Text = Format(ending_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If ending_vdepth = 4.5 Then decoresultgrid.Text = Format(ending_vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'dum = ppo2exposuretime(ending_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
decoresultgrid.Col = 5
decoresultgrid.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
decoresultgrid.Col = 6
decoresultgrid.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
decoresultgrid.Col = 8
decoresultgrid.Text = Format(rate * feetormeter_factor, "###")
If rate > 0.01 Then
  decoresultgrid.Text = "Descent        " + decoresultgrid.Text
Else
  If rate < -0.01 Then
    decoresultgrid.Text = "Ascent        " + decoresultgrid.Text
  Else
    decoresultgrid.Text = " ---- "
  End If
End If
decoresultgrid.Col = 4
If IsNumeric(SetPoint) Then
  If SetPoint < 0.21 Then decoresultgrid.Text = " " Else decoresultgrid.Text = Format((SetPoint), "0.00") + " "
Else
  decoresultgrid.Text = " "
End If
End Sub
Private Sub printtogridlite()

If rate > 0.01 Then
  decoresultgridlite.Rows = decoresultgridlite.Rows + 1
  decoresultgridlite.Row = decoresultgridlite.Rows - 1
  decoresultgridlite.Col = 8
  decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
  decoresultgridlite.Text = "Descent        " + decoresultgridlite.Text
Else
  If rate < -0.01 Then
    decoresultgridlite.Rows = decoresultgridlite.Rows + 1
    decoresultgridlite.Row = decoresultgridlite.Rows - 1
    decoresultgridlite.Col = 8
    decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
    decoresultgridlite.Text = "Ascent        " + decoresultgridlite.Text
 ' Else
 '   decoresultgridlite.Text = " ---- "
  End If
End If
'decoresultgridlite.Rows = decoresultgridlite.Rows + 1
'decoresultgridlite.Row = decoresultgridlite.Rows - 1
'MSFlexGrid3.RowHeight = 100
decoresultgridlite.Col = 7
decoresultgridlite.Text = CStr(vsegment_vnumber)
'decoresultgridlite.Col = 8
'decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
decoresultgridlite.Col = 1
If rate > 0.01 Then
  decoresultgridlite.Text = Format(vsegment_vtime, "###0.0")
  decoresultgridlite.Text = Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite.Text, 2)) * 60#), "00")
Else
  If rate < -0.01 Then
    decoresultgridlite.Text = Format(vsegment_vtime, "###0.0")
    decoresultgridlite.Text = Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite.Text, 2)) * 60#), "00")
  Else
  End If
End If
decoresultgridlite.Col = 2
decoresultgridlite.Text = Format(run_vtime, "###0.0")
decoresultgridlite.Text = Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite.Text, 2)) * 60#), "00")
decoresultgridlite.Col = 3
'decoresultgridlite.Text = CStr(vmix_vnumber - 1)
If IsEmpty(vmix_vnumber) Then
Else
  If txthelium(vmix_vnumber - 1).Text = "0" Then
    If txtoxygen(vmix_vnumber - 1).Text = "21" Then
      decoresultgridlite.Text = "  Air"
    Else
      decoresultgridlite.Text = " Nx" + CStr(txtoxygen(vmix_vnumber - 1).Text)
    End If
  Else
    decoresultgridlite.Text = "TX" + CStr(txtoxygen(vmix_vnumber - 1).Text + "/" + txthelium(vmix_vnumber - 1).Text)
  End If
End If
decoresultgridlite.Col = 0
decoresultgridlite.Text = Format(ending_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If ending_vdepth = 4.5 Then decoresultgridlite.Text = Format(ending_vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
dum = ppo2exposuretime(ending_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
decoresultgridlite.Col = 5
decoresultgridlite.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
decoresultgridlite.Col = 6
decoresultgridlite.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'decoresultgridlite.Col = 8
'decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
'If rate > 0.01 Then
'  decoresultgridlite.Text = "Descent        " + decoresultgridlite.Text
'Else
'  If rate < -0.01 Then
'    decoresultgridlite.Text = "Ascent        " + decoresultgridlite.Text
'  Else
'    decoresultgridlite.Text = " ---- "
'  End If
'End If
decoresultgridlite.Col = 4
If IsNumeric(SetPoint) Then
  If SetPoint < 0.21 Then decoresultgridlite.Text = " " Else decoresultgridlite.Text = Format((SetPoint), "0.00") + " "
Else
  decoresultgridlite.Text = " "
End If
End Sub
Private Sub printtogrid3()
decoresultgrid.Rows = decoresultgrid.Rows + 1
decoresultgrid.Row = decoresultgrid.Rows - 1

decoresultgrid.Cols = 10
decoresultgrid.Col = 7
decoresultgrid.Text = "No."
decoresultgrid.Col = 1
decoresultgrid.Text = "Duration"
decoresultgrid.Col = 2
decoresultgrid.Text = "RunTime"
decoresultgrid.Col = 3
decoresultgrid.Text = "Mix"
decoresultgrid.Col = 0
decoresultgrid.Text = "Depth"
decoresultgrid.Col = 5
decoresultgrid.Text = "CNS"
decoresultgrid.Col = 6
decoresultgrid.Text = "OTU"
decoresultgrid.Col = 8
decoresultgrid.Text = "Rate"
decoresultgrid.Col = 4
decoresultgrid.Text = "Set Point"

For K = 0 To 8
  decoresultgrid.Col = K
  decoresultgrid.CellBackColor = &H8000000F
Next K
End Sub
Private Sub printtogrid4()
decoresultgrid.Rows = decoresultgrid.Rows + 1
decoresultgrid.Row = decoresultgrid.Rows - 1
decoresultgrid.Col = 7
decoresultgrid.Text = CStr(vsegment_vnumber)
decoresultgrid.Col = 1
decoresultgrid.Text = Format(vsegment_vtime, "###0.0")
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 2
decoresultgrid.Text = Format(run_vtime, "###0.0")
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 3
If IsEmpty(vmix_vnumber) Then
Else
  If txthelium(vmix_vnumber - 1).Text = "0" Then
    If txtoxygen(vmix_vnumber - 1).Text = "21" Then
      decoresultgrid.Text = "  Air"
    Else
      decoresultgrid.Text = " Nx" + CStr(txtoxygen(vmix_vnumber - 1).Text)
    End If
  Else
    decoresultgrid.Text = "TX" + CStr(txtoxygen(vmix_vnumber - 1).Text + "/" + txthelium(vmix_vnumber - 1).Text)
  End If
End If
decoresultgrid.Col = 9
decoresultgrid.Text = CStr(vmix_vnumber)
decoresultgrid.Col = 0
decoresultgrid.Text = Format(vdeco_vstop_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then decoresultgrid.Text = Format(vdeco_vstop_vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'dum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
decoresultgrid.Col = 5
decoresultgrid.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
decoresultgrid.Col = 6
decoresultgrid.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
decoresultgrid.Col = 8
decoresultgrid.Text = Format(rate * feetormeter_factor, "###")
If rate > 0.01 Then
  decoresultgrid.Text = "Descent        " + decoresultgrid.Text
Else
  If rate < -0.01 Then
    decoresultgrid.Text = "Ascent        " + decoresultgrid.Text
  Else
    decoresultgrid.Text = " ---- "
  End If
End If
decoresultgrid.Col = 4
If IsNumeric(SetPoint) Then
  If SetPoint < 0.21 Then decoresultgrid.Text = " " Else decoresultgrid.Text = Format((SetPoint), "0.00") + " "
Else
  decoresultgrid.Text = " "
End If
For K = 0 To 8
  decoresultgrid.Col = K
  decoresultgrid.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid4lite()

If rate > 0.01 Then
  decoresultgridlite.Rows = decoresultgridlite.Rows + 1
  decoresultgridlite.Row = decoresultgridlite.Rows - 1
  decoresultgridlite.Col = 8
  decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
  decoresultgridlite.Text = "Descent        " + decoresultgridlite.Text
Else
  If rate < -0.01 Then
    decoresultgridlite.Rows = decoresultgridlite.Rows + 1
    decoresultgridlite.Row = decoresultgridlite.Rows - 1
    decoresultgridlite.Col = 8
    decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
    decoresultgridlite.Text = "Ascent        " + decoresultgridlite.Text
 ' Else
 '   decoresultgridlite.Text = " ---- "
  End If
End If


'decoresultgridlite.Rows = decoresultgridlite.Rows + 1
'decoresultgridlite.Row = decoresultgridlite.Rows - 1
decoresultgridlite.Col = 7
decoresultgridlite.Text = CStr(vsegment_vnumber)
decoresultgridlite.Col = 1
tempdurationlite = decoresultgridlite.Text
 decoresultgridlite.Text = Format(vsegment_vtime, "###0.0")
decoresultgridlite.Text = Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite.Text, 2)) * 60#), "00")
tempdurationlitesec = Right(tempdurationlite, 2)
Select Case Len(tempdurationlite)
Case 4
   tempdurationlitemins = Left(tempdurationlite, 1)
Case 5
   tempdurationlitemins = Left(tempdurationlite, 2)
Case 6
   tempdurationlitemins = Left(tempdurationlite, 3)
Case 7
   tempdurationlitemins = Left(tempdurationlite, 4)
End Select
tempgridlitesec = Right(decoresultgridlite.Text, 2)
Select Case Len(decoresultgridlite.Text)
Case 4
   tempgridlitemins = Left(decoresultgridlite.Text, 1)
Case 5
   tempgridlitemins = Left(decoresultgridlite.Text, 2)
Case 6
   tempgridlitemins = Left(decoresultgridlite.Text, 3)
Case 7
   tempgridlitemins = Left(decoresultgridlite.Text, 4)
End Select
If Val(tempgridlitesec) + Val(tempdurationlitesec) > 59 Then
   templeftsec = (CInt(tempgridlitesec) + CInt(tempdurationlitesec)) - 60
   templeftsec = Format(CDbl(templeftsec), "00")
   tempgridlitemins = CInt(tempgridlitemins) + 1
Else
   templeftsec = Format(CDbl(Val(tempgridlitesec) + Val(tempdurationlitesec)), "00")
End If
totalmins = CInt(tempgridlitemins) + CInt(tempdurationlitemins)
decoresultgridlite.Text = CDbl(totalmins) & ":" & templeftsec
decoresultgridlite.Col = 2
decoresultgridlite.Text = Format(run_vtime, "###0.0")
decoresultgridlite.Text = Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite.Text, 2)) * 60#), "00")
decoresultgridlite.Col = 3
If IsEmpty(vmix_vnumber) Then
Else
  If txthelium(vmix_vnumber - 1).Text = "0" Then
    If txtoxygen(vmix_vnumber - 1).Text = "21" Then
      decoresultgridlite.Text = "  Air"
    Else
      decoresultgridlite.Text = " Nx" + CStr(txtoxygen(vmix_vnumber - 1).Text)
    End If
  Else
    decoresultgridlite.Text = "TX" + CStr(txtoxygen(vmix_vnumber - 1).Text + "/" + txthelium(vmix_vnumber - 1).Text)
  End If
End If
decoresultgridlite.Col = 0
decoresultgridlite.Text = Format(vdeco_vstop_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then decoresultgridlite.Text = Format(vdeco_vstop_vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring) 'Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'dum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
decoresultgridlite.Col = 5
decoresultgridlite.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
decoresultgridlite.Col = 6
decoresultgridlite.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'decoresultgridlite.Col = 8
'decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
'If rate > 0.01 Then
'  decoresultgridlite.Text = "Descent        " + decoresultgridlite.Text
'Else
'  If rate < -0.01 Then
'    decoresultgridlite.Text = "Ascent        " + decoresultgridlite.Text
'  Else
'    decoresultgridlite.Text = " ---- "
'  End If
'End If
decoresultgridlite.Col = 4
If IsNumeric(SetPoint) Then
  If SetPoint < 0.21 Then decoresultgridlite.Text = " " Else decoresultgridlite.Text = Format((SetPoint), "0.00") + " "
Else
  decoresultgridlite.Text = " "
End If
For K = 0 To 8
  decoresultgridlite.Col = K
  decoresultgridlite.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid5()
decoresultgrid.Rows = decoresultgrid.Rows + 1
decoresultgrid.Row = decoresultgrid.Rows - 1
decoresultgrid.Col = 7
decoresultgrid.Text = CStr(vsegment_vnumber)
decoresultgrid.Col = 1
decoresultgrid.Text = Format(vsegment_vtime, "###0.0") 'CStr((CDbl(CInt(vsegment_vtime * 10# + 0.4999) / 10)))
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 2
decoresultgrid.Text = Format(run_vtime, "###0.0") 'CStr((CDbl(CInt(run_vtime * 10# + 0.4999) / 10)))
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 3
If IsEmpty(vmix_vnumber) Then
Else
  If txthelium(vmix_vnumber - 1).Text = "0" Then
    If txtoxygen(vmix_vnumber - 1).Text = "21" Then
      decoresultgrid.Text = "  Air"
    Else
      decoresultgrid.Text = " Nx" + CStr(txtoxygen(vmix_vnumber - 1).Text)
    End If
  Else
    decoresultgrid.Text = "TX" + CStr(txtoxygen(vmix_vnumber - 1).Text + "/" + txthelium(vmix_vnumber - 1).Text)
  End If
End If
decoresultgrid.Col = 9
decoresultgrid.Text = CStr(vmix_vnumber)
decoresultgrid.Col = 0
decoresultgrid.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.4999) / 10) * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then decoresultgrid.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.4999) / 10) * feetormeter_factor, "###0.0" & feetormeter_shortstring)
'zzzdum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
decoresultgrid.Col = 5
decoresultgrid.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
decoresultgrid.Col = 6
decoresultgrid.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
decoresultgrid.Col = 8
decoresultgrid.Text = Format(rate * feetormeter_factor, "###")
If rate > 0.01 Then
  decoresultgrid.Text = "Descent        " + decoresultgrid.Text
Else
  If rate < -0.01 Then
    decoresultgrid.Text = "Ascent        " + decoresultgrid.Text
  Else
    decoresultgrid.Text = " ---- "
  End If
End If
decoresultgrid.Col = 4
If IsNumeric(SetPoint) Then
  If SetPoint < 0.21 Then decoresultgrid.Text = " " Else decoresultgrid.Text = Format((SetPoint), "0.00") + " "
Else
  decoresultgrid.Text = " "
End If
For K = 0 To 8
  decoresultgrid.Col = K
  decoresultgrid.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid5lite()
If rate > 0.01 Then
  decoresultgridlite.Rows = decoresultgridlite.Rows + 1
  decoresultgridlite.Row = decoresultgridlite.Rows - 1
  decoresultgridlite.Col = 8
  decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
  decoresultgridlite.Text = "Descent        " + decoresultgridlite.Text
Else
  If rate < -0.01 Then
    decoresultgridlite.Rows = decoresultgridlite.Rows + 1
    decoresultgridlite.Row = decoresultgridlite.Rows - 1
    decoresultgridlite.Col = 8
    decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
    decoresultgridlite.Text = "Ascent        " + decoresultgridlite.Text
 ' Else
 '   decoresultgridlite.Text = " ---- "
  End If
End If


'decoresultgridlite.Rows = decoresultgridlite.Rows + 1
'decoresultgridlite.Row = decoresultgridlite.Rows - 1
decoresultgridlite.Col = 7
decoresultgridlite.Text = CStr(vsegment_vnumber)
decoresultgridlite.Col = 1
tempdurationlite = decoresultgridlite.Text
 decoresultgridlite.Text = Format(vsegment_vtime, "###0.0")
decoresultgridlite.Text = Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite.Text, 2)) * 60#), "00")
tempdurationlitesec = Right(tempdurationlite, 2)
Select Case Len(tempdurationlite)
Case 4
   tempdurationlitemins = Left(tempdurationlite, 1)
Case 5
   tempdurationlitemins = Left(tempdurationlite, 2)
Case 6
   tempdurationlitemins = Left(tempdurationlite, 3)
Case 7
   tempdurationlitemins = Left(tempdurationlite, 4)
End Select
tempgridlitesec = Right(decoresultgridlite.Text, 2)
Select Case Len(decoresultgridlite.Text)
Case 4
   tempgridlitemins = Left(decoresultgridlite.Text, 1)
Case 5
   tempgridlitemins = Left(decoresultgridlite.Text, 2)
Case 6
   tempgridlitemins = Left(decoresultgridlite.Text, 3)
Case 7
   tempgridlitemins = Left(decoresultgridlite.Text, 4)
End Select
If Val(tempgridlitesec) + Val(tempdurationlitesec) > 59 Then
   templeftsec = (CInt(tempgridlitesec) + CInt(tempdurationlitesec)) - 60
   templeftsec = Format(CDbl(templeftsec), "00")
   tempgridlitemins = CInt(tempgridlitemins) + 1
Else
   templeftsec = Format(CDbl(Val(tempgridlitesec) + Val(tempdurationlitesec)), "00")
End If
totalmins = CInt(tempgridlitemins) + CInt(tempdurationlitemins)
decoresultgridlite.Text = CDbl(totalmins) & ":" & templeftsec
decoresultgridlite.Col = 2
decoresultgridlite.Text = Format(run_vtime, "###0.0") 'CStr((CDbl(CInt(run_vtime * 10# + 0.4999) / 10)))
decoresultgridlite.Text = Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite.Text, 2)) * 60#), "00")
decoresultgridlite.Col = 3
If IsEmpty(vmix_vnumber) Then
Else
  If txthelium(vmix_vnumber - 1).Text = "0" Then
    If txtoxygen(vmix_vnumber - 1).Text = "21" Then
      decoresultgridlite.Text = "  Air"
    Else
      decoresultgridlite.Text = " Nx" + CStr(txtoxygen(vmix_vnumber - 1).Text)
    End If
  Else
    decoresultgridlite.Text = "TX" + CStr(txtoxygen(vmix_vnumber - 1).Text + "/" + txthelium(vmix_vnumber - 1).Text)
  End If
End If
decoresultgridlite.Col = 0
decoresultgridlite.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.4999) / 10) * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then decoresultgridlite.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.4999) / 10) * feetormeter_factor, "###0.0" & feetormeter_shortstring)
dum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
decoresultgridlite.Col = 5
decoresultgridlite.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
decoresultgridlite.Col = 6
decoresultgridlite.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'decoresultgridlite.Col = 8
'decoresultgridlite.Text = Format(rate * feetormeter_factor, "###")
'If rate > 0.01 Then
'  decoresultgridlite.Text = "Descent        " + decoresultgridlite.Text
'Else
'  If rate < -0.01 Then
'    decoresultgridlite.Text = "Ascent        " + decoresultgridlite.Text
'  Else
'    decoresultgridlite.Text = " ---- "
'  End If
'End If
decoresultgridlite.Col = 4
If IsNumeric(SetPoint) Then
  If SetPoint < 0.21 Then decoresultgridlite.Text = " " Else decoresultgridlite.Text = Format((SetPoint), "0.00") + " "
Else
  decoresultgridlite.Text = " "
End If
For K = 0 To 8
  decoresultgridlite.Col = K
  decoresultgridlite.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid6()
decoresultgrid.Rows = decoresultgrid.Rows + 1
decoresultgrid.Row = decoresultgrid.Rows - 1
decoresultgrid.Col = 7
decoresultgrid.Text = CStr(vsegment_vnumber)
decoresultgrid.Col = 1
decoresultgrid.Text = Format(vsegment_vtime, "###0.0") 'CStr((CDbl(CInt(vsegment_vtime * 10# + 0.999) / 10)))
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 2
decoresultgrid.Text = Format(run_vtime, "###0.0") 'CStr((CDbl(CInt(run_vtime * 10# + 0.999) / 10)))
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 3
If IsEmpty(vmix_vnumber) Then
Else
  If txthelium(vmix_vnumber - 1).Text = "0" Then
    If txtoxygen(vmix_vnumber - 1).Text = "21" Then
      decoresultgrid.Text = "  Air"
    Else
      decoresultgrid.Text = " Nx" + CStr(txtoxygen(vmix_vnumber - 1).Text)
    End If
  Else
    decoresultgrid.Text = "TX" + CStr(txtoxygen(vmix_vnumber - 1).Text + "/" + txthelium(vmix_vnumber - 1).Text)
  End If
End If
decoresultgrid.Col = 9
decoresultgrid.Text = CStr(vmix_vnumber)
decoresultgrid.Col = 0
decoresultgrid.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.999) / 10) * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdeco_vstop_vdepth = 4.5 Then decoresultgrid.Text = Format(CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.4999) / 10) * feetormeter_factor, "###0.0" & feetormeter_shortstring)
dum = ppo2exposuretime(vdeco_vstop_vdepth, vsegment_vtime) '
decoresultgrid.Col = 5
decoresultgrid.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
decoresultgrid.Col = 6
decoresultgrid.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
decoresultgrid.Col = 8
decoresultgrid.Text = Format(rate * feetormeter_factor, "###")
If rate > 0.01 Then
  decoresultgrid.Text = "Descent        " + decoresultgrid.Text
Else
  If rate < -0.01 Then
    decoresultgrid.Text = "Ascent        " + decoresultgrid.Text
  Else
    decoresultgrid.Text = " ---- "
  End If
End If
decoresultgrid.Col = 4
If IsNumeric(SetPoint) Then
  If SetPoint < 0.21 Then decoresultgrid.Text = " " Else decoresultgrid.Text = Format((SetPoint), "0.00") + " "
Else
  decoresultgrid.Text = " "
End If
For K = 0 To 8
  decoresultgrid.Col = K
  decoresultgrid.CellBackColor = vbYellow
Next K
End Sub
Private Sub printtogrid2()
decoresultgrid.Rows = decoresultgrid.Rows + 1
decoresultgrid.Row = decoresultgrid.Rows - 1
decoresultgrid.Col = 7
decoresultgrid.Text = CStr(vsegment_vnumber)
decoresultgrid.Col = 1
decoresultgrid.Text = Format(vsegment_vtime, "###0.0")
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 2
decoresultgrid.Text = Format(run_vtime, "###0.0")
decoresultgrid.Text = Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 2) + ":" + Format((CDbl(Right(decoresultgrid.Text, 2)) * 60#), "00")
decoresultgrid.Col = 3
If IsEmpty(vmix_vnumber) Then
Else
  If txthelium(vmix_vnumber - 1).Text = "0" Then
    If txtoxygen(vmix_vnumber - 1).Text = "21" Then
      decoresultgrid.Text = "  Air"
    Else
      decoresultgrid.Text = " Nx" + CStr(txtoxygen(vmix_vnumber - 1).Text)
    End If
  Else
    decoresultgrid.Text = "TX" + CStr(txtoxygen(vmix_vnumber - 1).Text + "/" + txthelium(vmix_vnumber - 1).Text)
  End If
End If
decoresultgrid.Col = 9
decoresultgrid.Text = CStr(vmix_vnumber)
decoresultgrid.Col = 0
decoresultgrid.Text = Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdepth = 4.5 Then decoresultgrid.Text = Format(vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring)
'yyydum = ppo2exposuretime(vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
decoresultgrid.Col = 5
decoresultgrid.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
decoresultgrid.Col = 6
decoresultgrid.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
decoresultgrid.Col = 4
If IsNumeric(SetPoint) Then
  If SetPoint < 0.21 Then decoresultgrid.Text = " " Else decoresultgrid.Text = Format((SetPoint), "0.00") + " "
Else
  decoresultgrid.Text = " "
End If
decoresultgrid.Col = 8
'decoresultgrid.Text = format(rate*feetormeter_factor,"###")
'If rate > 0.01 Then
'  decoresultgrid.text = "Descent        " + decoresultgrid.Text
'Else
'  If rate < -0.01 Then
'    decoresultgrid.text = "Ascent        " + decoresultgrid.Text
'  Else
    decoresultgrid.Text = " ---- "
'  End If
'End If

End Sub

Private Sub printtogrid2lite()
decoresultgridlite.Rows = decoresultgridlite.Rows + 1
decoresultgridlite.Row = decoresultgridlite.Rows - 1
decoresultgridlite.Col = 7
decoresultgridlite.Text = CStr(vsegment_vnumber)
decoresultgridlite.Col = 1
decoresultgridlite.Text = Format(vsegment_vtime, "###0.0")
decoresultgridlite.Text = Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite.Text, 2)) * 60#), "00")
decoresultgridlite.Col = 2
decoresultgridlite.Text = Format(run_vtime, "###0.0")
decoresultgridlite.Text = Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 2) + ":" + Format((CDbl(Right(decoresultgridlite.Text, 2)) * 60#), "00")
decoresultgridlite.Col = 3
If IsEmpty(vmix_vnumber) Then
Else
  If txthelium(vmix_vnumber - 1).Text = "0" Then
    If txtoxygen(vmix_vnumber - 1).Text = "21" Then
      decoresultgridlite.Text = "  Air"
    Else
      decoresultgridlite.Text = " Nx" + CStr(txtoxygen(vmix_vnumber - 1).Text)
    End If
  Else
    decoresultgridlite.Text = "TX" + CStr(txtoxygen(vmix_vnumber - 1).Text + "/" + txthelium(vmix_vnumber - 1).Text)
  End If
End If
decoresultgridlite.Col = 0
decoresultgridlite.Text = Format(vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
If vdepth = 4.5 Then decoresultgridlite.Text = Format(vdepth * feetormeter_factor, "###0.0" & feetormeter_shortstring)
xxxdum = ppo2exposuretime(vdepth, vsegment_vtime) 'cns_current = cns_current + (vdepth * vsegment_vtime)
decoresultgridlite.Col = 5
decoresultgridlite.Text = Format(cns_current, "###0") & "%" 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
'otu_current = otu_current + (vdepth * vsegment_vtime)
decoresultgridlite.Col = 6
decoresultgridlite.Text = Format(otu_current, "###0") & " " 'Format(starting_vdepth * feetormeter_factor, "###0" & feetormeter_shortstring)
decoresultgridlite.Col = 4
If IsNumeric(SetPoint) Then
  If SetPoint < 0.21 Then decoresultgridlite.Text = " " Else decoresultgridlite.Text = Format((SetPoint), "0.00") + " "
Else
  decoresultgridlite.Text = " "
End If
decoresultgridlite.Col = 8
'decoresultgridlite.Text = format(rate*feetormeter_factor,"###")
'If rate > 0.01 Then
'  decoresultgridlite.text = "Descent        " + decoresultgridlite.Text
'Else
'  If rate < -0.01 Then
'    decoresultgridlite.text = "Ascent        " + decoresultgridlite.Text
'  Else
    decoresultgridlite.Text = " ---- "
'  End If
'End If

End Sub

Private Sub cleardecogrid()
decoresultgrid.Rows = 1
End Sub

Private Sub lblgasindex_Click(Index As Integer)
Dim i As Integer


If Left(Cbogasused(Index).Text, 1) = "0" Then ' Or Left(Cbogasused(Index).Text, 1) = "4" Or Left(Cbogasused(Index).Text, 1) = "5" Then
  Exit Sub
End If

G = txtoxygen(Index).Index
current_index = G
tempgasused = Cbogasused(G).Text
If tempgasused <> "0 - Not Used" Then

If Option1(Index).Value = True Then
  lblo2.Caption = txtoxygen(G).Text
  lblhelium.Caption = txthelium(G).Text
  If cbogasindex.Text <> lblgasindex(G).Caption Then
   'txtdepth.Text = "0"
  End If
  If Option1(p).Value = True Then
    cbogasindex.Text = lblgasindex(G).Caption
  End If
End If
tempgasused = Cbogasused(G).Text
If InStr(tempgasused, "Closed C") Then
   Option3.Enabled = True
   Option3.Value = True
   Option3.Enabled = False
   txtppo2v.Enabled = True
   CMDPPO2PLUS.Enabled = True
   Command4.Enabled = True
   Option4.Enabled = True
      Option4.Value = False
      Option4.Enabled = False
Else
  txtppo2v = txtppo2(G).Text
   If InStr(tempgasused, "Open C") Then
      Option4.Enabled = True
      Option4.Value = True
      Option4.Enabled = False
      Option3.Enabled = True
      Option3.Value = False
      Option3.Enabled = False
      txtppo2v.Enabled = False
      CMDPPO2PLUS.Enabled = False
      Command4.Enabled = False
   Else
      Option3.Value = False
      Option4.Value = True
      Option3.Enabled = True
      Option4.Enabled = True
      txtppo2v.Enabled = False
      CMDPPO2PLUS.Enabled = False
      Command4.Enabled = False
   End If
End If
End If
For i = 0 To 9

  Decochk(i).BackColor = &HE0E0E0
  Label32(i).BackColor = &HE0E0E0
  Option1(i).BackColor = &HE0E0E0
  Check1(i).BackColor = &HE0E0E0
  Check2(i).BackColor = &HE0E0E0
  Label13(i).BackColor = &HE0E0E0
  Label28(i).BackColor = &HE0E0E0
  lblgasindex(i).BackColor = &HC0C0C0    '&HFFFFC0
  txtoxygen(i).BackColor = &HE0E0E0
  txthelium(i).BackColor = &HE0E0E0
  txtmaxd(i).BackColor = &HE0E0E0
  If InStr(1, Cbogasused(i), "Closed") Then txtppo2(i).BackColor = vbYellow Else txtppo2(i).BackColor = &HE0E0E0
  Cbogasused(i).BackColor = &HE0E0E0
  txtmaxdft(i).BackColor = &HE0E0E0
  
  lblgasindex(i).ForeColor = vbBlack '&H80FFFF
  txtoxygen(i).ForeColor = vbBlack
  txthelium(i).ForeColor = vbBlack
  txtmaxd(i).ForeColor = vbBlack
  txtppo2(i).ForeColor = vbBlack
  Cbogasused(i).ForeColor = vbBlack
  txtmaxdft(i).ForeColor = vbBlack

Next i
For i = 0 To 9
  If Option1(i).Value = True Then
    Decochk(i).BackColor = vbGreen
    Label32(i).BackColor = vbGreen
    Option1(i).BackColor = vbGreen
  
    Label13(i).BackColor = vbGreen
    Check1(i).BackColor = vbGreen
    Check2(i).BackColor = vbGreen
    Label28(i).BackColor = vbGreen
    lblgasindex(i).BackColor = vbGreen '&H80FFFF
    txtoxygen(i).BackColor = vbGreen
    txthelium(i).BackColor = vbGreen
    txtmaxd(i).BackColor = vbWhite
    txtppo2(i).BackColor = vbGreen
    Cbogasused(i).BackColor = vbGreen
    txtmaxdft(i).BackColor = vbWhite
    lblgasindex(i).ForeColor = vbBlack '&H80FFFF
    txtoxygen(i).ForeColor = vbBlack
    txthelium(i).ForeColor = vbBlack
    txtmaxd(i).ForeColor = vbBlack
    txtppo2(i).ForeColor = vbBlack
    Cbogasused(i).ForeColor = vbBlack
    txtmaxdft(i).ForeColor = vbBlack
    'update_gas_graph (Index)
  End If
  If Decochk(i).Value = 1 Then
    Decochk(i).BackColor = vbYellow
    Label32(i).BackColor = vbYellow
    Option1(i).BackColor = vbYellow
  
    Label13(i).BackColor = vbYellow
    Check1(i).BackColor = vbYellow
    Check2(i).BackColor = vbYellow
    Label28(i).BackColor = vbYellow
    lblgasindex(i).BackColor = vbYellow '&H80FFFF
    txtoxygen(i).BackColor = vbYellow
    txthelium(i).BackColor = vbYellow
    txtmaxd(i).BackColor = vbWhite
    txtppo2(i).BackColor = vbYellow
    Cbogasused(i).BackColor = vbYellow
    txtmaxdft(i).BackColor = vbWhite
    lblgasindex(i).ForeColor = vbBlack '&H80FFFF
    txtoxygen(i).ForeColor = vbBlack
    txthelium(i).ForeColor = vbBlack
    txtmaxd(i).ForeColor = vbBlack
    txtppo2(i).ForeColor = vbBlack
    Cbogasused(i).ForeColor = vbBlack
    txtmaxdft(i).ForeColor = vbBlack
    'update_gas_graph (i)
  End If
Next i
' MsgBox Cbogasused(Index).Text
  
  'Decochk(Index).BackColor = vbBlue
  'Label32(Index).BackColor = vbBlue
  
  'Label13(Index).BackColor = vbBlue
  'Check1(Index).BackColor = vbBlue
  'Check2(Index).BackColor = vbBlue
  'Label28(Index).BackColor = vbBlue
  'lblgasindex(Index).BackColor = vbBlue '&H80FFFF
  txtoxygen(Index).BackColor = vbBlue
  txthelium(Index).BackColor = vbBlue
  'txtmaxd(Index).BackColor = vbBlue
  'txtppo2(Index).BackColor = vbBlue
  'Cbogasused(Index).BackColor = vbBlue
  'txtmaxdft(Index).BackColor = vbBlue
  'lblgasindex(Index).ForeColor = vbWhite '&H80FFFF
  txtoxygen(Index).ForeColor = vbWhite
  txthelium(Index).ForeColor = vbWhite
  'txtmaxd(Index).ForeColor = vbWhite
  'txtppo2(Index).ForeColor = vbWhite
  'Cbogasused(Index).ForeColor = vbWhite
  'txtmaxdft(Index).ForeColor = vbWhite
  update_gas_graph (Index)
' MsgBox Cbogasused(Index).Text
  End Sub



Private Sub lblhelium_Change()
  If cmdgenerate.Visible = False Then cmdmodify.Visible = True
  lblgasvrupdate
End Sub

Private Sub lblo2_Change()
  If cmdgenerate.Visible = False Then cmdmodify.Visible = True
  lblgasvrupdate
End Sub

Private Sub mnufileexit_Click()
Unload Me
End Sub

Private Sub mnugasloadefault_Click()
On Error Resume Next
ans = MsgBox("Do you really want to load the factory default setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
    SQL = "SELECT * FROM dpfacgasdefault "
    SQL = SQL & " order by gasid"
    Set RS = DB.OpenRecordset(SQL)
    RS.MoveFirst
    T = 0
    While RS.EOF = False
       If CInt(T) < 10 Then
       i = T
          lblgasindex(i).Caption = RS("gasid")
          tempnitrogen = RS("gasnitrogen")
          txthelium(i) = RS("gashelium")
          txtoxygen(i).Text = 100 - CInt(txthelium(i).Text) - CInt(tempnitrogen)
          txtmaxd(i) = RS("gasmaxopdepth")
          txtmaxd(i) = txtmaxd(i) / 10
          temptxthe1 = txthe1
          temptxtmaxd1 = txtmaxd(i)
          txtppo2(i).Enabled = True
         ' txtppo2(i).Text = RS("dpgaspo2setpoint")
          txtppo2(i).Text = (CInt(txtoxygen(i).Text) / 100) * ((CInt(txtmaxd(i).Text) / 10) + 1)
          txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
          txtppo2(i).Enabled = False
          tempgasused = RS("gasused")
          Select Case tempgasused
          Case "0 - Air Dive"
             Cbogasused(i).ListIndex = 0
          Case "4 - Deco Open Circuit"
             Cbogasused(i).ListIndex = 4
          Case "1 - Open Circuit"
             Cbogasused(i).ListIndex = 1
          Case "2 - Closed Circuit"
             Cbogasused(i).ListIndex = 2
          Case "5 - Deco Closed Circuit"
             Cbogasused(i).ListIndex = 5
          Case "3 - Open & Closed"
             Cbogasused(i).ListIndex = 3
          End Select
       End If
  T = T + 1
  
  RS.MoveNext
  Wend
  If T < 10 Then
'    MsgBox T
    mnugasloadefault_Click
  End If
    
  lblgasindex_Click (0)
  lblgasindex_Click (1)
  lblgasindex_Click (2)
  lblgasindex_Click (3)
  lblgasindex_Click (4)
  lblgasindex_Click (5)
  lblgasindex_Click (6)
  lblgasindex_Click (7)
  lblgasindex_Click (8)
  lblgasindex_Click (9)
  Form_Load
' MsgBox "All value reset to factory default."
Case Else
 '  MsgBox "Request cancelled. "
End Select
End Sub

Private Sub mnugassetdefault_Click()
ans = MsgBox("Do you really want to set all the value as default ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
    SQL = "SELECT * FROM dpgasdefault"
    Set RS = DB.OpenRecordset(SQL)
    RS.MoveFirst
    While RS.EOF = False
       RS.Delete
       RS.MoveNext
    Wend
    RS.Close
    SQL = "SELECT * FROM dpgasdefault"
    Set RS = DB.OpenRecordset(SQL)
       For i = 0 To 9
          RS.AddNew
          RS!gasid = lblgasindex(i).Caption
          RS!gashelium = txthelium(i).Text
          tempnitrogen = 100 - CInt(txthelium(i).Text) - CInt(txtoxygen(i).Text)
          RS!gasnitrogen = tempnitrogen
          RS!gasmaxopdepth = CInt(txtmaxd(i).Text) * 10
          RS!gasused = Cbogasused(i).Text
          RS.Update
       Next i
Case Else
   'MsgBox "Request cancelled. "
End Select
End Sub

Private Sub mnuGenerate_Click()
  Command1_Click
End Sub

Private Sub mnuHelp_Click()
frmintro.Show
End Sub

Private Sub mnulite_Click()
decoresultgrid.Visible = False
decoresultgridlite.Visible = True
mnuprofessional.Checked = False
mnulite.Checked = True
decographversion = "Lite"
display_deco_graph (0)
systemversion = "Lite"
End Sub

Private Sub mnuloaddesetting_Click()
On Error Resume Next
ans = MsgBox("Do you really want to load the User default settings?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
    SQL = "SELECT * FROM dpgasdefault "
    SQL = SQL & " order by gasid "
    Set RS = DB.OpenRecordset(SQL)
    RS.MoveFirst
    T = 0
    While RS.EOF = False
       If CInt(T) < 10 Then
       i = T
          lblgasindex(i).Caption = RS("gasid")
          tempnitrogen = RS("gasnitrogen")
          txthelium(i) = RS("gashelium")
          txtoxygen(i).Text = 100 - CInt(txthelium(i).Text) - CInt(tempnitrogen)
          txtmaxd(i) = RS("gasmaxopdepth")
          txtmaxd(i) = txtmaxd(i) / 10
          temptxthe1 = txthe1
          temptxtmaxd1 = txtmaxd(i)
          txtppo2(i).Enabled = True
         ' txtppo2(i).Text = RS("dpgaspo2setpoint")
          txtppo2(i).Text = (CInt(txtoxygen(i).Text) / 100) * ((CInt(txtmaxd(i).Text) / 10) + 1)
          txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
          txtppo2(i).Enabled = False
          tempgasused = RS("gasused")
          Select Case tempgasused
          Case "0 - Not Used"
             Cbogasused(i).ListIndex = 0
             Case "4 - Deco Open Circuit"
             Cbogasused(i).ListIndex = 4
             Case "1 - Open Circuit"
             Cbogasused(i).ListIndex = 1
              Case "2 - Closed Circuit"
             Cbogasused(i).ListIndex = 2
             Case "5 - Deco Closed Circuit"
             Cbogasused(i).ListIndex = 5
              Case "3 - Open & Closed"
             Cbogasused(i).ListIndex = 3
          End Select
          
      End If
  T = T + 1
  RS.MoveNext
  Wend
  If T < 10 Then
'    MsgBox T
    mnuloaddesetting_Click
  End If
  Form_Load
 ' MsgBox "All value reset to User default."
Case Else
 '  MsgBox "Request cancelled. "
End Select
End Sub

Private Sub mnuPrint_Click()
Dim comptext As String
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim p As Integer
Dim q As Integer
Dim temptext As String
Dim temptext1 As String
Dim temptext2 As String
Dim temptext3 As String
Dim temptext4 As String
Dim temptext5 As String
Dim temptext6 As String
On Error GoTo ErrorHandler:


   CommonDialog2.ShowPrinter
   Text2.Text = ""
   comptext = "DO NOT DIVE USING THESE TABLES. BETA SOFTWARE TESTING ONLY"
   Text2.Text = Text2.Text + comptext + vbCrLf
   comptext = "Dive Plan No : " & txtserialno.Text
   Text2.Text = Text2.Text + comptext + vbCrLf
   comptext = "  Atmospheric : " & atmtext.Text
   Text2.Text = Text2.Text + comptext + vbCrLf
   comptext = "  Safety : " & safetytext.Text
   Text2.Text = Text2.Text + comptext + vbCrLf
   comptext = ""
   comptext = txtdecoalg.Text
   Text2.Text = Text2.Text + comptext + vbCrLf
   comptext = ""
   Text2.Text = Text2.Text + comptext + vbCrLf
   comptext = "Sequence Of the Dive Plan : "
   If feetormeter_feeton = 1 Then
      For K = 0 To MSFlexGrid3.Rows - 1
         For j = 1 To MSFlexGrid3.Cols - 1
            MSFlexGrid3.Row = K
            MSFlexGrid3.Col = j
            rowtext = MSFlexGrid3.Text
            comptext = comptext + "   " + (rowtext)
         Next j
         Text2.Text = Text2.Text + comptext + vbCrLf
         comptext = ""
      Next K
      Text2.Text = Text2.Text + comptext + vbCrLf
   Else
      For K = 0 To MSFlexGrid3.Rows - 1
         For j = 1 To MSFlexGrid3.Cols - 1
            MSFlexGrid3.Row = K
            MSFlexGrid3.Col = j
            rowtext = MSFlexGrid3.Text
            comptext = comptext + "   " + (rowtext)
         Next j
         Text2.Text = Text2.Text + comptext + vbCrLf
         comptext = ""
      Next K
      Text2.Text = Text2.Text + comptext + vbCrLf
  End If
   comptext = ""
   Text2.Text = Text2.Text + comptext + vbCrLf
    
    comptext = "Active" & vbTab & "Gas #" & vbTab & "O2" & vbTab & "He" & vbTab & "Depth" & vbTab & "PPO2" & vbTab & "CC" & vbTab & "Deco" & vbTab & "WC" & vbTab & "SAC" & "BarUse"
    Text2.Text = Text2.Text + comptext + vbCrLf
    comptext = ""
    For v = 0 To 9
      If Check1(v).Value = 1 Then
         temptext = "On"
      Else
         temptext = "Off"
      End If
      
      temptext2 = "Gas " + CStr(v)
      temptext3 = txtoxygen(v).Text
      temptext4 = txthelium(v).Text
      temptext5 = txtmaxdft(v).Text 'lbldepth(v).Caption
      temptext6 = txtppo2(v).Text
      If Check2(v).Value = 1 Then
         temptext7 = "CC"
      Else
         temptext7 = "OC"
      End If
      If Decochk(v).Value = 1 Then
         temptext8 = "Deco"
      Else
         temptext8 = ""
      End If
      If txtcylcap2(v).Visible = True Then
         temptext9 = txtcylcap2(v).Text
      Else
         temptext9 = ""
      End If
      If txtbreathrate2(v).Visible = True Then
         temptext10 = txtbreathrate2(v).Text
      Else
         temptext10 = ""
      End If
      If gasusage(v).Visible = True Then
         temptext11 = gasusage(v).Caption
      Else
           temptext11 = ""
      End If
      comptext = temptext & vbTab & temptext2 & vbTab & temptext3 & vbTab & temptext4 & vbTab & temptext5 & vbTab & temptext6
      comptext = comptext & vbTab & temptext7 & vbTab & temptext8 & vbTab & temptext9 & vbTab & temptext10 & vbTab & temptext11
      Text2.Text = Text2.Text + comptext + vbCrLf
      comptext = ""
    Next v
    comptext = ""
    Text2.Text = Text2.Text + comptext + vbCrLf
    'comptext = " " & Frame3.Caption
    'text2.text=text2.text +  comptext
    comptext = ""
    Text2.Text = Text2.Text + comptext + vbCrLf
    comptext = ""
    If systemversion = "Pro" Then
    For K = 0 To decoresultgrid.Rows - 1
      For p = 0 To decoresultgrid.Cols - 1
        decoresultgrid.Row = K
        decoresultgrid.Col = p
        rowtext = CStr(decoresultgrid.Text) 'Format(decoresultgrid.Text, "")
        rowtext = Left(rowtext, 8)
        If Len(rowtext) < 7 Then
          rowtext = rowtext + vbTab
        End If
        comptext = comptext + (rowtext + vbTab)
      Next p
     
      Text2.Text = Text2.Text + comptext + vbCrLf
      comptext = ""
    Next K
    Else
    For K = 0 To decoresultgridlite.Rows - 1
      For p = 0 To decoresultgridlite.Cols - 1
        decoresultgridlite.Row = K
        decoresultgridlite.Col = p
        rowtext = CStr(decoresultgridlite.Text) 'Format(decoresultgrid.Text, "")
        rowtext = Left(rowtext, 8)
        If Len(rowtext) < 7 Then
          rowtext = rowtext + vbTab
        End If
        comptext = comptext + (rowtext + vbTab)
      Next p
     
      Text2.Text = Text2.Text + comptext + vbCrLf
      comptext = ""
    Next K
    End If
    Printer.Print Text2.Text
    Printer.EndDoc
ErrorHandler:
   MsgBox "Printer error !!"

End Sub

Private Sub mnuprofessional_Click()
decoresultgrid.Visible = True
decoresultgridlite.Visible = False
mnuprofessional.Checked = True
mnulite.Checked = False
decographversion = "Pro"
systemversion = "Pro"
display_deco_graph (0)
End Sub

Private Sub mnuStep_Click()
mnuStep.Checked = True
mnuStep15.Checked = False
mnuStep2.Checked = False
laststop_index = 1
cmdgenerate_Click
End Sub

Private Sub mnuStep15_Click()
mnuStep.Checked = False
mnuStep15.Checked = True
mnuStep2.Checked = False
laststop_index = 2
cmdgenerate_Click
End Sub

Private Sub mnuStep2_Click()
mnuStep.Checked = False
mnuStep15.Checked = False
mnuStep2.Checked = True
laststop_index = 3
cmdgenerate_Click
End Sub

Private Sub mnuVPMB_Click(Index As Integer)
  mnuVPMB(0).Checked = False
  mnuVPMB(1).Checked = False
  mnuVPMB(2).Checked = False
  mnuVPMB(Index).Checked = True
  mnuVPMBdef.Caption = mnuVPMB(Index).Caption
  If buhl_mode = Index Then
  Else
    buhl_mode = Index
    SQL = "SELECT * FROM dpserialno"
    Set RS = DB.OpenRecordset(SQL)
    RS.Edit
    RS!buhl = CStr(buhl_mode)
    RS.Update
    Command1_Click
  End If
  txtdecoalg.Text = "Deco Algorithm: " + mnuVPMB(buhl_mode).Caption
'  If buhl_mode = 0 Then
'    vpmBuhl.Caption = "VPMB only" ' + CStr(buhl_mode)
'  End If
'  If buhl_mode = 1 Then
'    vpmBuhl.Caption = "VPMB+Buhl" ' + CStr(buhl_mode)
'  End If
'  If buhl_mode = 2 Then
'    vpmBuhl.Caption = "Buhlmann only" + CStr(buhl_mode)
'  End If
End Sub

Private Sub MSFlexGrid3_Click()
rowindentified = MSFlexGrid3.Row
tempcolselect2 = MSFlexGrid3.Col
msflexgriddoclick
tempcolselect = tempcolselect2
If tempcolselect = "1" Then
'   txtdepth.SetFocus
End If
If tempcolselect = "2" Then
'   txttime.SetFocus
End If
If tempcolselect = "3" Then
'   txtoxygen(Right(cbogasindex.Text, 1)).SetFocus
End If
If tempcolselect = "4" Then
'   txthelium(Right(cbogasindex.Text, 1)).SetFocus
End If
If tempcolselect = "8" Then
   If txtdepthft.Visible = False Then
   Else
      txtdepthft.SetFocus
   End If
End If
SendKeys "{HOME}+{END}"
End Sub

Private Function msflexgriddoclick()
If MSFlexGrid3.Rows < 2 Then Exit Function
'Text1.Visible = True
numrow = MSFlexGrid3.Rows
Totalcount = numrow - 1
For K = 0 To Totalcount
  For p = 0 To 0
    MSFlexGrid3.Row = K
    MSFlexGrid3.Col = p
    If MSFlexGrid3.CellBackColor = vbBlue Then
      For H = 0 To 8
        MSFlexGrid3.Row = K
        MSFlexGrid3.Col = H
        MSFlexGrid3.CellForeColor = MSFlexGrid3.ForeColor
        MSFlexGrid3.CellBackColor = MSFlexGrid3.BackColor
      Next H
    End If
  Next p
Next K
For q = 0 To 8
   MSFlexGrid3.Row = rowindentified
   MSFlexGrid3.Col = q
   MSFlexGrid3.CellForeColor = vbWhite
   MSFlexGrid3.CellBackColor = vbBlue
Next q

MSFlexGrid3.Col = 0
  tempseq = MSFlexGrid3.Text
  SQL = "SELECT * FROM seqdpprofile"
  SQL = SQL & " where dpprofileid = '" & tempserialno & "' and dpnumseq = '" & tempseq & "' "
  Set RS = DB.OpenRecordset(SQL)
  txtppo2v.Text = RS("po2")
MSFlexGrid3.Row = rowindentified
 MSFlexGrid3.Col = 1
 If txtdepth.Text <> MSFlexGrid3.Text Then
   'txtdepth.Text = "0"
End If
 MSFlexGrid3.Row = rowindentified
 MSFlexGrid3.Col = 7
 cbogasindex.Text = MSFlexGrid3.Text
For p = 1 To 7

  MSFlexGrid3.Row = rowindentified
  MSFlexGrid3.Col = p
  Select Case p
  Case 1
     txtdepth.Text = MSFlexGrid3.Text
  Case 2
     txttime.Text = MSFlexGrid3.Text
  Case 3
     lblo2.Caption = MSFlexGrid3.Text
  Case 4
     lblhelium.Caption = MSFlexGrid3.Text
  Case 5
    txtppo2v.Enabled = True
    CMDPPO2PLUS.Enabled = True
    Command4.Enabled = True
    txtppo2v.Text = MSFlexGrid3.Text
  Case 6
    MSFlexGrid3.Row = rowindentified
    MSFlexGrid3.Col = 7
    lblgasindex_Click (CInt(Right(MSFlexGrid3.Text, 1)))
    MSFlexGrid3.Col = p
    If MSFlexGrid3.Text = "Closed Circuit" Then
              Option3.Value = True
              Option4.Value = False
              txtppo2v.Enabled = True
              CMDPPO2PLUS.Enabled = True
              Command4.Enabled = True
              MSFlexGrid3.Col = 5
              txtppo2v.Text = MSFlexGrid3.Text
    Else
              Option4.Value = True
              Option3.Value = False
              txtppo2v.Enabled = False
              CMDPPO2PLUS.Enabled = True
              Command4.Enabled = True
    End If
    MSFlexGrid3.Col = 7
    cbogasindex.Text = MSFlexGrid3.Text
    For i = 0 To 9
              If lblgasindex(i).Caption = cbogasindex.Text Then
                 tempgasused = Cbogasused(i).Text
              End If
    Next i
           
    Select Case tempgasused
             Case "3 - Open & Closed"
                'Option4.Enabled = True
                'Option3.Enabled = True
             Case "1 - Open Circuit"
                Option4.Enabled = True
                Option3.Enabled = False
                txtppo2v.Enabled = False
                CMDPPO2PLUS.Enabled = False
                Command4.Enabled = False
             Case "2 - Closed Circuit"
                Option4.Enabled = False
                Option3.Enabled = True
                txtppo2v.Enabled = True
                CMDPPO2PLUS.Enabled = True
                Command4.Enabled = True
             Case "4 - Deco Open Circuit"
                Option4.Enabled = True
                Option3.Enabled = False
                txtppo2v.Enabled = False
                CMDPPO2PLUS.Enabled = False
                Command4.Enabled = False
             Case "5 - Deco Closed Circuit"
             
                Option4.Enabled = False
                Option3.Enabled = True
                txtppo2v.Enabled = True
                CMDPPO2PLUS.Enabled = True
                Command4.Enabled = True
                End Select
             Case 7
                cbogasindex.Text = MSFlexGrid3.Text
    End Select
Next p
 ' cmdinsert.Visible = True
  cmdremove.Enabled = True
  MSFlexGrid3.Row = rowindentified
'  MSFlexGrid3.Col = 7
  lblgasindex_Click (CInt(Right(MSFlexGrid3.Text, 1)))
  cmdmodify.Visible = False
  
End Function

Private Sub Option1_Click(Index As Integer)
Dim i As Integer
  For i = 0 To 9
    If Decochk(i).Value = 1 Then
    Else
      Check1(i).Value = 0
    End If
  Next i
  For i = 0 To 9
    Decochk(i).Enabled = True
  Next i
  Check1(Index).Value = 1
  Decochk(Index).Value = 0
  Decochk(Index).Enabled = False
  Check1(Index).Value = 1
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
   txtppo2v.Enabled = True
   CMDPPO2PLUS.Enabled = True
   Command4.Enabled = True
End If
  If cmdgenerate.Visible = False Then cmdmodify.Visible = True
End Sub

Private Sub Picture1_DblClick()
detectdouclick = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, Xmouse As Single, Ymouse As Single)
Dim changedata As Integer

  If Xstart = -1 Then Exit Sub
  Me.MousePointer = 0
  If decoresultgrid.Visible = True Then
     If decoresultgrid.CellBackColor = vbYellow Then Exit Sub
  Else
     If decoresultgridlite.CellBackColor = vbYellow Then Exit Sub
  End If
  changedata = 0
  If Ymouse > Ystart + 100 Then
    Me.MousePointer = 7
    changedata = 1
  End If

  If Ymouse + 100 < Ystart Then
    Me.MousePointer = 7
    changedata = 1
  End If

  If Xmouse > Xstart + 100 Then
    Me.MousePointer = 9
    changedata = changedata + 2
  End If

  If Xmouse + 100 < Xstart Then
    Me.MousePointer = 9 '7
    changedata = changedata + 2
  End If

  If changedata > 2 Then
        Me.MousePointer = 5
  Else
        If changedata = 0 Then Me.MousePointer = 0
  End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, Xmouse As Single, Ymouse As Single)
Dim changedata As Integer
Dim i As Integer
Dim xc As Single
Dim timet As Single


Me.MousePointer = 0
changedata = 0
If detectdouclick = False Then
If inc_depth < 0.2 Then inc_depth = 1#
If inc_depth < 0.4 Then inc_depth = 3#
deco_grid_display_rowlast = deco_grid_display_rowlast
If decoresultgrid.Visible = True Then
   If deco_grid_display_rowlast + 1 > decoresultgrid.Rows Then Exit Sub
Else
   If deco_grid_display_rowlast + 1 > decoresultgridlite.Rows Then Exit Sub
End If
If decoresultgrid.Visible = True Then
   decoresultgrid.Row = deco_grid_display_rowlast
   decoresultgrid.Col = 3
   If decoresultgrid.CellBackColor = vbYellow Then Exit Sub
Else
   decoresultgridlite.Row = deco_grid_display_rowlast
   decoresultgridlite.Col = 3
   If decoresultgridlite.CellBackColor = vbYellow Then Exit Sub
End If
  If Ymouse > Ystart + 100 Then
    If Ymouse > Ystart + 400 Then inc_depth = inc_depth * 10#
    cmddepthup_Click
    If Ymouse > Ystart + 400 Then inc_depth = inc_depth / 10#
    changedata = 1
  End If

  If Ymouse + 100 < Ystart Then
    If Ymouse + 400 < Ystart Then inc_depth = inc_depth * 10#
    cmddepthdown_Click
    If Ymouse + 400 < Ystart Then inc_depth = inc_depth / 10#
    changedata = 1
  End If


  inc_time = 1#
  If Xmouse > Xstart + 100 Then
    If Xmouse > Xstart + 400 Then inc_time = 10#
    cmdtimeup_Click
    changedata = 1
  End If

  If Xmouse + 100 < Xstart Then
    If Xmouse + 400 < Xstart Then inc_time = 10#
    cmdtimedown_Click
    changedata = 1
  End If
  
  inc_time = 1#
  
  If changedata = 1 Then
    cmdmodify_Click
    If decoresultgrid.Rows < 4 Then Exit Sub

    xc = Xstart
    xc = xc / Picture1.Width
    timet = xc * runtime_graph
    For i = 1 To row_count - 1
      If X(i) > timet Then Exit For
    Next i
  
  
'  msgres = MsgBox("Time: " & Format(timet, "###0mins") & vbCrLf & "Depth: " & Format(Y(i) * feetormeter_factor, "###0" & feetormeter_shortstring), vbOKOnly, "Dive Point")
  
    If deco_grid_display_last >= 0 Then
      If decoresultgrid.Visible = True Then
         If decoresultgrid.Rows < 4 Then Exit Sub
         decoresultgrid.Row = deco_grid_display_rowlast
         decoresultgrid.Col = 0
         decoresultgrid.CellBackColor = deco_grid_display_celllast
      Else
         If decoresultgridlite.Rows < 4 Then Exit Sub
         decoresultgridlite.Row = deco_grid_display_rowlast
         decoresultgridlite.Col = 0
         decoresultgridlite.CellBackColor = deco_grid_display_celllast
      End If
    End If
    If decoresultgrid.Visible = True Then
       If i <= 4 Then decoresultgrid.TopRow = 1 Else decoresultgrid.TopRow = i - 4
    Else
       If i <= 4 Then decoresultgridlite.TopRow = 1 Else decoresultgridlite.TopRow = i - 4
    End If
    If decoresultgrid.Visible = True Then
       decoresultgrid.Col = 0
       decoresultgrid.Row = i
       deco_grid_display_last = 0 'deco_grid_display
       deco_grid_display_rowlast = (i)
       deco_grid_display_celllast = decoresultgrid.CellBackColor
       decoresultgrid.CellBackColor = vbBlue
    Else
       decoresultgridlite.Col = 0
       decoresultgridlite.Row = i
       deco_grid_display_last = 0 'deco_grid_display
       deco_grid_display_rowlast = (i)
       deco_grid_display_celllast = decoresultgridlite.CellBackColor
       decoresultgridlite.CellBackColor = vbBlue
    End If
    rowindentified = Fix((deco_grid_display_rowlast + 1) / 2)
    If rowindentified >= MSFlexGrid3.Rows Then rowindentified = 1 'MSFlexGrid3.Rows - 1
    msflexgriddoclick
    xc = xc
    'picture1.m
    Text1.Visible = True
    If decoresultgrid.Visible = True Then
       decoresultgrid.Col = 0
       decoresultgrid.Row = i
       Text1.Text = decoresultgrid.Text + " "
       decoresultgrid.Col = 1
       Text1.Text = Text1.Text + decoresultgrid.Text
    Else
       decoresultgridlite.Col = 0
       decoresultgridlite.Row = i
       Text1.Text = decoresultgridlite.Text + " "
       decoresultgridlite.Col = 1
       Text1.Text = Text1.Text + decoresultgridlite.Text
    End If
    Text1.Top = Ystart - Text1.Height
    Text1.Left = Xstart
    Text1.Width = 900
  End If
'End If
  Xstart = -1
  Ystart = -1
End If
End Sub

Private Sub Picture4_Click()
  Frame4.Visible = False
End Sub

Private Sub safetytext_Change()
  If IsNumeric(safetytext.Text) Then
  Else
    safetytext.Text = "0"
  End If
End Sub

Private Sub singlelevel_Click()
If MSFlexGrid3.Rows > 2 Then
   MsgBox "Too many depth profile, Please remove some depth profile"
Else

Shape4.Visible = False
Shape5.Visible = False
Image1.Visible = True
txtserialno.Top = 440
txtserialno.BackColor = &HFFFFFF
txtserialno.ForeColor = &H808000
Picture3.Top = 240
Picture3.Height = 4215
cmdremove.Visible = False
cmdclearall.Visible = False
MSFlexGrid3.Visible = False
MSFlexGrid3.Left = 2650
MSFlexGrid3.Top = 950
txtdepth.Left = 960
txtdepthft.Left = 960
txtdepth.Width = 1275
txtdepthft.Width = 1275
txtdepthft.Top = 1440
txtdepth.Top = 1440
txttime.Top = 1440
txttime.Left = 3000
txttime.Width = 1275
Label36.Visible = False
Label37.Visible = False
Label38.Visible = False
Label11.Visible = True
Label11.Left = 1800
Label11.Top = 1200
Cmdadd.Visible = False
cmdinsert.Visible = False
cmdmodify.Visible = False
cmddepthup.Top = 1200
cmddepthdown.Top = 1200
cmddepthup.Left = 1320
cmddepthdown.Left = 1560
cmdtimeup.Left = 3240
cmdtimedown.Left = 3480
cmdtimeup.Top = 1200
cmdtimedown.Top = 1200
Label1.Left = 1800
lblminutes.Left = 3720
txtppo2v.Left = 1960
txtppo2v.Width = 1275
txtppo2v.Top = 2420
txtppo2v.FontSize = 16
txtppo2v.Height = 435
CMDPPO2PLUS.Left = 2320
CMDPPO2PLUS.Top = 2200
Command4.Left = 2560
Command4.Top = 2200
Label29.Left = 2800
Label29.Top = 2200
lblgasvr.Left = 2040
lblgasvr.Top = 740
Option4.Left = 3240
Option4.Top = 2280
Option3.Left = 3240
Option3.Top = 2560
cmdgenerate.Visible = True
If buhl_mode = 2 Then Else lbllevel.Visible = True
singlelevel.Visible = False
cmdgeneratem.Visible = False
End If
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

Private Sub txthe_Change(Index As Integer)

End Sub




Private Sub Option4_Click()
  If cmdgenerate.Visible = False Then cmdmodify.Visible = True
   txtppo2v.Enabled = False
   CMDPPO2PLUS.Enabled = False
   Command4.Enabled = False
End Sub

Private Sub Timer2_Timer()
timer2buffer = timer2buffer + 1
If timer2buffer < 10 Then Exit Sub
 If depthcount > 0 Then
  checkgasselected = False
  backcolortored
  checkgasindex
  'MsgBox Cint(txtdepth)
  If checkgasselected = True Then
        If CInt(txtdepth) >= 0 And CInt(txtdepth) < 2000 Then
           If depthcount = 1 Then txtdepth.Text = CStr(CDbl(txtdepth.Text) + (inc_depth / feetormeter_factor))
           If depthcount = 2 Then txtdepth.Text = CStr(CDbl(txtdepth.Text) - (inc_depth / feetormeter_factor))
           If Option4.Value = True Then
              txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
              txtppo2v = Format(txtppo2v, "###.00")
           End If
        Else
            txtdepth.Text = "10"
        End If
  Else
    Title = "Error on System Validation.."
    MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
  End If
 Else
  If timecount = 1 Then
   checkgasselected = False
   backcolortored
   checkgasindex
   If checkgasselected = True Then
     If CInt(txttime) >= 0 And CInt(txttime) < 9999 Then
        txttime = txttime + 1
     Else
        txttime = "10"
     End If
   Else
    Title = "Error on System Validation.."
    MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
   End If
  End If
  If timecount = 2 Then
   checkgasselected = False
   backcolortored
   checkgasindex
   If checkgasselected = True Then
      If CInt(txttime) > 1 And CInt(xttime) < 9999 Then
         txttime = txttime - 1
      Else
         txttime = "10"
      End If
   Else
      Title = "Error on System Validation.."
      MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
   End If
  End If
  If ppo2count = 1 Then
   checkgasselected = False
   backcolortored
   checkgasindex
   If checkgasselected = True Then
    If CDbl(txtppo2v) >= 0.15 And CDbl(txtppo2v) < 2.01 Then
      txtppo2v = CDbl(txtppo2v) + 0.01
      backcolortored
    End If
   Else
    Title = "Error on System Validation.."
    MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
   End If
  End If
  If ppo2count = 2 Then
   checkgasselected = False
   backcolortored
   checkgasindex
   If checkgasselected = True Then
    If CDbl(txtppo2v) > 0.15 And CDbl(txtppo2v) < 2.01 Then
      txtppo2v = CDbl(txtppo2v) - 0.01
      backcolortored
  '  Else
  '    MsgBox " PO2 value out of range !"
    End If
   Else
      Title = "Error on System Validation.."
      MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
   End If
  End If
  If gascount > 0 Then
    If gascount > 2 Then
      Cmdplus_Click (gascount)
    Else
      Cmdminus_Click (gascount + 2)
    End If
  End If
  If vhmxcount > 0 Then
    If vhmxcount >= 100 Then
      If vhmxcount = 105 Then
        vhmx_up_Click (4)
      Else
        vhmx_down_Click (4)
      End If
    Else
      If vhmxcount >= 5 Then
        vhmx_up_Click (vhmxcount - 5)
      Else
        vhmx_down_Click (vhmxcount - 1)
      End If
    End If
  End If
End If
txtppo2v.Text = Format(txtppo2v.Text, "0.00")
End Sub

Private Sub Timer3_Timer()
  Frame4.Visible = False
  Timer3.Enabled = False
End Sub

Private Sub txtbreathrate_Change()
Dim i As Integer
  If IsNumeric(txtcylcap.Text) And IsNumeric(txtbreathrate.Text) Then
    If CDbl(txtcylcap.Text) < 1# Or CDbl(txtbreathrate.Text) < 5# Or CDbl(txtbreathrate.Text) > 50# Then
      clear_baruse
      Exit Sub
    End If
    For i = 0 To 9
      If IsNumeric(txtcylcap.Text) And IsNumeric(txtbreathrate.Text) Then gasusageupdate (i)
    Next
  Else
    clear_baruse
  End If
End Sub

Private Function clear_baruse()
Dim i As Integer

  For i = 0 To 9
    gasusage(i).Caption = " "
  Next
End Function

Private Sub txtbreathrate2_Change(Index As Integer)
  gasusageupdate (Index)
End Sub

Private Sub txtbreathratecuft_Change()
  txtbreathrate.Text = CStr(CDbl(txtbreathratecuft.Text) / 0.035)
End Sub

Private Sub txtcylcap_Change()
Dim i As Integer
  If IsNumeric(txtcylcap.Text) And IsNumeric(txtbreathrate.Text) Then
    If CDbl(txtcylcap.Text) < 1# Or CDbl(txtbreathrate.Text) < 5# Or CDbl(txtbreathrate.Text) > 50# Then
      clear_baruse
      Exit Sub
    End If
    For i = 0 To 9
      If IsNumeric(txtcylcap.Text) And IsNumeric(txtbreathrate.Text) Then gasusageupdate (i)
    Next
  Else
    clear_baruse
  End If
End Sub

Private Sub txtcylcap2_Change(Index As Integer)
  gasusageupdate (Index)
End Sub

Private Sub txtdepth_Change()
  If Len(txtdepth.Text) < 1 Then Exit Sub
 
  If IsNumeric(Right(cbogasindex.Text, 1)) = False Or IsNumeric(txtdepth.Text) = False Then
     lengthtxtdepth = Len(txtdepth)
     For K = 1 To lengthtxtdepth '- 1
     If Asc(Mid$(txtdepth, K, 1)) > 45 And Asc(Mid$(txtdepth, K, 1)) < 59 Then
        tempcode = tempcode & Mid$(txtdepth, K, 1)
     Else
        tempcode = tempcode
     End If
  Next
   txtdepth.Text = tempcode
  SendKeys "{END}"
  Exit Sub
  End If
  
 For i = 0 To 9
   If Option1(i).Value = True Then Exit For
 Next i
 If CDbl(txtdepth.Text) > CDbl(txtmaxd(i)) Then '(CInt(Right(cbogasindex.Text, 1)))) Then
    'MsgBox "Depth too deep for gas - change gas or depth depth"
    txtdepth.Text = txtmaxd(CInt(Right(cbogasindex.Text, 1))) '"10"
    txtdepthft.Text = Format(CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) * feetormeter_factor, "###") '"30"
    SendKeys "{END}"
    Exit Sub
  End If
  If CInt(txtdepth) < 1 Or CInt(txtdepth) > 2000 Then
     If MSFlexGrid3.Rows > 1 Then
        MSFlexGrid3.Col = 1
        txtdepth.Text = MSFlexGrid3.Text
        SendKeys "{END}"
        Exit Sub
     End If
  Else
     If Option4.Value = True And CInt(txtdepth) <> 0 Then
      If IsNumeric(lblo2.Caption) Then
        txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
    '    txtppo2v = Format(txtppo2v, "###.00")
        'txttime.SetFocus
      Else
'        MsgBox "No ppo2"
      End If
     End If
  End If
  
  
  
  If cmdgenerate.Visible = False Then cmdmodify.Visible = True
  If txtdepthft_focus = 1 Or Len(txtdepth.Text) < 1 Or IsNumeric(txtdepth.Text) = False Then Exit Sub
  txtdepth_focus = 1
  txtdepthft.Text = Format(CStr(CDbl(txtdepth.Text) * feetormeter_factor), "###0")
  txtdepth_focus = 0
 ' MsgBox Cbogasused(p).Text & "feetmt"
  '
 '
 ' If txtdepthft_focus = 0 And Len(txtdepth.Text) > 0 Then txtdepthft.Text = Format(CStr(CDbl(txtdepth.Text) * feetormeter_factor), "###0")
End Sub

Private Sub txtdepth_GotFocus()
 checkgasselected = False
 checkgasindex
     If checkgasselected = False Then
        Title = "Error on System Validation.."
        MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
     End If
End Sub

Private Sub txtdepth_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  If CInt(txtdepth) < 0 Or CInt(txtdepth) > 2000 Then
     MsgBox " Depth value out of range !"
  Else
     If Option4.Value = True And CInt(txtdepth) <> 0 Then
        txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
        txtppo2v = Format(txtppo2v, "###.00")
        txttime.SetFocus
        
     End If
  End If
Else
   If KeyAscii <> 8 Then
      If KeyAscii < 46 Or KeyAscii > 57 Then
         'MsgBox "Sorry, Only numeric characters allowed !"
         'txtdepth.SetFocus
         SendKeys "{END}"
      End If
   End If
End If
End Sub

Private Sub txtdepth_LostFocus()
If Trim(cbogasindex) <> "Gas Index" Then
If CInt(txtdepth) < 0 Or CInt(txtdepth) > 2000 Then
     MsgBox " Depth value out of range !"
     txtdepth.SetFocus
     SendKeys "{END}"
Else
   If Option4.Value = True And CInt(txtdepth) <> 0 Then
      txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
     End If
End If
End If
End Sub

Private Sub txtdepthft_Change()

     lengthtxtdepthft = Len(txtdepthft)
     For K = 1 To lengthtxtdepthft '- 1
     If Asc(Mid$(txtdepthft, K, 1)) > 45 And Asc(Mid$(txtdepthft, K, 1)) < 59 Then
        tempcode = tempcode & Mid$(txtdepthft, K, 1)
     Else
        tempcode = tempcode
     End If
  Next
txtdepthft.Text = tempcode
SendKeys "{END}"

   If txtdepth_focus = 1 Or Len(txtdepthft.Text) < 1 Or IsNumeric(txtdepthft.Text) = False Then Exit Sub
   txtdepthft_focus = 1
   txtdepth.Text = txtdepthft.Text / feetormeter_factor
   txtdepth.Text = Format(CDbl(txtdepthft.Text / feetormeter_factor), "###0.00")
   txtdepth.Text = Format(CStr(CDbl(txtdepthft.Text) / feetormeter_factor), "###0.00")
   txtdepthft_focus = 0
'MsgBox Cbogasused(Index).Text & "feetft"
'  If Len(txtdepth.Text) > 0 Then txtdepthft.Text = Format(CStr(CDbl(txtdepth.Text) / feetormeter_factor), "####.0")
End Sub

Private Sub txtdepthft_GotFocus()
'  txtdepthft_focus = 1
End Sub

Private Sub txtdepthft_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If CInt(txtdepthft) < 0 Or CInt(txtdepthft) > 2000 Then
     MsgBox " Depth value out of range !"
  Else
'     If Option4.Value = True And Cint(txtdepth) <> 0 Then
'        txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
'        txtppo2v = Format(txtppo2v, "###.00")
'        txttime.SetFocus
'
'     End If
  End If
Else
   If KeyAscii <> 8 Then
      If KeyAscii < 46 Or KeyAscii > 57 Then
         MsgBox "Sorry, Only numeric characters allowed !"
         txtdepthft.SetFocus
         MSFlexGrid3_Click
         SendKeys "{END}"
      End If
   End If
End If
End Sub

Private Sub txtdepthft_LostFocus()
 ' txtdepthft_focus = 0
End Sub

Private Sub txthelium_Change(Index As Integer)
lengthtxthelium = Len(txthelium(p))
For K = 1 To lengthtxthelium '- 1
      If Asc(Mid$(txthelium(p), K, 1)) > 45 And Asc(Mid$(txthelium(p), K, 1)) < 59 Then
         tempcode = tempcode & Mid$(txthelium(p), K, 1)
      Else
         tempcode = tempcode
      End If
   Next
txthelium(p).Text = tempcode
txthelium_LostFocus (Index)
SendKeys "{END}"
End Sub

Private Sub txthelium_GotFocus(Index As Integer)
   If Left(Cbogasused(Index).Text, 1) = "0" Then Exit Sub
   lblgasindex_Click (Index)
End Sub

Private Sub txthelium_KeyPress(Index As Integer, KeyAscii As Integer)
p = txthelium(Index).Index
current_index = p

If KeyAscii = 13 Then
  If IsNumeric(txthelium(Index)) Then
    validatehelium
    update_gas_graph (current_index)
  Else
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      temphelium = RS2("dpgashelium")
       txthelium(p).Text = temphelium
'       txthelium(p).SetFocus
       SendKeys "{HOME}+{END}"
  End If
End If
End Sub


Private Sub txthelium_LostFocus(Index As Integer)
   If Left(Cbogasused(Index).Text, 1) = "0" Then Exit Sub
   p = txthelium(Index).Index
   current_index = p
   validatehelium
   
   update_gas_graph (current_index)
End Sub

Private Sub txtmaxd_Change(Index As Integer)
lengthtxtmaxd = Len(txtmaxd(p))
For K = 1 To lengthtxtmaxd '- 1
      If Asc(Mid$(txtmaxd(p), K, 1)) > 45 And Asc(Mid$(txtmaxd(p), K, 1)) < 59 Then
         tempcode = tempcode & Mid$(txtmaxd(p), K, 1)
      Else
         tempcode = tempcode
      End If
   Next
txtmaxd(p).Text = tempcode
SendKeys "{END}"
  If txtmaxdft_focus = 1 Or Len(txtmaxd(Index).Text) < 1 Or IsNumeric(txtmaxd(Index).Text) = False Then Exit Sub
  txtmaxd_focus = 1
  txtmaxdft(Index).Text = Format(CStr(CDbl(txtmaxd(Index).Text) * feetormeter_factor), "###0")
  txtmaxd_focus = 0
End Sub

Private Sub txtmaxd_GotFocus(Index As Integer)
   'txtmaxd_focus = 1
   If Left(Cbogasused(Index).Text, 1) = "0" Then Exit Sub
   lblgasindex_Click (Index)
End Sub

Private Sub txtmaxd_KeyPress(Index As Integer, KeyAscii As Integer)
p = txtmaxd(Index).Index
If KeyAscii = 13 Then
  
  validatemaxdepth
Else
   If KeyAscii <> 8 Then
      If KeyAscii < 46 Or KeyAscii > 58 Then
       '  MsgBox "Sorry, Only numeric characters allowed !"
         'txtdepth.SetFocus
         'SendKeys "{HOME}+{END}"
      End If
   End If
End If
End Sub
Private Sub validateoxygen()
If (txthelium(p).Text <> "") And (txtoxygen(p).Text <> "") And (txtmaxd(p).Text <> "") Then
   If CInt(txtoxygen(p).Text) >= 0 And CInt(txtoxygen(p).Text) <= 100 And ((CInt(txtoxygen(p).Text) + CInt(txthelium(p).Text)) < 101) Then
     'txtppo2(p).Enabled = True
    temptextpo2 = (CInt(txtoxygen(p).Text) / 100) * ((CInt(txtmaxd(p).Text) / 10) + 1)
    temptextpo2 = Format(temptextpo2, "###.00")
    If Cbogasused(p).Text = "5 - Deco Closed Circuit" Then
'       If (Cdbl(temptextpo2) <> Cdbl(txtppo2(p).Text)) Then
'         ans = MsgBox("Do you want to replace PO2 value with new setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
'         Select Case ans
'           Case vbYes
'              txtppo2(p).Enabled = True
'              txtppo2(p).Text = temptextpo2
'           Case Else
'              MsgBox "PPO2 value not replaced "
'         End Select
'      End If
    Else
        txtppo2(p).Enabled = True
        txtppo2(p).Text = temptextpo2
        txtppo2(p).Enabled = False
    End If
    tempgasindex = lblgasindex(p).Caption
    tempgasused = Cbogasused(p).Text
    If Left(Cbogasused(p).Text, 1) = "0" Or Left(Cbogasused(p).Text, 1) = "4" Or Left(Cbogasused(p).Text, 1) = "5" Then
    Else
     txtppo2v.Text = txtppo2(p).Text
     lblhelium.Caption = txthelium(p).Text
     lblo2.Caption = txtoxygen(p).Text
     If Option1(p).Value = True Then
       cbogasindex.Text = lblgasindex(p).Caption
     End If
     If InStr(tempgasused, "Closed C") Then
        Option3.Enabled = True
        Option3.Value = True
        Option3.Enabled = False
        txtppo2v.Enabled = True
        CMDPPO2PLUS.Enabled = True
        Command4.Enabled = True
        Option4.Enabled = True
        Option4.Value = False
        Option4.Enabled = False
     Else
        If InStr(tempgasused, "Open C") Then
           Option4.Enabled = True
           Option4.Value = True
           Option4.Enabled = False
           Option3.Enabled = True
           Option3.Value = False
           Option3.Enabled = False
           txtppo2v.Enabled = False
           CMDPPO2PLUS.Enabled = False
           Command4.Enabled = False
        Else
           Option3.Value = False
           Option4.Value = True
           Option3.Enabled = True
           Option4.Enabled = True
           txtppo2v.Enabled = False
           CMDPPO2PLUS.Enabled = False
           Command4.Enabled = False
        End If
     End If
     If Option4.Value = True Then
        txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
        txtppo2v = Format(txtppo2v, "###.00")
     End If
    End If
     SQL = "SELECT * FROM dpmaingaslist "
     SQL = SQL & " where dpmainid = '" & tempserialno & "' "
     SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     tempnitrogen = 100 - CInt(txtoxygen(p).Text) - CInt(txthelium(p).Text)
     RS2.Edit
     RS2!dpgasnitrogen = tempnitrogen
     RS2.Update
     SQL = "SELECT * FROM seqdpprofile "
     SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
     SQL = SQL & " and gasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     tempnitrogen = 100 - CInt(txtoxygen(p).Text) - CInt(txthelium(p).Text)
     While RS2.EOF = False
        tempdepth = CDbl(RS2("depth"))
        tempcircuit = RS2("dpcircuit")
        RS2.Edit
        If InStr(tempcircuit, "Open C") Then
           tempoxygen = (CDbl(txtoxygen(p).Text) / 100#) * ((tempdepth / 10#) + 1)
           RS2!po2 = Format(CStr(tempoxygen), "###.00")
        End If
        RS2!dpo2 = txtoxygen(p).Text
'
        RS2.Update
        RS2.MoveNext
     Wend
     If MSFlexGrid3.Rows > 1 Then
        profilerecordexist = True
        reloadgriddata
     End If
   Else
      Title = "Error on System Validation.."
      MsgBox "(oxygen " & p & " value can not be less then 0 or more than 100) or " & Chr(13) & "(Oxygen + Helium value can not be more than 100)", 48, Title
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      tempoxy = RS2("dpgasnitrogen")
      tempoxygen = 100 - CInt(txthelium(p).Text) - CInt(tempoxy)
      txtoxygen(p).Text = tempoxygen
'      txtoxygen(p).SetFocus
      SendKeys "{END}"
      Timer2.Enabled = False
   End If
Else
'   txtoxygen(p).SetFocus
   SendKeys "{END}"
End If
End Sub

Private Sub validategasused()

'MsgBox Cbogasused(Index).Text
If CInt(txtoxygen(p).Text) >= 0 And CInt(txtoxygen(p).Text) <= 100 And ((CInt(txtoxygen(p).Text) + CInt(txthelium(p).Text)) < 101) Then
     temptextpo2 = (CInt(txtoxygen(p).Text) / 100) * ((CInt(txtmaxd(p).Text) / 10) + 1)
     temptextpo2 = Format(temptextpo2, "###.00")
     
     If Cbogasused(p).Text = "5 - Deco Closed Circuit" Then
        txtppo2(p).Enabled = True
        SendKeys "{END}"
        txtppo2(p).SetFocus
     Else
        txtppo2(p).Enabled = True
        txtppo2(p).Text = temptextpo2
        txtppo2(p).Enabled = False
     End If
     If CDbl(txtppo2v.Text) < 0.4 Or CDbl(txtppo2v.Text) > 2# Then
       txtppo2v.Text = txtppo2(p).Text
     End If
     If Option1(p).Value = True Then
       lblhelium.Caption = txthelium(p).Text
       lblo2.Caption = txtoxygen(p).Text
     End If
     If Option1(p).Value = True Then
       cbogasindex.Text = lblgasindex(p).Caption
     End If
     tempgasindex = lblgasindex(p).Caption
     tempgasused = Cbogasused(p).Text
     If InStr(tempgasused, "Closed C") Then
        Option3.Enabled = True
        Option3.Value = True
        Option3.Enabled = False
        txtppo2v.Enabled = True
        CMDPPO2PLUS.Enabled = True
        Command4.Enabled = True
        Option4.Enabled = True
        Option4.Value = False
        Option4.Enabled = False
     Else
        If InStr(tempgasused, "Open C") Then
           Option4.Enabled = True
           Option4.Value = True
           Option4.Enabled = False
           Option3.Enabled = True
           Option3.Value = False
           Option3.Enabled = False
           txtppo2v.Enabled = False
           CMDPPO2PLUS.Enabled = False
           Command4.Enabled = False
        Else
           Option3.Value = False
           Option4.Value = True
           Option3.Enabled = True
           Option4.Enabled = True
           txtppo2v.Enabled = False
           CMDPPO2PLUS.Enabled = False
           Command4.Enabled = False
        End If
     End If
     If Option4.Value = True And (CDbl(txtppo2v.Text) < 0.4 Or CDbl(txtppo2v.Text) > 2#) Then
        txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
        txtppo2v = Format(txtppo2v, "###.00")
     End If
     'Update gaslist for gas choose
     SQL = "SELECT * FROM dpmaingaslist "
     SQL = SQL & " where dpmainid = '" & tempserialno & "' "
     SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     RS2.Edit
     RS2!dpgasused = Cbogasused(p).Text
     RS2.Update
     'update profile sequence with gas selected
     SQL = "SELECT * FROM seqdpprofile "
     SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
     SQL = SQL & " and gasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     tempnitrogen = 100 - CInt(txtoxygen(p).Text) - CInt(txthelium(p).Text)
     While RS2.EOF = False
        tempdepth = CDbl(RS2("depth"))
        tempcircuit = RS2("dpcircuit")
        RS2.Edit
        If Option3.Value = True Then
           tempcircuit = "Closed Circuit"
        Else
           tempcircuit = "Open Circuit"
        End If
        If InStr(tempcircuit, "Open C") Then
           tempoxygen = (CDbl(txtoxygen(p).Text) / 100#) * ((tempdepth / 10#) + 1)
           RS2!po2 = Format(CStr(tempoxygen), "###.00")
        Else
           RS2!po2 = txtppo2(p).Text
           
        End If
        RS2!dpo2 = txtoxygen(p).Text
        RS2!dpcircuit = tempcircuit
        RS2.Update
        RS2.MoveNext
     Wend
     
     If MSFlexGrid3.Rows > 1 Then
        profilerecordexist = True
        reloadgriddata
     End If
     
   Else
      Title = "Error on System Validation.."
      MsgBox "(Maximum depth " & p & " value can not be less then 0 or more than 1000) or " & Chr(13) & "(Oxygen + Helium value can not be more than 100)", 48, Title
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      tempoxy = RS2("dpgasnitrogen")
      tempoxygen = 100 - CInt(txthelium(p).Text) - CInt(tempoxy)
      txtoxygen(p).Text = tempoxygen
'      txtoxygen(p).SetFocus
      SendKeys "{END}"
      Timer2.Enabled = False
   End If
End Sub
Private Sub validatemaxdepth()
If txtmaxd(p).Text <> "" Then
   If CInt(txtmaxd(p).Text) >= 0 And CInt(txtmaxd(p).Text) < 1000 And ((CInt(txtoxygen(p).Text) + CInt(txthelium(p).Text)) < 101) Then
    temptextpo2 = (CInt(txtoxygen(p).Text) / 100) * ((CInt(txtmaxd(p).Text) / 10) + 1)
    temptextpo2 = Format(temptextpo2, "###.00")
    If Cbogasused(p).Text = "5 - Deco Closed Circuit" Then
       If (CDbl(temptextpo2) <> CDbl(txtppo2(p).Text)) Then
         ans = MsgBox("Do you want to replace PO2 value with new setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
         Select Case ans
           Case vbYes
              txtppo2(p).Enabled = True
              txtppo2(p).Text = temptextpo2
           Case Else
'              MsgBox "PPO2 value not replaced "
         End Select
       End If
    Else
        txtppo2(p).Enabled = True
        txtppo2(p).Text = temptextpo2
        txtppo2(p).Enabled = False
    End If
    tempgasindex = lblgasindex(p).Caption
    tempgasused = Cbogasused(p).Text
    If Left(Cbogasused(p).Text, 1) = "0" Or Left(Cbogasused(p).Text, 1) = "4" Or Left(Cbogasused(p).Text, 1) = "5" Then
    Else
     txtppo2v.Text = txtppo2(p).Text
     lblo2.Caption = txtoxygen(p).Text
     lblhelium.Caption = txthelium(p).Text
     If Option1(p).Value = True Then
       cbogasindex.Text = lblgasindex(p).Caption
     End If
     If InStr(tempgasused, "Closed C") Then
        Option3.Enabled = True
        Option3.Value = True
        Option3.Enabled = False
        txtppo2v.Enabled = True
        CMDPPO2PLUS.Enabled = True
        Command4.Enabled = True
        Option4.Enabled = True
        Option4.Value = False
        Option4.Enabled = False
     Else
        If InStr(tempgasused, "Open C") Then
           Option4.Enabled = True
           Option4.Value = True
           Option4.Enabled = False
           Option3.Enabled = True
           Option3.Value = False
           Option3.Enabled = False
           txtppo2v.Enabled = False
           CMDPPO2PLUS.Enabled = False
           Command4.Enabled = False
        Else
           Option3.Value = False
           Option4.Value = True
           Option3.Enabled = True
           Option4.Enabled = True
           txtppo2v.Enabled = False
           CMDPPO2PLUS.Enabled = False
           Command4.Enabled = False
        End If
     End If
    End If
     SQL = "SELECT * FROM dpmaingaslist "
     SQL = SQL & " where dpmainid = '" & tempserialno & "' "
     SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     RS2.Edit
     RS2!dpgasmaxopdepth = CInt(txtmaxd(p).Text) * 10
     RS2.Update
     SQL = "SELECT * FROM seqdpprofile "
     SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
     SQL = SQL & " and gasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     tempnitrogen = 100 - CInt(txtoxygen(p).Text) - CInt(txthelium(p).Text)
     While RS2.EOF = False
        tempdepth = CDbl(RS2("depth"))
        tempcircuit = RS2("dpcircuit")
        RS2.Edit
        
        If InStr(tempcircuit, "Open C") Then
           tempoxygen = (CDbl(txtoxygen(p).Text) / 100#) * ((tempdepth / 10#) + 1)
           RS2!po2 = Format(CStr(tempoxygen), "###.00")
        End If
        RS2!dpo2 = txtoxygen(p).Text
'
        RS2.Update
        RS2.MoveNext
     Wend
     If MSFlexGrid3.Rows > 1 Then
        profilerecordexist = True
        reloadgriddata
     End If
   Else
      Title = "Error on System Validation.."
      MsgBox "(Max.Depth " & p & " value can not be less then 0 or more than 1000) or " & Chr(13) & "(Oxygen + Helium value can not be more than 100)", 48, Title
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      tempmaxdepth = RS2("dpgasmaxopdepth")
      txtmaxd(p).Text = tempmaxdepth
      txtmaxd(p).SetFocus
      SendKeys "{END}"
      Timer2.Enabled = False
   End If
Else
   txtmaxd(p).SetFocus
   SendKeys "{END}"
End If
   
End Sub
Private Sub validateppo2()
 If CDbl(txtppo2(p).Text) >= 0.15 And CDbl(txtppo2(p).Text) < 2# Then
     tempgasindex = lblgasindex(p).Caption
     SQL = "SELECT * FROM dpmaingaslist "
     SQL = SQL & " where dpmainid = '" & tempserialno & "' "
     SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
     RS2.Edit
     RS2!dpgaspo2setpoint = txtppo2(p).Text
     RS2.Update
     tempgasindex = lblgasindex(p).Caption
     tempgasused = Cbogasused(p).Text
    If Left(Cbogasused(p).Text, 1) = "0" Or Left(Cbogasused(p).Text, 1) = "4" Or Left(Cbogasused(p).Text, 1) = "5" Then
    Else
     txtppo2v.Text = txtppo2(p).Text
     lblhelium.Caption = txthelium(p).Text
     lblo2.Caption = txtoxygen(p).Text
     If Option1(p).Value = True Then
       cbogasindex.Text = lblgasindex(p).Caption
     End If
     If InStr(tempgasused, "Closed C") Then
        Option3.Enabled = True
        Option3.Value = True
        Option3.Enabled = False
        txtppo2v.Enabled = True
        CMDPPO2PLUS.Enabled = True
        Command4.Enabled = True
        Option4.Enabled = True
        Option4.Value = False
        Option4.Enabled = False
     Else
        If InStr(tempgasused, "Open C") Then
           Option4.Enabled = True
           Option4.Value = True
           Option4.Enabled = False
           Option3.Enabled = True
           Option3.Value = False
           Option3.Enabled = False
           txtppo2v.Enabled = False
           CMDPPO2PLUS.Enabled = False
           Command4.Enabled = False
        Else
           Option3.Value = False
           Option4.Value = True
           Option3.Enabled = True
           Option4.Enabled = True
           txtppo2v.Enabled = False
           CMDPPO2PLUS.Enabled = False
           Command4.Enabled = False
        End If
     End If
        If InStr(tempgasused, "5") = False Then
           If Option4.Value = True Then
              txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
              txtppo2v = Format(txtppo2v, "###.00")
           End If
        Else
           If Option4.Value = True Then
              txtppo2v = Format(txtppo2v, "###.00")
           End If
        End If
    End If
    
'     SQL = "SELECT * FROM dpmaingaslist "
'     SQL = SQL & " where dpmainid = '" & tempserialno & "' "
'     SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
'     Set RS2 = DB.OpenRecordset(SQL)
'     RS2.Edit '????
'     RS2!dpgaspo2setpoint = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
'     RS2.Update
     SQL = "SELECT * FROM seqdpprofile "
     SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
     SQL = SQL & " and gasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     While RS2.EOF = False
        RS2.Edit
        RS2!po2 = Format(txtppo2v.Text, "0.00") 'txtppo2v.Text
        RS2.Update
        RS2.MoveNext
     Wend
      If MSFlexGrid3.Rows > 1 Then
        profilerecordexist = True
        reloadgriddata
     End If
   Else
      Title = "Error on System Validation.."
      MsgBox "PPO2 value for Gas " & p & " can not be more than 2  ", 48, Title
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      tempppo2 = RS2("dpgaspo2setpoint")
      txtppo2(p).Text = tempppo2
      txtppo2(p).SetFocus
      SendKeys "{HOME}+{END}"
   End If
End Sub
Private Sub validatehelium()
If (txthelium(p).Text <> "") And (txtoxygen(p).Text <> "") And (txtmaxd(p).Text <> "") Then
  If CInt(txthelium(p).Text) >= 0 And CInt(txthelium(p).Text) <= 100 And ((CInt(txtoxygen(p).Text) + CInt(txthelium(p).Text)) < 101) Then
   If Cbogasused(p).Text = "5 - Deco Closed Circuit" Then
   Else
     txtppo2(p).Enabled = True
     S = txtoxygen(p).Text
     S = txthelium(p).Text
     txtppo2(p).Text = (CInt(txtoxygen(p).Text) / 100) * ((CInt(txtmaxd(p).Text) / 10) + 1)
     txtppo2(p).Text = Format(txtppo2(p).Text, "###.00")
     'txtppo2(p).Enabled = False
   End If
     tempgasindex = lblgasindex(p).Caption
     tempnitrogen = 100 - CInt(txtoxygen(p).Text) - CInt(txthelium(p).Text)
   If Left(Cbogasused(p).Text, 1) = "0" Or Left(Cbogasused(p).Text, 1) = "4" Or Left(Cbogasused(p).Text, 1) = "5" Then
   Else
     lblo2.Caption = txtoxygen(p).Text
     lblhelium.Caption = txthelium(p).Text
   End If
     SQL = "SELECT * FROM dpmaingaslist "
     SQL = SQL & " where dpmainid = '" & tempserialno & "' "
     SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     RS2.Edit
     RS2!dpgasnitrogen = tempnitrogen
     RS2!dpgashelium = txthelium(p).Text
     RS2.Update
      SQL = "SELECT * FROM seqdpprofile "
     SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
     SQL = SQL & " and gasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     While RS2.EOF = False
        tempdepth = CDbl(RS2("depth"))
        tempcircuit = RS2("dpcircuit")
        RS2.Edit
        RS2!dphe = txthelium(p).Text
        RS2.Update
        RS2.MoveNext
     Wend
     If MSFlexGrid3.Rows > 1 Then
        profilerecordexist = True
        reloadgriddata
     End If
   Else
      Title = "Error on System Validation.."
      MsgBox "(Helium " & p & " value can not be less then 0 or more than 100) or " & Chr(13) & "(Oxygen + Helium value can not be more than 100)", 48, Title
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      temphelium = RS2("dpgashelium")
       txthelium(p).Text = temphelium
'       txthelium(p).SetFocus
       SendKeys "{HOME}+{END}"
       Timer2.Enabled = False
   End If
Else
'   txthelium(p).SetFocus
   SendKeys "{END}"
End If
   
End Sub
Private Sub txtmaxd_LostFocus(Index As Integer)
   'txtmaxd_focus = 0
   If Left(Cbogasused(Index).Text, 1) = "0" Then Exit Sub
   p = txtmaxd(Index).Index
   validatemaxdepth
End Sub

Private Sub txtmaxdft_Change(Index As Integer)
lengthtxtmaxdft = Len(txtmaxdft(p))
For K = 1 To lengthtxtmaxdft '- 1
      If Asc(Mid$(txtmaxdft(p), K, 1)) > 45 And Asc(Mid$(txtmaxdft(p), K, 1)) < 59 Then
         tempcode = tempcode & Mid$(txtmaxdft(p), K, 1)
      Else
         tempcode = tempcode
      End If
   Next
txtmaxdft(p).Text = tempcode
SendKeys "{END}"
  If txtmaxd_focus = 1 Or Len(txtmaxdft(Index).Text) < 1 Or IsNumeric(txtmaxdft(Index).Text) = False Then Exit Sub
  txtmaxdft_focus = 1
  txtmaxd(Index).Text = Format(CStr(CDbl(txtmaxdft(Index).Text) / feetormeter_factor), "###")
  validatemaxdepth
  txtmaxdft_focus = 0
End Sub

Private Sub txtmaxdft_GotFocus(Index As Integer)
  'txtmaxdft_focus = 1
End Sub

Private Sub txtmaxdft_KeyPress(Index As Integer, KeyAscii As Integer)
p = txtmaxdft(Index).Index
If KeyAscii = 13 Then
  If CInt(txtmaxdft(Index).Text) < 0 Or CInt(txtmaxdft(Index).Text) > 2000 Then
     MsgBox " Depth value out of range !"
  End If
  
End If
End Sub

Private Sub txtmaxdft_LostFocus(Index As Integer)
  'txtmaxdft_focus = 0
End Sub

Private Sub txtoxygen_Change(Index As Integer)
lengthtxtoxygen = Len(txtoxygen(p))
For K = 1 To lengthtxtoxygen '- 1
      If Asc(Mid$(txtoxygen(p), K, 1)) > 45 And Asc(Mid$(txtoxygen(p), K, 1)) < 59 Then
         tempcode = tempcode & Mid$(txtoxygen(p), K, 1)
      Else
         tempcode = tempcode
      End If
   Next
txtoxygen(p).Text = tempcode
txtoxygen_LostFocus (Index)
SendKeys "{END}"
End Sub

Private Sub txtoxygen_GotFocus(Index As Integer)
   If Left(Cbogasused(Index).Text, 1) = "0" Then Exit Sub
   lblgasindex_Click (Index)
End Sub

Private Sub txtoxygen_KeyPress(Index As Integer, KeyAscii As Integer)
 p = txtoxygen(Index).Index
 current_index = p
If KeyAscii = 13 Then
   validateoxygen
   update_gas_graph (current_index)
End If
End Sub

Private Sub txtoxygen_LostFocus(Index As Integer)
   If Left(Cbogasused(Index).Text, 1) = "0" Then Exit Sub
   p = txtoxygen(Index).Index
   current_index = p
   validateoxygen
   update_gas_graph (current_index)
End Sub

Private Sub txtppo2_Change(Index As Integer)
lengthtxtppo2 = Len(txtppo2(p))
For K = 1 To lengthtxtppo2 '- 1
   If Asc(Mid$(txtppo2(p), K, 1)) > 45 And Asc(Mid$(txtppo2(p), K, 1)) < 59 Then
      tempcode = tempcode & Mid$(txtppo2(p), K, 1)
   Else
      tempcode = tempcode
   End If
   Next
txtppo2(p).Text = tempcode
SendKeys "{END}"
End Sub

Private Sub txtppo2_GotFocus(Index As Integer)
   If Left(Cbogasused(Index).Text, 1) = "0" Then Exit Sub
   lblgasindex_Click (Index)
End Sub

Private Sub txtppo2_KeyPress(Index As Integer, KeyAscii As Integer)
 p = txtppo2(Index).Index
If KeyAscii = 13 Then
 
  validateppo2
End If
End Sub

Private Sub txtppo2_LostFocus(Index As Integer)
   If Left(Cbogasused(Index).Text, 1) = "0" Then Exit Sub
   p = txtppo2(Index).Index
   validateppo2
End Sub

Private Sub txtppo2v_Change()
  If cmdgenerate.Visible = False Then cmdmodify.Visible = True
'  lengthtxtppo2v = Len(txtppo2v)
'For K = 1 To lengthtxtppo2v '- 1
'     If Asc(Mid$(txtppo2v, K, 1)) > 45 And Asc(Mid$(txtppo2v, K, 1)) < 59 Then
'         tempcode = tempcode & Mid$(txtppo2v, K, 1)
'      Else
'         tempcode = tempcode
'      End If
'   Next
'txtppo2v.Text = tempcode
'txtppo2v.Text = Format(txtppo2v, "###.00")
''SendKeys "{END}"
''MsgBox Cbogasused(p).Text & "ppo2v"
End Sub

Private Sub txtppo2v_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If CDbl(txtppo2v) < 0.15 Or CDbl(txtppo2v) > 2# Then
     MsgBox " PO2 value out of range !"
     txtppo2v.Text = "1.20"
  End If
  If cmdgenerate.Visible = False Then cmdmodify.Visible = True
  lengthtxtppo2v = Len(txtppo2v)
  For K = 1 To lengthtxtppo2v '- 1
     If Asc(Mid$(txtppo2v, K, 1)) > 45 And Asc(Mid$(txtppo2v, K, 1)) < 59 Then
         tempcode = tempcode & Mid$(txtppo2v, K, 1)
      Else
         tempcode = tempcode
      End If
   Next
  txtppo2v.Text = tempcode
  txtppo2v.Text = Format(txtppo2v, "###.00")
Else
   If KeyAscii <> 8 Then
   If KeyAscii < 46 Or KeyAscii > 57 Then
    '  MsgBox "Sorry, Only numeric characters allowed !"
    '  txtppo2v.SetFocus
    '  SendKeys "{HOME}+{END}"
   End If
   End If
End If
End Sub

Private Sub txttime_Change()
If cmdgenerate.Visible = False Then cmdmodify.Visible = True
If Len(txttime.Text) < 1 Then Exit Sub
lengthtxttime = Len(txttime)
 If lengthtxttime <> 0 Then
   For K = 1 To lengthtxttime '- 1
      If Asc(Mid$(txttime, K, 1)) > 45 And Asc(Mid$(txttime, K, 1)) < 59 Then
         tempcode = tempcode & Mid$(txttime, K, 1)
      Else
         tempcode = tempcode
      End If
   Next

txttime.Text = tempcode
End If
SendKeys "{END}"
If CInt(txttime) < 1 Or CInt(txttime) > 2000 Then
 If MSFlexGrid3.Rows > 1 Then
        MSFlexGrid3.Col = 2
        txttime.Text = MSFlexGrid3.Text
        SendKeys "{END}"
        Exit Sub
 Else
     txttime = "10"
     End If
End If
End Sub

Private Sub txttime_GotFocus()
     SendKeys "{HOME}+{END}"
     checkgasselected = False
     checkgasindex
     If checkgasselected = False Then
        Title = "Error on System Validation.."
        MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
        cbogasindex.SetFocus
     End If
End Sub

Private Sub txttime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If CInt(txttime) < 0 Or CInt(txttime) > 3000 Then
     MsgBox " Time value out of range !"
  End If
Else
   If KeyAscii <> 8 Then
      If KeyAscii < 46 Or KeyAscii > 57 Then
         'MsgBox "Sorry, Only numeric characters allowed !"
         'txtdepth.SetFocus
         SendKeys "{END}"
      End If
   End If
End If
End Sub

Private Sub txttime_LostFocus()
If Trim(cbogasindex) <> "Gas Index" Then
   If CInt(txttime) < 0 Or CInt(txttime) > 9999 Then
       MSFlexGrid3.Col = 2
       txttime.Text = MSFlexGrid3.Text
       SendKeys "{END}"
   End If
End If
End Sub

'Nick code start here

Private Sub Command1_Click()
  'atmtext.Text
  'safetytext.Text
  vimportdb_data
  'Sequence_deco
  display_deco_graph (0)
  mnufile.Enabled = True
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
 vhelium_half_vtime(1) = 1.51: vhelium_half_vtime(2) = 3.02: vhelium_half_vtime(3) = 4.72: vhelium_half_vtime(4) = 6.99: vhelium_half_vtime(5) = 10.21: vhelium_half_vtime(6) = 14.48: vhelium_half_vtime(7) = 20.53: vhelium_half_vtime(8) = 29.11: vhelium_half_vtime(9) = 41.2: vhelium_half_vtime(10) = 55.19: vhelium_half_vtime(11) = 70.69: vhelium_half_vtime(12) = 90.34: vhelium_half_vtime(13) = 115.29: vhelium_half_vtime(14) = 147.42: vhelium_half_vtime(15) = 188.24: vhelium_half_vtime(16) = 240.03
 vnitrogen_half_vtime(1) = 4#: vnitrogen_half_vtime(2) = 8#: vnitrogen_half_vtime(3) = 12.5: vnitrogen_half_vtime(4) = 18.5: vnitrogen_half_vtime(5) = 27#: vnitrogen_half_vtime(6) = 38.3: vnitrogen_half_vtime(7) = 54.3: vnitrogen_half_vtime(8) = 77#: vnitrogen_half_vtime(9) = 109#: vnitrogen_half_vtime(10) = 146#: vnitrogen_half_vtime(11) = 187#: vnitrogen_half_vtime(12) = 239#: vnitrogen_half_vtime(13) = 305#: vnitrogen_half_vtime(14) = 390#: vnitrogen_half_vtime(15) = 498#: vnitrogen_half_vtime(16) = 635#
  '4.00, 8.00, 12.50, 18.50, 27.00, 38.30, 54.30, 77.00, 109.0, 146.0, 187.0, 239.0, 305.0, 390.0, 498.0, 635.0,
  '1.51, 3.02, 4.72, 6.99, 10.21, 14.48, 20.53, 29.11, 41.2, 55.19, 70.69, 90.34, 115.29, 147.42, 188.24, 240.03
 
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
 minimum_vdeco_vstop_vtime = 1.0000001  '0.5
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
              Me.MousePointer = 0
              Xstart = -1
              Ystart = -1
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
            t1print (vbCrLf)
           ' If systemversion = "Lite" Then
               printtogrid2lite
           ' Else
               printtogrid2
           ' End If
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
        '    Else
               printtogrid4
        '    End If
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
            If (critical_volume_comparison <= 1#) Or buhl_mode = 2 Then
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
            vhmx_first_stop = 0
            vhmx_last_stop = 0
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
                printtogrid5lite
                printtogrid5
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
                test = vdeco_vceiling_vi
                If vhmx_first_stop = 0 Then
                  vhmx_first_stop = vdeco_vceiling_vi
                  vhmx_first_stop_depth = vdeco_vstop_vdepth
                  vhmx_tol_pressure_first_stop = vhmx_tol_pressure
                  vhmx_tol_pressure_bnorm_first = vhmx_tol_pressure_bnorm
                  vhmx_tol_pressure_bvhmx_first = vhmx_tol_pressure_bvhmx
                  vhmx_ptissue_first = vhmx_ptissue
'                  gf_1 = (vhmx_tol_pressure_bvhmx - vdeco_vstop_vdepth) / (vhmx_tol_pressure_bnorm - vdeco_vstop_vdepth)
                  gf_1 = (vhmx_ptissue - (vdeco_vstop_vdepth + barometric_vpressure)) / (vhmx_tol_pressure_bnorm - (vdeco_vstop_vdepth + barometric_vpressure))
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
                    printtogrid5lite
                    printtogrid5
                    
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
                   ' If systemversion = "Lite" Then
                       printtogrid5lite
                   ' Else
                       printtogrid5
                   ' End If
                    rate = ratelasttemp
                End If
                starting_vdepth = vdeco_vstop_vdepth
                next_vstop = vdeco_vstop_vdepth - vstep_size
                vdeco_vstop_vdepth = next_vstop
                last_run_vtime = run_vtime
                If vdeco_vstop_vdepth < 6 Then
                  If vhmx_last_stop = 0 Then
                    vhmx_last_stop = vdeco_vceiling_vi
                    vhmx_tol_pressure_last_stop = vhmx_tol_pressure
                    vhmx_tol_pressure_first_stop = vhmx_tol_pressure
                    vhmx_tol_pressure_bnorm_last = vhmx_tol_pressure_bnorm
                    vhmx_tol_pressure_bvhmx_last = vhmx_tol_pressure_bvhmx
                    vhmx_ptissue_first = vhmx_ptissue
'                    gf_2 = (vhmx_tol_pressure_bvhmx - vdeco_vstop_vdepth) / (vhmx_tol_pressure_bnorm - vdeco_vstop_vdepth)
                    gf_2 = (vhmx_ptissue - (vdeco_vstop_vdepth + barometric_vpressure)) / (vhmx_tol_pressure_bnorm - (vdeco_vstop_vdepth + barometric_vpressure))
                    Label6.Caption = "EGF" + vbCrLf + Format(gf_1 * 100, "0") + "/" + Format(gf_2 * 100, "0") 'un tweaked value
                    Label6.Caption = "EGF" + vbCrLf + Format(gf_1 * 100 + ((0.5 - gf_1) * 7 / 0.4), "0") + "/" + Format(gf_2 * 100 - ((gf_2 - 0.5) * 7 / 0.5), "0")
                    Cls
                  End If
                End If
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
            dum = ppo2exposuretime(0#, surface_interval_vtime * 0.01)   ' cns_current = cns_current - (barometric_vpressure * surface_interval_vtime * 60)
            'If cns_current < 0# Then cns_current = 0#
            'otu_current = otu_current - (barometric_vpressure * surface_interval_vtime * 60)
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
    
    If (buhl_mode = 2) Then
            nfrac = vnitrogen_vpressure(i) / gas_loading
            hefrac = vhelium_vpressure(i) / gas_loading
            If (hefrac = 0) Then hefrac = 0.005
            If (nfrac = 0) Then nfrac = 0.005
            atotal = (nfrac * an2(i - 1) + hefrac * ahe(i - 1)) / (nfrac + hefrac)
            btotal = (nfrac * bn2(i - 1) + hefrac * bhe(i - 1)) / (nfrac + hefrac)
            amultfactorcalc (i - 1)
            btotal = btotal * amultfreal
            'atotal = atotal * 0.9 * (1# - (CDbl(safetytext.Text) / 300#))
            'btotal = btotal * 1.1 * (1# + CDbl(safetytext.Text) / 300#)
            
            tolerated_ambient_vpressure = btotal * (vnitrogen_vpressure(i) + vhelium_vpressure(i) - (atotal * units_factor))
            If (tolerated_ambient_vpressure < 0#) Then
              tolerated_ambient_vpressure = 0#
            End If
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
            amultfactorcalc (i - 1)
            btotal = btotal * amultfreal
        
            'atotal = atotal * 0.9 * (1# - (CDbl(safetytext.Text) / 300#))
            'btotal = btotal * 1.1 * (1# + CDbl(safetytext.Text) / 300#)
        
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
        vhmx_compartment_vdeco_vceiling_bnorm(i - 1) = btotal * (vnitrogen_vpressure(i) + vhelium_vpressure(i) - atotal * units_factor)
        amultfactorcalc (i - 1)
        btotal = btotal * amultfreal
        
        ''atotal = atotal * 0.9 * (1# - (CDbl(safetytext.Text) / 300#))
        'btotal = btotal * 1.1 * (1# + CDbl(safetytext.Text) / 300#)
        
        buhlptoln2(i - 1) = btotal * (vnitrogen_vpressure(i) + vhelium_vpressure(i) - atotal * units_factor)
        vhmx_compartment_vdeco_vceiling_bvhmx(i - 1) = buhlptoln2(i - 1)
        vhmx_gastotal_pressure(i - 1) = vnitrogen_vpressure(i) + vhelium_vpressure(i)
        
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
        vhmx_compartment_vdeco_vceiling(i - 1) = tolerated_ambient_vpressure
    Next i
    vdeco_vceiling_vdepth = compartment_vdeco_vceiling(1)
    vdeco_vceiling_vi = 0
    For i = 2 To 16
        If (vdeco_vceiling_vdepth < compartment_vdeco_vceiling(i)) Then
          vdeco_vceiling_vi = i - 1
          vhmx_tol_pressure = vhmx_compartment_vdeco_vceiling(i - 1)
          vhmx_tol_pressure_bvhmx = vhmx_compartment_vdeco_vceiling_bvhmx(i - 1)
          nfrac = vnitrogen_vpressure(i) / (vnitrogen_vpressure(i) + vhelium_vpressure(i))
          hefrac = vhelium_vpressure(i) / (vnitrogen_vpressure(i) + vhelium_vpressure(i))
          If (hefrac = 0) Then hefrac = 0.005
          If (nfrac = 0) Then nfrac = 0.005
          atotal = (nfrac * an2(i - 1) + hefrac * ahe(i - 1)) / (nfrac + hefrac)
          btotal = (nfrac * bn2(i - 1) + hefrac * bhe(i - 1)) / (nfrac + hefrac)
          vhmx_tol_pressure_bnorm = (vhmx_compartment_vdeco_vceiling_bvhmx(i - 1) / btotal) + atotal * units_factor
          vhmx_ptissue = vhmx_gastotal_pressure(i - 1)
        End If
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
' MsgBox tempserialno
  diveplan_num = tempserialno
'  MSFlexGrid1.Col = 2
  surface_interval_vtime = 1#  'MSFlexGrid1.Text
  
  SQL = "SELECT * FROM seqdpprofile"
  SQL = SQL & " where dpprofileid = '" & diveplan_num & "' "
  SQL = SQL & " order by dpnumseq "
  Set RS = DB.OpenRecordset(SQL)
  If RS.EOF Then
'    MsgBox "Add Profile Plan Points before calculating decompression"
    vimportdb_data = 99
    Exit Function
  End If
  
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
     barused(i + 1) = 0#
  Next i
  'Exit Function
'  MsgBox Plan_Gas_list_n2(1)
  'If j < Number_Dives Then repetitive_dive_flag = 1 Else repetitive_dive_flag = 0
  repetitive_dive_flag = 0
  'If j = 1 Then repetitive_dive_flag = -1 'initialise data
  repetitive_dive_flag = -1 'initialise data
  barometric_vpressure = CDbl(atmtext.Text) / 100# 'this gives a value of 10 for sea level..
  Sequence_deco
  If no_deco_found > 0 Then
    decoresultgrid.Rows = 0
    Picture1.Cls
    MsgBox "Bad depth points!! Re-enter sensible values!"
  End If
  For i = 0 To 9
    gasusageupdate (i) 'gasusage(i).Caption = Format(barused(i + 1) * feetormeter_factor * CDbl(txtbreathrate.Text) / CDbl(txtcylcap.Text), "####")
    If barused(i + 1) < 1# Then
       gasusage(i).Visible = False
    Else
       If Check2(i).Value = 0 Then
          gasusage(i).Visible = True
       Else
          gasusage(i).Visible = False
       End If
    End If
    
  Next
  ' Decoresult(j - 1).Text = Text10.Text
  'T = j
  'Duplicategrid
 'Next j
Screen.MousePointer = 0
End Function

'Private Sub vhighlite_line(Index As Integer)
'  MSFlexGrid3.Row = Index
'  MSFlexGrid3_Click
'End Sub

Private Sub display_deco_text()
Dim j As Integer
   If rowindentified = "" Then
     Exit Sub
   Else
     If rowindentified = 0 Then Exit Sub
   End If
  For j = 0 To 9
    
   '   decoresultgrid.Visible = False
   '   decoresultgridlite.Visible = False
      decoresultgrid.Enabled = True
      decoresultgridlite.Enabled = True
    'decoresultgrid(j).Locked = True
    If rowindentified = j Then
      Frame3.Caption = "Decompression Result for " & "Dive " & CStr(j) & " - " & txtplanno.Text

      MSFlexGrid4.Visible = False
      If decoresultgrid.Visible = True Then
         decoresultgrid.Visible = True
         decoresultgridlite.Visible = False
      Else
         decoresultgrid.Visible = False
         decoresultgridlite.Visible = True
      End If
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

Private Sub initialgrid4()
decoresultgrid.Cols = 10
decoresultgrid.Col = 7
decoresultgrid.Rows = 1
decoresultgrid.Row = 0
decoresultgrid.Text = "No."
decoresultgrid.Col = 1
decoresultgrid.Text = "Duration"
decoresultgrid.Col = 2
decoresultgrid.Text = "RunTime"
decoresultgrid.Col = 3
decoresultgrid.Text = "Mix"
decoresultgrid.Col = 0
decoresultgrid.Text = "Depth"
decoresultgrid.Col = 5
decoresultgrid.Text = "CNS"
decoresultgrid.Col = 6
decoresultgrid.Text = "OTU"
decoresultgrid.Col = 8
decoresultgrid.Text = "Rate"
decoresultgrid.Col = 4
decoresultgrid.Text = "SPO2"
decoresultgrid.ColWidth(0) = 580
decoresultgrid.ColWidth(1) = 720
decoresultgrid.ColWidth(2) = 720
decoresultgrid.ColWidth(3) = 750
decoresultgrid.ColWidth(4) = 550
decoresultgrid.ColWidth(5) = 550
decoresultgrid.ColWidth(6) = 550
decoresultgrid.ColWidth(7) = 0 '450
decoresultgrid.ColWidth(8) = 1000 '650
decoresultgridlite.Cols = 9
decoresultgridlite.Col = 7
decoresultgridlite.Rows = 1
decoresultgridlite.Row = 0
decoresultgridlite.Text = "No."
decoresultgridlite.Col = 1
decoresultgridlite.Text = "Duration"
decoresultgridlite.Col = 2
decoresultgridlite.Text = "RunTime"
decoresultgridlite.Col = 3
decoresultgridlite.Text = "Mix"
decoresultgridlite.Col = 0
decoresultgridlite.Text = "Depth"
decoresultgridlite.Col = 5
decoresultgridlite.Text = "CNS"
decoresultgridlite.Col = 6
decoresultgridlite.Text = "OTU"
decoresultgridlite.Col = 8
decoresultgridlite.Text = "Rate"
decoresultgridlite.Col = 4
decoresultgridlite.Text = "SPO2"
decoresultgridlite.ColWidth(0) = 580
decoresultgridlite.ColWidth(1) = 720
decoresultgridlite.ColWidth(2) = 720
decoresultgridlite.ColWidth(3) = 750
decoresultgridlite.ColWidth(4) = 550
decoresultgridlite.ColWidth(5) = 550
decoresultgridlite.ColWidth(6) = 550
decoresultgridlite.ColWidth(7) = 0 '450
decoresultgridlite.ColWidth(8) = 1000 '650
End Sub

Private Sub vhighlite_line(Index As Integer)
  MSFlexGrid3.Row = Index
  MSFlexGrid3_Click
End Sub



'nick code end here

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, Xmouse As Single, Ymouse As Single)
Dim i As Integer
Dim xc As Single
Dim timet As Single
detectdouclick = False
  If decoresultgrid.Rows < 4 Then Exit Sub

  xc = Xmouse
  xc = xc / Picture1.Width
  timet = xc * runtime_graph
  For i = 1 To row_count
    If X(i) > timet Then Exit For
  Next i
  
  
'  msgres = MsgBox("Time: " & Format(timet, "###0mins") & vbCrLf & "Depth: " & Format(Y(i) * feetormeter_factor, "###0" & feetormeter_shortstring), vbOKOnly, "Dive Point")
  
  If deco_grid_display_last >= 0 Then
    If decoresultgrid.Rows < 4 Then Exit Sub
    If decoresultgrid.Visible = True Then
       decoresultgrid.Row = deco_grid_display_rowlast
       decoresultgrid.Col = 0
       decoresultgrid.CellBackColor = deco_grid_display_celllast
    Else
       decoresultgridlite.Row = deco_grid_display_rowlast
       decoresultgridlite.Col = 0
       decoresultgridlite.CellBackColor = deco_grid_display_celllast
    End If
  End If
  If decoresultgrid.Visible = True Then
     If i <= 4 Then decoresultgrid.TopRow = 1 Else decoresultgrid.TopRow = i - 4
     decoresultgrid.Col = 0
     decoresultgrid.Row = i
     deco_grid_display_last = 0 'deco_grid_display
     deco_grid_display_rowlast = (i)
     deco_grid_display_celllast = decoresultgrid.CellBackColor
     decoresultgrid.CellBackColor = vbBlue
  Else
     i = (i / 2) + 1
     If i <= 4 Then decoresultgridlite.TopRow = 1 Else decoresultgridlite.TopRow = i - 4
     decoresultgridlite.Col = 0
     decoresultgridlite.Row = i
     deco_grid_display_last = 0 'deco_grid_display
     deco_grid_display_rowlast = (i)
     deco_grid_display_celllast = decoresultgridlite.CellBackColor
     decoresultgridlite.CellBackColor = vbBlue
  End If
rowindentified = Fix((deco_grid_display_rowlast + 1) / 2)
If rowindentified >= MSFlexGrid3.Rows Then rowindentified = 1 'MSFlexGrid3.Rows - 1
msflexgriddoclick
xc = xc
  'picture1.m
Text1.Visible = True


Xstart = Xmouse
Ystart = Ymouse
If decoresultgrid.Visible = True Then
   decoresultgrid.Col = 0
   decoresultgrid.Row = i
   Text1.Text = decoresultgrid.Text + " "
   decoresultgrid.Col = 1
   Text1.Text = Text1.Text + decoresultgrid.Text
Else
   decoresultgridlite.Col = 0
   decoresultgridlite.Row = i
   Text1.Text = decoresultgridlite.Text + " "
   decoresultgridlite.Col = 1
   Text1.Text = Text1.Text + decoresultgridlite.Text
End If
Text1.Top = Ymouse - Text1.Height
Text1.Left = Xmouse
Text1.Width = 900
End Sub

Private Sub safetytext_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    If CInt(safetytext.Text) < 0 Or CInt(safetytext.Text) > 50 Then
     MsgBox "Value must be between 0 and 50 !"
     safetytext.Text = "0"
    Else
         safetytext.SetFocus
         SendKeys "{HOME}+{END}"
         Command1_Click
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
'  Else
'    Command1_Click
  End If
End Sub


Private Sub display_deco_graph(grid_num As Integer)
Dim i As Integer
Dim K As Integer
Dim maxd As Integer
Dim dive_time As Integer
Dim deco_section As Integer
Dim C As Long

On Error Resume Next
'Add by Goh on 08/11/2004 12:30pm
If decoresultgrid.Rows = 0 Then
  deco_update = 0
Else
  deco_update = 1
End If
'Add by Goh on 08/11/2004 12:30pm
  If deco_update = 0 Then
  '  For i = 0 To 9
  '    decoresultgrid(i).Rows = 0
  '  Next i
    Picture1.Visible = False
    mnugaslist_Click
    Exit Sub
  Else
    Picture1.Visible = True
    'deco_grid_display = grid_num
    mnugraph_Click
  End If
  row_count = 1
  deco_section = 0 '4 '6
  Picture1.Cls
  'Picture1.Line (0, 0)-(0, 0), vbBlue
  maxd = 0
  dive_time = 0
 ' If decographversion = "Pro" Then
     For i = 1 To decoresultgrid.Rows - 1
       decoresultgrid.Row = i
       decoresultgrid.Col = 7
       If IsNumeric(decoresultgrid.Text) = True Then
          decoresultgrid.Col = 2
          X(row_count) = CSng(Left(decoresultgrid.Text, Len(decoresultgrid.Text) - 3))
          decoresultgrid.Col = deco_section
          Y(row_count) = CSng(CDbl(Left(decoresultgrid.Text, Len(decoresultgrid.Text) - Len(feetormeter_shortstring))) / feetormeter_factor)
          If Y(row_count) > maxd Then maxd = Y(row_count)
          row_count = row_count + 1
       Else
          If CStr(decoresultgrid.Text) = "" Then Exit For
          If InStr(1, (decoresultgrid.Text), "No") Then deco_section = 4
       End If
     Next i
 ' Else
  '   For i = 1 To decoresultgridlite.Rows - 1
  '     decoresultgridlite.Row = i
  '     decoresultgridlite.Col = 7
  '     If IsNumeric(decoresultgridlite.Text) = True Then
  '        decoresultgridlite.Col = 2
   '       X(row_count) = CSng(Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - 3))
  '        decoresultgridlite.Col = deco_section
  '        Y(row_count) = CSng(CDbl(Left(decoresultgridlite.Text, Len(decoresultgridlite.Text) - Len(feetormeter_shortstring))) / feetormeter_factor)
  '        If Y(row_count) > maxd Then maxd = Y(row_count)
  '        row_count = row_count + 1
  '     Else
  '        If CStr(decoresultgridlite.Text) = "" Then Exit For
  '        If InStr(1, (decoresultgridlite.Text), "No") Then deco_section = 4
  '     End If
  '   Next i
  ' End If
  If row_count < 4 Then Exit Sub
  row_count = row_count - 1
  lbly(0).Caption = "0"
  lbly(1).Caption = Fix(maxd * feetormeter_factor / 3)
  lbly(2).Caption = Fix(maxd * feetormeter_factor * 2 / 3)
  lbly(3).Caption = Fix(maxd * feetormeter_factor)
  lblx(0).Caption = "0"
  lblx(1).Caption = Fix(X(row_count) / 3)
  lblx(2).Caption = Fix(X(row_count) * 2 / 3)
  lblx(3).Caption = Fix(X(row_count))
  yscale = ((Picture1.Height - 150) / maxd)
  xscale = ((Picture1.Width - 150) / X(row_count))
  runtime_graph = X(row_count)
  'ysacle = Picture1.ScaleHeight
  'xscale = Picture1.ScaleWidth
  
  X(0) = 0
  Y(0) = 0
'  Picture1.Line (0, 0)-(x(1) * xscale, y(2) * yscale), vbBlue
'  Picture1.Line (x(1), y(2))-(x(2) * xscale, y(2) * yscale), vbBlue
'  Picture1.Line (0, 0)-(x(1) * xscale, y(2) * yscale), vbBlue
'  Picture1.Line (0, 0)-(x(1) * xscale, y(2) * yscale), vbBlue
'  Picture1.Line (0, 0)-(x(1) * xscale, y(2) * yscale), vbBlue
  decoresultgrid.Col = 9
  For i = 1 To row_count
     decoresultgrid.Row = i
     If InStr(1, decoresultgrid.Text, "0") Then C = vbWhite
     If InStr(1, decoresultgrid.Text, "1") Then C = vbYellow
     If InStr(1, decoresultgrid.Text, "2") Then C = vbCyan
     If InStr(1, decoresultgrid.Text, "3") Then C = &H808080
     If InStr(1, decoresultgrid.Text, "4") Then C = vbGreen
     If InStr(1, decoresultgrid.Text, "5") Then C = vbMagenta
     If InStr(1, decoresultgrid.Text, "6") Then C = vbBlue
     If InStr(1, decoresultgrid.Text, "7") Then C = &HFF8080
     If InStr(1, decoresultgrid.Text, "8") Then C = &H8080FF
     If InStr(1, decoresultgrid.Text, "9") Then C = vbRed
     Picture1.Line (X(i - 1) * xscale, Y(i - 1) * yscale)-(X(i) * xscale, Y(i) * yscale), C
  Next i
  Frame8.Visible = True
  
End Sub

Private Function lblgasvrupdate()
  If IsNumeric(lblhelium.Caption) Then Else lblhelium.Caption = "0"
  If Fix(lblhelium.Caption) > 0 Then
    lblgasvr.Caption = "TX" + lblo2.Caption + "/" + lblhelium.Caption
  Else
    If lblo2.Caption = "21" Then
      lblgasvr.Caption = "AIR"
    Else
      lblgasvr.Caption = "NX" + lblo2.Caption
    End If
  End If
End Function

Private Function gasusageupdate(i As Integer)
On Error Resume Next
gasusage(i).Caption = Format(barused(i + 1) * psiorbar_factor * CDbl(txtbreathrate2(i).Text) / CDbl(txtcylcap2(i).Text), "####") 'er_factor * CDbl(txtbreathrate.Text) / CDbl(txtcylcap.Text), "####")
'printf(stitle, " Res  %s   SAC %4.0f %4.0f  %5.1f", cyl_or_vcf, hprealfloat_reserve[HP_RES]*psifactor, hprealfloat_reserve[HP_CYLINDER_SIZE], hprate_float_surface_cylinder*10.00*ltrorcuft);
End Function

Private Function validate_data()
  If CDbl(txtdepth.Text) > CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) Then
    'MsgBox "Depth too deep for gas - change gas or depth depth"
    txtdepth.Text = txtmaxd(CInt(Right(cbogasindex.Text, 1))) '"10"
    txtdepthft.Text = Format(CDbl(txtmaxd(CInt(Right(cbogasindex.Text, 1)))) * feetormeter_factor, "###") '"30"
    SendKeys "{HOME}+{END}"
    Exit Function
  End If
  If CInt(txtdepth) < 0 Or CInt(txtdepth) > 2000 Then
     MsgBox " Depth value out of range !"
  Else
     If Option4.Value = True And CInt(txtdepth) <> 0 Then
      If IsNumeric(lblo2.Caption) Then
        txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
        txtppo2v = Format(txtppo2v, "###.00")
        'txttime.SetFocus
      Else
'        MsgBox "No ppo2"
      End If
     End If
  End If
  If cmdgenerate.Visible = False Then cmdmodify.Visible = True
  If txtdepthft_focus = 1 Or Len(txtdepth.Text) < 1 Or IsNumeric(txtdepth.Text) = False Then Exit Function
  txtdepth_focus = 1
  txtdepthft.Text = Format(CStr(CDbl(txtdepth.Text) * feetormeter_factor), "###0")
  txtdepth_focus = 0

End Function
Private Sub checkgasuse(i As Integer)
Select Case Cbogasused(i).Text
     Case "0 - Not Used"
        Check1(i).Value = 0
        Check2(i).Value = 0
        Decochk(i).Value = 0
        gasusage(i).Visible = False
        txthelium(i).Enabled = False
        txtoxygen(i).Enabled = False
        txtmaxd(i).Enabled = False
        txtppo2(i).Enabled = False
        Check2(i).Enabled = False
        Decochk(i).Enabled = True 'False
        txtcylcap2(i).Enabled = False
        txtbreathrate2(i).Enabled = False
        
     Case "1 - Open Circuit"
        Check1(i).Value = 1
        Check2(i).Value = 0
        Decochk(i).Value = 0
        Decochk(i).Enabled = False
        gasusage(i).Visible = True
        txtcylcap2(i).Visible = True
        txtbreathrate2(i).Visible = True
     Case "2 - Closed Circuit"
        Check1(i).Value = 1
        Check2(i).Value = 1
        Decochk(i).Value = 0
        Decochk(i).Enabled = False
        gasusage(i).Visible = False
        txtcylcap2(i).Visible = False
        txtbreathrate2(i).Visible = False
     Case "4 - Deco Open Circuit"
        Check1(i).Value = 1
        Check2(i).Value = 0
        Decochk(i).Value = 1
        gasusage(i).Visible = True
        txtcylcap2(i).Visible = True
        txtbreathrate2(i).Visible = True
     Case "3 - Open & Closed"
        Check1(i).Value = 1
        Check2(i).Value = 1
        Decochk(i).Value = 0
        
     Case "5 - Deco Closed Circuit"
        Check1(i).Value = 1
        Check2(i).Value = 1
        Decochk(i).Value = 1
        gasusage(i).Visible = False
        txtcylcap2(i).Visible = False
        txtbreathrate2(i).Visible = False
   End Select
     
End Sub


Private Function amultfactorcalc(i As Integer)

Dim timehalf As Double
Dim time_surf As Double
Dim maxd24 As Double
Dim maxd24t As Double
Dim afact As Double
Dim ff As Double
Dim seconds_total_t As Long
Dim j As Integer
Dim K As Integer
Dim l As Integer

Dim amultf As Double

timehalf = 6#
time_surf = 1#
afact = 0#
afact2 = 0#

ff = 0#

seconds_total_t = 0

amultf = 1#
maxdstart30 = 30
maxdnorm80 = 40

vhmx_maxd_factor = 1# + 0.01 * CDbl(vhmx_text(0).Text)
vhmx_mid_factor = 1# + 0.01 * CDbl(vhmx_text(1).Text)
vhmx_stop_factor = 1# + 0.01 * CDbl(vhmx_text(2).Text)
vhmx_safe_factor = 1# + 0.01 * CDbl(vhmx_text(3).Text)

  MSFlexGrid3.Col = 1
  maxd24 = CDbl(MSFlexGrid3.Text) ' CDbl(txtdepth.Text) ' change maxdepth;
  MSFlexGrid3.Col = 2
  divemins_bottom_vgm = CDbl(MSFlexGrid3.Text) 'CInt(txttime.Text) 'divemins
  MSFlexGrid3.Col = 1

  If maxd24 > (maxdstart30 + maxdnorm80) Then
 '   maxd24 = maxdstart30 + maxdnorm80 '; //150.0) maxd24=150.0;
  Else
    If maxd24 < 1# Then maxd24 = 1#
  End If
  afact = (maxd24 - maxdstart30) / maxdnorm80
  If (afact < 0#) Then afact = 0#
  
  timedstart30 = 30
  timednorm80 = 30 '60
  afact2 = (divemins_bottom_vgm - timedstart30) / timednorm80
  If (afact2 < 0#) Then afact2 = 0#
  afact = afact + afact2
  If (afact > 1.5) Then
    afact = 1.5 'limit
  End If
  
  amultf = (amult(i) - 1#) * afact '; //twist here!!!!!!!!!!!


  amultf = amultf + 1#
  If (amultf < 0.5) Then amultf = 0.5
  amultf = amultf * vhmx_safe_factor


  afact = (maxd24 - 80#) / 40# '; //mid tweak for 80+ dives
  If (afact < 0#) Then afact = 0#
  If (afact > 1#) Then afact = 1#
  If (afact2 < 0#) Then afact2 = 0#
  If (afact2 > 1#) Then afact2 = 1#
  afact = afact + afact2
  afact = afact / 10# '; //make same order as vhmx factors so make max=0.10
  If (afact > 0.1) Then afact = 0.1 '; //max of 10% add

  If (i <= 7) Then '{ //fast tissues
    ff = ((vhmx_mid_factor + afact) - vhmx_maxd_factor) * (i / 7#) + vhmx_maxd_factor
  Else '{ //slow tissues
    ff = (vhmx_stop_factor - (vhmx_mid_factor + afact)) * (((i - 8)) / 7#) + (vhmx_mid_factor + afact)
  End If
  amultf = amultf * ff
  

  If (amultf < 0.8) Then amultf = 0.8 '; //50;
  amultfreal = amultf
  
  If amultfreal > 1.99 Then amultfreal = 1.99
  
'  amultfreal = 1

End Function

Private Sub vhmx_text_Change(Index As Integer)
  If IsNumeric(vhmx_text(Index).Text) Then
  Else
      vhmx_text(Index).Text = "0"
  End If
  If CInt(vhmx_text(Index).Text) > 20 Then vhmx_text(Index).Text = "20"
  If CInt(vhmx_text(0).Text) > 10 Then vhmx_text(0).Text = "10"
  If CInt(vhmx_text(Index).Text) < -20 Then vhmx_text(Index).Text = "-20"
  vhmx_text(Index).BackColor = vbWhite
  If CInt(vhmx_text(Index).Text) <= -1 Then vhmx_text(Index).BackColor = vbRed
  If CInt(vhmx_text(Index).Text) >= 5 Then vhmx_text(Index).BackColor = vbGreen
'  If CInt(vhmx_text(index).Text) < 0 Then vhmx_text(index).BackColor = vbRed
End Sub

Private Sub vhmx_up_Click(Index As Integer)
  If Index = 4 Then
    vhmx_text(0).Text = CStr(CInt(vhmx_text(0).Text) + 1)
    vhmx_text(1).Text = CStr(CInt(vhmx_text(1).Text) + 1)
    vhmx_text(2).Text = CStr(CInt(vhmx_text(2).Text) + 1)
  Else
    vhmx_text(Index).Text = CStr(CInt(vhmx_text(Index).Text) + 1)
  End If
  For Index = 0 To 2
    If CInt(vhmx_text(Index).Text) > 20 Then vhmx_text(Index).Text = "20"
    If CInt(vhmx_text(0).Text) > 10 Then vhmx_text(0).Text = "10"
    If CInt(vhmx_text(Index).Text) < -20 Then vhmx_text(Index).Text = "-20"
  Next Index
  Cls
  If Timer2.Enabled = False Then cmdgenerate_Click
End Sub

Private Sub vhmx_up_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 0
timecount = 0
ppo2count = 0
gascount = 0
Timer2.Enabled = True
timer2buffer = 0
If Index = 4 Then
  vhmxcount = 105
Else
  vhmxcount = Index + 5
End If
End Sub

Private Sub vhmx_up_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer2.Enabled = False
  Cls
  cmdgenerate_Click
End Sub

Private Sub vhmx_down_Click(Index As Integer)
  If Index = 4 Then
    vhmx_text(0).Text = CStr(CInt(vhmx_text(0).Text) - 1)
    vhmx_text(1).Text = CStr(CInt(vhmx_text(1).Text) - 1)
    vhmx_text(2).Text = CStr(CInt(vhmx_text(2).Text) - 1)
  Else
    vhmx_text(Index).Text = CStr(CInt(vhmx_text(Index).Text) - 1)
  End If
  For Index = 0 To 2
    If CInt(vhmx_text(Index).Text) > 20 Then vhmx_text(Index).Text = "20"
    If CInt(vhmx_text(0).Text) > 10 Then vhmx_text(0).Text = "10"
    If CInt(vhmx_text(Index).Text) < -20 Then vhmx_text(Index).Text = "-20"
  Next Index
  Cls
  If Timer2.Enabled = False Then cmdgenerate_Click
End Sub

Private Sub vhmx_down_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
depthcount = 0
timecount = 0
ppo2count = 0
gascount = 0
Timer2.Enabled = True
timer2buffer = 0
If Index = 4 Then
  vhmxcount = 100
Else
  vhmxcount = Index + 1
End If
End Sub

Private Sub vhmx_down_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer2.Enabled = False
  Cls
  cmdgenerate_Click
End Sub


