VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Planprofile 
   BackColor       =   &H80000013&
   Caption         =   "Diveplan Profile"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   -1095
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
   Icon            =   "Planprofile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.CommandButton cmdsaveas 
      Caption         =   "Save As"
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
      Left            =   5160
      TabIndex        =   119
      ToolTipText     =   "Save as sequential dive"
      Top             =   7620
      Width           =   975
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
      Height          =   3380
      Left            =   6120
      TabIndex        =   26
      Top             =   3840
      Width           =   5760
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   100
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   45
         Text            =   "Planprofile.frx":2CFA
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000013&
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
      Left            =   9960
      MaskColor       =   &H80000013&
      TabIndex        =   9
      Top             =   7560
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000013&
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
      Left            =   10920
      MaskColor       =   &H80000013&
      TabIndex        =   8
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
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
      Left            =   7560
      TabIndex        =   11
      Top             =   7620
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
      Left            =   6240
      TabIndex        =   10
      Top             =   7620
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
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "View the single dive plan list"
      Top             =   7620
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
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Save the dive plan"
      Top             =   7620
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
      Left            =   3240
      TabIndex        =   5
      ToolTipText     =   "Print the details"
      Top             =   7620
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
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Key in the details"
      Top             =   7620
      Width           =   855
   End
   Begin VB.CommandButton cmdremove 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3840
      TabIndex        =   20
      Top             =   3315
      Width           =   935
   End
   Begin VB.CommandButton cmdinsert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   19
      Top             =   3315
      Width           =   975
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
      Left            =   1635
      Picture         =   "Planprofile.frx":2D01
      TabIndex        =   12
      Top             =   1305
      Width           =   190
   End
   Begin VB.TextBox txttime 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   805
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
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Generate the Decompression result"
      Top             =   7620
      Width           =   975
   End
   Begin VB.TextBox txtdepth 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   795
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
      Left            =   1635
      TabIndex        =   13
      Top             =   1515
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
      Height          =   7380
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   12015
      Begin VB.Frame Frame2 
         Caption         =   "Gas Profile"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         TabIndex        =   50
         Top             =   3840
         Width           =   6030
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   9
            Left            =   3185
            TabIndex        =   116
            Top             =   3000
            Width           =   800
         End
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   3185
            TabIndex        =   115
            Top             =   2715
            Width           =   800
         End
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   3185
            TabIndex        =   114
            Top             =   2445
            Width           =   800
         End
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   3185
            TabIndex        =   113
            Top             =   2160
            Width           =   800
         End
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   3185
            TabIndex        =   112
            Top             =   1875
            Width           =   800
         End
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   3185
            TabIndex        =   111
            Top             =   1590
            Width           =   800
         End
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   3185
            TabIndex        =   110
            Top             =   1320
            Width           =   800
         End
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   3185
            TabIndex        =   109
            Top             =   1035
            Width           =   800
         End
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   3185
            TabIndex        =   108
            Top             =   765
            Width           =   800
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   9
            Left            =   960
            TabIndex        =   107
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   960
            TabIndex        =   106
            Top             =   2715
            Width           =   735
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   960
            TabIndex        =   105
            Top             =   2445
            Width           =   735
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   960
            TabIndex        =   104
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   960
            TabIndex        =   103
            Top             =   1875
            Width           =   735
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   960
            TabIndex        =   102
            Top             =   1590
            Width           =   735
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   960
            TabIndex        =   101
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   960
            TabIndex        =   100
            Top             =   1035
            Width           =   735
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   960
            TabIndex        =   99
            Top             =   765
            Width           =   735
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   9
            Left            =   1680
            TabIndex        =   98
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   1680
            TabIndex        =   97
            Top             =   2715
            Width           =   735
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   1680
            TabIndex        =   96
            Top             =   2445
            Width           =   735
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   1680
            TabIndex        =   95
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   1680
            TabIndex        =   94
            Top             =   1875
            Width           =   735
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   1680
            TabIndex        =   93
            Top             =   1590
            Width           =   735
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   1680
            TabIndex        =   92
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   1680
            TabIndex        =   91
            Top             =   1035
            Width           =   735
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   1680
            TabIndex        =   90
            Top             =   765
            Width           =   735
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   3000
            Width           =   2020
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   2715
            Width           =   2020
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   2445
            Width           =   2020
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   2160
            Width           =   2020
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   1875
            Width           =   2020
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   1590
            Width           =   2020
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   1320
            Width           =   2020
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   1035
            Width           =   2020
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   765
            Width           =   2020
         End
         Begin VB.ComboBox Cbogasused 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   3980
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   480
            Width           =   2020
         End
         Begin VB.TextBox txtppo2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   3185
            TabIndex        =   67
            Top             =   480
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   9
            Left            =   2400
            TabIndex        =   65
            Top             =   3000
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   2400
            TabIndex        =   64
            Top             =   2715
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   2400
            TabIndex        =   63
            Top             =   2445
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   2400
            TabIndex        =   62
            Top             =   2160
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   2400
            TabIndex        =   61
            Top             =   1875
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   2400
            TabIndex        =   60
            Top             =   1590
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   2400
            TabIndex        =   59
            Top             =   1320
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   2400
            TabIndex        =   58
            Top             =   1035
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   2400
            TabIndex        =   57
            Top             =   765
            Width           =   800
         End
         Begin VB.TextBox txtmaxd 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   2400
            TabIndex        =   56
            Top             =   480
            Width           =   800
         End
         Begin VB.TextBox txthelium 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1680
            TabIndex        =   54
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtoxygen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   960
            TabIndex        =   51
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   83
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   9
            Left            =   120
            TabIndex        =   89
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   8
            Left            =   120
            TabIndex        =   88
            Top             =   2715
            Width           =   855
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   7
            Left            =   120
            TabIndex        =   87
            Top             =   2445
            Width           =   855
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   120
            TabIndex        =   86
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   120
            TabIndex        =   85
            Top             =   1875
            Width           =   855
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   120
            TabIndex        =   84
            Top             =   1590
            Width           =   855
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   82
            Top             =   1035
            Width           =   855
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   81
            Top             =   765
            Width           =   855
         End
         Begin VB.Label lblgasindex 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gas Index"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   120
            TabIndex        =   79
            Top             =   195
            Width           =   855
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gas Used"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   3975
            TabIndex        =   69
            Top             =   195
            Width           =   2030
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PPO2"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   3180
            TabIndex        =   66
            Top             =   195
            Width           =   795
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Depth"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   2400
            TabIndex        =   55
            Top             =   195
            Width           =   795
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "He"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   1680
            TabIndex        =   53
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "O2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   960
            TabIndex        =   52
            Top             =   195
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Height          =   3620
         Left            =   50
         TabIndex        =   24
         Top             =   240
         Width           =   6000
         Begin VB.CommandButton cmdadd 
            Caption         =   "Add to End"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   100
            TabIndex        =   44
            Top             =   3120
            Width           =   1355
         End
         Begin VB.CommandButton cmdmodify 
            Caption         =   "Modify"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2760
            TabIndex        =   43
            Top             =   3120
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Open circuit"
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
            Left            =   4320
            MaskColor       =   &H00FFFF80&
            TabIndex        =   42
            Top             =   1230
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Close Circuit"
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
            Left            =   2760
            MaskColor       =   &H00FFFF80&
            TabIndex        =   41
            Top             =   1230
            Width           =   1695
         End
         Begin VB.ComboBox cbogasindex 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3840
            Sorted          =   -1  'True
            TabIndex        =   40
            Text            =   "Gas Index"
            Top             =   1755
            Width           =   2055
         End
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
            Left            =   5320
            TabIndex        =   38
            Top             =   2790
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
            Left            =   5320
            TabIndex        =   37
            Top             =   2580
            Width           =   190
         End
         Begin VB.TextBox txtppo2v 
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
            Left            =   4400
            TabIndex        =   36
            Top             =   2580
            Width           =   915
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
            Left            =   1600
            TabIndex        =   34
            Top             =   2040
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
            Left            =   1600
            TabIndex        =   33
            Top             =   1830
            Width           =   190
         End
         Begin VB.CommandButton cmdclearall 
            Caption         =   "Clear All"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4900
            TabIndex        =   25
            Top             =   3120
            Width           =   1005
         End
         Begin VB.Label lblserialno 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   450
            Left            =   2160
            TabIndex        =   118
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label13 
            Caption         =   "Dive Plan Serial No :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   117
            Top             =   360
            Width           =   1935
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000000&
            X1              =   0
            X2              =   6000
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   6000
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   2475
            X2              =   2475
            Y1              =   960
            Y2              =   2400
         End
         Begin VB.Label Label19 
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
            Left            =   2640
            TabIndex        =   49
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblhelium 
            Alignment       =   2  'Center
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
            Height          =   330
            Left            =   2400
            TabIndex        =   48
            Top             =   2595
            Width           =   1215
         End
         Begin VB.Label lblo2 
            Alignment       =   2  'Center
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
            Height          =   330
            Left            =   600
            TabIndex        =   47
            Top             =   2595
            Width           =   1095
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   2640
            TabIndex        =   46
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label16 
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
            Left            =   5555
            TabIndex        =   39
            Top             =   2640
            Width           =   420
         End
         Begin VB.Label Label15 
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
            Left            =   3750
            TabIndex        =   35
            Top             =   2610
            Width           =   615
         End
         Begin VB.Label Label12 
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
            Left            =   1830
            TabIndex        =   32
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "meter"
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
            TabIndex        =   31
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label10 
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
            Left            =   1920
            TabIndex        =   30
            Top             =   2610
            Width           =   375
         End
         Begin VB.Label Label9 
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
            Left            =   120
            TabIndex        =   29
            Top             =   2610
            Width           =   495
         End
         Begin VB.Label Label8 
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
            Left            =   120
            TabIndex        =   28
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label7 
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
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   15
         Left            =   0
         TabIndex        =   23
         Top             =   2760
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   3585
         Left            =   6070
         TabIndex        =   0
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   6324
         _Version        =   393216
         Rows            =   50
         Cols            =   6
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
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   960
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   10
      RTSEnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlgchart 
      Left            =   5280
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   6240
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
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
      TabIndex        =   21
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilesave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnufilesaveas 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuspclose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnugas 
      Caption         =   "&Gas"
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
End
Attribute VB_Name = "Planprofile"
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
'dim j as integer                                               'loop as integer
Dim vmonth As Integer
Dim vday As Integer
Dim vyear As Integer
Dim clock_hour As Integer
Dim vminute As Integer
Dim vnumber_of_vmixes As Integer
Dim vnumber_of_changes As Integer
Dim vprofile_code As Integer
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
Dim ending_ambient_vpressure As Double
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
'common /block_34/ vdeco_gradient_he, vdeco_gradient_n2
'=======================================================================
'     namelist for subroutine settings (read in from ascii text file)
'=======================================================================
'=======================================================================
'     assign half-time values to buhlmann compartment arrays
'=======================================================================
Dim Plan_Depth(100) As Double
Dim Plan_Time(100) As Double
Dim Plan_o2(100) As Double
Dim Plan_he(100) As Double
Dim Plan_OpenClosed(100) As String
Dim Plan_GasID(100) As Integer
Dim Plan_PPo2(100) As Double
Dim Plan_Gas_list_o2(100) As Double
Dim Plan_Gas_list_he(100) As Double
Dim Plan_Gas_list_n2(100) As Double
Dim Plan_Gas_list_mod(100) As Double
Dim Plan_Gas_list_used(100) As Integer
Dim Plan_Gas_list_deco(100) As Integer
Dim Plan_Gas_list_numgasdeco As Integer
Dim Number_of_planpoints As Long

Dim current_vdepth As Double
Dim current_vmix_number As Double

'Nick data end here

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
Dim newseqserialno As String
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
Dim W(4) As Boolean
Dim zxt As Double
Dim CT As Integer
Dim Auto_run As Integer
Dim ppo2changed, tempdepth, temptime, tempo2, temppo2, temphe, tempgasindex, tempcircuit As String
Dim zoom As Integer
Dim pan As Integer


Private Sub cmd02up_Click()
txto2 = txto2 + 1
End Sub



Private Sub cbogasindex_Change()
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
   formstarted = False
Else
   validategasused
End If
End If
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
End Sub

Private Sub cmdadd_Click()
checkgasindex
If checkgasselected = True And Val(txtdepth) > 0 And Val(txttime) > 0 Then
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
MSFlexGrid3.Text = lblo2.Caption
MSFlexGrid3.Col = 4
MSFlexGrid3.Text = lblhelium
MSFlexGrid3.Col = 5
MSFlexGrid3.Text = txtppo2v
MSFlexGrid3.Col = 6
If Option3.Value = True Then
   MSFlexGrid3.Text = "Closed Circuit"
Else
   MSFlexGrid3.Text = "Open Circuit"
End If
MSFlexGrid3.Col = 7
MSFlexGrid3.Text = cbogasindex.Text
saveprorecord
If Val(MSFlexGrid3.Rows) > 1 Then
     cmdinsert.Enabled = True
     cmdadd.Caption = "Add To end"
  Else
     cmdinsert.Enabled = False
     cmdadd.Caption = "Create"
  End If
Else
   Title = "Error on System Validation.."
   MsgBox "Incomplete Profile Data !", 48, Title
End If
End Sub
Private Sub saveprorecord()
SQL = "SELECT * FROM dpprofile"
Set RS = DB.OpenRecordset(SQL)
   RS.AddNew
   RS!dpprofileid = tempserialno
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


Private Sub cmdclearall_Click()
ans = MsgBox("Do you really want to clear all the Dive sequence setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
   removerecord
   cleargriddata2
   setdatadefault
   MsgBox "All Dive sequence remove."
Case Else
   MsgBox "Request cancelled. "
End Select
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
   If Val(MSFlexGrid3.Rows) > 1 Then
     cmdinsert.Enabled = True
     cmdadd.Caption = "Add To End"
  Else
     cmdinsert.Enabled = False
     cmdadd.Caption = "Create"
  End If
End Sub

Private Sub CMDCLOSE_Click(Index As Integer)
tempsnfound = "False"
Screen.MousePointer = 11
SQL = "select * FROM dpmain "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("diveplanid")
   If tempdpid Like "T*" Then
      tempsnfound = "True"
   End If
   RS3.MoveNext
Wend
If tempsnfound = "True" Then
Title = "Plan List not Save.."
ans = MsgBox("You have Dive plan that was not saved, " & Chr(13) & "Press No will remove all previous unsaved plans !" & Chr(13) & "Do you want to save now ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
   savedpmain
   MsgBox "Dive plan Saved !!"
Case vbNo
   deletedpmain
End Select
End If
Screen.MousePointer = 0
Unload Me
End Sub

Private Sub cmddepthdown_Click()
checkgasselected = False
checkgasindex
     If checkgasselected = True Then
        If Val(txtdepth) > 1 And Val(txtdepth) < 2001 Then
           txtdepth = txtdepth - 1
           If Option4.Value = True Then
              txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
              txtppo2v = Format(txtppo2v, "###.00")
           End If
        Else
           MsgBox " Depth value out of range !"
        End If
     Else
        Title = "Error on System Validation.."
        MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
     End If
End Sub

Private Sub cmddepthup_Click()
  checkgasselected = False
  checkgasindex
     If checkgasselected = True Then
        If Val(txtdepth) >= 0 And Val(txtdepth) < 2000 Then
           txtdepth = txtdepth + 1
           If Option4.Value = True Then
              txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
              txtppo2v = Format(txtppo2v, "###.00")
           End If
        Else
            MsgBox " Depth value out of range !"
        End If
  Else
    Title = "Error on System Validation.."
    MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
  End If
End Sub
Private Sub checkgasindex()
   checkgasselected = False
   If cbogasindex.Text = "Gas Index" Then
      checkgasselected = False
   Else
      checkgasselected = True
   End If
End Sub
Private Sub checkgasused()
   checkgasusedselected = False
   If Cbogasused(p).Text = "0 - Not Used" Then
      SQL = "SELECT COUNT(*) FROM dpprofile "
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
Private Sub checkseqserialno()
  SQL = "SELECT * FROM dpserialno "
  Set RS = DB.OpenRecordset(SQL)
  tempseqdiveno2 = RS("seqdiveserialno")
  tempseqdiveno = Right(tempseqdiveno2, 8)
  newseqdiveno = Val(tempseqdiveno) + 1
  tempseqdiveno = Val(tempseqdiveno) + 1
  lengthsn = Len(tempseqdiveno)
  Select Case lengthsn
  Case 1
     newseqdiveno = "SM0000000" & newseqdiveno
  Case 2
     newseqdiveno = "SM000000" & newseqdiveno
  Case 3
     newseqdiveno = "SM00000" & newseqdiveno
  Case 4
     newseqdiveno = "SM0000" & newseqdiveno
  Case 5
     newseqdiveno = "SM000" & newseqdiveno
  Case 6
     newseqdiveno = "SM00" & newseqdiveno
  Case 7
     newseqdiveno = "SM0" & newseqdiveno
  Case 8
     newseqdiveno = "SM" & newseqdiveno
 End Select
 tempseqdiveno = newseqdiveno
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

Private Sub cmdinsert_Click()
 rowchanged = rowindentified - 1
 If rowindentified <> "0" Then
   totalrow = MSFlexGrid3.Rows - 1
   MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
   For i = totalrow To rowchanged Step -1
      If Val(i) = Val(rowchanged) Then
         MSFlexGrid3.Row = rowchanged + 1
         MSFlexGrid3.Col = 0
         MSFlexGrid3.Text = i + 1
         MSFlexGrid3.Col = 1
         MSFlexGrid3.Text = txtdepth
         MSFlexGrid3.Col = 2
         MSFlexGrid3.Text = txttime
         MSFlexGrid3.Col = 3
         MSFlexGrid3.Text = lblo2.Caption
         MSFlexGrid3.Col = 4
         MSFlexGrid3.Text = lblhelium.Caption
         MSFlexGrid3.Col = 5
         MSFlexGrid3.Text = txtppo2v
         MSFlexGrid3.Col = 6
         If Option3.Value = True Then
            MSFlexGrid3.Text = "Closed Circuit"
         End If
         If Option4.Value = True Then
            MSFlexGrid3.Text = "Open Circuit"
         End If
         MSFlexGrid3.Col = 7
         MSFlexGrid3.Text = cbogasindex.Text
      Else
         readprerowval ' read previous row value
         saveprerowdata
      End If
   Next i
   removerecord
   savechangerecord
Else
   Title = "Dive Profile"
   MsgBox "You must selected a record in the list to insert the sequence", 48, Title
End If
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
End Sub
Private Sub updateserialno()
   SQL = "SELECT * FROM dpserialno "
   Set RS = DB.OpenRecordset(SQL)
  tempsqserialno = RS("lastseqdserialno")
  tempsqserialno = Right(tempsqserialno, 8)
  newseqserialno = Val(tempsqserialno) + 1
  lengthsn = Len(newseqserialno)
  Select Case lengthsn
  Case 1
     newseqserialno = "SP0000000" & newseqserialno
  Case 2
     newseqserialno = "SP000000" & newseqserialno
  Case 3
     newseqserialno = "SP00000" & newseqserialno
  Case 4
     newseqserialno = "SP0000" & newseqserialno
  Case 5
     newseqserialno = "SP000" & newseqserialno
  Case 6
     newseqserialno = "SP00" & newseqserialno
  Case 7
     newseqserialno = "SP0" & newseqserialno
  Case 8
     newseqserialno = "SP" & newseqserialno
 End Select
   RS.Edit
  RS!lastseqdserialno = newseqserialno
  lblserialno.Caption = newseqserialno
  RS.Update
  End Sub
Private Sub saveseqdpmain()
  updateserialno
  updategaslist
  updateseqdpprofile
  updateseqdpmain
  MsgBox "Record Saved !!"
End Sub
Private Sub updateseqdpmain()
   SQL = "SELECT * FROM seqdpmain "
   Set RS = DB.OpenRecordset(SQL)
   RS.AddNew
   RS!diveplanid = newseqserialno
   RS.Update
   RS.Close
End Sub
Private Sub updategaslist()
   SQL = "SELECT * FROM dpmaingaslist "
   Set RS = DB.OpenRecordset(SQL)
   For i = 0 To 9
   RS.AddNew
   RS!dpmainid = newseqserialno
   RS!dpgasid = lblgasindex(i).Caption
   RS!dpgashelium = txthelium(i).Text
   tempnitrogen = 100 - Val(txthelium(i).Text) - Val(txtoxygen(i).Text)
   RS!dpgasnitrogen = tempnitrogen
   RS!dpgasmaxopdepth = Val(txtmaxd(i).Text) * 10
   RS!dpgasused = Cbogasused(i).Text
   RS.Update
 Next i
 RS.Close
 End Sub
 Private Sub updateseqdpprofile()
 If MSFlexGrid3.Rows > 1 Then
 For K = 1 To MSFlexGrid3.Rows - 1
    MSFlexGrid3.Row = K
    SQL = "SELECT * FROM seqdpprofile"
    Set RS = DB.OpenRecordset(SQL)
    RS.AddNew
    RS!dpprofileid = newseqserialno
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
  Next K
    
 RS.Close
 End If
 End Sub
Private Sub savedpmain()
SQL = "select * FROM dpmain "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("diveplanid")
   If tempdpid Like "T*" Then
      tempdpid2 = Right(tempdpid, 9)
      tempdpid2 = "D" & tempdpid2
      RS3.Edit
      RS3!diveplanid = tempdpid2
      RS3.Update
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & "where "
      SQL = SQL & " dpmainid = '" & tempdpid & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      While RS2.EOF = False
         RS2.Edit
         RS2!dpmainid = tempdpid2
         RS2.Update
         RS2.MoveNext
      Wend
      SQL = "SELECT * FROM dpprofile "
      SQL = SQL & "where "
      SQL = SQL & " dpprofileid = '" & tempdpid & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      While RS2.EOF = False
         RS2.Edit
         RS2!dpprofileid = tempdpid2
         RS2.Update
         RS2.MoveNext
      Wend
      SQL = "select * FROM dpserialno "
      Set RS3 = DB.OpenRecordset(SQL)
      RS3.Edit
      RS3!dplanserialno = tempdpid2
      RS3.Update
      End If
   RS3.MoveNext
Wend
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
End Sub
Private Sub cmdmodify_Click()
DataChanged = "False"
ppo2changed = "False"
checkforchanges
If ppo2changed = "True" Then
   MSFlexGrid3.Row = rowindentified
   MSFlexGrid3.Col = 5
   tempppo2 = MSFlexGrid3.Text
End If
If DataChanged = "True" Then
   MSFlexGrid3.Row = rowindentified
   MSFlexGrid3.Col = 1
   MSFlexGrid3.Text = txtdepth
   MSFlexGrid3.Col = 2
   MSFlexGrid3.Text = txttime
   MSFlexGrid3.Col = 3
   MSFlexGrid3.Text = lblo2.Caption
   MSFlexGrid3.Col = 4
   MSFlexGrid3.Text = lblhelium.Caption
   MSFlexGrid3.Col = 5
   MSFlexGrid3.Text = txtppo2v
   MSFlexGrid3.Col = 6
   If Option3.Value = True Then
      MSFlexGrid3.Text = "Closed Circuit"
   End If
   If Option4.Value = True Then
      MSFlexGrid3.Text = "Open Circuit"
   End If
   MSFlexGrid3.Col = 7
   MSFlexGrid3.Text = cbogasindex.Text
   SQL = "SELECT * FROM dpprofile"
   SQL = SQL & " where dpprofileid = '" & tempserialno & "' and dpnumseq = '" & rowindentified & "' "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   MSFlexGrid3.Col = 1
   RS("depth") = MSFlexGrid3.Text
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
   RS("po2") = txtppo2v.Text
   RS.Update
   End If
   
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

Private Sub cmdplan_Click()
Unload Me
planmain.Show
End Sub

Private Sub CMDPPO2PLUS_Click()
If Val(txtppo2v) >= 0.15 And Val(txtppo2v) < 2.01 Then
   txtppo2v = Val(txtppo2v) + 0.01
Else
   MsgBox " PO2 value out of range !"
End If
End Sub

Private Sub cmdremove_Click()
numrow = MSFlexGrid3.Rows
Totalcount = numrow - 1
For K = 0 To Totalcount
    MSFlexGrid3.Row = K
    If MSFlexGrid3.CellBackColor = vbBlue Then
       MSFlexGrid3.Col = 0
       tempseq = MSFlexGrid3.Text
    End If
Next K
SQL = "SELECT * FROM dpprofile "
SQL = SQL & "where dpnumseq = '" & tempseq & "'  and dpprofileid = '" & tempserialno & "' "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   RS3.Delete
   RS3.MoveNext
Wend
griddataexist
reloadgriddata
removerecord
savechangerecord
cmdinsert.Enabled = False
cmdmodify.Enabled = False
cmdremove.Enabled = False

End Sub
Private Sub savechangerecord()
  For K = 1 To MSFlexGrid3.Rows - 1
    MSFlexGrid3.Row = K
    saveprorecord
  Next K
End Sub
Private Sub removerecord()
  SQL = "SELECT * FROM dpprofile"
  SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
  Set RS = DB.OpenRecordset(SQL)
  While RS.EOF = False
   RS.Delete
   RS.MoveNext
  Wend
End Sub
Private Sub removerecordgasindex()
  SQL = "SELECT * FROM dpprofile "
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
For p = 0 To 7
   MSFlexGrid3.Col = p
   MSFlexGrid3.Row = K
   MSFlexGrid3.Text = ""
Next p
Next K
End Sub
Private Sub cleargriddata2()
 i = MSFlexGrid3.Rows
 i = i - 1
For K = i To 1 Step -1
For p = 0 To 7
   MSFlexGrid3.Col = p
   MSFlexGrid3.Row = K
   MSFlexGrid3.Text = ""
Next p
MSFlexGrid3.Rows = MSFlexGrid3.Rows - 1
Next K
End Sub
Private Sub griddataexist()
profilerecordexist = False
SQL = "SELECT COUNT(*) FROM dpprofile "
  SQL = SQL & " WHERE "
  SQL = SQL & " dpprofileid ='" & Trim(tempserialno) & "'"
  Set RS3 = DB.OpenRecordset(SQL)
  If RS3.Fields(0) <> 0 Then
     profilerecordexist = True
  End If
End Sub
Private Sub reloadgriddata()
cleargriddata
If profilerecordexist = True Then
SQL = "SELECT * FROM dpprofile"
SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
SQL = SQL & " order by dpnumseq "
Set RS = DB.OpenRecordset(SQL)
RS.MoveFirst
MSFlexGrid3.Rows = 1
While RS.EOF = False
     MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
     K = MSFlexGrid3.Rows
     MSFlexGrid3.Row = K - 1
     MSFlexGrid3.Col = 0
     MSFlexGrid3.Text = K - 1
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
     RS.MoveNext
Wend
Else
   cmdinsert.Enabled = False
   cmdmodify.Enabled = False
   cmdremove.Enabled = False
   cmdadd.Caption = "Create"
End If
  
End Sub
Private Sub reloadgriddata2()
cleargriddata
If profilerecordexist = True Then
SQL = "SELECT * FROM dpprofile"
SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
SQL = SQL & " order by dpnumseq "
Set RS = DB.OpenRecordset(SQL)
RS.MoveFirst
MSFlexGrid3.Rows = 1
While RS.EOF = False
     MSFlexGrid3.Rows = MSFlexGrid3.Rows + 1
     K = MSFlexGrid3.Rows
     MSFlexGrid3.Row = K - 1
     MSFlexGrid3.Col = 0
     MSFlexGrid3.Text = K - 1
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
     RS.MoveNext
Wend
Else
   cmdinsert.Enabled = False
   cmdmodify.Enabled = False
   cmdremove.Enabled = False
   cmdadd.Caption = "Create"
End If
  
End Sub
Private Sub cmdsave_Click()
If tempchoice = "NP" Then
   
   SQL = "select * FROM dpmain "
   SQL = SQL & " where diveplanid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!diveplanid = newserialno
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpmaingaslist "
   SQL = SQL & " where dpmainid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!dpmainid = newserialno
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpprofile "
   SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!dpprofileid = newserialno
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpserialno "
   Set RS3 = DB.OpenRecordset(SQL)
   RS3.Edit
   RS3!dplanserialno = newserialno
   RS3.Update
End If
MsgBox "Record Saved!"
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
tempsnfound = "False"
SQL = "select * FROM dpmain "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("diveplanid")
   If tempdpid Like "T*" Then
      tempsnfound = "True"
   End If
   RS3.MoveNext
Wend
If tempsnfound = "True" Then
   Title = "Plan List not Save.."
   ans = MsgBox("Current Dive plan was not save on last setting, " & Chr(13) & "Do you want to save now ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
   Select Case ans
   Case vbYes
      savedpmain
      saveseqdpmain
      checkseqserialno
      frmseqdive.Show
   Case vbNo
      Title = "Plan cannot Save.."
      MsgBox "Dive Plan must be save first inorder to use it for sequential dive " & Chr(13) & "Press Save As again to save the dive !", 48, Title
   End Select
Else
   saveseqdpmain
   checkseqserialno
   frmseqdive.Show
End If
End Sub

Private Sub cmdtimedown_Click()
checkgasselected = False
checkgasindex
   If checkgasselected = True Then
      If Val(txttime) > 1 And Val(xttime) < 9999 Then
         txttime = txttime - 1
      Else
         MsgBox " Duration value out of range !"
      End If
   Else
      Title = "Error on System Validation.."
      MsgBox "Please select at least one gas. " & Chr(13) & "You can select from the Gas Index List or Click on the Gas Index from the Gas Profile !", 48, Title
   End If
End Sub

Private Sub cmdtimeup_Click()
checkgasselected = False
checkgasindex
 If checkgasselected = True Then
     If Val(txttime) >= 0 And Val(txttime) < 9999 Then
        txttime = txttime + 1
     Else
        MsgBox " Duration value out of range !"
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

'Private Sub Command2_Click(Index As Integer) 'nick
'  Select Case Index
'    Case 0
'    If Totalcount > 100 Then
'      If zoom = 10 Then
'        zoom = 1
'        pan = 1
'        Command2(1).Enabled = False
'        Command2(2).Enabled = False
'      Else
'        zoom = 10
'        pan = 1
'        Command2(1).Enabled = True
'        Command2(2).Enabled = True
'      End If
'    End If
'    Case 1
'      pan = pan - 1
'      If pan < 1 Then pan = 1
'    Case 2
'      pan = pan + 1
'      If pan > 10 Then pan = 10
'  End Select
'  If zoom = 10 Then
'    If pan = 1 Then Command2(1).Enabled = False Else Command2(1).Enabled = True
'    If pan = 10 Then Command2(2).Enabled = False Else Command2(2).Enabled = True
'  End If
'  Command2(0).Caption = "Zoom " + CStr(zoom)
'  plotgraph
'End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command6_Click()

End Sub
Private Sub deletedpmain()
SQL = "select * FROM dpmain "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("diveplanid")
   If tempdpid Like "T*" Then
      RS3.Delete
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & "where "
      SQL = SQL & " dpmainid = '" & tempdpid & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      While RS2.EOF = False
         RS2.Delete
         RS2.MoveNext
      Wend
      SQL = "SELECT * FROM dpprofile "
      SQL = SQL & "where "
      SQL = SQL & " dpprofileid = '" & tempdpid & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      While RS2.EOF = False
         RS2.Delete
         RS2.MoveNext
      Wend
   End If
   RS3.MoveNext
Wend
End Sub
'Private Sub Command5_Click()
'ans = MsgBox("Do you really want to load the factory default setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
'Select Case ans
'Case vbYes
'    SQL = "SELECT * FROM dpfacgasdefault"
'    Set RS = DB.OpenRecordset(SQL)
'    RS.MoveFirst
'    i = 0
'    While RS.EOF = False
'       If Val(i) < 10 Then
'          gasindex(i).Caption = RS("gasid")
'          tempnitrogen = RS("gasnitrogen")
'          txthelium(i) = RS("gashelium")
'          txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
'          txtmaxd(i) = RS("gasmaxopdepth")
'          txtmaxd(i) = txtmaxd(i) / 10
'          temptxthe1 = txthe1
'          temptxtmaxd1 = txtmaxd(i)
'          Cbogasused(i) = RS("gasused")
'          txtppo2(i).Enabled = True
'          txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
'          txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
'          txtppo2(i).Enabled = False
'      End If
'  i = i + 1
'  RS.MoveNext
'  Wend
'  MsgBox "All value reset to factory default."
'Case Else
'   MsgBox "Request cancelled. "
'End Select
'End Sub

Private Sub Command4_Click()
If Val(txtppo2v) > 0.15 And Val(txtppo2v) < 2.01 Then
   txtppo2v = Val(txtppo2v) - 0.01
Else
   MsgBox " PO2 value out of range !"
End If
End Sub

Private Sub Form_Load()
formstarted = True
Top = 30
Me.Left = (Screen.Width - Me.Width) / 2
fgactivate = "0"
colselected = "false"
cols0activated = "false"
initialgrid
txtdepth = 0
txttime = 0
For i = 0 To 9
  Cbogasused(i).AddItem "0 - Not Used"
  Cbogasused(i).AddItem "1 - Open Circuit"
  Cbogasused(i).AddItem "2 - Closed Circuit"
  Cbogasused(i).AddItem "3 - Open & Closed"
  Cbogasused(i).AddItem "4 - Deco Open Circuit"
  Cbogasused(i).AddItem "5 - Deco Closed Circuit"
Next i
If tempchoice = "NP" Then
  mnuloaddesetting.Enabled = True
  mnugasloadefault.Enabled = True
  lblserialno.Caption = "   " & newserialno
  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & " where dpmainid = '" & tempserialno & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  For i = 0 To 9
     lblgasindex(i).Caption = RS("dpgasid")
     tempnitrogen = RS("dpgasnitrogen")
     txthelium(i) = RS("dpgashelium")
     txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
     txtmaxd(i) = RS("dpgasmaxopdepth")
     txtmaxd(i) = txtmaxd(i) / 10
     temptxthe1 = txthelium(i)
     temptxtmaxd1 = txtmaxd(i)
     Cbogasused(i) = RS("dpgasused")
     txtppo2(i).Enabled = True
     txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
     txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     txtppo2(i).Enabled = False
     If Cbogasused(i).Text <> "0 - Not Used" Then
        cbogasindex.AddItem lblgasindex(i).Caption
     End If
     RS.MoveNext
  Next i
  cmdsaveas.Enabled = False
  
 End If
'Dim p As Integer
If tempchoice = "PP" Or tempchoice = "GP" Then
  mnuloaddesetting.Enabled = False
  mnugasloadefault.Enabled = False
  lblserialno.Caption = "   " & oldserialno
  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & " where dpmainid = '" & tempserialno & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  For i = 0 To 9
      lblgasindex(i).Caption = RS("dpgasid")
     
     tempnitrogen = RS("dpgasnitrogen")
     txthelium(i) = RS("dpgashelium")
     txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
     
     txtmaxd(i) = RS("dpgasmaxopdepth")
     
     txtmaxd(i) = txtmaxd(i) / 10
     
     temptxthe1 = txthelium(i)
     
     temptxtmaxd1 = txtmaxd(i)
     Cbogasused(i) = RS("dpgasused")
     
     txtppo2(i).Enabled = True
     txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
     txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     txtppo2(i).Enabled = False
     If Cbogasused(i).Text <> "0 - Not Used" Then
        cbogasindex.AddItem lblgasindex(i).Caption
     End If
     
     RS.MoveNext
  Next i
  SQL = "SELECT COUNT(*) FROM dpprofile "
  SQL = SQL & " WHERE "
  SQL = SQL & " dpprofileid ='" & Trim(tempserialno) & "'"
  Set RS3 = DB.OpenRecordset(SQL)
  If RS3.Fields(0) <> 0 Then
     loaddpprofiledata
  End If
   
  rowindentified = 0
  
 End If
 If Val(MSFlexGrid3.Rows) > 1 Then
    cmdadd.Caption = "Add To End"
 Else
    cmdadd.Caption = "Create"
 End If
 'Cbogasused(1).Style = 2
 '
 For i = 0 To 9
    ' Cbogasused(i).Style = 2
  Next i
  cmdinsert.Enabled = False
    cmdremove.Enabled = False
    cmdmodify.Enabled = False
    formstarted = False
End Sub
Private Sub loaddpprofiledata()
  SQL = "SELECT * FROM dpprofile"
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
     RS.MoveNext
   Wend
End Sub
Private Sub initialgrid()
MSFlexGrid3.Cols = 8
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
MSFlexGrid3.Col = 5
MSFlexGrid3.Text = "Po2"
MSFlexGrid3.Col = 6
MSFlexGrid3.Text = "Circuit"
MSFlexGrid3.Col = 7
MSFlexGrid3.Text = "Gas Index"
MSFlexGrid3.ColWidth(0) = 465
MSFlexGrid3.ColWidth(1) = 630
MSFlexGrid3.ColWidth(2) = 630
MSFlexGrid3.ColWidth(3) = 580
MSFlexGrid3.ColWidth(4) = 580
MSFlexGrid3.ColWidth(5) = 630
MSFlexGrid3.ColWidth(6) = 1200
MSFlexGrid3.ColWidth(7) = 1020
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
End Sub

Private Sub lblgasindex_Click(Index As Integer)
G = txtoxygen(Index).Index
tempgasused = Cbogasused(G).Text
If tempgasused <> "0 - Not Used" Then
lblo2.Caption = txtoxygen(G).Text
lblhelium.Caption = txthelium(G).Text
cbogasindex.Text = lblgasindex(G).Caption
tempgasused = Cbogasused(G).Text
txtppo2v = txtppo2(G).Text
If InStr(tempgasused, "Closed C") Then
   Option3.Enabled = True
   Option3.Value = True
   Option3.Enabled = False
   txtppo2v.Enabled = True
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
   Else
      Option3.Value = False
      Option4.Value = True
      Option3.Enabled = True
      Option4.Enabled = True
      txtppo2v.Enabled = False
   End If
End If
End If

End Sub



Private Sub mnufilesave_Click()
If tempchoice = "NP" Then
   
   SQL = "select * FROM dpmain "
   SQL = SQL & " where diveplanid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!diveplanid = newserialno
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpmaingaslist "
   SQL = SQL & " where dpmainid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!dpmainid = newserialno
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpprofile "
   SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
   Set RS3 = DB.OpenRecordset(SQL)
   While RS3.EOF = False
      RS3.Edit
      RS3!dpprofileid = newserialno
      RS3.Update
      RS3.MoveNext
   Wend
   SQL = "select * FROM dpserialno "
   Set RS3 = DB.OpenRecordset(SQL)
   RS3.Edit
   RS3!dplanserialno = newserialno
   RS3.Update
End If
MsgBox "Record Saved!"
End Sub

Private Sub mnugasloadefault_Click()
ans = MsgBox("Do you really want to load the factory default setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
    SQL = "SELECT * FROM dpfacgasdefault "
    SQL = SQL & " order by gasid"
    Set RS = DB.OpenRecordset(SQL)
    RS.MoveFirst
    T = 0
    While RS.EOF = False
       If Val(T) < 10 Then
       i = T
          lblgasindex(i).Caption = RS("gasid")
          tempnitrogen = RS("gasnitrogen")
          txthelium(i) = RS("gashelium")
          txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
          txtmaxd(i) = RS("gasmaxopdepth")
          txtmaxd(i) = txtmaxd(i) / 10
          temptxthe1 = txthe1
          temptxtmaxd1 = txtmaxd(i)
          txtppo2(i).Enabled = True
          txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
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
  MsgBox "All value reset to factory default."
Case Else
   MsgBox "Request cancelled. "
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
          tempnitrogen = 100 - Val(txthelium(i).Text) - Val(txtoxygen(i).Text)
          RS!gasnitrogen = tempnitrogen
          RS!gasmaxopdepth = Val(txtmaxd(i).Text) * 10
          RS!gasused = Cbogasused(i).Text
          RS.Update
       Next i
Case Else
   MsgBox "Request cancelled. "
End Select
End Sub

Private Sub mnuloaddesetting_Click()
ans = MsgBox("Do you really want to load the factory default setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
    SQL = "SELECT * FROM dpgasdefault "
    SQL = SQL & " order by gasid "
    RS.MoveFirst
    i = 0
    While RS.EOF = False
       If Val(T) < 10 Then
       i = T
          lblgasindex(i).Caption = RS("gasid")
          tempnitrogen = RS("gasnitrogen")
          txthelium(i) = RS("gashelium")
          txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
          txtmaxd(i) = RS("gasmaxopdepth")
          txtmaxd(i) = txtmaxd(i) / 10
          temptxthe1 = txthe1
          temptxtmaxd1 = txtmaxd(i)
          txtppo2(i).Enabled = True
          txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
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
  MsgBox "All value reset to User default default."
Case Else
   MsgBox "Request cancelled. "
End Select
End Sub

Private Sub mnuspclose_Click()
tempsnfound = "False"
Screen.MousePointer = 11
SQL = "select * FROM dpmain "
Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("diveplanid")
   If tempdpid Like "T*" Then
      tempsnfound = "True"
   End If
   RS3.MoveNext
Wend
If tempsnfound = "True" Then
Title = "Plan List not Save.."
ans = MsgBox("You have Dive plan that was not saved on last setting, " & Chr(13) & "Press No will remove all previous unsaved plans !" & Chr(13) & "Do you want to save now ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
   savedpmain
   MsgBox "Dive plan Saved !!"
Case vbNo
   deletedpmain
End Select
End If
Screen.MousePointer = 0
Unload Me
End Sub

Private Sub MSFlexGrid3_Click()
numrow = MSFlexGrid3.Rows
Totalcount = numrow - 1
rowindentified = MSFlexGrid3.Row
For K = 0 To Totalcount
  For p = 0 To 0
    MSFlexGrid3.Row = K
    MSFlexGrid3.Col = p
    If MSFlexGrid3.CellBackColor = vbBlue Then
      For H = 0 To 7
        MSFlexGrid3.Row = K
        MSFlexGrid3.Col = H
        MSFlexGrid3.CellForeColor = MSFlexGrid3.ForeColor
        MSFlexGrid3.CellBackColor = MSFlexGrid3.BackColor
      Next H
    End If
  Next p
Next K
For q = 0 To 7
   MSFlexGrid3.Row = rowindentified
   MSFlexGrid3.Col = q
   MSFlexGrid3.CellForeColor = vbWhite
   MSFlexGrid3.CellBackColor = vbBlue
Next q

MSFlexGrid3.Col = 0
  tempseq = MSFlexGrid3.Text
  SQL = "SELECT * FROM dpprofile"
  SQL = SQL & " where dpprofileid = '" & tempserialno & "' and dpnumseq = '" & tempseq & "' "
  Set RS = DB.OpenRecordset(SQL)
  txtppo2v.Text = RS("po2")
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
    txtppo2v.Text = MSFlexGrid3.Text
  Case 6
            If MSFlexGrid3.Text = "Closed Circuit" Then
              Option3.Value = True
              Option4.Value = False
              txtppo2v.Enabled = True
           Else
              Option4.Value = True
              Option3.Value = False
              txtppo2v.Enabled = False
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
                Option4.Enabled = True
                Option3.Enabled = True
             Case "1 - Open Circuit"
                Option4.Enabled = True
                Option3.Enabled = False
                txtppo2v.Enabled = False
             Case "2 - Closed Circuit"
                Option4.Enabled = False
                Option3.Enabled = True
                txtppo2v.Enabled = True
             Case "4 - Deco Open Circuit"
                Option4.Enabled = True
                Option3.Enabled = False
                txtppo2v.Enabled = False
             Case "5 - Deco Closed Circuit"
             
                Option4.Enabled = False
                Option3.Enabled = True
                txtppo2v.Enabled = True
                End Select
        Case 7
     cbogasindex.Text = MSFlexGrid3.Text
  End Select
Next p
  cmdinsert.Enabled = True
  cmdremove.Enabled = True
  cmdmodify.Enabled = True
  
  
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
   txtppo2v.Enabled = True
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
  If Val(txtdepth) < 0 Or Val(txtdepth) > 2000 Then
     MsgBox " Depth value out of range !"
  Else
     If Option4.Value = True And Val(txtdepth) <> 0 Then
        txtppo2v = CStr(((CDbl(txtdepth) / 100#) + 1) * (CDbl(lblo2.Caption) / 100))
        txtppo2v = Format(txtppo2v, "###.00")
     End If
  End If
End If
End Sub

Private Sub txtdepth_LostFocus()
If Trim(cbogasindex) <> "Gas Index" Then
If Val(txtdepth) < 0 Or Val(txtdepth) > 2000 Then
     MsgBox " Depth value out of range !"
     txtdepth.SetFocus
     SendKeys "{HOME}+{END}"
Else
   If Option4.Value = True And Val(txtdepth) <> 0 Then
      txtppo2v = CStr(((CDbl(txtdepth) / 100#) + 1) * (CDbl(lblo2.Caption) / 100))
      txtppo2v = Format(txtppo2v, "###.00")
   End If
End If
End If
End Sub

Private Sub txthelium_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
  p = txthelium(Index).Index
  validatehelium
End If
End Sub


Private Sub txthelium_LostFocus(Index As Integer)
  p = txthelium(Index).Index
  validatehelium
End Sub

Private Sub txtmaxd_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
  p = txtmaxd(Index).Index
  validatemaxdepth
End If
End Sub
Private Sub validateoxygen()
   If Val(txtoxygen(p).Text) >= 0 And Val(txtoxygen(p).Text) <= 100 And ((Val(txtoxygen(p).Text) + Val(txthelium(p).Text)) < 101) Then
     txtppo2(p).Enabled = True
     txtppo2(p).Text = (Val(txtoxygen(p).Text) / 100) * ((Val(txtmaxd(p).Text) / 10) + 1)
     txtppo2(p).Text = Format(txtppo2(p).Text, "###.00")
     txtppo2(p).Enabled = False
     txtppo2v.Text = txtppo2(p).Text
     lblhelium.Caption = txthelium(p).Text
     lblo2.Caption = txtoxygen(p).Text
     cbogasindex.Text = lblgasindex(p).Caption
     tempgasindex = lblgasindex(p).Caption
     tempgasused = Cbogasused(p).Text
     If InStr(tempgasused, "Closed C") Then
        Option3.Enabled = True
        Option3.Value = True
        Option3.Enabled = False
        txtppo2v.Enabled = True
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
        Else
           Option3.Value = False
           Option4.Value = True
           Option3.Enabled = True
           Option4.Enabled = True
           txtppo2v.Enabled = False
        End If
     End If
     If Option4.Value = True Then
        txtppo2v = CStr(((CDbl(txtdepth) / 10#) + 1) * (CDbl(lblo2.Caption) / 100))
        txtppo2v = Format(txtppo2v, "###.00")
     End If
     SQL = "SELECT * FROM dpmaingaslist "
     SQL = SQL & " where dpmainid = '" & tempserialno & "' "
     SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     tempnitrogen = 100 - Val(txtoxygen(p).Text) - Val(txthelium(p).Text)
     RS2.Edit
     RS2!dpgasnitrogen = tempnitrogen
     RS2.Update
     SQL = "SELECT * FROM dpprofile "
     SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
     SQL = SQL & " and gasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     tempnitrogen = 100 - Val(txtoxygen(p).Text) - Val(txthelium(p).Text)
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
      MsgBox "(oxygen " & p & " value can not be less then 0 or more then 100) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      tempoxy = RS2("dpgasnitrogen")
      tempoxygen = 100 - Val(txthelium(p).Text) - Val(tempoxy)
      txtoxygen(p).Text = tempoxygen
      txtoxygen(p).SetFocus
      SendKeys "{HOME}+{END}"
   End If
End Sub
Private Sub validategasused()
If Val(txtoxygen(p).Text) >= 0 And Val(txtoxygen(p).Text) <= 100 And ((Val(txtoxygen(p).Text) + Val(txthelium(p).Text)) < 101) Then
     txtppo2(p).Enabled = True
     txtppo2(p).Text = (Val(txtoxygen(p).Text) / 100) * ((Val(txtmaxd(p).Text) / 10) + 1)
     txtppo2(p).Text = Format(txtppo2(p).Text, "###.00")
     txtppo2(p).Enabled = False
     txtppo2v.Text = txtppo2(p).Text
     lblhelium.Caption = txthelium(p).Text
     lblo2.Caption = txtoxygen(p).Text
     cbogasindex.Text = lblgasindex(p).Caption
     tempgasindex = lblgasindex(p).Caption
     tempgasused = Cbogasused(p).Text
     If InStr(tempgasused, "Closed C") Then
        Option3.Enabled = True
        Option3.Value = True
        Option3.Enabled = False
        txtppo2v.Enabled = True
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
        Else
           Option3.Value = False
           Option4.Value = True
           Option3.Enabled = True
           Option4.Enabled = True
           txtppo2v.Enabled = False
        End If
     End If
     If Option4.Value = True Then
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
     SQL = "SELECT * FROM dpprofile "
     SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
     SQL = SQL & " and gasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     tempnitrogen = 100 - Val(txtoxygen(p).Text) - Val(txthelium(p).Text)
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
      MsgBox "(Maximum depth " & p & " value can not be less then 0 or more then 1000) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      tempoxy = RS2("dpgasnitrogen")
      tempoxygen = 100 - Val(txthelium(p).Text) - Val(tempoxy)
      txtoxygen(p).Text = tempoxygen
      txtoxygen(p).SetFocus
      SendKeys "{HOME}+{END}"
   End If
End Sub
Private Sub validatemaxdepth()
   If Val(txtmaxd(p).Text) >= 0 And Val(txtmaxd(p).Text) < 1000 And ((Val(txtoxygen(p).Text) + Val(txthelium(p).Text)) < 101) Then
     txtppo2(p).Enabled = True
     txtppo2(p).Text = (Val(txtoxygen(p).Text) / 100) * ((Val(txtmaxd(p).Text) / 10) + 1)
     txtppo2(p).Text = Format(txtppo2(p).Text, "###.00")
     txtppo2(p).Enabled = False
       txtppo2v.Text = txtppo2(p).Text
     lblo2.Caption = txtoxygen(p).Text
     lblhelium.Caption = txthelium(p).Text
     cbogasindex.Text = lblgasindex(p).Caption
     tempgasindex = lblgasindex(p).Caption
     tempgasused = Cbogasused(p).Text
     If InStr(tempgasused, "Closed C") Then
        Option3.Enabled = True
        Option3.Value = True
        Option3.Enabled = False
        txtppo2v.Enabled = True
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
        Else
           Option3.Value = False
           Option4.Value = True
           Option3.Enabled = True
           Option4.Enabled = True
           txtppo2v.Enabled = False
        End If
     End If
     
     SQL = "SELECT * FROM dpmaingaslist "
     SQL = SQL & " where dpmainid = '" & tempserialno & "' "
     SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     RS2.Edit
     RS2!dpgasmaxopdepth = Val(txtmaxd(p).Text) * 10
     RS2.Update
     SQL = "SELECT * FROM dpprofile "
     SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
     SQL = SQL & " and gasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     tempnitrogen = 100 - Val(txtoxygen(p).Text) - Val(txthelium(p).Text)
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
      MsgBox "(Max.Depth " & p & " value can not be less then 0 or more then 1000) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      tempmaxdepth = RS2("dpgasmaxopdepth")
      txtmaxd(p).Text = tempmaxdepth
      txtmaxd(p).SetFocus
      SendKeys "{HOME}+{END}"
   End If
End Sub
Private Sub validatehelium()
  If Val(txthelium(p).Text) >= 0 And Val(txthelium(p).Text) <= 100 And ((Val(txtoxygen(p).Text) + Val(txthelium(p).Text)) < 101) Then
     txtppo2(p).Enabled = True
     txtppo2(p).Text = (Val(txtoxygen(p).Text) / 100) * ((Val(txtmaxd(p).Text) / 10) + 1)
     txtppo2(p).Text = Format(txtppo2(p).Text, "###.00")
     txtppo2(p).Enabled = False
     tempgasindex = lblgasindex(p).Caption
     lblo2.Caption = txtoxygen(p).Text
     lblhelium.Caption = txthelium(p).Text
     SQL = "SELECT * FROM dpmaingaslist "
     SQL = SQL & " where dpmainid = '" & tempserialno & "' "
     SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
     Set RS2 = DB.OpenRecordset(SQL)
     tempnitrogen = 100 - Val(txtoxygen(p).Text) - Val(txthelium(p).Text)
     RS2.Edit
     RS2!dpgasnitrogen = tempnitrogen
     RS2!dpgashelium = txthelium(p).Text
     RS2.Update
      SQL = "SELECT * FROM dpprofile "
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
      MsgBox "(Helium " & p & " value can not be less then 0 or more then 100) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
      tempgasindex = lblgasindex(p).Caption
      SQL = "SELECT * FROM dpmaingaslist "
      SQL = SQL & " where dpmainid = '" & tempserialno & "' "
      SQL = SQL & " and dpgasid = '" & tempgasindex & "' "
      Set RS2 = DB.OpenRecordset(SQL)
      temphelium = RS2("dpgashelium")
       txthelium(p).Text = temphelium
       txthelium(p).SetFocus
       SendKeys "{HOME}+{END}"
   End If
End Sub
Private Sub txtmaxd_LostFocus(Index As Integer)
 p = txtmaxd(Index).Index
  validatemaxdepth
End Sub

Private Sub txtoxygen_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
  p = txtoxygen(Index).Index
  validateoxygen
End If
End Sub

Private Sub txtoxygen_LostFocus(Index As Integer)
  p = txtoxygen(Index).Index
  validateoxygen
End Sub

Private Sub txtppo2v_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Val(txtppo2v) < 0.15 And Val(txtppo2v) > 2# Then
     MsgBox " PO2 value out of range !"
  End If
End If
End Sub

Private Sub txttime_GotFocus()
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
  If Val(txttime) < 0 Or Val(txttime) > 9999 Then
     MsgBox " Duration value out of range !"
  End If
End If
End Sub

Private Sub txttime_LostFocus()
If Trim(cbogasindex) <> "Gas Index" Then
   If Val(txttime) < 0 Or Val(txttime) > 9999 Then
       MsgBox "Duration value out of range !"
       txttime.SetFocus
       SendKeys "{HOME}+{END}"
   End If
End If
End Sub
'Nick code start here

Private Sub Command1_Click()
Dim Planpoint As Integer
If vimportdb_data > 0 Then Exit Sub

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
'cleardecogrid
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
valtitude_dive_valgorithm = "off"
minimum_vdeco_vstop_vtime = 0.1
critical_radius_vn2_microns = 0.6
critical_radius_vhe_microns = 0.5
critical_volume_valgorithm = "on"
crit_volume_parameter_lambda = 7500#
gradient_onset_of_imperm_atm = 8.2
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
'    MsgBox "subroutine terminated"
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
'    MsgBox "subroutine terminated"
'End If
If ((critical_radius_vn2_microns < 0.2) Or (critical_radius_vn2_microns > 1.35)) Then
    MsgBox "subroutine terminated"
End If
If ((critical_radius_vhe_microns < 0.2) Or (critical_radius_vhe_microns > 1.35)) Then
    MsgBox "subroutine terminated"
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
'    MsgBox "subroutine terminated"
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
If (valtitude_dive_valgorithm_off) Then
    valtitude_of_dive = 0#
    Call calc_barometric_vpressure(valtitude_of_dive)            'su
    t1print CStr("Alt")
    t1print CStr(valtitude_of_dive)
    t1print CStr("Atmmospheric")
    t1print CStr(barometric_vpressure)
    t1print (vbCrLf)
    For i = 1 To 16
        adjusted_critical_radius_n2(i) = initial_critical_radius_n2(i)
        adjusted_critical_radius_he(i) = initial_critical_radius_he(i)
        vhelium_vpressure(i) = 0#
        vnitrogen_vpressure(i) = (barometric_vpressure - water_vapor_vpressure) * 0.79
    Next i
Else
    Call vpm_valtitude_dive_valgorithm                           'su
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
Do While (True)                     'loop will run continuous
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
            MsgBox "subroutine terminated"
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
            rate = 20
          Else
            starting_vdepth = Plan_Depth(Planpoint - 1)
            ending_vdepth = Plan_Depth(Planpoint)
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
            t1print8 CStr(vmix_vnumber)
            t1print8 CStr(vdepth)
            t1print8 CStr(starting_vdepth)
            t1print8 CStr(ending_vdepth)
            t1print8 CStr(rate)
            t1print (vbCrLf)
'            t1print (CStr(vsegment_vnumber) + CStr(vsegment_vtime) + CStr(run_vtime) + CStr(vmix_vnumber) + CStr(starting_vdepth) + CStr(ending_vdepth) + CStr(rate) + vbCrLf)
        ElseIf (vprofile_code = 2) Then
            'vdepth = 80#
            'run_vtime_end_of_vsegment = 30#
            'vmix_vnumber = 1
            If (run_vtime_end_of_vsegment - run_vtime) <= 0 Then
              MsgBox "Segment time too short Ascent/Descent at Segment: " + CStr(CInt(vsegment_vnumber / 2) + 1)
              vhighlite_line (CInt(vsegment_vnumber / 2) + 1)
              Exit Sub
            End If
            Call gas_loadings_constant_vdepth(vdepth, run_vtime_end_of_vsegment)
            t1print8 CStr(vsegment_vnumber)
            t1print8 CStr(vsegment_vtime)
            t1print8 CStr(run_vtime)
            t1print8 CStr(vmix_vnumber)
            t1print8 CStr(vdepth)
'            t1print8 cstr(rate)
            t1print (vbCrLf)
            Planpoint = Planpoint + 1
        ElseIf (vprofile_code = 99) Then
                Exit Do
        Else
              MsgBox "subroutine terminated"
        End If
        current_vdepth = vdepth
        current_vmix_number = vmix_vnumber
        
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
          vmix_change(i) = current_vmix_number
          rate_change(i) = -10#
          vstep_size_change(i) = 3
        Else
          vdepth_change(i) = Plan_Gas_list_mod(Plan_Gas_list_deco(i)) '33#
          vmix_change(i) = Plan_Gas_list_deco(i)
          rate_change(i) = -10#
          vstep_size_change(i) = 3
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
               MsgBox "subroutine terminated"
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
            MsgBox "subroutine terminated"
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
            t1print ("Decompression" + vbCr + "  #     Dur     RT      Mix     Depth   Rate    ")
            t1print8 CStr(vsegment_vnumber)
            t1print8 CStr(vsegment_vtime)
            t1print8 CStr(run_vtime)
            t1print8 CStr(vmix_vnumber)
            t1print8 CStr(vdeco_vstop_vdepth)
            t1print8 CStr(rate)
            t1print (vbCrLf)
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
                    End If
                Next i
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
            last_run_vtime = 0#
            '=======================================================================
            '     vdeco vstop loop block for final vdecompression schedule
            '=======================================================================
            t1print ("Decompression" + vbCrLf + "  " + "   #     Dur      RT     Mix   Depth    Rate   Stime" + vbCrLf)
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
                t1print8dbl (CDbl(CInt(run_vtime * 10# + 0.999) / 10))
                t1print8 CStr(vmix_vnumber)
                t1print8dbl (CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.999) / 10))
                t1print8 CStr(rate)
                t1print (vbCrLf)
                If vdeco_vstop_vdepth <= 0# Then Exit Do   ' .le. 0.0) exit                !exit a
                If (vnumber_of_changes > 1) Then
                    vdepth_change_new = 9999
                    For i = 2 To vnumber_of_changes
                        If (vdepth_change(i) >= vdeco_vstop_vdepth) And vdepth_change(i) < vdepth_change_new Then
                            vmix_vnumber = vmix_change(i)
                            rate = rate_change(i)
                            vstep_size = vstep_size_change(i)
                            vdepth_change_new = vdepth_change(i)
                        End If
                    Next i
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
                    t1print CStr(vmix_vnumber)
                    t1print CStr(CInt(vdeco_vstop_vdepth))
                    t1print CStr(CInt(vstop_vtime))
                    t1print CStr(CInt(run_vtime))
                    t1print (vbCrLf)
                Else
                    t1print8 CStr(vsegment_vnumber)
                    t1print8dbl (CDbl(CInt(vsegment_vtime * 10# + 0.999) / 10))
                    t1print8dbl (CDbl(CInt(run_vtime * 10# + 0.999) / 10))
                    t1print8 CStr(vmix_vnumber)
                    t1print8dbl (CDbl(CInt(vdeco_vstop_vdepth * 10# + 0.999) / 10))
                    t1print8 CStr(rate)
                    t1print8dbl (CDbl(CInt(vstop_vtime * 10#) / 10))
                    t1print (vbCrLf)
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
        repetitive_dive_flag = 0
        If (repetitive_dive_flag = 0) Then
            GoTo L330 'Exit Do '          exit                                        !exit repetitive
            'at line 330
            '=======================================================================
            '     if there is a repetitive dive, compute gas loadings (off-gassing)
            '     surface interval time.  adjust critical radii using vpm repetitive
            '     valgorithm.  re-initialize selected variables and return to start o
            '     repetitive loop at line 30.
            '=======================================================================
            ElseIf (repetitive_dive_flag = 1) Then
            surface_interval_vtime = 60
            Call gas_loadings_surface_interval(surface_interval_vtime)  'su
            Call vpm_repetitive_valgorithm(surface_interval_vtime)       'su
            For i = 1 To 16
                max_crushing_vpressure_he(i) = 0#
                max_crushing_vpressure_n2(i) = 0#
                max_actual_gradient(i) = 0#
            Next i
            run_vtime = 0#
            vsegment_vnumber = 0
            '          cycle      !return to start of repetitive loop to process ano
            '=======================================================================
            '     write error message and terminate subroutine if there is an error in
            '     input file for the repetitive dive flag
            '=======================================================================
            Else
            MsgBox "subroutine terminated"
        End If
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
Dim initial_inspired_vhe_vpressure As Double
Dim initial_inspired_vn2_vpressure As Double
Dim last_run_vtime As Double
Dim vhelium_rate As Double
Dim vnitrogen_rate As Double
Dim starting_ambient_vpressure As Double
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
Dim initial_inspired_vhe_vpressure As Double
Dim initial_inspired_vn2_vpressure As Double
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
For j = 1 To 100
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
For i = 1 To 100
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
Dim inspired_vhelium_vpressure As Double
Dim inspired_vnitrogen_vpressure As Double
Dim ambient_vpressure As Double
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
Dim initial_inspired_vhe_vpressure As Double
Dim initial_inspired_vn2_vpressure As Double
Dim time_to_start_of_vdeco_zone As Double
Dim vhelium_rate As Double
Dim vnitrogen_rate As Double
Dim starting_ambient_vpressure As Double
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
        MsgBox "root not in brackets"
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
    For j = 1 To 100
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
    MsgBox "root exceed"
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
        Dim initial_inspired_vhe_vpressure As Double
        Dim initial_inspired_vn2_vpressure As Double
        Dim vhelium_rate As Double
        Dim vnitrogen_rate As Double
        Dim starting_ambient_vpressure As Double
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
Dim ambient_vpressure As Double
Dim inspired_vhelium_vpressure As Double
Dim inspired_vnitrogen_vpressure As Double
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
round_up_operation = CDbl(CLng((last_run_vtime / minimum_vdeco_vstop_vtime))) * minimum_vdeco_vstop_vtime
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
'=======================================================================
'     check to make sure that subroutine won't lock up if unable to vdecompr
'     to the next vstop.  if so, write error message and terminate progra
'=======================================================================
For i = 1 To 16
    If ((inspired_vhelium_vpressure + inspired_vnitrogen_vpressure) > 0#) Then
        weighted_allowable_gradient = (vdeco_gradient_he(i) * inspired_vhelium_vpressure + vdeco_gradient_n2(i) * inspired_vnitrogen_vpressure) / (inspired_vhelium_vpressure + inspired_vnitrogen_vpressure)
        If ((inspired_vhelium_vpressure + inspired_vnitrogen_vpressure + constant_vpressure_other_gases - weighted_allowable_gradient) > (next_vstop + barometric_vpressure)) Then
            MsgBox "subroutine terminated"
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
    If temp_vsegment_vtime >= 100 Then
      i = 1
    End If
    If (vdeco_vceiling_vdepth > next_vstop) And temp_vsegment_vtime < 100 Then
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
Sub gas_loadings_surface_interval(surface_interval_vtime As Double)
    '      implicit none
    '=======================================================================
    '     arguments
    '=======================================================================
    ' sub parameter : do not dim ! Dim surface_interval_vtime as double
    '=======================================================================
    '     local variables
    '=======================================================================
    Dim i As Integer                                                  'loop as integer
    Dim inspired_vhelium_vpressure As Double
    Dim inspired_vnitrogen_vpressure As Double
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
    Dim starting_ambient_vpressure As Double
    Dim ending_ambient_vpressure As Double
    Dim initial_inspired_vn2_vpressure As Double
    Dim rate As Double
    Dim vnitrogen_rate As Double
    Dim inspired_vnitrogen_vpressure As Double
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
        MsgBox "subroutine terminated"
    End If
    If ((units_equal_msw) And (valtitude_of_dive > 9144#)) Then
        MsgBox "subroutine terminated"
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
'        MsgBox "subroutine terminated"
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
            MsgBox "subroutine terminated"
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
  Text10.Text = Text10.Text + "  " + CStr(t1)
End Sub

Private Sub t1print8(t1 As String)
Dim S As String

  S = CStr(t1)
  Do While (Len(S) < 6)
    S = " " + S
  Loop
  If (Len(S) > 6) Then
    S = Left(S, 6)
  End If
  
  Text10.Text = Text10.Text + S + "  "
End Sub

Private Sub t1print8dbl(t1 As Double)
Dim S As String

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
  
  Text10.Text = Text10.Text + S + "  "
End Sub

Private Function vimportdb_data() As Integer
Dim i As Integer
  
  SQL = "SELECT * FROM dpprofile"
  SQL = SQL & " where dpprofileid = '" & tempserialno & "' "
  SQL = SQL & " order by dpnumseq "
  Set RS = DB.OpenRecordset(SQL)
  If RS.EOF Then
    MsgBox "Add Profile Plan Points before calculating decompression"
    vimportdb_data = 99
    Exit Function
  End If
  
  vimportdb_data = 0
  
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
     Plan_OpenClosed(i) = RS("dpcircuit")
'     MSFlexGrid3.Col = 6
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

  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & " where dpmainid = '" & tempserialno & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  Plan_Gas_list_numgasdeco = 1 ' make first gas the last bottom gas when main deco calc done. the extra deco gases are then added below
  For i = 0 To 9
     Plan_Gas_list_n2(i + 1) = CDbl(RS("dpgasnitrogen")) / 100#
     Plan_Gas_list_he(i + 1) = CDbl(RS("dpgashelium")) / 100#
     Plan_Gas_list_mod(i + 1) = CDbl(RS("dpgasmaxopdepth")) / 10#
     Plan_Gas_list_used(i + 1) = CInt(Left(CStr(RS("dpgasused")), 1))
     If Plan_Gas_list_used(i + 1) > 3 Then
       Plan_Gas_list_numgasdeco = Plan_Gas_list_numgasdeco + 1
       Plan_Gas_list_deco(Plan_Gas_list_numgasdeco) = i + 1
     End If
     RS.MoveNext
  Next i

End Function

Private Sub vhighlite_line(Index As Integer)
  MSFlexGrid3.Row = Index
  MSFlexGrid3_Click
End Sub



'nick code end here
