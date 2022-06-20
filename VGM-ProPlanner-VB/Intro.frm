VERSION 5.00
Begin VB.Form frmintro 
   Caption         =   "Introduction"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7815
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9735
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   7815
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   9735
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Height          =   7815
            Index           =   2
            Left            =   480
            TabIndex        =   32
            Top             =   2040
            Width           =   9735
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   5655
               Left            =   240
               Picture         =   "Intro.frx":0000
               ScaleHeight     =   5655
               ScaleWidth      =   9255
               TabIndex        =   33
               Top             =   1320
               Width           =   9255
               Begin VB.Label Label27 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "This is the Dive Series No."
                  ForeColor       =   &H000000C0&
                  Height          =   495
                  Left            =   3240
                  TabIndex        =   36
                  Top             =   3120
                  Width           =   1455
               End
               Begin VB.Line Line18 
                  BorderWidth     =   2
                  X1              =   3720
                  X2              =   3840
                  Y1              =   2640
                  Y2              =   3120
               End
               Begin VB.Line Line17 
                  BorderWidth     =   2
                  X1              =   3720
                  X2              =   3650
                  Y1              =   2640
                  Y2              =   2800
               End
               Begin VB.Line Line16 
                  BorderWidth     =   2
                  X1              =   3720
                  X2              =   3840
                  Y1              =   2640
                  Y2              =   2760
               End
               Begin VB.Line Line15 
                  BorderWidth     =   2
                  X1              =   5475
                  X2              =   5595
                  Y1              =   2880
                  Y2              =   3360
               End
               Begin VB.Line Line14 
                  BorderWidth     =   2
                  X1              =   5470
                  X2              =   5400
                  Y1              =   2880
                  Y2              =   3040
               End
               Begin VB.Line Line13 
                  BorderWidth     =   2
                  X1              =   5475
                  X2              =   5595
                  Y1              =   2880
                  Y2              =   3000
               End
               Begin VB.Label Label26 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "This is the individual Dive Plan No. "
                  ForeColor       =   &H000000C0&
                  Height          =   495
                  Left            =   4800
                  TabIndex        =   35
                  Top             =   3360
                  Width           =   1455
               End
               Begin VB.Label Label25 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "You are editing the gas profile of this Plan serial no  "
                  ForeColor       =   &H000000C0&
                  Height          =   495
                  Left            =   6840
                  TabIndex        =   34
                  Top             =   1560
                  Width           =   2295
               End
               Begin VB.Line Line12 
                  BorderWidth     =   2
                  X1              =   7635
                  X2              =   7755
                  Y1              =   840
                  Y2              =   960
               End
               Begin VB.Line Line11 
                  BorderWidth     =   2
                  X1              =   7630
                  X2              =   7560
                  Y1              =   840
                  Y2              =   1000
               End
               Begin VB.Line Line10 
                  BorderWidth     =   2
                  X1              =   7620
                  X2              =   7870
                  Y1              =   840
                  Y2              =   1560
               End
            End
            Begin VB.Label Label41 
               BackColor       =   &H00FFFFFF&
               Caption         =   $"Intro.frx":AF842
               Height          =   735
               Left            =   360
               TabIndex        =   40
               Top             =   480
               Width           =   9135
            End
            Begin VB.Label Label42 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Library Dive : Gas Profile"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   120
               TabIndex        =   41
               Top             =   120
               Width           =   7695
            End
            Begin VB.Label Label30 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Click on the blue text for explaination....."
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   495
               Left            =   2760
               TabIndex        =   39
               Top             =   7200
               Width           =   5055
            End
            Begin VB.Label Label29 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ">>>"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   495
               Left            =   8580
               TabIndex        =   38
               Top             =   7250
               Width           =   615
            End
            Begin VB.Label Label28 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Main Menu"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   450
               Left            =   720
               TabIndex        =   37
               Top             =   7320
               Width           =   1215
            End
            Begin VB.Shape Shape4 
               BackColor       =   &H00FFFFC0&
               BackStyle       =   1  'Opaque
               Height          =   495
               Left            =   8160
               Shape           =   2  'Oval
               Top             =   7200
               Width           =   1215
            End
            Begin VB.Shape Shape3 
               FillColor       =   &H00FFFFC0&
               FillStyle       =   0  'Solid
               Height          =   495
               Left            =   480
               Shape           =   2  'Oval
               Top             =   7200
               Width           =   1455
            End
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   5295
            Left            =   480
            Picture         =   "Intro.frx":AF909
            ScaleHeight     =   5295
            ScaleWidth      =   8895
            TabIndex        =   7
            Top             =   1800
            Width           =   8895
            Begin VB.Line Line9 
               BorderWidth     =   2
               X1              =   7020
               X2              =   7270
               Y1              =   2880
               Y2              =   3600
            End
            Begin VB.Line Line8 
               BorderWidth     =   2
               X1              =   7030
               X2              =   6960
               Y1              =   2880
               Y2              =   3040
            End
            Begin VB.Line Line7 
               BorderWidth     =   2
               X1              =   7035
               X2              =   7155
               Y1              =   2880
               Y2              =   3000
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "Record highlighted in blue was selected for Edit and delete."
               ForeColor       =   &H000000C0&
               Height          =   495
               Left            =   6480
               TabIndex        =   28
               Top             =   3600
               Width           =   2295
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "This is the individual Dive Plan No. "
               ForeColor       =   &H000000C0&
               Height          =   495
               Left            =   4800
               TabIndex        =   16
               Top             =   3360
               Width           =   1455
            End
            Begin VB.Line Line6 
               BorderWidth     =   2
               X1              =   5475
               X2              =   5595
               Y1              =   2880
               Y2              =   3000
            End
            Begin VB.Line Line5 
               BorderWidth     =   2
               X1              =   5470
               X2              =   5400
               Y1              =   2880
               Y2              =   3040
            End
            Begin VB.Line Line4 
               BorderWidth     =   2
               X1              =   5475
               X2              =   5595
               Y1              =   2880
               Y2              =   3360
            End
            Begin VB.Line Line3 
               BorderWidth     =   2
               X1              =   3720
               X2              =   3840
               Y1              =   2640
               Y2              =   2760
            End
            Begin VB.Line Line2 
               BorderWidth     =   2
               X1              =   3720
               X2              =   3650
               Y1              =   2640
               Y2              =   2800
            End
            Begin VB.Line Line1 
               BorderWidth     =   2
               X1              =   3720
               X2              =   3840
               Y1              =   2640
               Y2              =   3120
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "This is the Dive Series No."
               ForeColor       =   &H000000C0&
               Height          =   495
               Left            =   3240
               TabIndex        =   15
               Top             =   3120
               Width           =   1455
            End
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Main Menu"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   720
            TabIndex        =   31
            Top             =   7320
            Width           =   1215
         End
         Begin VB.Label Label23 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ">>>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   8520
            TabIndex        =   30
            Top             =   7200
            Width           =   615
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Click on the bule text for explaination....."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   2760
            TabIndex        =   29
            Top             =   7200
            Width           =   5055
         End
         Begin VB.Label Label20 
            BackColor       =   &H00FFFFFF&
            Caption         =   "- You can delete the dive series here. "
            Height          =   255
            Left            =   6480
            TabIndex        =   27
            Top             =   1335
            Width           =   2775
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Delete this Series"
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
            Left            =   4800
            TabIndex        =   26
            Top             =   1320
            Width           =   3735
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFFFFF&
            Caption         =   "- You can edit the dive series here. "
            Height          =   255
            Left            =   6150
            TabIndex        =   25
            Top             =   975
            Width           =   2655
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edit this Series  "
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
            Left            =   4800
            TabIndex        =   24
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFFFF&
            Caption         =   "- Create new Dive Series here. "
            Height          =   255
            Left            =   7180
            TabIndex        =   23
            Top             =   630
            Width           =   2175
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Make a New Dive Series "
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
            Left            =   4800
            TabIndex        =   22
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFFFF&
            Caption         =   "- You can delete the dive plan here. "
            Height          =   255
            Left            =   1920
            TabIndex        =   21
            Top             =   1335
            Width           =   2775
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Delete this Dive Dive "
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
            Left            =   360
            TabIndex        =   20
            Top             =   1320
            Width           =   3735
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "- You can edit the dive plan here. "
            Height          =   255
            Left            =   1560
            TabIndex        =   19
            Top             =   975
            Width           =   2415
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edit this Dive "
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
            Left            =   360
            TabIndex        =   18
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "- You can create a new dive plan here. "
            Height          =   255
            Left            =   1920
            TabIndex        =   17
            Top             =   615
            Width           =   2775
         End
         Begin VB.Label lblmain 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sequential Dive Plan list"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   120
            Width           =   7695
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Plan a New Dive "
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
            Left            =   360
            TabIndex        =   8
            Top             =   600
            Width           =   3735
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00FFFFC0&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   480
            Shape           =   2  'Oval
            Top             =   7200
            Width           =   1455
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   1  'Opaque
            Height          =   495
            Left            =   8160
            Shape           =   2  'Oval
            Top             =   7200
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   2415
         Left            =   5160
         Picture         =   "Intro.frx":1F1A2B
         ScaleHeight     =   2355
         ScaleWidth      =   3675
         TabIndex        =   12
         Top             =   840
         Width           =   3735
      End
      Begin VB.PictureBox Picture8 
         Height          =   2295
         Left            =   5280
         Picture         =   "Intro.frx":20DECD
         ScaleHeight     =   2235
         ScaleWidth      =   3555
         TabIndex        =   4
         Top             =   3840
         Width           =   3615
      End
      Begin VB.PictureBox Picture6 
         Height          =   2415
         Left            =   360
         Picture         =   "Intro.frx":44DF0F
         ScaleHeight     =   2355
         ScaleWidth      =   3555
         TabIndex        =   3
         Top             =   3840
         Width           =   3615
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   240
         Picture         =   "Intro.frx":5B3365
         ScaleHeight     =   2415
         ScaleWidth      =   3735
         TabIndex        =   2
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Click on the graphical image to go for details explaination."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   1080
         TabIndex        =   14
         Top             =   6720
         Width           =   7455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gas Profile - Create your own gas setting, save as default for your dive planing."
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   5160
         TabIndex        =   13
         Top             =   3270
         Width           =   3735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dive Series - Create and plan the series of dive, generate the deco result and graphical dive pattern."
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   5280
         TabIndex        =   11
         Top             =   6195
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Library Dive Editor - You can create and specify the parameters for individual  dive."
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   6240
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dive Plan List - Brain of the system, you can plan, edit and delete of any library and series dive ."
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DivePlan - An intelligent software for Sequential Dive Plan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   7695
      End
   End
End
Attribute VB_Name = "frmintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Frame1(1).Visible = False
'Frame1(0).Visible = True
Frame1(2).Visible = True
End Sub

Private Sub Label28_Click()
Frame1(0).Visible = True
End Sub

Private Sub Picture1_Click()

Frame1(1).Visible = False
Frame1(0).Visible = False
Frame1(2).Visible = True
End Sub

Private Sub Picture5_Click()
Frame1(1).Visible = True
Frame1(2).Visible = False

End Sub
