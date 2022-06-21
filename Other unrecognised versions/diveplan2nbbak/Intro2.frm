VERSION 5.00
Begin VB.Form frmintro 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Walk Through the system......"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "main"
      Height          =   9015
      Left            =   120
      TabIndex        =   118
      Top             =   120
      Width           =   11775
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   3105
         Left            =   720
         Picture         =   "Intro2.frx":0000
         ScaleHeight     =   3105
         ScaleWidth      =   4260
         TabIndex        =   123
         Top             =   1680
         Width           =   4260
      End
      Begin VB.PictureBox Picture6 
         Height          =   3135
         Left            =   6240
         Picture         =   "Intro2.frx":2B12E
         ScaleHeight     =   3075
         ScaleWidth      =   4395
         TabIndex        =   122
         Top             =   2520
         Width           =   4455
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00FFFFFF&
         Height          =   2895
         Left            =   600
         Picture         =   "Intro2.frx":59498
         ScaleHeight     =   2835
         ScaleWidth      =   4635
         TabIndex        =   121
         Top             =   5280
         Width           =   4695
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Close"
         Height          =   375
         Left            =   8040
         TabIndex        =   120
         Top             =   8640
         Width           =   1335
      End
      Begin VB.CommandButton cmdnext7 
         Caption         =   "Next"
         Height          =   375
         Left            =   9720
         TabIndex        =   119
         Top             =   8640
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the pictures above for more details..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   600
         TabIndex        =   129
         Top             =   8400
         Width           =   7455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dive Series - Create and plan the series of dive, generate the deco result and graphical dive pattern."
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
         Height          =   495
         Left            =   5400
         TabIndex        =   128
         Top             =   7440
         Width           =   4695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8D65E
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
         Height          =   735
         Left            =   6240
         TabIndex        =   127
         Top             =   5760
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8D6E5
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
         Height          =   2895
         Left            =   5160
         TabIndex        =   126
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Pro Planner - Intelligent software for Dive Series Planning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   2040
         TabIndex        =   125
         Top             =   120
         Width           =   8535
      End
      Begin VB.Label Label137 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8D773
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   720
         TabIndex        =   124
         Top             =   600
         Width           =   11055
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "seqplan2"
      Height          =   9135
      Left            =   120
      TabIndex        =   18
      Top             =   -360
      Width           =   12015
      Begin VB.CommandButton cmdhome 
         Caption         =   "Home"
         Height          =   375
         Left            =   9480
         TabIndex        =   102
         Top             =   8520
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "Next"
         Height          =   375
         Left            =   10920
         TabIndex        =   101
         Top             =   8520
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   5175
         Left            =   0
         Top             =   0
         Width           =   12075
      End
      Begin VB.Label Label136 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on another dive plan no in the list to see details for that dive."
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
         Height          =   495
         Left            =   6720
         TabIndex        =   81
         Top             =   7560
         Width           =   5055
      End
      Begin VB.Label Label135 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6240
         TabIndex        =   80
         Top             =   7500
         Width           =   375
      End
      Begin VB.Label Label134 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Note the changes of the Dive deteails when the you switch from one dive to another dive."
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
         Height          =   495
         Left            =   6720
         TabIndex        =   79
         Top             =   6840
         Width           =   5055
      End
      Begin VB.Label Label133 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6240
         TabIndex        =   78
         Top             =   6840
         Width           =   375
      End
      Begin VB.Label Label132 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8D92A
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
         Height          =   735
         Left            =   720
         TabIndex        =   77
         Top             =   7845
         Width           =   5055
      End
      Begin VB.Label Label131 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6240
         TabIndex        =   76
         Top             =   5640
         Width           =   375
      End
      Begin VB.Label Label130 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8D9B8
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
         Height          =   1215
         Left            =   6720
         TabIndex        =   75
         Top             =   5640
         Width           =   5055
      End
      Begin VB.Label Label129 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Data format - Data highlighted in yellow is the decompression result of the selected dive plan."
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
         Height          =   495
         Left            =   720
         TabIndex        =   74
         Top             =   6840
         Width           =   5055
      End
      Begin VB.Label Label128 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "The screen above shows the decompression result in graphical and text formats."
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
         Height          =   735
         Left            =   120
         TabIndex        =   73
         Top             =   6000
         Width           =   5415
      End
      Begin VB.Label Label127 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Plan New Dive Series :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   72
         Top             =   5640
         Width           =   3135
      End
      Begin VB.Label Label126 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   71
         Top             =   7800
         Width           =   375
      End
      Begin VB.Label Label125 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   70
         Top             =   6800
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Seqplan1"
      Height          =   9255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   12255
      Begin VB.CommandButton cmdhome6 
         Caption         =   "Home"
         Height          =   375
         Left            =   9840
         TabIndex        =   112
         Top             =   8640
         Width           =   1095
      End
      Begin VB.CommandButton cmdnext6 
         Caption         =   "Next"
         Height          =   375
         Left            =   11160
         TabIndex        =   111
         Top             =   8640
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   3255
         Left            =   9240
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label Label120 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   89
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label115 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   88
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label122 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   87
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label116 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   86
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label124 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   85
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label123 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   84
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label139 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   83
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label138 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   82
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label121 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   69
         Top             =   5280
         Width           =   375
      End
      Begin VB.Label Label114 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click here to delete the dive plan."
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
         Height          =   375
         Left            =   1080
         TabIndex        =   68
         Top             =   8760
         Width           =   5175
      End
      Begin VB.Label Label113 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   67
         Top             =   8760
         Width           =   375
      End
      Begin VB.Label Label112 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click here to edit the dive plan."
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
         Height          =   375
         Left            =   1080
         TabIndex        =   66
         Top             =   8400
         Width           =   5175
      End
      Begin VB.Label Label111 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   65
         Top             =   8400
         Width           =   375
      End
      Begin VB.Label Label110 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click here to duplicate the plan no as new plan no."
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
         Height          =   375
         Left            =   1080
         TabIndex        =   64
         Top             =   8040
         Width           =   5175
      End
      Begin VB.Label Label109 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   63
         Top             =   8040
         Width           =   375
      End
      Begin VB.Label Label108 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8DA8F
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
         Height          =   735
         Left            =   600
         TabIndex        =   62
         Top             =   4800
         Width           =   8055
      End
      Begin VB.Label Label107 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   61
         Top             =   5610
         Width           =   375
      End
      Begin VB.Label Label106 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   60
         Top             =   6240
         Width           =   375
      End
      Begin VB.Label Label105 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   59
         Top             =   6825
         Width           =   375
      End
      Begin VB.Label Label104 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Plan New Dive Series :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   4500
         Width           =   2655
      End
      Begin VB.Label Label103 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "This screen is used to create a new dive series. Choose Create new Dive series from main screen."
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
         Height          =   495
         Left            =   2880
         TabIndex        =   57
         Top             =   4560
         Width           =   9375
      End
      Begin VB.Label Label102 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Select a dive plan which you have created ealier,  plan selected will highlight in green color."
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
         Height          =   615
         Left            =   1080
         TabIndex        =   56
         Top             =   5640
         Width           =   6615
      End
      Begin VB.Label Label101 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Check through each depth point, interval, plan no and gas profile to confirm the settings are correct."
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
         Height          =   615
         Left            =   1080
         TabIndex        =   55
         Top             =   6240
         Width           =   6855
      End
      Begin VB.Label Label100 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the Add button to add into the list."
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
         Left            =   1080
         TabIndex        =   54
         Top             =   6840
         Width           =   5055
      End
      Begin VB.Label Label98 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dive Plan selected will appear into the list here."
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
         Height          =   375
         Left            =   1080
         TabIndex        =   52
         Top             =   7245
         Width           =   5055
      End
      Begin VB.Label Label97 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   51
         Top             =   7680
         Width           =   375
      End
      Begin VB.Label Label96 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "To add another plan into the list, follow the step from 1 to 4."
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
         Height          =   375
         Left            =   1080
         TabIndex        =   50
         Top             =   7680
         Width           =   5175
      End
      Begin VB.Shape Shape34 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   600
         Shape           =   2  'Oval
         Top             =   5610
         Width           =   375
      End
      Begin VB.Shape Shape35 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   600
         Shape           =   2  'Oval
         Top             =   6240
         Width           =   375
      End
      Begin VB.Shape Shape36 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   600
         Shape           =   2  'Oval
         Top             =   6840
         Width           =   375
      End
      Begin VB.Label Label99 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   53
         Top             =   7200
         Width           =   375
      End
      Begin VB.Shape Shape37 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   600
         Shape           =   2  'Oval
         Top             =   7200
         Width           =   375
      End
      Begin VB.Shape Shape38 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   600
         Shape           =   2  'Oval
         Top             =   7680
         Width           =   375
      End
      Begin VB.Shape Shape39 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   600
         Shape           =   2  'Oval
         Top             =   8040
         Width           =   375
      End
      Begin VB.Shape Shape40 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   600
         Shape           =   2  'Oval
         Top             =   8400
         Width           =   375
      End
      Begin VB.Shape Shape41 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   600
         Shape           =   2  'Oval
         Top             =   8760
         Width           =   375
      End
      Begin VB.Shape Shape48 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   8760
         Shape           =   2  'Oval
         Top             =   5280
         Width           =   375
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   1080
         Shape           =   2  'Oval
         Top             =   650
         Width           =   375
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   1800
         Shape           =   2  'Oval
         Top             =   645
         Width           =   375
      End
      Begin VB.Shape Shape22 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   360
         Shape           =   2  'Oval
         Top             =   650
         Width           =   375
      End
      Begin VB.Shape Shape17 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   480
         Shape           =   2  'Oval
         Top             =   2160
         Width           =   375
      End
      Begin VB.Shape Shape49 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   2400
         Shape           =   2  'Oval
         Top             =   3480
         Width           =   375
      End
      Begin VB.Shape Shape43 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   2400
         Shape           =   2  'Oval
         Top             =   2280
         Width           =   375
      End
      Begin VB.Shape Shape47 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   8640
         Shape           =   2  'Oval
         Top             =   840
         Width           =   375
      End
      Begin VB.Shape Shape42 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   7560
         Shape           =   2  'Oval
         Top             =   3720
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   4335
         Left            =   240
         Top             =   120
         Width           =   12135
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "diveplan2"
      Height          =   9135
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton cmdhome5 
         Caption         =   "Home"
         Height          =   375
         Left            =   9600
         TabIndex        =   110
         Top             =   8760
         Width           =   975
      End
      Begin VB.CommandButton cmdnext5 
         Caption         =   "Next"
         Height          =   375
         Left            =   10920
         TabIndex        =   109
         Top             =   8760
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8DB80
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
         Height          =   1215
         Left            =   600
         TabIndex        =   117
         Top             =   7005
         Width           =   5295
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   116
         Top             =   6960
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   1215
         Left            =   7200
         Top             =   4800
         Width           =   3615
      End
      Begin VB.Image Image4 
         Height          =   4215
         Left            =   120
         Top             =   120
         Width           =   12015
      End
      Begin VB.Label Label88 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8DC13
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
         Height          =   1095
         Left            =   6840
         TabIndex        =   49
         Top             =   6405
         Width           =   5055
      End
      Begin VB.Label Label85 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6360
         TabIndex        =   48
         Top             =   6360
         Width           =   375
      End
      Begin VB.Label Label83 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   5280
         Width           =   375
      End
      Begin VB.Label Label82 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   5880
         Width           =   375
      End
      Begin VB.Label Label81 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Plan new dive :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   4320
         Width           =   4095
      End
      Begin VB.Label Label75 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "The screen above is the decompression result in graphical and spreadsheet text format."
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
         Height          =   615
         Left            =   120
         TabIndex        =   44
         Top             =   4680
         Width           =   5895
      End
      Begin VB.Label Label74 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Spreadsheet text format - Data highlighted in yellow color is the decompression result."
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
         Height          =   615
         Left            =   600
         TabIndex        =   43
         Top             =   5325
         Width           =   5055
      End
      Begin VB.Label Label73 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8DCDA
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
         Height          =   1215
         Left            =   600
         TabIndex        =   42
         Top             =   5925
         Width           =   5175
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "plandive "
      Height          =   9135
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   12255
      Begin VB.CommandButton cmdhome4 
         Caption         =   "Home"
         Height          =   375
         Left            =   9360
         TabIndex        =   108
         Top             =   8640
         Width           =   975
      End
      Begin VB.CommandButton cmdnext4 
         Caption         =   "Next"
         Height          =   375
         Left            =   10560
         TabIndex        =   107
         Top             =   8640
         Width           =   1095
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   94
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   97
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   96
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   95
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label140 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         TabIndex        =   93
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label119 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   92
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label118 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   91
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label117 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   90
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label53 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "You have successfully create the dive plan, to enjoy for multi level dive plan, simply click the "" Next "" button."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   6480
         TabIndex        =   41
         Top             =   6960
         Width           =   5055
      End
      Begin VB.Label Label78 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "You will now see the ""multi level ?"" link appear, click to switch your view to muti level environment."
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
         Height          =   735
         Left            =   6940
         TabIndex        =   40
         Top             =   6180
         Width           =   5055
      End
      Begin VB.Label Label77 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8DDB1
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
         Height          =   735
         Left            =   6915
         TabIndex        =   39
         Top             =   5400
         Width           =   5055
      End
      Begin VB.Label Label76 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Modify the safety factor from here."
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
         Height          =   375
         Left            =   6915
         TabIndex        =   38
         Top             =   4980
         Width           =   5055
      End
      Begin VB.Label Label72 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   37
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   36
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label Label70 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   35
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label Label69 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Modify the Atmospheric value from here"
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
         Height          =   375
         Left            =   675
         TabIndex        =   34
         Top             =   8460
         Width           =   4095
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   8400
         Width           =   375
      End
      Begin VB.Label Label67 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Modify the PPO2 value from here, If system allow the field to be modify."
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
         Height          =   615
         Left            =   675
         TabIndex        =   32
         Top             =   7920
         Width           =   5055
      End
      Begin VB.Label Label66 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   7920
         Width           =   375
      End
      Begin VB.Label Label65 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8DE55
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
         Height          =   735
         Left            =   675
         TabIndex        =   30
         Top             =   7035
         Width           =   5055
      End
      Begin VB.Label Label64 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8DEFC
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
         Height          =   735
         Left            =   675
         TabIndex        =   29
         Top             =   6240
         Width           =   5055
      End
      Begin VB.Label Label63 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":8DF84
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
         Height          =   735
         Left            =   675
         TabIndex        =   28
         Top             =   5400
         Width           =   5055
      End
      Begin VB.Label Label62 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "the above screen is shown. Follow the steps below to create the new dive."
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
         Height          =   615
         Left            =   240
         TabIndex        =   27
         Top             =   4840
         Width           =   5415
      End
      Begin VB.Label Label61 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "When you click on the Plan New Dive,"
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
         Height          =   495
         Left            =   1920
         TabIndex        =   26
         Top             =   4590
         Width           =   4575
      End
      Begin VB.Label Label60 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Plan New Dive :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label59 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   7035
         Width           =   375
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   6240
         Width           =   375
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   5400
         Width           =   375
      End
      Begin VB.Shape Shape12 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   240
         Shape           =   2  'Oval
         Top             =   5400
         Width           =   375
      End
      Begin VB.Shape Shape13 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   240
         Shape           =   2  'Oval
         Top             =   6240
         Width           =   375
      End
      Begin VB.Shape Shape14 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   240
         Shape           =   2  'Oval
         Top             =   7035
         Width           =   375
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   240
         Shape           =   2  'Oval
         Top             =   7920
         Width           =   375
      End
      Begin VB.Shape Shape16 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   240
         Shape           =   2  'Oval
         Top             =   8400
         Width           =   375
      End
      Begin VB.Shape Shape18 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   6480
         Shape           =   2  'Oval
         Top             =   4920
         Width           =   375
      End
      Begin VB.Shape Shape19 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   6480
         Shape           =   2  'Oval
         Top             =   5520
         Width           =   375
      End
      Begin VB.Shape Shape20 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   6480
         Shape           =   2  'Oval
         Top             =   6120
         Width           =   375
      End
      Begin VB.Shape Shape50 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   6600
         Shape           =   2  'Oval
         Top             =   1440
         Width           =   375
      End
      Begin VB.Shape Shape46 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   8640
         Shape           =   2  'Oval
         Top             =   1440
         Width           =   375
      End
      Begin VB.Shape Shape45 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   3000
         Shape           =   2  'Oval
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape44 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   9840
         Shape           =   2  'Oval
         Top             =   2880
         Width           =   375
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   4680
         Shape           =   2  'Oval
         Top             =   3600
         Width           =   375
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   7680
         Shape           =   2  'Oval
         Top             =   3120
         Width           =   375
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   0
         Shape           =   2  'Oval
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape Shape10 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   7560
         Shape           =   2  'Oval
         Top             =   2400
         Width           =   375
      End
      Begin VB.Image Image6 
         Height          =   4035
         Left            =   240
         Picture         =   "Intro2.frx":8E02E
         Top             =   120
         Width           =   11265
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Gas"
      Height          =   9255
      Left            =   120
      TabIndex        =   14
      Top             =   -120
      Width           =   12015
      Begin VB.CommandButton Cmdhome3 
         Caption         =   "Home"
         Height          =   375
         Left            =   7200
         TabIndex        =   106
         Top             =   8640
         Width           =   1095
      End
      Begin VB.CommandButton cmdnext3 
         Caption         =   "Next"
         Height          =   375
         Left            =   8640
         TabIndex        =   105
         Top             =   8640
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11040
         TabIndex        =   156
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label79 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   155
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label89 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   154
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label Label87 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   153
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label Label86 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   152
         Top             =   6315
         Width           =   375
      End
      Begin VB.Label Label84 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Plan New Dive with multi level:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   151
         Top             =   3840
         Width           =   4335
      End
      Begin VB.Label Label80 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "When you click on the ""Multi level"" Link, the screen above will appear, you can now enjoy the multi level dive planning."
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
         Height          =   495
         Left            =   3720
         TabIndex        =   150
         Top             =   3840
         Width           =   8175
      End
      Begin VB.Label Label56 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":122300
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
         Height          =   1095
         Left            =   675
         TabIndex        =   149
         Top             =   4680
         Width           =   5055
      End
      Begin VB.Label Label55 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the ""Insert"" button will insert the level above the level selected(highlighted in blue color row in spreadsheet)"
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
         Height          =   735
         Left            =   675
         TabIndex        =   148
         Top             =   5760
         Width           =   5415
      End
      Begin VB.Label Label42 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":1223DE
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
         Height          =   735
         Left            =   675
         TabIndex        =   147
         Top             =   6315
         Width           =   5415
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   146
         Top             =   7080
         Width           =   375
      End
      Begin VB.Label Label27 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the ""Delete"" button will delete the selected level. (Selected level will highlighted in blue color row in spreadsheet)"
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
         Height          =   855
         Left            =   675
         TabIndex        =   145
         Top             =   7080
         Width           =   5295
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   144
         Top             =   7920
         Width           =   375
      End
      Begin VB.Label Label25 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the ""DClear All"" button will delete all level. "
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
         Height          =   375
         Left            =   675
         TabIndex        =   143
         Top             =   7920
         Width           =   5295
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   142
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   141
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label Label21 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":122471
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
         Height          =   1095
         Left            =   6915
         TabIndex        =   140
         Top             =   4740
         Width           =   5055
      End
      Begin VB.Label Label19 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":122520
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
         Height          =   735
         Left            =   6945
         TabIndex        =   139
         Top             =   5760
         Width           =   5055
      End
      Begin VB.Label Label18 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "You have successfully create the dive plan, to view the decompression result and the graphical plotting, press the "" Next ""  ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   6720
         TabIndex        =   138
         Top             =   6720
         Width           =   5055
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   137
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   136
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         TabIndex        =   135
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   134
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   133
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   132
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   131
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   130
         Top             =   2640
         Width           =   375
      End
      Begin VB.Line Line28 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   7080
         X2              =   7320
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   0
         Shape           =   2  'Oval
         Top             =   2640
         Width           =   375
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   6240
         Shape           =   2  'Oval
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   6240
         Shape           =   2  'Oval
         Top             =   1680
         Width           =   375
      End
      Begin VB.Shape Shape9 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   8280
         Shape           =   2  'Oval
         Top             =   600
         Width           =   375
      End
      Begin VB.Shape Shape11 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   8280
         Shape           =   2  'Oval
         Top             =   1320
         Width           =   375
      End
      Begin VB.Shape Shape21 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   8280
         Shape           =   2  'Oval
         Top             =   2040
         Width           =   375
      End
      Begin VB.Shape Shape23 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   9120
         Shape           =   2  'Oval
         Top             =   2880
         Width           =   375
      End
      Begin VB.Shape Shape24 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   10080
         Shape           =   2  'Oval
         Top             =   2880
         Width           =   375
      End
      Begin VB.Shape Shape25 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   8280
         Shape           =   2  'Oval
         Top             =   2880
         Width           =   375
      End
      Begin VB.Shape Shape26 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   11040
         Shape           =   2  'Oval
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Image7 
         Height          =   3660
         Left            =   120
         Top             =   0
         Width           =   11835
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "mainmenu"
      Height          =   9255
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton cmdhome2 
         Caption         =   "Home"
         Height          =   375
         Left            =   8280
         TabIndex        =   104
         Top             =   8640
         Width           =   1095
      End
      Begin VB.CommandButton cmdnext1 
         Caption         =   "Next"
         Height          =   375
         Left            =   9720
         TabIndex        =   103
         Top             =   8640
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Click On this icon will switch the view from Dive Series to Dive Plan"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   115
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Line Line24 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3135
         X2              =   3000
         Y1              =   3000
         Y2              =   2880
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3000
         X2              =   3120
         Y1              =   3120
         Y2              =   3000
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3120
         X2              =   600
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   600
         X2              =   600
         Y1              =   3000
         Y2              =   3240
      End
      Begin VB.Image Image12 
         Height          =   975
         Left            =   240
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "This is the Dive Plan No."
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5160
         TabIndex        =   114
         Top             =   7560
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Click On this icon will switch the view from Dive Plan to Dive Series"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   360
         TabIndex        =   113
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3015
         X2              =   2880
         Y1              =   6240
         Y2              =   6120
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   2880
         X2              =   3000
         Y1              =   6360
         Y2              =   6240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3000
         X2              =   480
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   480
         X2              =   480
         Y1              =   6240
         Y2              =   6480
      End
      Begin VB.Image Image11 
         Height          =   975
         Left            =   240
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Line Line23 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   7080
         X2              =   7080
         Y1              =   6600
         Y2              =   7440
      End
      Begin VB.Line Line22 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   6975
         X2              =   7075
         Y1              =   6720
         Y2              =   6600
      End
      Begin VB.Line Line21 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   7080
         X2              =   7180
         Y1              =   6600
         Y2              =   6720
      End
      Begin VB.Line Line20 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   5880
         X2              =   5980
         Y1              =   6720
         Y2              =   6840
      End
      Begin VB.Line Line19 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   5775
         X2              =   5875
         Y1              =   6840
         Y2              =   6720
      End
      Begin VB.Line Line18 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   5880
         X2              =   5880
         Y1              =   6840
         Y2              =   7440
      End
      Begin VB.Line Line13 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3840
         X2              =   3940
         Y1              =   6600
         Y2              =   6720
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3720
         X2              =   3820
         Y1              =   6720
         Y2              =   6600
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3840
         X2              =   3840
         Y1              =   6600
         Y2              =   7440
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3960
         X2              =   4060
         Y1              =   4080
         Y2              =   4200
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   6225
         X2              =   6325
         Y1              =   4200
         Y2              =   4320
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   7905
         X2              =   8005
         Y1              =   4200
         Y2              =   4320
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   7800
         X2              =   7900
         Y1              =   4320
         Y2              =   4200
      End
      Begin VB.Line Line45 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3840
         X2              =   3940
         Y1              =   4200
         Y2              =   4080
      End
      Begin VB.Line Line44 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   6105
         X2              =   6205
         Y1              =   4320
         Y2              =   4200
      End
      Begin VB.Line Line15 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   6240
         X2              =   6240
         Y1              =   4200
         Y2              =   4440
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3960
         X2              =   3960
         Y1              =   4080
         Y2              =   4440
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   7920
         X2              =   7920
         Y1              =   4200
         Y2              =   4560
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "ppo2 value"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7560
         TabIndex        =   100
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "This is the maximum depth"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5640
         TabIndex        =   99
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "This is the Dive Plan No."
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3120
         TabIndex        =   98
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Image Image9 
         Height          =   2055
         Left            =   3240
         Top             =   5040
         Width           =   8775
      End
      Begin VB.Label Label46 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the text highlighted in blue color for more explanation..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1440
         TabIndex        =   21
         Top             =   8160
         Width           =   8895
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Surface time between two dive"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6600
         TabIndex        =   20
         Top             =   7560
         Width           =   1335
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "This is the Dive Series No."
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3240
         TabIndex        =   19
         Top             =   7560
         Width           =   1215
      End
      Begin VB.Label Label40 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "- You can delete the dive series here. "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6600
         TabIndex        =   13
         Top             =   1455
         Width           =   2775
      End
      Begin VB.Label Label39 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
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
         Left            =   4920
         MouseIcon       =   "Intro2.frx":1225BA
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label38 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "- You can edit the dive series here. "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6600
         TabIndex        =   11
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label37 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
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
         Left            =   4920
         MouseIcon       =   "Intro2.frx":1228C4
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label36 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "- Create new Dive Series here. "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6600
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label35 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Make New Series "
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
         Left            =   4920
         MouseIcon       =   "Intro2.frx":122BCE
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label34 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "- You can delete the dive plan here. "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   1455
         Width           =   2775
      End
      Begin VB.Label Label33 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Delete this Dive  "
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
         Left            =   480
         MouseIcon       =   "Intro2.frx":122ED8
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label32 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "- You can edit the dive plan here. "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   1095
         Width           =   2415
      End
      Begin VB.Label Label31 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
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
         Left            =   480
         MouseIcon       =   "Intro2.frx":1231E2
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label30 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "- Click to create a new dive plan."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   735
         Width           =   2775
      End
      Begin VB.Label Label29 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Pro Planner Main Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   7695
      End
      Begin VB.Label Label28 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
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
         Left            =   480
         MouseIcon       =   "Intro2.frx":1234EC
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.Image Image10 
         Height          =   2175
         Left            =   3240
         Top             =   1920
         Width           =   8655
      End
   End
End
Attribute VB_Name = "frmintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub viewframe1()
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame2.Visible = False
Frame1.Visible = True
End Sub
Private Sub viewframe2()
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame2.Visible = True
Image11.Picture = LoadPicture(App.Path & "\" & "wt4" + ".bmp")
Image12.Picture = LoadPicture(App.Path & "\" & "wt5" + ".bmp")
Image9.Picture = LoadPicture(App.Path & "\" & "wt3" + ".bmp")
Image10.Picture = LoadPicture(App.Path & "\" & "wt2" + ".bmp")
Frame1.Visible = False
End Sub
Private Sub viewframe3()
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = True
Image7.Picture = LoadPicture(App.Path & "\" & "wt8" + ".bmp")
Frame2.Visible = False
Frame1.Visible = False
End Sub
Private Sub viewframe4()
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = True
Frame3.Visible = False
Frame2.Visible = False
Frame1.Visible = False
Image6.Picture = LoadPicture(App.Path & "\" & "wt6" + ".bmp")
End Sub
Private Sub viewframe5()
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = True
Frame4.Visible = False
Frame3.Visible = False
Frame2.Visible = False
Frame1.Visible = False
Image4.Picture = LoadPicture(App.Path & "\" & "wt7" + ".bmp")
Image5.Picture = LoadPicture(App.Path & "\" & "diveplan3" + ".bmp")
End Sub
Private Sub viewframe6()
Frame7.Visible = False
Frame6.Visible = True
Image2.Picture = LoadPicture(App.Path & "\" & "seqplan6" + ".bmp")
Image3.Picture = LoadPicture(App.Path & "\" & "seqplan7" + ".bmp")
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame2.Visible = False
Frame1.Visible = False
End Sub
Private Sub viewframe7()
Frame7.Visible = True
Image1.Picture = LoadPicture(App.Path & "\" & "seqplan8" + ".bmp")
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame2.Visible = False
Frame1.Visible = False
End Sub

Private Sub CMDCLOSE_Click()
Unload Me
End Sub

Private Sub cmdhome_Click()
viewframe1
End Sub

Private Sub cmdhome2_Click()
viewframe1
End Sub

Private Sub Cmdhome3_Click()
viewframe1
End Sub

Private Sub cmdhome4_Click()
viewframe1
End Sub

Private Sub cmdhome5_Click()
viewframe1
End Sub

Private Sub cmdhome6_Click()
viewframe1
End Sub

Private Sub cmdnext_Click()
viewframe1
End Sub

Private Sub cmdnext1_Click()
viewframe4
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdnext3_Click()
viewframe5
End Sub

Private Sub cmdnext4_Click()
viewframe3

End Sub

Private Sub cmdnext5_Click()
viewframe6
End Sub

Private Sub cmdnext6_Click()
viewframe7
End Sub

Private Sub cmdnext7_Click()
viewframe2
End Sub

Private Sub Form_Load()
Top = 20
Left = 1300

viewframe1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Splanmain.Show
End Sub

Private Sub Label140_Click()
splmain.chow
End Sub

Private Sub Label10_Click()
viewframe1
End Sub

Private Sub Label16_Click()
viewframe3
End Sub

Private Sub Label18_Click()
viewframe4
End Sub

Private Sub Label28_Click()
viewframe4
End Sub

Private Sub Label31_Click()
viewframe4
End Sub

Private Sub Label33_Click()
viewframe4
End Sub

Private Sub Label35_Click()
viewframe6
End Sub

Private Sub Label37_Click()
viewframe6
End Sub

Private Sub Label39_Click()
viewframe6
End Sub

Private Sub Picture11_Click()
 viewframe1
End Sub

Private Sub Picture12_Click()

End Sub

Private Sub Picture13_Click()
 viewframe1
End Sub

Private Sub Picture14_Click()
viewframe1
End Sub

Private Sub Picture15_Click()
End Sub

Private Sub Picture16_Click()
End Sub

Private Sub Picture1_Click()
viewframe3
End Sub

Private Sub Picture5_Click()
viewframe2
End Sub

Private Sub Picture6_Click()
viewframe4
End Sub

Private Sub Picture8_Click()
viewframe6
End Sub
