VERSION 5.00
Begin VB.Form frmgasprofile2 
   BackColor       =   &H00808000&
   Caption         =   "Library Dive : Gas Profile"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   ForeColor       =   &H00000000&
   Icon            =   "frmgasprofile2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdcreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   4320
      TabIndex        =   78
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "Reset Factory"
      Height          =   315
      Left            =   4125
      TabIndex        =   61
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Cmdsetdefault 
      Caption         =   "Set as Default"
      Height          =   315
      Left            =   5685
      TabIndex        =   60
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   1350
      TabIndex        =   59
      Top             =   1080
      Width           =   1080
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1350
      TabIndex        =   58
      Top             =   1440
      Width           =   1080
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   1350
      TabIndex        =   57
      Top             =   1800
      Width           =   1080
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   1350
      TabIndex        =   56
      Top             =   2160
      Width           =   1080
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   1350
      TabIndex        =   55
      Top             =   2520
      Width           =   1080
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   1350
      TabIndex        =   54
      Top             =   2880
      Width           =   1080
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   1350
      TabIndex        =   53
      Top             =   3240
      Width           =   1080
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   1350
      TabIndex        =   52
      Top             =   3600
      Width           =   1080
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   1350
      TabIndex        =   51
      Top             =   3960
      Width           =   1080
   End
   Begin VB.TextBox txtoxygen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   9
      Left            =   1350
      TabIndex        =   50
      Top             =   4320
      Width           =   1080
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   2445
      TabIndex        =   49
      Top             =   1080
      Width           =   1200
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   2445
      TabIndex        =   48
      Top             =   1440
      Width           =   1200
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   2445
      TabIndex        =   47
      Top             =   1800
      Width           =   1200
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   2445
      TabIndex        =   46
      Top             =   2160
      Width           =   1200
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   2445
      TabIndex        =   45
      Top             =   2520
      Width           =   1200
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   2445
      TabIndex        =   44
      Top             =   2880
      Width           =   1200
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   2445
      TabIndex        =   43
      Top             =   3240
      Width           =   1200
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   2445
      TabIndex        =   42
      Top             =   3600
      Width           =   1200
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   2445
      TabIndex        =   41
      Top             =   3960
      Width           =   1200
   End
   Begin VB.TextBox txthelium 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   9
      Left            =   2445
      TabIndex        =   40
      Top             =   4320
      Width           =   1200
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   3660
      TabIndex        =   39
      Top             =   1080
      Width           =   1300
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   3660
      TabIndex        =   38
      Top             =   1440
      Width           =   1300
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   3660
      TabIndex        =   37
      Top             =   1800
      Width           =   1300
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   3660
      TabIndex        =   36
      Top             =   2160
      Width           =   1300
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   3660
      TabIndex        =   35
      Top             =   2520
      Width           =   1300
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   3660
      TabIndex        =   34
      Top             =   2880
      Width           =   1300
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   3660
      TabIndex        =   33
      Top             =   3240
      Width           =   1300
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   3660
      TabIndex        =   32
      Top             =   3600
      Width           =   1300
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   3660
      TabIndex        =   31
      Top             =   3960
      Width           =   1300
   End
   Begin VB.TextBox txtmaxd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   9
      Left            =   3660
      TabIndex        =   30
      Top             =   4320
      Width           =   1300
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   4980
      TabIndex        =   29
      Top             =   1080
      Width           =   1005
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   4980
      TabIndex        =   28
      Top             =   1440
      Width           =   1005
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   4980
      TabIndex        =   27
      Top             =   1800
      Width           =   1005
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   4980
      TabIndex        =   26
      Top             =   2160
      Width           =   1005
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   4980
      TabIndex        =   25
      Top             =   2520
      Width           =   1005
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   4980
      TabIndex        =   24
      Top             =   2880
      Width           =   1005
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   4980
      TabIndex        =   23
      Top             =   3240
      Width           =   1005
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   4980
      TabIndex        =   22
      Top             =   3600
      Width           =   1005
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   4980
      TabIndex        =   21
      Top             =   3960
      Width           =   1005
   End
   Begin VB.TextBox txtppo2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   9
      Left            =   4980
      TabIndex        =   20
      Top             =   4320
      Width           =   1005
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   0
      Left            =   6005
      Sorted          =   -1  'True
      TabIndex        =   19
      Text            =   "cbogasused"
      Top             =   1080
      Width           =   2850
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   1
      Left            =   6005
      TabIndex        =   18
      Text            =   "cbogasused"
      Top             =   1440
      Width           =   2850
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   2
      Left            =   6005
      TabIndex        =   17
      Text            =   "cbogasused"
      Top             =   1800
      Width           =   2850
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   3
      Left            =   6005
      TabIndex        =   16
      Text            =   "cbogasused"
      Top             =   2160
      Width           =   2850
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   4
      Left            =   6005
      TabIndex        =   15
      Text            =   "cbogasused"
      Top             =   2520
      Width           =   2850
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   5
      Left            =   6005
      TabIndex        =   14
      Text            =   "cbogasused"
      Top             =   2880
      Width           =   2850
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   6
      Left            =   6005
      TabIndex        =   13
      Text            =   "cbogasused"
      Top             =   3240
      Width           =   2850
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   7
      Left            =   6005
      TabIndex        =   12
      Text            =   "cbogasused"
      Top             =   3600
      Width           =   2850
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   8
      Left            =   6005
      TabIndex        =   11
      Text            =   "cbogasused"
      Top             =   3960
      Width           =   2850
   End
   Begin VB.ComboBox cbogasused 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   9
      Left            =   6005
      TabIndex        =   10
      Text            =   "cbogasused"
      Top             =   4320
      Width           =   2850
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   315
      Left            =   165
      TabIndex        =   9
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Cmdcancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2805
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Cmdplanprofile 
      Caption         =   "Plan Profile"
      Height          =   315
      Left            =   1485
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   4
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtmaxd2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   3
      Top             =   6285
      Width           =   1030
   End
   Begin VB.TextBox txthe2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9495
      TabIndex        =   2
      Top             =   6285
      Width           =   1030
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gas Index"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   77
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gas Used"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6020
      TabIndex        =   76
      Top             =   720
      Width           =   2845
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "PPO2"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4980
      TabIndex        =   75
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Max. Depth"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3660
      TabIndex        =   74
      Top             =   720
      Width           =   1300
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Helium"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2445
      TabIndex        =   73
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Oxygen"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1355
      TabIndex        =   72
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 0"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   71
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 1"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   70
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 2"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   69
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 3"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   68
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 4"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   67
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 5"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   66
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 6"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   65
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 7"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   64
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 8"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   8
      Left            =   120
      TabIndex        =   63
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label gasindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gas 9"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   9
      Left            =   120
      TabIndex        =   62
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblsediveno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sequential Serial No"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblserialno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan Serial No"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmgasprofile2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temptxtmaxd1, temptxtmaxd2, temptxtmaxd3, temptxtmaxd4, temptxtmaxd5, temptxtmaxd6, temptxtmaxd, temptxtmaxd8, temptxtmaxd9, temptxtmaxd10
Dim temptxtni1, temptxtni2, temptxtni3, temptxtni4, temptxtni5, temptxtni6, temptxtni7, temptxtni8, temptxtni9, temptxtni10
Dim temptxthe1, temptxthe2, temptxthe3, temptxthe4, temptxthe5, temptxthe6, temptxthe7, temptxthe8, temptxthe9, temptxthe10


Private Sub Cbogasused_Change(index As Integer)
p = Cbogasused(index).index
If Cbogasused(p).Text = "5 - Deco Closed Circuit" Then
  txtppo2(p).Enabled = True
Else
txtppo2(p).Enabled = False
End If
End Sub

Private Sub Cbogasused_Click(index As Integer)
p = Cbogasused(index).index
If Cbogasused(p).Text = "5 - Deco Closed Circuit" Then
  txtppo2(p).Enabled = True
Else
  txtppo2(p).Enabled = False
End If
End Sub

Private Sub cmdcancel_Click()
Splanmain.Show
End Sub



Private Sub Cmdcreate_Click()
gasvalidate
If gasvalidation = "failed" Then
   MsgBox "Data cannot be blank, please fill in data !"
Else
   Select Case tempchoice
   Case "GP"
     savenewrecord
     Unload Me
     Planprofile.Show
   Case "NP"
     savenewrecord
     Unload Me
     Planprofile.Show
   Case "NSP"
     saveseqnewrecord
     Unload Me
     Planprofile2.Show
   Case "NPP"
     saveseqnewrecord
     Planprofile2.Show
     Unload Me
     'Unload Planprofile2
   Case "GSP"
     saveseqnewrecord
     Unload Me
     Planprofile2.Show
   End Select
End If
End Sub
Private Sub savenewrecord()
SQL = "SELECT * FROM dpmaingaslist "
Set RS = DB.OpenRecordset(SQL)
For i = 0 To 9
   RS.AddNew
   RS!dpmainid = tempserialno
   RS!dpgasid = gasindex(i).Caption
   RS!dpgashelium = txthelium(i).Text
   tempnitrogen = 100 - Val(txthelium(i).Text) - Val(txtoxygen(i).Text)
   RS!dpgasnitrogen = tempnitrogen
   RS!dpgasmaxopdepth = Val(txtmaxd(i).Text) * 10
   RS!dpgaspo2setpoint = txtppo2(i).Text
   RS!dpgasused = Cbogasused(i).Text
   RS.Update
Next i
SQL = "SELECT * FROM dpmain"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!diveplanid = tempserialno
RS.Update
End Sub
Private Sub saveseqnewrecord()
SQL = "SELECT * FROM dpmaingaslist"
Set RS = DB.OpenRecordset(SQL)
For i = 0 To 9
   RS.AddNew
   RS!dpmainid = tempserialno
   RS!dpgasid = gasindex(i).Caption
   RS!dpgashelium = txthelium(i).Text
   tempnitrogen = 100 - Val(txthelium(i).Text) - Val(txtoxygen(i).Text)
   RS!dpgasnitrogen = tempnitrogen
   RS!dpgasmaxopdepth = Val(txtmaxd(i).Text) * 10
   RS!dpgaspo2setpoint = txtppo2(i).Text
   RS!dpgasused = Cbogasused(i).Text
   RS.Update
Next i
SQL = "SELECT * FROM seqdpmain"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!diveplanid = tempserialno
RS!Plandate = Now
RS.Update
End Sub
Private Sub saveseqcurrentrecord()
SQL = "SELECT * FROM dpmaingaslist"
Set RS = DB.OpenRecordset(SQL)
For i = 0 To 9
   RS.AddNew
   RS!dpmainid = tempserialno
   RS!dpgasid = gasindex(i).Caption
   RS!dpgashelium = txthelium(i).Text
   tempnitrogen = 100 - Val(txthelium(i).Text) - Val(txtoxygen(i).Text)
   RS!dpgasnitrogen = tempnitrogen
   RS!dpgasmaxopdepth = Val(txtmaxd(i).Text) * 10
   RS!dpgaspo2setpoint = txtppo2(i).Text
   RS!dpgasused = Cbogasused(i).Text
   RS.Update
Next i
SQL = "SELECT * FROM seqdpmain"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!diveplanid = tempserialno
RS.Update
End Sub
Private Sub savecurrentrecord()
SQL = "SELECT * FROM dpmaingaslist "
SQL = SQL & " where dpmainid = '" & tempserialno & "' "
SQL = SQL & " order by dpgasid "
Set RS = DB.OpenRecordset(SQL)
RS.MoveFirst
   For i = 0 To 9
   RS.Edit
   RS!dpgashelium = txthelium(i).Text
   tempnitrogen = 100 - Val(txthelium(i).Text) - Val(txtoxygen(i).Text)
   RS!dpgasnitrogen = tempnitrogen
   RS!dpgasmaxopdepth = Val(txtmaxd(i).Text) * 10
   RS!dpgaspo2setpoint = txtppo2(i).Text
   RS!dpgasused = Cbogasused(i).Text
   RS.Update
   RS.MoveNext
 Next i
 RS.Close
End Sub
Private Sub Cmdplanprofile_Click()
Select Case tempchoice
   Case "GP"
     Planprofile.Show
   Case "NP"
     Planprofile.Show
   Case "NSP"
     Planprofile2.Show
   Case "GSP"
     Planprofile2.Show
End Select
End Sub

Private Sub cmdreset_Click()
ans = MsgBox("Do you really want to load the factory default setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
    SQL = "SELECT * FROM dpfacgasdefault "
    SQL = SQL & " order by gasid"
    Set RS = DB.OpenRecordset(SQL)
    RS.MoveFirst
    i = 0
    While RS.EOF = False
       If Val(i) < 10 Then
          p = i
          gasindex(i).Caption = RS("gasid")
          tempnitrogen = RS("gasnitrogen")
          txthelium(i) = RS("gashelium")
          txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
          txtmaxd(i) = Val(RS("gasmaxopdepth")) / 10
         ' txtmaxd(i) = txtmaxd(i) / 10
          temptxthe1 = txthe1
          temptxtmaxd1 = txtmaxd(i)
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
          txtppo2(i).Enabled = True
          txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
          txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
          If Cbogasused(i).Text <> "5 - Deco Closed Circuit" Then
             txtppo2(i).Enabled = False
          End If
      End If
  i = i + 1
  RS.MoveNext
  Wend
  MsgBox "All value reset to factory default."
Case Else
   'MsgBox "Request cancelled. "
End Select
End Sub

Private Sub cmdsave_Click()
gasvalidate
If gasvalidation = "failed" Then
   MsgBox "Data cannot be blank, please fill in data !"
Else
  savecurrentrecord
  Select Case tempchoice
  Case "GP"
    Unload Me
    Planprofile.Show
  Case "GSP"
    Unload Me
    Splanmain.Show
  Case "NP"
    Unload Me
    Planprofile.Show
  Case "NSP"
    Unload Me
    Planprofile2.Show
  End Select
End If
End Sub

Private Sub cmdseperate_Click()
 Unload Me
End Sub



Private Sub Cmdsetdefault_Click()
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
          RS!gasid = gasindex(i).Caption
          RS!gashelium = txthelium(i).Text
          tempnitrogen = 100 - Val(txthelium(i).Text) - Val(txtoxygen(i).Text)
          RS!gasnitrogen = tempnitrogen
          RS!gasmaxopdepth = Val(txtmaxd(i).Text) * 10
          RS!gasused = Cbogasused(i).Text
          RS.Update
       Next i
Case Else
   'MsgBox "Request cancelled. "
End Select
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 30
If Trim(tempserialno) <> "" Then
   oldserialno = tempserialno
End If
Label8.Visible = False
lblsediveno.Visible = False
For i = 0 To 9
  Cbogasused(i).AddItem "0 - Not Used"
  Cbogasused(i).AddItem "1 - Open Circuit"
  Cbogasused(i).AddItem "2 - Closed Circuit"
  Cbogasused(i).AddItem "3 - Open & Closed"
  Cbogasused(i).AddItem "4 - Deco Open Circuit"
  Cbogasused(i).AddItem "5 - Deco Closed Circuit"
Next i
Select Case tempchoice
Case "GP"
  cmdSave.Visible = True
  Cmdcreate.Visible = False
  lblserialno.Caption = "  " & oldserialno
  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & "where "
  SQL = SQL & " dpmainid = '" & tempserialno & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  For i = 0 To 9
     gasindex(i).Caption = RS("dpgasid")
     tempnitrogen = RS("dpgasnitrogen")
     txthelium(i) = RS("dpgashelium")
     txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
     txtmaxd(i) = RS("dpgasmaxopdepth")
     txtmaxd(i) = txtmaxd(i) / 10
     temptxthe1 = txthe1
     temptxtmaxd1 = txtmaxd(i)
     Cbogasused(i) = RS("dpgasused")
     txtppo2(i).Enabled = True
     txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
     txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     If Cbogasused(i).Text <> "5 - Deco Closed Circuit" Then
        txtppo2(i).Enabled = False
     End If
     RS.MoveNext
  Next i
Case "GSP"
  Label8.Visible = True
  lblsediveno.Visible = True
  cmdSave.Visible = True
  Cmdcreate.Visible = False
  lblserialno.Caption = "  " & oldserialno
  'lblsediveno.Caption = "   " & tempdiveserialno
  SQL = "SELECT * FROM dpmaingaslist "
  SQL = SQL & "where "
  SQL = SQL & " dpmainid = '" & tempserialno & "' "
  SQL = SQL & " order by dpgasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  For i = 0 To 9
     gasindex(i).Caption = RS("dpgasid")
     tempnitrogen = RS("dpgasnitrogen")
     txthelium(i) = RS("dpgashelium")
     txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
     txtmaxd(i) = RS("dpgasmaxopdepth")
     txtmaxd(i) = txtmaxd(i) / 10
     temptxthe1 = txthe1
     temptxtmaxd1 = txtmaxd(i)
     Cbogasused(i) = RS("dpgasused")
     txtppo2(i).Enabled = True
     txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
     txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     If Cbogasused(i).Text <> "5 - Deco Closed Circuit" Then
        txtppo2(i).Enabled = False
     Else
        txtppo2(i).Text = RS("dpgaspo2setpoint")
        txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     End If
     RS.MoveNext
  Next i
Case "NP"
  cmdSave.Visible = False
  Cmdcreate.Visible = True
  frmgasprofile2.Caption = "Dive Plan : Gas Profile - System Default"
  SQL = "SELECT * FROM dpserialno "
  Set RS = DB.OpenRecordset(SQL)
  tempserialno2 = RS("dplanserialno")
  tempserialno = Right(tempserialno2, 8)
  newserialno = Val(tempserialno) + 1
  tempserialno = Val(tempserialno) + 1
  lengthsn = Len(tempserialno)
  Select Case lengthsn
  Case 1
     tempserialno = "TP0000000" & tempserialno
     newserialno = "DP0000000" & newserialno
  Case 2
     tempserialno = "TP000000" & tempserialno
     newserialno = "DP000000" & newserialno
  Case 3
     tempserialno = "TP00000" & tempserialno
     newserialno = "DP00000" & newserialno
  Case 4
     tempserialno = "TP0000" & tempserialno
     newserialno = "DP0000" & newserialno
  Case 5
     tempserialno = "TP000" & tempserialno
     newserialno = "DP000" & newserialno
  Case 6
     tempserialno = "TP00" & tempserialno
     newserialno = "DP00" & newserialno
  Case 7
     tempserialno = "TP0" & tempserialno
     newserialno = "DP0" & newserialno
  Case 8
     tempserialno = "TP" & tempserialno
     newserialno = "DP" & newserialno
 End Select
  lblserialno.Caption = "  " & newserialno
  SQL = "SELECT * FROM dpgasdefault "
  SQL = SQL & " order by gasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  For i = 0 To 9
     gasindex(i).Caption = RS("gasid")
     tempnitrogen = RS("gasnitrogen")
     txthelium(i) = RS("gashelium")
     txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
     txtmaxd(i) = RS("gasmaxopdepth")
     txtmaxd(i) = txtmaxd(i) / 10
     temptxthe1 = txthe1
     temptxtmaxd1 = txtmaxd1
     Cbogasused(i) = RS("gasused")
     txtppo2(i).Enabled = True
     txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
     txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     If Cbogasused(i).Text <> "5 - Deco Closed Circuit" Then
        txtppo2(i).Enabled = False
     End If
     RS.MoveNext
  Next i
Case "NSP"
  Label8.Visible = True
  lblsediveno.Visible = True
  cmdSave.Visible = False
  Cmdcreate.Visible = True
  frmgasprofile2.Caption = "Dive Plan : Gas Profile - System Default"
  SQL = "SELECT * FROM dpserialno "
  Set RS = DB.OpenRecordset(SQL)
  tempserialno2 = RS("lastseqdserialno")
  tempserialno = Right(tempserialno2, 8)
  newserialno = Val(tempserialno) + 1
  tempserialno = Val(tempserialno) + 1
  lengthsn = Len(tempserialno)
  Select Case lengthsn
  Case 1
     tempserialno = "TP0000000" & tempserialno
     newserialno = "SP0000000" & newserialno
  Case 2
     tempserialno = "TP000000" & tempserialno
     newserialno = "SP000000" & newserialno
  Case 3
     tempserialno = "TP00000" & tempserialno
     newserialno = "SP00000" & newserialno
  Case 4
     tempserialno = "TP0000" & tempserialno
     newserialno = "SP0000" & newserialno
  Case 5
     tempserialno = "TP000" & tempserialno
     newserialno = "SP000" & newserialno
  Case 6
     tempserialno = "TP00" & tempserialno
     newserialno = "SP00" & newserialno
  Case 7
     tempserialno = "TP0" & tempserialno
     newserialno = "SP0" & newserialno
  Case 8
     tempserialno = "TP" & tempserialno
     newserialno = "SP" & newserialno
 End Select
 lblserialno = "   " & newserialno
 lblsediveno.Visible = False
 Label8.Visible = False
 Case "NPP"
  'Unload Planprofile2
  Label8.Visible = True
  lblsediveno.Visible = True
  cmdSave.Visible = False
  Cmdcreate.Visible = True
  frmgasprofile2.Caption = "Dive Plan : Gas Profile - System Default"
  SQL = "SELECT * FROM dpserialno "
  Set RS = DB.OpenRecordset(SQL)
  If IsNull(RS("lastseqdserialno")) Then
    tempserialno2 = "00000001"
  Else
    tempserialno2 = RS("lastseqdserialno")
  End If
  tempserialno = Right(tempserialno2, 8)
  newserialno = Val(tempserialno) + 1
  tempserialno = Val(tempserialno) + 1
  'newserialno = "1" '"000000001" 'Val(tempserialno) + 1
  'tempserialno = "0" '"000000002" 'Val(tempserialno) + 1
  lengthsn = Len(tempserialno)
  Select Case lengthsn
  Case 1
     tempserialno = "TP0000000" & tempserialno
     newserialno = "SP0000000" & newserialno
  Case 2
     tempserialno = "TP000000" & tempserialno
     newserialno = "SP000000" & newserialno
  Case 3
     tempserialno = "TP00000" & tempserialno
     newserialno = "SP00000" & newserialno
  Case 4
     tempserialno = "TP0000" & tempserialno
     newserialno = "SP0000" & newserialno
  Case 5
     tempserialno = "TP000" & tempserialno
     newserialno = "SP000" & newserialno
  Case 6
     tempserialno = "TP00" & tempserialno
     newserialno = "SP00" & newserialno
  Case 7
     tempserialno = "TP0" & tempserialno
     newserialno = "SP0" & newserialno
  Case 8
     tempserialno = "TP" & tempserialno
     newserialno = "SP" & newserialno
 End Select
 lblserialno = "   " & newserialno
 lblsediveno.Visible = False
 Label8.Visible = False
  SQL = "SELECT * FROM dpgasdefault "
  SQL = SQL & " order by gasid "
  Set RS = DB.OpenRecordset(SQL)
  RS.MoveFirst
  For i = 0 To 9
     gasindex(i).Caption = RS("gasid")
     tempnitrogen = RS("gasnitrogen")
     txthelium(i) = RS("gashelium")
     txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
     txtmaxd(i) = RS("gasmaxopdepth")
     txtmaxd(i) = txtmaxd(i) / 10
     temptxthe1 = txthe1
     temptxtmaxd1 = txtmaxd1
     Cbogasused(i) = RS("gasused")
     txtppo2(i).Enabled = True
     txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
     txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
     If Cbogasused(i).Text <> "5 - Deco Closed Circuit" Then
        txtppo2(i).Enabled = False
     End If
     RS.MoveNext
  Next i
  Cmdcreate_Click
  End Select
  
End Sub
Function gasvalidate()
  gasvalidation = "Passed"
  For i = 0 To 9
    If txtoxygen(i).Text = "" Or txthelium(i).Text = "" Or txtmaxd(i).Text = "" Or txtppo2(i).Text = "" Or Cbogasused(i).Text = "" Then
       gasvalidation = "failed"
    End If
  Next i
End Function


Private Sub Form_Unload(Cancel As Integer)
  If tempchoice = "GSP" Then Splanmain.Show
End Sub

Private Sub txthelium_Change(index As Integer)

p = txthelium(index).index

lengthtxthelium = Len(txthelium(p))
  For K = 1 To lengthtxthelium '- 1
      If Asc(Mid$(txthelium(p), K, 1)) > 47 And Asc(Mid$(txthelium(p), K, 1)) < 59 Then
         tempcode = tempcode & Mid$(txthelium(p), K, 1)
      Else
         tempcode = tempcode
      End If
   Next
txthelium(p).Text = tempcode
SendKeys "{END}"
End Sub

Private Sub txthelium_KeyPress(index As Integer, KeyAscii As Integer)

p = txthelium(index).index
If KeyAscii = 13 Then
  p = txthelium(index).index
  validatehelium
Else
If KeyAscii = 8 Or (KeyAscii < 59 And KeyAscii > 47) Then
   
Else

   lengthtxthelium = Len(txthelium(p))
   For K = 1 To lengthtxthelium '- 1
      tempcode = tempcode & Mid$(txthelium(p), K, 1)
   Next
   txthelium(p).Text = tempcode
End If
End If
txthelium(p).SetFocus


End Sub
Private Sub txthelium_LostFocus(index As Integer)
p = txthelium(index).index
  validatehelium
End Sub

Private Sub txtmaxd_Change(index As Integer)
 lengthtxtmaxd = Len(txtmaxd(p))
'MsgBox txtmaxd(p)
  For K = 1 To lengthtxtmaxd '- 1
    'MsgBox Asc(Mid$(txtmaxd(p), K, 1))
      If Asc(Mid$(txtmaxd(p), K, 1)) > 45 And Asc(Mid$(txtmaxd(p), K, 1)) < 59 Then
         tempcode = tempcode & Mid$(txtmaxd(p), K, 1)
      Else
         tempcode = tempcode
      End If
   Next
txtmaxd(p).Text = tempcode
SendKeys "{END}"
End Sub

Private Sub txtmaxd_KeyPress(index As Integer, KeyAscii As Integer)
p = txtmaxd(index).index
If KeyAscii = 13 Then
  p = txtmaxd(index).index
  validatemaxdepth
  Else
If KeyAscii = 8 Or (KeyAscii < 59 And KeyAscii > 45) Then
   
Else

   lengthtxtmaxd = Len(txtmaxd(p))
   For K = 1 To lengthtxtmaxd '- 1
      tempcode = tempcode & Mid$(txtmaxd(p), K, 1)
   Next
   txtmaxd(p).Text = tempcode
End If
End If
txtmaxd(p).SetFocus
End Sub

Private Sub txtmaxd_LostFocus(index As Integer)
  p = txtmaxd(index).index
  validatemaxdepth
End Sub

Private Sub txtoxygen_Change(index As Integer)
lengthtxtoxygen = Len(txtoxygen(p))
  For K = 1 To lengthtxtoxygen '- 1
      If Asc(Mid$(txtoxygen(p), K, 1)) > 47 And Asc(Mid$(txtoxygen(p), K, 1)) < 59 Then
         tempcode = tempcode & Mid$(txtoxygen(p), K, 1)
      Else
         tempcode = tempcode
      End If
   Next
txtoxygen(p).Text = tempcode
SendKeys "{END}"
End Sub

Private Sub txtoxygen_KeyPress(index As Integer, KeyAscii As Integer)
p = txtoxygen(index).index
If KeyAscii = 13 Then
  p = txtoxygen(index).index
  validateoxygen
Else
If KeyAscii = 8 Or (KeyAscii < 59 And KeyAscii > 47) Then
   
Else

   lengthtxtoxygen = Len(txtoxygen(p))
   For K = 1 To lengthtxtoxygen '- 1
      tempcode = tempcode & Mid$(txtoxygen(p), K, 1)
   Next
   txtoxygen(p).Text = tempcode
End If
End If
txtoxygen(p).SetFocus
End Sub
Private Sub validateoxygen()
   If Val(txtoxygen(p).Text) >= 0 And Val(txtoxygen(p).Text) <= 100 And ((Val(txtoxygen(p).Text) + Val(txthelium(p).Text)) < 101) Then
     'txtppo2(p).Enabled = True
     temptextpo2 = (Val(txtoxygen(p).Text) / 100) * ((Val(txtmaxd(p).Text) / 10) + 1)
     temptextpo2 = Format(temptextpo2, "###.00")
     If Cbogasused(p).Text = "5 - Deco Closed Circuit" And (Val(temptextpo2) <> Val(txtppo2(p).Text)) Then
        ans = MsgBox("Do you really want to load the factory default setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
        Select Case ans
           Case vbYes
              txtppo2(p).Enabled = True
              txtppo2(p).Text = temptextpo2
           Case Else
              MsgBox "PPO2 value not replace "
        End Select
     Else
        txtppo2(p).Enabled = True
        txtppo2(p).Text = temptextpo2
        txtppo2(p).Enabled = False
     End If
   Else
      Title = "Error on System Validation.."
      MsgBox "(oxygen " & p & " value can not be less then 0 or more then 100) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
      txtoxygen(p).SetFocus
      SendKeys "{HOME}+{END}"
   End If
End Sub
Private Sub validatemaxdepth()
   If Val(txtmaxd(p).Text) >= 0 And Val(txtmaxd(p).Text) <= 1000 Then
     temptextpo2 = (Val(txtoxygen(p).Text) / 100) * ((Val(txtmaxd(p).Text) / 10) + 1)
     temptextpo2 = Format(temptextpo2, "###.00")
     If Cbogasused(p).Text = "5 - Deco Closed Circuit" And (Val(temptextpo2) <> Val(txtppo2(p).Text)) Then
        ans = MsgBox("Do you really want to load the factory default setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
        Select Case ans
           Case vbYes
              txtppo2(p).Enabled = True
              txtppo2(p).Text = temptextpo2
           Case Else
              MsgBox "PPO2 value not replace "
        End Select
     Else
        txtppo2(p).Enabled = True
        txtppo2(p).Text = temptextpo2
        txtppo2(p).Enabled = False
     End If
   Else
      Title = "Error on System Validation.."
      MsgBox "(Max.Depth " & p & " value can not be less then 0 or more then 1000) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
      txtmaxd(p).SetFocus
      SendKeys "{HOME}+{END}"
   End If
End Sub
Private Sub validatehelium()
  If Val(txthelium(p).Text) >= 0 And Val(txthelium(p).Text) <= 100 And ((Val(txtoxygen(p).Text) + Val(txthelium(p).Text)) < 101) Then
     temptextpo2 = (Val(txtoxygen(p).Text) / 100) * ((Val(txtmaxd(p).Text) / 10) + 1)
     temptextpo2 = Format(temptextpo2, "###.00")
     If Cbogasused(p).Text = "5 - Deco Closed Circuit" And (Val(temptextpo2) <> Val(txtppo2(p).Text)) Then
        ans = MsgBox("Do you really want to load the factory default setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
        Select Case ans
           Case vbYes
              txtppo2(p).Enabled = True
              txtppo2(p).Text = temptextpo2
           Case Else
              MsgBox "PPO2 value not replace "
        End Select
     Else
        txtppo2(p).Enabled = True
        txtppo2(p).Text = temptextpo2
        txtppo2(p).Enabled = False
     End If
   Else
      Title = "Error on System Validation.."
      MsgBox "(Helium " & p & " value can not be less then 0 or more then 100) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
       txthelium(p).SetFocus
      SendKeys "{HOME}+{END}"
   End If
End Sub

Private Sub txtoxygen_LostFocus(index As Integer)
 p = txtoxygen(index).index
 validateoxygen
 
End Sub

Private Sub validatepo2()
  If Val(txtppo2(p).Text) <= 0.14 Then
    MsgBox "PO2 value out of range "
    txtppo2(p).Enabled = True
    txtppo2(p).SetFocus
  Else
    If Val(txtppo2(p).Text) > 2 Then
       MsgBox "PO2 value out of range "
       txtppo2(p).Enabled = True
       txtppo2(p).SetFocus
       SendKeys "{HOME}+{END}"
    End If
  End If
End Sub


Private Sub txtppo2_LostFocus(index As Integer)
p = txtppo2(index).index
validatepo2
End Sub
