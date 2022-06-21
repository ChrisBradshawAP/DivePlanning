VERSION 5.00
Begin VB.Form frmgasprofile2 
   BackColor       =   &H80000013&
   Caption         =   "Dive Plan : Gas Profile"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   ForeColor       =   &H00000000&
   Icon            =   "frmgasprofile2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   9615
      Begin VB.CommandButton Cmdplanprofile 
         Caption         =   "Plan Profile"
         Height          =   375
         Left            =   1680
         TabIndex        =   76
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton Cmdcancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3000
         TabIndex        =   75
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         Height          =   375
         Left            =   360
         TabIndex        =   74
         Top             =   4440
         Width           =   1095
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   9
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   3960
         Width           =   2975
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   8
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   3600
         Width           =   2975
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   7
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   3240
         Width           =   2975
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   6
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   2880
         Width           =   2975
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   5
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2520
         Width           =   2975
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   4
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   2160
         Width           =   2975
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   3
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   1800
         Width           =   2975
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   2
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1440
         Width           =   2975
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   1
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   1080
         Width           =   2975
      End
      Begin VB.ComboBox cbogasused 
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
         Index           =   0
         Left            =   6300
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   720
         Width           =   2975
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   9
         Left            =   5170
         TabIndex        =   63
         Top             =   3960
         Width           =   1105
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   8
         Left            =   5170
         TabIndex        =   62
         Top             =   3600
         Width           =   1105
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   5170
         TabIndex        =   61
         Top             =   3240
         Width           =   1105
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   5170
         TabIndex        =   60
         Top             =   2880
         Width           =   1105
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   5170
         TabIndex        =   59
         Top             =   2520
         Width           =   1105
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   5170
         TabIndex        =   58
         Top             =   2160
         Width           =   1105
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   5170
         TabIndex        =   57
         Top             =   1800
         Width           =   1105
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   5170
         TabIndex        =   56
         Top             =   1440
         Width           =   1105
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   5170
         TabIndex        =   55
         Top             =   1080
         Width           =   1105
      End
      Begin VB.TextBox txtppo2 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   5170
         TabIndex        =   54
         Top             =   720
         Width           =   1105
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   9
         Left            =   3850
         TabIndex        =   53
         Top             =   3960
         Width           =   1300
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   8
         Left            =   3850
         TabIndex        =   52
         Top             =   3600
         Width           =   1300
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   3850
         TabIndex        =   51
         Top             =   3240
         Width           =   1300
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   3850
         TabIndex        =   50
         Top             =   2880
         Width           =   1300
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   3850
         TabIndex        =   49
         Top             =   2520
         Width           =   1300
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   3850
         TabIndex        =   48
         Top             =   2160
         Width           =   1300
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   3850
         TabIndex        =   47
         Top             =   1800
         Width           =   1300
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   3850
         TabIndex        =   46
         Top             =   1440
         Width           =   1300
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   3850
         TabIndex        =   45
         Top             =   1080
         Width           =   1300
      End
      Begin VB.TextBox txtmaxd 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   3850
         TabIndex        =   44
         Top             =   720
         Width           =   1300
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   9
         Left            =   2640
         TabIndex        =   43
         Top             =   3960
         Width           =   1200
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   8
         Left            =   2640
         TabIndex        =   42
         Top             =   3600
         Width           =   1200
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   2640
         TabIndex        =   41
         Top             =   3240
         Width           =   1200
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   2640
         TabIndex        =   40
         Top             =   2880
         Width           =   1200
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   2640
         TabIndex        =   39
         Top             =   2520
         Width           =   1200
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   2640
         TabIndex        =   38
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   2640
         TabIndex        =   37
         Top             =   1800
         Width           =   1200
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   2640
         TabIndex        =   36
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   2640
         TabIndex        =   35
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox txthelium 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   2640
         TabIndex        =   34
         Top             =   720
         Width           =   1200
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   9
         Left            =   1540
         TabIndex        =   33
         Top             =   3960
         Width           =   1080
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   8
         Left            =   1540
         TabIndex        =   32
         Top             =   3600
         Width           =   1080
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   1540
         TabIndex        =   31
         Top             =   3240
         Width           =   1080
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   1540
         TabIndex        =   30
         Top             =   2880
         Width           =   1080
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   1540
         TabIndex        =   29
         Top             =   2520
         Width           =   1080
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   1540
         TabIndex        =   28
         Top             =   2160
         Width           =   1080
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   1540
         TabIndex        =   27
         Top             =   1800
         Width           =   1080
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   1540
         TabIndex        =   26
         Top             =   1440
         Width           =   1080
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   1540
         TabIndex        =   25
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox txtoxygen 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   1540
         TabIndex        =   24
         Top             =   720
         Width           =   1080
      End
      Begin VB.CommandButton Cmdcreate 
         Caption         =   "Create"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton Cmdsetdefault 
         Caption         =   "Set as Default"
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton cmdreset 
         Caption         =   "Reset Factory"
         Height          =   375
         Left            =   4320
         TabIndex        =   5
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 9"
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
         Index           =   9
         Left            =   320
         TabIndex        =   23
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 8"
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
         Index           =   8
         Left            =   320
         TabIndex        =   22
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 7"
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
         Index           =   7
         Left            =   320
         TabIndex        =   21
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 6"
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
         Index           =   6
         Left            =   320
         TabIndex        =   20
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 5"
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
         Index           =   5
         Left            =   320
         TabIndex        =   19
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 4"
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
         Index           =   4
         Left            =   320
         TabIndex        =   18
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 3"
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
         Index           =   3
         Left            =   320
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 2"
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
         Index           =   2
         Left            =   320
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 1"
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
         Index           =   1
         Left            =   320
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label gasindex 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas 0"
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
         Index           =   0
         Left            =   320
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Oxygen"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1530
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Helium"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Max. Depth"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   3855
         TabIndex        =   11
         Top             =   360
         Width           =   1300
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PPO2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Width           =   1145
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas Used"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   6295
         TabIndex        =   9
         Top             =   360
         Width           =   2975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gas Index"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   320
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox txtmaxd2 
      Alignment       =   2  'Center
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
      Left            =   3000
      TabIndex        =   3
      Top             =   2090
      Width           =   1030
   End
   Begin VB.TextBox txthe2 
      Alignment       =   2  'Center
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
      Left            =   1935
      TabIndex        =   2
      Top             =   2090
      Width           =   1030
   End
   Begin VB.Label lblserialno 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   450
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gas Profile "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   450
      Left            =   240
      TabIndex        =   0
      Top             =   360
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


Private Sub cmdcancel_Click()
Unload Me
End Sub



Private Sub Cmdcreate_Click()
gasvalidate
If gasvalidation = "failed" Then
   MsgBox "Data cannot be blank, please fill in data !"
Else
   savenewrecord
   Planprofile.Show
End If
End Sub
Private Sub savenewrecord()
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
   RS!dpgasused = Cbogasused(i).Text
   RS.Update
Next i
SQL = "SELECT * FROM dpmain"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!diveplanid = tempserialno
RS.Update
End Sub
Private Sub savecurrentrecord()
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
   RS!dpgasused = Cbogasused(i).Text
   RS.Update
Next i
SQL = "SELECT * FROM dpmain"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!diveplanid = tempserialno
RS.Update
End Sub
Private Sub Cmdplanprofile_Click()
  Planprofile.Show
End Sub

Private Sub cmdreset_Click()
ans = MsgBox("Do you really want to load the factory default setting ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
    SQL = "SELECT * FROM dpfacgasdefault"
    Set RS = DB.OpenRecordset(SQL)
    RS.MoveFirst
    i = 0
    While RS.EOF = False
       If Val(i) < 10 Then
          gasindex(i).Caption = RS("gasid")
          tempnitrogen = RS("gasnitrogen")
          txthelium(i) = RS("gashelium")
          txtoxygen(i).Text = 100 - Val(txthelium(i).Text) - Val(tempnitrogen)
          txtmaxd(i) = RS("gasmaxopdepth")
          txtmaxd(i) = txtmaxd(i) / 10
          temptxthe1 = txthe1
          temptxtmaxd1 = txtmaxd(i)
          Cbogasused(i) = RS("gasused")
          txtppo2(i).Enabled = True
          txtppo2(i).Text = (Val(txtoxygen(i).Text) / 100) * ((Val(txtmaxd(i).Text) / 10) + 1)
          txtppo2(i).Text = Format(txtppo2(i).Text, "###.00")
          txtppo2(i).Enabled = False
      End If
  i = i + 1
  RS.MoveNext
  Wend
  MsgBox "All value reset to factory default."
Case Else
   MsgBox "Request cancelled. "
End Select
End Sub

Private Sub cmdsave_Click()
gasvalidate
If gasvalidation = "failed" Then
   MsgBox "Data cannot be blank, please fill in data !"
Else
  tempserialno = Right(tempserialno, 8)
  tempserialno = Val(tempserialno) + 1
  lengthsn = Len(tempserialno)
  Select Case lengthsn
  Case 1
     tempserialno = "TP0000000" & tempserialno
  Case 2
     tempserialno = "TP000000" & tempserialno
  Case 3
     tempserialno = "TP00000" & tempserialno
  Case 4
     tempserialno = "TP0000" & tempserialno
  Case 5
     tempserialno = "TP000" & tempserialno
  Case 6
     tempserialno = "TP00" & tempserialno
  Case 7
     tempserialno = "TP0" & tempserialno
  Case 8
     tempserialno = "TP" & tempserialno
  End Select
  savecurrentrecord
  Planprofile.Show
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
   MsgBox "Request cancelled. "
End Select
End Sub

Private Sub Form_Load()
If Trim(tempserialno) <> "" Then
   oldserialno = tempserialno
End If
Me.Top = 50
Me.Left = (Screen.Width - Me.Width) / 2
For i = 0 To 9
  Cbogasused(i).AddItem "0 - Not Used"
  Cbogasused(i).AddItem "1 - Open Circuit"
  Cbogasused(i).AddItem "2 - Closed Circuit"
  Cbogasused(i).AddItem "3 - Open & Closed"
Next i
Select Case tempchoice
Case "GP"
  cmdSave.Visible = True
  Cmdcreate.Visible = False
  lblserialno.Caption = "  " & oldpserialno
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
     txtppo2(i).Enabled = False
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
     txtppo2(i).Enabled = False
     RS.MoveNext
  Next i
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
Private Sub validateoxygen()
   If Val(txtoxygen(p).Text) >= 0 And Val(txtoxygen(p).Text) <= 100 And ((Val(txtoxygen(p).Text) + Val(txthelium(p).Text)) < 101) Then
     txtppo2(p).Enabled = True
     txtppo2(p).Text = (Val(txtoxygen(p).Text) / 100) * ((Val(txtmaxd(p).Text) / 10) + 1)
     txtppo2(p).Text = Format(txtppo2(p).Text, "###.00")
     txtppo2(p).Enabled = False
   Else
      Title = "Error on System Validation.."
      MsgBox "(oxygen " & p & " value can not be less then 0 or more then 100) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
      txtoxygen(p).SetFocus
      SendKeys "{HOME}+{END}"
   End If
End Sub
Private Sub validatemaxdepth()
   If Val(txtmaxd(p).Text) >= 0 And Val(txtmaxd(p).Text) <= 100 And ((Val(txtoxygen(p).Text) + Val(txthelium(p).Text)) < 101) Then
     txtppo2(p).Enabled = True
     txtppo2(p).Text = (Val(txtoxygen(p).Text) / 100) * ((Val(txtmaxd(p).Text) / 10) + 1)
     txtppo2(p).Text = Format(txtppo2(p).Text, "###.00")
     txtppo2(p).Enabled = False
   Else
      Title = "Error on System Validation.."
      MsgBox "(Max.Depth " & p & " value can not be less then 0 or more then 100) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
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
   Else
      Title = "Error on System Validation.."
      MsgBox "(Helium " & p & " value can not be less then 0 or more then 100) or " & Chr(13) & "(Oxygen + Helium value can not more then 100)", 48, Title
       txthelium(p).SetFocus
      SendKeys "{HOME}+{END}"
   End If
End Sub

Private Sub txtoxygen_LostFocus(Index As Integer)
 p = txtoxygen(Index).Index
 validateoxygen
 
End Sub
