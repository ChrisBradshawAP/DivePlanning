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
      Index           =   2
      Left            =   480
      TabIndex        =   25
      Top             =   0
      Width           =   11775
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   6795
         Index           =   2
         Left            =   1920
         Picture         =   "Intro2.frx":0000
         ScaleHeight     =   6795
         ScaleWidth      =   8115
         TabIndex        =   28
         Top             =   1800
         Width           =   8115
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   6
            X1              =   0
            X2              =   1680
            Y1              =   960
            Y2              =   480
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   5
            X1              =   0
            X2              =   3240
            Y1              =   4920
            Y2              =   5400
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   4
            X1              =   6120
            X2              =   8160
            Y1              =   5040
            Y2              =   3240
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   3
            X1              =   5880
            X2              =   8160
            Y1              =   2520
            Y2              =   2160
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   2
            X1              =   8040
            X2              =   4680
            Y1              =   480
            Y2              =   2520
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   1
            X1              =   0
            X2              =   1680
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   0
            X1              =   720
            X2              =   840
            Y1              =   5640
            Y2              =   6140
         End
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Close"
         Height          =   375
         Index           =   2
         Left            =   8040
         TabIndex        =   27
         Top             =   8640
         Width           =   1335
      End
      Begin VB.CommandButton cmdnext7 
         Caption         =   "Next"
         Height          =   375
         Index           =   2
         Left            =   9720
         TabIndex        =   26
         Top             =   8640
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pull down menu to change units to feet or metrers"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   8
         Left            =   240
         TabIndex        =   39
         Top             =   7320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pull down menu to change units to feet or metrers"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   38
         Top             =   7200
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pull down menu to change units to feet or metrers"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   6
         Left            =   360
         TabIndex        =   37
         Top             =   7440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pull down menu to change units to feet or metrers"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   36
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Information on Highlighted Dive"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4. Double Click dive in list to edit"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   3
         Left            =   10200
         TabIndex        =   34
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3. Click to Delete Highlighted Dive"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   10200
         TabIndex        =   33
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2. Click to Edit Highlighted Dive"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   10080
         TabIndex        =   32
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "VGM Dive Plan PC software"
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
         Index           =   2
         Left            =   600
         TabIndex        =   31
         Top             =   120
         Width           =   8535
      End
      Begin VB.Label Label137 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":B7142
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
         Index           =   2
         Left            =   480
         TabIndex        =   30
         Top             =   600
         Width           =   11055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1. Click to Plan New Dive"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   4080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "main"
      Height          =   9015
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton cmdnext7 
         Caption         =   "Next"
         Height          =   375
         Index           =   1
         Left            =   9720
         TabIndex        =   13
         Top             =   8640
         Width           =   1215
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Close"
         Height          =   375
         Index           =   1
         Left            =   8040
         TabIndex        =   12
         Top             =   8640
         Width           =   1335
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   4035
         Index           =   1
         Left            =   840
         Picture         =   "Intro2.frx":B733C
         ScaleHeight     =   4035
         ScaleWidth      =   9465
         TabIndex        =   11
         Top             =   3240
         Width           =   9465
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   8
            X1              =   8400
            X2              =   7800
            Y1              =   3600
            Y2              =   4320
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   7
            X1              =   5160
            X2              =   5040
            Y1              =   3500
            Y2              =   4080
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   6
            X1              =   9360
            X2              =   8760
            Y1              =   4080
            Y2              =   2860
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   5
            X1              =   7560
            X2              =   5520
            Y1              =   0
            Y2              =   1200
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   4
            X1              =   5880
            X2              =   3360
            Y1              =   0
            Y2              =   600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   3
            X1              =   3960
            X2              =   4275
            Y1              =   0
            Y2              =   560
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   1
            X1              =   240
            X2              =   360
            Y1              =   0
            Y2              =   500
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   0
            X1              =   480
            X2              =   0
            Y1              =   3720
            Y2              =   4080
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   2
            X1              =   2160
            X2              =   2640
            Y1              =   0
            Y2              =   600
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7. If required, Adjust VGM safety parameters to tweak decompression."
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   8
         Left            =   8040
         TabIndex        =   24
         Top             =   7275
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7. If required, Adjust VGM safety parameters to tweak decompression."
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   7
         Left            =   5160
         TabIndex        =   23
         Top             =   7275
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6. Click Calculate Deco"
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   6
         Left            =   9840
         TabIndex        =   22
         Top             =   7275
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5. Click on bottom gas"
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   5
         Left            =   7680
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4. Choose which gasses are used in deco only. Click Deco box"
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   4
         Left            =   5880
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3. Adjust oxygen and helium percentages of all enabled gasses."
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   3
         Left            =   4080
         TabIndex        =   19
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2. Choose Closed circuit option if planning a rebreather dive"
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   2
         Left            =   2160
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1. Enable gasses for dive. Click Enable box"
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "You have now planned a dive using DivePlan VGM"
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   0
         Left            =   600
         TabIndex        =   16
         Top             =   7280
         Width           =   1575
      End
      Begin VB.Label Label137 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":133BC6
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
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   600
         Width           =   11055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "VGM Dive Plan PC software"
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
         Index           =   1
         Left            =   600
         TabIndex        =   14
         Top             =   120
         Width           =   8535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "main"
      Height          =   9015
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   3075
         Index           =   0
         Left            =   480
         Picture         =   "Intro2.frx":133DC0
         ScaleHeight     =   3075
         ScaleWidth      =   3945
         TabIndex        =   4
         Top             =   2640
         Width           =   3945
      End
      Begin VB.PictureBox Picture6 
         Height          =   3135
         Index           =   0
         Left            =   6240
         Picture         =   "Intro2.frx":15CAFA
         ScaleHeight     =   3075
         ScaleWidth      =   4395
         TabIndex        =   3
         Top             =   2640
         Width           =   4455
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Close"
         Height          =   375
         Index           =   0
         Left            =   8040
         TabIndex        =   2
         Top             =   8640
         Width           =   1335
      End
      Begin VB.CommandButton cmdnext7 
         Caption         =   "Next"
         Height          =   375
         Index           =   0
         Left            =   9720
         TabIndex        =   1
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
         Index           =   0
         Left            =   600
         TabIndex        =   9
         Top             =   8400
         Width           =   7455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   " Dive Plan  Editor - You can create and specify the parameters and settings for gas, depth, time and PPO2"
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
         Index           =   0
         Left            =   6240
         TabIndex        =   8
         Top             =   5880
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":18AE64
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
         Height          =   1335
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   5760
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "VGM Dive Plan PC software"
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
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   120
         Width           =   8535
      End
      Begin VB.Label Label137 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro2.frx":18AEEE
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
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   11055
      End
   End
End
Attribute VB_Name = "frmintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim linevisible As Integer


Private Sub viewframe1()
Frame1(0).Visible = True
Frame1(1).Visible = False
Frame1(2).Visible = False
End Sub
Private Sub viewframe2()
Frame1(1).Visible = True
Frame1(0).Visible = False
Frame1(2).Visible = False
End Sub
Private Sub viewframe3()
Frame1(2).Visible = True
Frame1(1).Visible = False
Frame1(0).Visible = False
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

Private Sub CMDCLOSE_Click(index As Integer)
  Unload Me
End Sub

Private Sub cmdnext7_Click(index As Integer)
  Line1(0).Visible = False
  Line1(1).Visible = False
  Line1(2).Visible = False
  Line1(3).Visible = False
  Line1(4).Visible = False
  Line1(5).Visible = False
  Line1(6).Visible = False
  Line1(7).Visible = False
  Line1(8).Visible = False
  Line2(0).Visible = False
  Line2(1).Visible = False
  Line2(2).Visible = False
  Line2(3).Visible = False
  Line2(4).Visible = False
  Line2(5).Visible = False
  Line2(6).Visible = False
'  Line2(7).Visible = False
'  Line2(8).Visible = False
  If index = 0 Then
    viewframe2
    linevisible = 1
    Line1(linevisible).Visible = True
  End If
  If index = 1 Then
    viewframe2
    linevisible = linevisible + 1
    If linevisible = 9 Then
      viewframe1
    Else
      If linevisible > 8 Then linevisible = 1
      Line1(linevisible).Visible = True
    End If
  End If
  If index = 2 Then
    viewframe3
    linevisible = linevisible + 1
    If linevisible = 7 Then
      viewframe1
    Else
      If linevisible > 6 Then linevisible = 1
      Line2(linevisible).Visible = True
    End If
  End If
End Sub

Private Sub Form_Load()
Top = 20
Left = 1300
Label3(1).Caption = "1. Enable gasses for dive. Click Enable box"

Label3(2).Caption = "2. Choose Closed circuit option if planning a rebreather dive"

Label3(3).Caption = "3. Adjust oxygen and helium percentages of all enabled gasses." & vbCrLf & "Click on Gas name to make active in the gas mix adjuster" & vbCrLf & "Click +- to adjust oxygen" & vbCrLf & "Click +- to adjust helium"

Label3(4).Caption = "4. Choose which gasses are used in deco only. Click Deco box." & vbCrLf & vbCrLf & "Unclick deco to edit mix."

Label3(5).Caption = "5. Click on bottom gas" & vbCrLf & "Adjust bottom depth" & vbCrLf & "Adjust bottom time" & vbCrLf & "If closed circuit dive, adjust PPO2"

Label3(6).Caption = "6. Click Calculate Deco" & vbCrLf & "Decompression schedule and dive profile are updated"

Label3(7).Caption = "7. If required, Adjust VGM safety parameters to tweak decompression."

Label3(8).Caption = "Equivalent Gradient Factor shows approximate gradient factors for comparisson"

Label3(0).Caption = "You have now planned a dive using DivePlan VGM"

viewframe1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Planprofile2.Visible = False Then
    Splanmain.Show
  End If
End Sub

Private Sub Label140_Click()
splmain.chow
End Sub

Private Sub Label10_Click()
viewframe1
End Sub

Private Sub Picture11_Click()
 viewframe1
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

Private Sub Picture5_Click(index As Integer)
  If index = 0 Then
    viewframe3
    linevisible = 1
    Line2(linevisible).Visible = True
  End If
  If index = 1 Then viewframe1
End Sub

Private Sub Picture6_Click(index As Integer)
  If index = 0 Then
    viewframe2
    linevisible = 1
    Line1(linevisible).Visible = True
  End If
  If index = 1 Then viewframe1
End Sub
