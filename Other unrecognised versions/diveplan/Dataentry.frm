VERSION 5.00
Begin VB.Form rbdetails 
   BackColor       =   &H80000013&
   Caption         =   "RB Dive details"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   Icon            =   "Dataentry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtsuit 
      Height          =   375
      Left            =   5160
      TabIndex        =   31
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Txtdiver 
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox TxtMaxdepth 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Txtduration 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtDivedate 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cbowhether 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.ComboBox Cbosite 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox Cbodepartment 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtdiveid 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dive Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3255
      Left            =   0
      TabIndex        =   21
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtlocation 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   37
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Start Date :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "Suit :"
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
         Left            =   4560
         TabIndex        =   30
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Weather :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Location :"
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
         Left            =   4200
         TabIndex        =   28
         Top             =   2680
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Site :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Department :"
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
         Left            =   3960
         TabIndex        =   26
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Max. Depth :"
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
         Left            =   240
         TabIndex        =   25
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Duration :"
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
         Left            =   480
         TabIndex        =   24
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Diver :"
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
         Left            =   720
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Dive ID :"
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
         Left            =   600
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ComboBox cbobuddy 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8760
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buddy"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3255
      Left            =   7680
      TabIndex        =   16
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton cmdremove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2640
         TabIndex        =   36
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
         Height          =   375
         Left            =   1440
         TabIndex        =   35
         Top             =   2760
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1815
         ItemData        =   "Dataentry.frx":030A
         Left            =   1080
         List            =   "Dataentry.frx":030C
         TabIndex        =   34
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Buddy :"
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
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Cmdclose 
      Caption         =   "Dive Profile"
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txtcomments 
      Height          =   3015
      Left            =   7920
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox txtsitenote 
      Height          =   3015
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox txtdivenote 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4080
      Width           =   3255
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9480
      TabIndex        =   4
      Text            =   "Combo4"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   9960
      TabIndex        =   0
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Note"
      Height          =   4215
      Left            =   0
      TabIndex        =   17
      Top             =   3450
      Width           =   11535
      Begin VB.Label Label11 
         Caption         =   "Comments :"
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
         Left            =   8040
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Site Note :"
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
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Dive Note :"
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
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "rbdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub cmdadd_Click()
List1.AddItem cbobuddy
End Sub

Private Sub cmdclose_Click()
Unload Me
rbinterface.Show
'Unload Me
'rbmain.Show
'For i = 0 To 3
' List1.ListIndex = i
'TEMP2 = List1.Text
'Next
End Sub

Private Sub cmdremove_Click()
 i = List1.ListIndex
 List1.RemoveItem i
End Sub

Private Sub cmdsave_Click()
SQL = "SELECT * FROM main "
SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
Set RS3 = DB.OpenRecordset(SQL)
RS3.Edit
RS3!UserName = Txtdiver
RS3!Department = Cbodepartment
RS3!site = Cbosite
RS3!Location = cbolocation
RS3!whether = cbowhether
RS3!suit = txtsuit
RS3!divenote = txtdivenote
RS3!sitenote = txtsitenote
RS3!Coments = txtcomments
RS3.Update

 For i = 0 To List1.ListCount - 1
   SQL = "SELECT * FROM buddylist"
   Set RS = DB.OpenRecordset(SQL)
   RS.AddNew
   List1.ListIndex = i
   templist = List1.Text
   RS!DiveID = tempserialno
   RS!buddylist = templist
   RS.Update
 Next
 MsgBox "close"
End Sub

Private Sub Form_Activate()
Txtdiver.SetFocus
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 30
SQL = "SELECT * FROM main "
SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
Set RS2 = DB.OpenRecordset(SQL)
txtdiveid = tempserialno
txtDivedate = RS2("startdate")
txtduration = RS2("duration")
txtmaxdepth = RS2("maxdepth")
txtlocation = RS2("location")
getdepartment
getsite
getwhether
getbuddy
End Sub

Private Sub getdepartment()
  SQL = "SELECT * FROM DEPARTMENT "
  Set RS2 = DB.OpenRecordset(SQL)
  If RS2.BOF And RS2.EOF Then
    Cbodepartment = ""
  Else
    RS2.MoveFirst
    While RS2.EOF = False
       templist = RS2("DEPARTMENT")
       Cbodepartment.AddItem templist
       RS2.MoveNext
    Wend
  End If
End Sub
Private Sub getsite()
  SQL = "SELECT * FROM site "
  Set RS2 = DB.OpenRecordset(SQL)
  If RS2.BOF And RS2.EOF Then
    Cbosite = ""
  Else
    RS2.MoveFirst
    While RS2.EOF = False
       templist = RS2("Site")
       Cbosite.AddItem templist
       RS2.MoveNext
    Wend
  End If
End Sub

Private Sub getwhether()
  SQL = "SELECT * FROM whether "
  Set RS2 = DB.OpenRecordset(SQL)
  If RS2.BOF And RS2.EOF Then
    cbowhether = ""
  Else
    RS2.MoveFirst
    While RS2.EOF = False
       templist = RS2("whether")
       cbowhether.AddItem templist
       RS2.MoveNext
    Wend
  End If
End Sub
Private Sub getbuddy()
  SQL = "SELECT * FROM buddy "
  Set RS2 = DB.OpenRecordset(SQL)
  If RS2.BOF And RS2.EOF Then
    cbobuddy = ""
  Else
    RS2.MoveFirst
    While RS2.EOF = False
       templist = RS2("buddy")
       cbobuddy.AddItem templist
       RS2.MoveNext
    Wend
  End If
End Sub

Private Sub Txtdiver_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  Txtdiver = UCase(Txtdiver)
  Txtdiver = Trim(Txtdiver)
  Cbodepartment.SetFocus
End If
End Sub

Private Sub Txtdiver_LostFocus()
    Txtdiver = UCase(Txtdiver)
    Txtdiver = Trim(Txtdiver)
    Cbodepartment.SetFocus
End Sub

Private Sub txtsuit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtsuit = UCase(txtsuit)
   txtsuit = Trim(txtsuit)
End If
End Sub
