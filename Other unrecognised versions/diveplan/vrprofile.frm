VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form vrinterface 
   BackColor       =   &H80000013&
   Caption         =   "VR Interface - Profile"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "vrprofile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Pan >"
      Enabled         =   0   'False
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
      Left            =   8640
      TabIndex        =   43
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pan <"
      Enabled         =   0   'False
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
      Left            =   7560
      TabIndex        =   42
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zoom"
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
      Left            =   6480
      TabIndex        =   41
      Top             =   7560
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgchart 
      Left            =   840
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdsavegraph 
      Caption         =   "Save Graph"
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
      Left            =   5160
      TabIndex        =   32
      Top             =   7560
      Width           =   1215
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
      Left            =   10680
      TabIndex        =   30
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
      Left            =   9840
      TabIndex        =   29
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Dive List"
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
      Left            =   4080
      TabIndex        =   28
      Top             =   7560
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
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
      Height          =   375
      Left            =   3120
      TabIndex        =   27
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton Cmdtissue 
      Caption         =   "Tissue"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   2160
      TabIndex        =   17
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton cmdgas 
      Caption         =   "Gas"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   1200
      TabIndex        =   16
      Top             =   7560
      Width           =   855
   End
   Begin VB.TextBox txtinterval 
      Height          =   300
      Left            =   1320
      TabIndex        =   15
      Top             =   1560
      Width           =   1035
   End
   Begin VB.TextBox txtversion 
      Height          =   300
      Left            =   1320
      TabIndex        =   14
      Top             =   840
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   120
      TabIndex        =   10
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox txtdiveid 
      DataField       =   "DiveID"
      Height          =   300
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   1875
   End
   Begin VB.TextBox txttimedown 
      Height          =   300
      Left            =   1320
      TabIndex        =   8
      Top             =   2280
      Width           =   1875
   End
   Begin VB.TextBox txttimeup 
      DataField       =   "Finisheddate"
      Height          =   300
      Left            =   1320
      TabIndex        =   7
      Top             =   2880
      Width           =   1875
   End
   Begin VB.TextBox txtduration 
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtmaxdepth 
      DataField       =   "MaxDepth"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   6500
      Cols            =   51
      FixedCols       =   0
      RowHeightMin    =   30
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame3 
      Caption         =   "Profile Plotter"
      Height          =   4695
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.TextBox Text5 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5500
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   4300
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   4300
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4300
         TabIndex        =   37
         Text            =   "Text4"
         Top             =   4300
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   36
         Text            =   "Text3"
         Top             =   4300
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1700
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   4300
         Width           =   960
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   4380
         Left            =   120
         OleObjectBlob   =   "vrprofile.frx":030A
         TabIndex        =   31
         Top             =   240
         Width           =   6735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Draw Dive"
         Height          =   375
         Left            =   9120
         TabIndex        =   2
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Depth"
         Height          =   255
         Left            =   7020
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   720
         TabIndex        =   33
         Top             =   4200
         Width           =   735
      End
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
   Begin VB.Frame Frame1 
      Caption         =   "Dive Info"
      Height          =   4695
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   3255
      Begin VB.Label lblmaxdepth 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   40
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "seconds"
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
         Left            =   2400
         TabIndex        =   35
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Duration :"
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Max. Depth :"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Time Up :"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Time Down :"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Interval :"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Version :"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Dive ID :"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "minutes"
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
         Left            =   2400
         TabIndex        =   19
         Top             =   4200
         Width           =   855
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   9000
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File2 
      Height          =   675
      Left            =   6360
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtsearch 
      Height          =   495
      Left            =   7200
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000013&
      Caption         =   "Dive ID :"
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
      Left            =   4935
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
End
Attribute VB_Name = "vrinterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim T(4) As String
Dim T2(4) As String
Dim S As String
Dim TS As String
Dim K As Integer
Dim H As String
Dim W(4) As Boolean
Dim zxt As Double
Dim CT As Integer
Dim Auto_run As Integer

Dim zoom As Integer
Dim pan As Integer

Private Sub Check1_Click()
  plotgraph
End Sub
Private Sub Check2_Click()
   plotgraph
End Sub
Private Sub Check3_Click()
  plotgraph
End Sub
Private Sub Check4_Click()
 plotgraph
End Sub
Private Sub Check5_Click()
   plotgraph
End Sub
Private Sub Check6_Click()
   plotgraph
End Sub
Private Sub Check7_Click()
  plotgraph
End Sub

Private Sub Check8_Click()
   plotgraph
End Sub
Private Sub Check9_Click()
   plotgraph
End Sub
Private Sub Check10_Click()
 plotgraph
End Sub
Private Sub Check11_Click()
 plotgraph
End Sub
Private Sub Combo1_Change()

End Sub




Private Sub cmdclose_Click()
Unload Me
rbmain.Show
End Sub

Private Sub cmdgas_Click()
Unload Me
rbgas.Show
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

Private Sub cmdsave_Click()
 On Error GoTo ErrorHandler2
    CommonDialog1.Action = 2
 filefilter = "Text Files (*.csv)|*.csv|All Files (*.*)|*.*"
 CommonDialog1.Filter = filefilter
 Open CommonDialog1.FileName For Output As #1
 MSFlexGrid1.Row = 0
 
    For j = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Col = j
        rowtext = MSFlexGrid1.Text
        cOMPTEXT = cOMPTEXT + (rowtext + ",")
    Next j
    Print #1, cOMPTEXT
    cOMPTEXT = ""
          Select Case displaydefaulted
             Case "Feet"
                displaydefault = "Feet"
             Case "Meter"
                displaydefault = "Meter"
             Case Else
                SQL = "SELECT * FROM Display "
                Set RS4 = DB.OpenRecordset(SQL)
                displaydefault = RS4("display")
              End Select
          
    RS.MoveFirst
    Do Until RS.EOF
      For j = 0 To RS.Fields.Count - 1
         If IsNull(RS(j)) Then
            rowtext = ""
         Else
            rowtext = CStr(RS(j))
         End If
         If displaydefault = "Feet" Then
            If j = 2 Or j = 9 Or j = 7 Or j = 15 Or j = 13 Then
               cOMPTEXT = Trim(cOMPTEXT)
            Else
               cOMPTEXT = Trim(cOMPTEXT)
               cOMPTEXT = cOMPTEXT + (rowtext) & ","
            End If
         Else
            If j = 3 Or j = 10 Or j = 8 Or j = 16 Or j = 14 Then
              cOMPTEXT = Trim(cOMPTEXT)
               'cOMPTEXT = cOMPTEXT + (rowtext) & ","
            Else
                cOMPTEXT = Trim(cOMPTEXT)
               cOMPTEXT = cOMPTEXT + (rowtext) & ","
              
            
            End If
         End If
        Next j
           Print #1, cOMPTEXT
        cOMPTEXT = ""
        RS.MoveNext
    Loop
    Close #1

ErrorHandler2:
    If Err = 32755 Then
        MsgBox "Data is not saved to a file....!!"
        Exit Sub
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

Private Sub Cmdtissue_Click()
Unload Me
rbtissue.Show
End Sub

Private Sub Command1_Click()
  Unload Me
  rbdetails.Show
  
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

Private Sub Command4_Click()
  Open F1 For Binary As #2
  Text9.Text = F1
  MSComm1.Output = "L"
  For I1 = 0 To 10
    Cls
    S = MSComm1.Input
    Text8.Text = S
  Next I1
  'wait
  For I1 = 0 To 2816
    Get #2, , c1
    MSComm1.Output = Chr$(c1)
    Cls
    Text8.Text = Chr$(c1)
  Next I1
  MSComm1.Output = vbCrLf
  Close #2
End Sub

Private Sub Command2_Click(Index As Integer) 'nick
  Select Case Index
    Case 0
    If Totalcount > 100 Then
      If zoom = 10 Then
        zoom = 1
        pan = 1
        Command2(1).Enabled = False
        Command2(2).Enabled = False
      Else
        zoom = 10
        pan = 1
        Command2(1).Enabled = True
        Command2(2).Enabled = True
      End If
    End If
    Case 1
      pan = pan - 1
      If pan < 1 Then pan = 1
    Case 2
      pan = pan + 1
      If pan > 10 Then pan = 10
  End Select
  If zoom = 10 Then
    If pan = 1 Then Command2(1).Enabled = False Else Command2(1).Enabled = True
    If pan = 10 Then Command2(2).Enabled = False Else Command2(2).Enabled = True
  End If
  Command2(0).Caption = "Zoom " + CStr(zoom)
  plotgraph
End Sub

Private Sub Form_Load()
Dim X As Single
Dim Y As Single
Dim p As Integer

Me.Top = 30
Screen.MousePointer = 11
Me.Left = (Screen.Width - Me.Width) / 2
If tempserialno = "" Then
   SQL = "SELECT DiveID, systemid FROM main "
   SQL = SQL & " WHERE SystemID = '" & tempsysviewid & "' "
   SQL = SQL & " order by DiveID "
   Set RS = DB.OpenRecordset(SQL)
   RS.MoveLast
   tempserialno = RS("DiveID")
   txtdiveid = tempserialno
Else
   txtdiveid = tempserialno
End If
SQL = "SELECT * FROM main "
SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
Set RS = DB.OpenRecordset(SQL)

txtduration = RS("duration")
  Select Case displaydefaulted
    Case "Feet"
      displaydefault = "Feet"
    Case "Meter"
      displaydefault = "Meter"
    Case Else
      SQL = "SELECT * FROM Display "
      Set RS4 = DB.OpenRecordset(SQL)
      displaydefault = RS4("display")
    End Select
    If displaydefault = "Feet" Then
      txtmaxdepth = RS("maxdepth")
      txtmaxdepth = txtmaxdepth * 3.28084
      txtmaxdepth = Format(txtmaxdepth, "########.00")
      lblmaxdepth = "Feet"
      txttimedown = RS("startdate")
      txttimeup = RS("finisheddate")
    Else
      txtmaxdepth = RS("maxdepth")
      lblmaxdepth = "Meters"
      txttimedown = Format$(RS("startdate"), "dd/mm/yyyy hh:nn:ss")
     ' txttimeup = RS("finisheddate")
      txttimeup = Format$(RS("finisheddate"), "dd/mm/yyyy hh:nn:ss")
    End If
txtinterval = RS("interval")
txtversion = RS("version")
SQL = "SELECT * FROM profile "
SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
SQL = SQL & " ORDER BY item1 "
Set RS = DB.OpenRecordset(SQL)
MSFlexGrid1.FontSize = 7
MSFlexGrid1.FontBold = False
  If RS.BOF And RS.EOF Then
    Screen.MousePointer = 0
    MsgBox "Empty record !", 48, Title
    Unload Me
    Exit Sub
  End If
  p = 1
  MSFlexGrid1.Col = 0
  MSFlexGrid1.Row = 0
  MSFlexGrid1.Text = "Dive ID"
  For i = 0 To RS.Fields.Count - 1
    MSFlexGrid1.Cols = RS.Fields.Count
    MSFlexGrid1.Col = Val(p)
    MSFlexGrid1.ColWidth(i) = 1000
    tempname = RS.Fields(i).Name
    If IsNull(tempname) = True Or i = 0 Then
    Else
      Select Case displaydefaulted
      Case "Feet"
         displaydefault = "Feet"
      Case "Meter"
         displaydefault = "Meter"
      Case Else
         SQL = "SELECT * FROM Display "
         Set RS4 = DB.OpenRecordset(SQL)
         displaydefault = RS4("display")
    End Select
    K = 1
    If displaydefault = "Feet" Then
       Option1 = True
    Else
       Option2 = True
    End If
      If checkpfexist(tempname) = True Then
            SQL = "SELECT * FROM vrpfindex "
            SQL = SQL & "WHERE itemname =  '" & tempname & "'"
            Set RS2 = DB.OpenRecordset(SQL)
            p = p + 1
               itemheader = RS2("ITACNAME")
               MSFlexGrid1 = itemheader
               If itemheader = "depth(m)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  Else
                     MSFlexGrid1 = itemheader
                   End If
                End If
               
               If itemheader = "depth(f)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = itemheader
                  Else
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  End If
                End If
               
               If itemheader = "temperature(c)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  Else
                     MSFlexGrid1 = itemheader
                  End If
                End If
                
                 If itemheader = "temperature(f)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = itemheader
                  Else
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  End If
                End If
              If itemheader = "extdepth(f)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = itemheader
                  Else
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  End If
                End If
                
                 If itemheader = "extdepth(m)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  Else
                     MSFlexGrid1 = itemheader
                  End If
                End If
                If itemheader = "HPDil(p)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = itemheader
                  Else
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  End If
                End If
                
                 If itemheader = "HPDil(b)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  Else
                     MSFlexGrid1 = itemheader
                  End If
                End If
                If itemheader = "HPO2(p)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = itemheader
                  Else
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  End If
                End If
                
                 If itemheader = "HPO2(b)" Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  Else
                     MSFlexGrid1 = itemheader
                  End If
                End If
               
           End If
          End If
        Next i
            SQL = "SELECT * FROM profile "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            SQL = SQL & " ORDER BY item1 "
            Set RS = DB.OpenRecordset(SQL)
            RS.MoveFirst
            p = 0
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
               p = 0
              'On Error Resume Next
              MSFlexGrid1.Row = i
              test = RS.Fields.Count
               
              For j = 0 To RS.Fields.Count - 1
                  
                  If IsNull(RS(j)) = True Then
                    If j > 6 And j < 50 Then
                      MSFlexGrid1.Col = MSFlexGrid1.Col + 1
                      MSFlexGrid1.Text = ""
                    End If
                 Else
                    MSFlexGrid1.Col = p
                    MSFlexGrid1 = CStr(RS(j))
                   
                 End If
                 p = p + 1
               If j = 2 Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                  Else
                     TEMPVALUE = CStr(RS(j))
                     MSFlexGrid1 = TEMPVALUE
                     
                   End If
                End If
               
               If j = 3 Then
                  If displaydefault = "Feet" Then
                     MSFlexGrid1 = CStr(RS(j))
                  Else
                     MSFlexGrid1 = ""
                     p = p - 1
                     MSFlexGrid1.Col = p
                     
                  End If
                End If
             
              Next j
              RS.MoveNext
            Next i
            MSFlexGrid1.Rows = numrow
    Totalcount = numrow - 1
    determinexaxis
    K = K + 1
    Check1.Value = 1
    Screen.MousePointer = 0
End Sub
Private Sub plotdchart()
  
  '  Dchart1.Xaxis.Interval = 4

End Sub
Private Sub Check_input()
End Sub
Function checkpfexist(ByVal tempname As String) As Boolean
SQL = "SELECT COUNT(*) FROM vrpfindex "
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
Dim xx(1 To 5000) As Single
Dim yy(1 To 10) As Single
'Const test
 Totalcount = Val(Totalcount)
 determinexaxis 'nick
 SQL = "SELECT * FROM profile "
 SQL = SQL & " where DiveID = '" & tempserialno & "'"
 SQL = SQL & " ORDER BY item1 "
 Set RS = DB.OpenRecordset(SQL)
 ReDim ARRAYDATA(Totalcount, 11)  '1 To 237, 1 To 1) ' As Integer
 r = 0
 If zoom < 1 Then zoom = 1 'nick
 If pan < 1 Then pan = 1 'nick
 ReDim ARRAYDATAB(Totalcount / zoom, 11) 'nick
 For q = 1 To 11
   For i = 1 To Totalcount
     Select Case q
       Case 1
         If Check1.Value = 1 Then
           If displaydefault = "Feet" Then
              tempreading = RS("item3")
              xx(i) = Val("-" & tempreading)
              ARRAYDATA(i, q) = xx(i)
              tempreading = Format(tempreading, "########.00")
              If i = 1 Then
                Maxdepthvalue = tempreading
                Mindepthvalue = tempreading
              End If
            If Val(tempreading) > Val(Maxdepthvalue) Then
               Maxdepthvalue = tempreading
            End If
            If Val(tempreading) < Val(Mindepthvalue) Then
               Mindepthvalue = tempreading
            End If
            RS.MoveNext
          Else
            tempreading = RS("item2")
            xx(i) = Val("-" & tempreading)
            ARRAYDATA(i, q) = xx(i)
            tempreading = Format(tempreading, "########.00")
            If i = 1 Then
               Mindepthvalue = tempreading
            End If
            If Val(tempreading) > Val(Maxdepthvalue) Then
               Maxdepthvalue = tempreading
            End If
            If Val(tempreading) < Val(Mindepthvalue) Then
               Mindepthvalue = tempreading
            End If
              RS.MoveNext
            End If
          End If
        
        
  
           
     End Select
   Next i
   RS.MoveFirst
   Next q

 
  If zoom = 10 Then 'nick
    For q = 1 To 11
      For i = 1 To (Totalcount / zoom)
          ARRAYDATAB(i, q) = ARRAYDATA(i + ((pan - 1) * (Totalcount / zoom)), q)
      Next i
    Next q
    MSChart1 = ARRAYDATAB 'nick
  Else
    MSChart1 = ARRAYDATA
  End If
End Function


Private Sub MSChart1_PointSelected(Series As Integer, Datapoint As Integer, MouseFlags As Integer, Cancel As Integer)
Dim Datapointtemp
  If zoom = 10 Then
    Datapointtemp = Datapoint - 1 + ((pan - 1) * (Totalcount / zoom))
  Else
    Datapointtemp = Datapoint
  End If
  If Datapointtemp <> previouspoint Then
    MSFlexGrid1.Row = previouspoint
    For p = 1 To 14
      MSFlexGrid1.Col = p
      MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
      MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor '
    Next p
  End If
  With MSChart1
      .Row = Datapoint
      On Error GoTo error_handle
      MSFlexGrid1.Row = Datapointtemp
      MSFlexGrid1.Col = 1
      MSFlexGrid1.TopRow = MSFlexGrid1.Row - 4
      MSFlexGrid1.LeftCol = MSFlexGrid1.Col
  End With
For p = 1 To 14
  MSFlexGrid1.Col = p
  MSFlexGrid1.CellForeColor = vbWhite
  MSFlexGrid1.CellBackColor = vbBlue
Next p
previouspoint = Datapointtemp
error_handle:
End Sub

Private Sub MSFlexGrid1_EnterCell()

 ' MSFlexGrid1.CellForeColor = vbWhite
  'MSFlexGrid1.CellBackColor = vbBlue
End Sub

Private Sub MSFlexGrid1_LeaveCell()
 '   MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
 '   MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor '
End Sub


Private Sub Option1_Click()
   If K <> 1 Then
     displaydefaulted = "Feet"
     SQL = "SELECT * FROM Display "
     Set RS = DB.OpenRecordset(SQL)
     RS.Edit
     RS!Display = displaydefaulted
     RS.Update
     Unload Me
     vrinterface.Show
   End If
End Sub

Private Sub Option2_Click()
   If K <> 1 Then
     displaydefaulted = "Meter"
     SQL = "SELECT * FROM Display "
     Set RS = DB.OpenRecordset(SQL)
     RS.Edit
     RS!Display = displaydefaulted
     RS.Update
     Unload Me
     vrinterface.Show
   End If
End Sub

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

