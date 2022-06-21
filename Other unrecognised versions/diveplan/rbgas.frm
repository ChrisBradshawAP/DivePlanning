VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form rbgas 
   BackColor       =   &H80000013&
   Caption         =   "RB Interface - Gas"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "rbgas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Dive Profile"
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
      TabIndex        =   27
      Top             =   7560
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   7680
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
      Left            =   3240
      TabIndex        =   26
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox txtinterval 
      Height          =   300
      Left            =   1320
      TabIndex        =   16
      Top             =   1200
      Width           =   2000
   End
   Begin VB.TextBox txtversion 
      Height          =   300
      Left            =   1320
      TabIndex        =   15
      Top             =   720
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dive Details"
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
      Width           =   1335
   End
   Begin VB.TextBox txtdiveid 
      DataField       =   "DiveID"
      Height          =   300
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   2000
   End
   Begin VB.TextBox txttimedown 
      Height          =   300
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   2000
   End
   Begin VB.TextBox txttimeup 
      DataField       =   "Finisheddate"
      Height          =   300
      Left            =   1320
      TabIndex        =   7
      Top             =   2160
      Width           =   2000
   End
   Begin VB.TextBox txtduration 
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   3140
      Width           =   1095
   End
   Begin VB.TextBox txtmaxdepth 
      DataField       =   "MaxDepth"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   2000
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   16384
      Cols            =   21
      FixedCols       =   0
      RowHeightMin    =   30
      BackColorBkg    =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Profile Plotter"
      Height          =   3615
      Left            =   3600
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   3135
         Left            =   240
         OleObjectBlob   =   "rbgas.frx":030A
         TabIndex        =   28
         Top             =   240
         Width           =   7815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Draw Dive"
         Height          =   375
         Left            =   9120
         TabIndex        =   1
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   7920
         TabIndex        =   2
         Top             =   4440
         Width           =   615
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
      Height          =   3615
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   3375
      Begin VB.Label Label21 
         Caption         =   "Duration :"
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Max. Depth :"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Time Up :"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2180
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Time Down :"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Interval :"
         Height          =   255
         Left            =   440
         TabIndex        =   21
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Version :"
         Height          =   255
         Left            =   410
         TabIndex        =   20
         Top             =   740
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Dive ID :"
         Height          =   255
         Left            =   380
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   3120
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
      Left            =   6720
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9000
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtsearch 
      Height          =   495
      Left            =   7200
      TabIndex        =   14
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
Attribute VB_Name = "rbgas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim DB As Database
Dim Y(20) As Integer
Dim X As Integer
Dim fp As UserDocument
'Dim fpp As CFileBinaryReadable
'Dim FileMgr As New CFileManager
'Dim TxtFile As CFileTextReadable
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
Private Sub cmdclose_Click()
If tempsysviewid = "Rebrether" Then
   rbinterface.Show
Else
   vrinterface.Show
End If
Unload Me

End Sub

Private Sub cmdgas_Click()

End Sub

Private Sub cmdsave_Click()
 On Error GoTo ErrorHandler2
 CommonDialog1.Action = 2
 Open CommonDialog1.FileName For Output As #1
 MSFlexGrid1.Row = 0
 
    For j = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Col = j
        rowtext = MSFlexGrid1.Text
        cOMPTEXT = cOMPTEXT + (rowtext + ",")
    Next j
    Print #1, cOMPTEXT
    cOMPTEXT = ""
    RS.MoveFirst
    Do Until RS.EOF
        For j = 0 To RS.Fields.Count - 1
            If IsNull(RS(j)) Then
               rowtext = ""
            Else
               rowtext = CStr(RS(j))
            End If
             cOMPTEXT = Trim(cOMPTEXT)
             cOMPTEXT = cOMPTEXT + (rowtext) & ","
                       
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

Private Sub Cmdtissue_Click()

End Sub

Private Sub Command1_Click()
Unload Me
  rbdetails.Show
  
 End Sub

Private Sub Command2_Click()
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

Private Sub Form_Activate()
cmdclose.SetFocus
End Sub

Private Sub Form_Load()
Dim X As Single
Dim Y As Single
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
Screen.MousePointer = 11
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 30
SQL = "SELECT * FROM main "
SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
Set RS = DB.OpenRecordset(SQL)
txttimedown = RS("startdate")
txttimeup = RS("finisheddate")
txtduration = RS("duration")
txtmaxdepth = RS("maxdepth")
txtinterval = RS("interval")
txtversion = RS("version")
txtdiveid = tempserialno
SQL = "SELECT * FROM GAS "
SQL = SQL & " where DiveID = '" & tempserialno & "'"
Set RS = DB.OpenRecordset(SQL)
MSFlexGrid1.FontSize = 11
MSFlexGrid1.FontBold = False
If RS.BOF And RS.EOF Then
  Screen.MousePointer = 0
  MsgBox "Empty record !", 48, Title
  Unload Me
  Exit Sub
End If
 For i = 0 To RS.Fields.Count - 1
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Cols = i + 1
    MSFlexGrid1.Col = i
    MSFlexGrid1.ColWidth(i) = 1500
    tempname = RS.Fields(i).Name
    If IsNull(tempname) = True Then
      MSFlexGrid1 = ""
    Else
      MSFlexGrid1 = tempname
    End If
  Next i
    RS.MoveFirst
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
              'On Error Resume Next
         MSFlexGrid1.Row = i
         For j = 0 To RS.Fields.Count - 1
             MSFlexGrid1.Col = j
                 If IsNull(RS(j)) Then
                    MSFlexGrid1.Text = ""
                 Else
                    TEMPVALUE = CStr(RS(j))
                    If Val(TEMPVALUE) < 1 And Val(TEMPVALUE) > 0 Then
                      TEMPVALUE = "0" & TEMPVALUE
                      MSFlexGrid1.Text = TEMPVALUE
                    Else
                      MSFlexGrid1.Text = TEMPVALUE
                    End If
                 End If
              Next j
              RS.MoveNext
            Next i
            MSFlexGrid1.Rows = numrow
    Totalcount = numrow - 1
    MSFlexGrid1.ColPosition(MSFlexGrid1.MouseCol) = 4

   ' plotdchart
    Screen.MousePointer = 0
End Sub
Private Sub plotdchart()
 SQL = "SELECT * FROM GAS "
SQL = SQL & " where DiveID = '" & tempserialno & "'"
Set RS = DB.OpenRecordset(SQL)
End Sub
Private Sub Check_input()

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



