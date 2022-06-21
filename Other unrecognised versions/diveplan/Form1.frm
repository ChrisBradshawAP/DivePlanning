VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB4C3E7-C689-11D2-B737-0060084D6C9E}#34.0#0"; "DChart.ocx"
Begin VB.Form rbinterface 
   BackColor       =   &H80000013&
   Caption         =   "RB Interface"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin DChartocx.Dchart Dchart1 
      Height          =   3255
      Left            =   3720
      TabIndex        =   40
      Top             =   240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5741
      MaxTime         =   480
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LineWidth       =   1
      SelectPointWidth=   1
      GraticuleColour =   12632064
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   3840
      TabIndex        =   41
      Top             =   7560
      Width           =   975
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
      Left            =   2640
      TabIndex        =   29
      Top             =   7560
      Width           =   975
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
      Left            =   1320
      TabIndex        =   28
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox txtinterval 
      Height          =   300
      Left            =   1320
      TabIndex        =   27
      Top             =   1200
      Width           =   2000
   End
   Begin VB.TextBox txtversion 
      Height          =   300
      Left            =   1320
      TabIndex        =   26
      Top             =   720
      Width           =   2000
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cmd6"
      Height          =   375
      Left            =   10680
      TabIndex        =   25
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SQL "
      Height          =   375
      Left            =   9960
      TabIndex        =   23
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "cmd4"
      Height          =   375
      Left            =   9240
      TabIndex        =   22
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "More..."
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
      TabIndex        =   18
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox txtdiveid 
      DataField       =   "DiveID"
      Height          =   300
      Left            =   1320
      TabIndex        =   17
      Top             =   240
      Width           =   2000
   End
   Begin VB.TextBox txttimedown 
      Height          =   300
      Left            =   1320
      TabIndex        =   16
      Top             =   1680
      Width           =   2000
   End
   Begin VB.TextBox txttimeup 
      DataField       =   "Finisheddate"
      Height          =   300
      Left            =   1320
      TabIndex        =   15
      Top             =   2160
      Width           =   2000
   End
   Begin VB.TextBox txtduration 
      Height          =   300
      Left            =   1320
      TabIndex        =   14
      Top             =   3140
      Width           =   1095
   End
   Begin VB.TextBox txtmaxdepth 
      DataField       =   "MaxDepth"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   1320
      TabIndex        =   13
      Top             =   2640
      Width           =   2000
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   11
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Profile Plotter"
      Height          =   3615
      Left            =   3600
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CheckBox chkpo2 
         Caption         =   "PO2"
         Height          =   255
         Left            =   6480
         TabIndex        =   39
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Draw Dive"
         Height          =   375
         Left            =   9120
         TabIndex        =   9
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Depth"
         Height          =   255
         Left            =   6480
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Temperature"
         Height          =   255
         Left            =   6480
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Gas"
         Height          =   255
         Left            =   6480
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check4 
         Caption         =   "HP"
         Height          =   255
         Left            =   6480
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Battery"
         Height          =   255
         Left            =   6480
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "PO2 - a"
         Height          =   255
         Left            =   6480
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox Check7 
         Caption         =   "PO2 - b"
         Height          =   255
         Left            =   6480
         TabIndex        =   2
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CheckBox Check8 
         Caption         =   "PO2 - c"
         Height          =   255
         Left            =   6480
         TabIndex        =   1
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   7920
         TabIndex        =   10
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
      TabIndex        =   30
      Top             =   0
      Width           =   3375
      Begin VB.Label Label21 
         Caption         =   "Duration :"
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Max. Depth :"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Time Up :"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   2180
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Time Down :"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Interval :"
         Height          =   255
         Left            =   440
         TabIndex        =   34
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Version :"
         Height          =   255
         Left            =   410
         TabIndex        =   33
         Top             =   740
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Dive ID :"
         Height          =   255
         Left            =   380
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   3120
         Width           =   855
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   9000
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File2 
      Height          =   675
      Left            =   6720
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9000
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtsearch 
      Height          =   495
      Left            =   7200
      TabIndex        =   24
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
      TabIndex        =   12
      Top             =   3000
      Width           =   1335
   End
End
Attribute VB_Name = "rbinterface"
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
Dim I As Integer
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

Private Sub Check1_Click()
  Getduration
  
  'plotgraph
End Sub

Private Sub Check2_Click()
  plotgraph
End Sub


Private Sub Check3_Click()
  Command3_Click
End Sub

Private Sub Check4_Click()
  Command3_Click
End Sub

Private Sub Check5_Click()
  Command3_Click
End Sub

Private Sub Check6_Click()
  Command3_Click
End Sub

Private Sub Check7_Click()
  Command3_Click
End Sub

Private Sub Check8_Click()
  Command3_Click
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub chkpo2_Click()
 plotgraph
End Sub

Private Sub cmdsave_Click()
' On Error GoTo ErrorHandler2
    CommonDialog1.Action = 2
    'MsgBox "FileName = " & CommonDialog1.Filename
    'CommonDialog1.InitDir = "\"
 '   CommonDialog1.FileName = ""
  '  filefilter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
   'CommonDialog1.Filter = filefilter
 '  Dim filefilter As String
'CommonDialog1.FileName = ""
' filefilter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
' CommonDialog1.Filter = filefilter
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

Private Sub Command1_Click()
  rbdetails.Show
  
 End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command3_Click()
  Picture1.Cls
  Open f For Binary As #1
  'Set TxtFile = FileMgr.OpenAsText("\c600\divelog.txt", afFileModeOpen)
  profilefound = 0
  Text1.Text = ""
  For ii = 1 To 10
    Get #1, , CB
    Text1.Text = Text1.Text + Chr$(CB) 'CStr(cb)
  Next ii
 
  j = 0
  ypmax1 = 1
  ypmax2 = 1
  ypmax3 = 1
  lpp = 0
  For LP = 0 To 999999
   lpp = lpp + 1
   If lpp = 100 Then
     lpp = 0
     Cls
     Print LP
   End If
   If profilefound = 1 Then
     'Text1.Text = TxtFile.Peek
     Get #1, , CB
     If CB = 13 Then
       Get #1, , CB
       'Get #1, , CB
     End If
       If CB = 69 Then
         LP = 9999999
       Else
         
         j = j + 1
         xp(j) = j 'CInt(txt2(0).Text)
         yp(j) = CInt(txt2(1))
         If xp(j) > xpmax Then xpmax = xp(j)
         'If yp(j) > ypmax Then ypmax = yp(j)
         yp1(j) = CInt(txt2(1))
         If yp1(j) > ypmax1 Then ypmax1 = yp1(j)
         yp2(j) = CInt(txt2(2))
         If yp2(j) > ypmax2 Then ypmax2 = yp2(j)
         If I > 2 Then
           yp3(j) = CInt(txt2(3))
           If yp3(j) > ypmax3 Then ypmax3 = yp3(j)
         End If
         If I > 3 Then
           yp4(j) = CInt(txt2(4))
           If yp4(j) > ypmax4 Then ypmax4 = yp4(j)
         End If
         If I > 4 Then
           yp5(j) = CInt(txt2(5))
           If yp5(j) > ypmax5 Then ypmax5 = yp5(j)
          End If
         If I > 5 Then
           yp6(j) = CInt(txt2(6))
           If yp6(j) > ypmax6 Then ypmax6 = yp6(j)
         End If
         If I > 6 Then
           yp7(j) = CInt(txt2(7))
           If yp7(j) > ypmax7 Then ypmax7 = yp7(j)
         End If
         If I > 7 Then
           yp8(j) = CInt(txt2(8))
           If yp8(j) > ypmax8 Then ypmax8 = yp8(j)
         End If
         If I > 8 Then
           yp9(j) = CInt(txt2(9))
           If yp9(j) > ypmax9 Then ypmax9 = yp9(j)
         End If
       End If
       I = 0
       cleartext2
     End If
     C = CStr(CB)
     If CB = 44 Then
       I = I + 1
     Else
       If CB > 57 Or CB < 48 Then
'         i = 0
'         cleartext2
       Else
         txt2(I) = txt2(I) + Chr$(CB)
       End If
     End If
   Else
     Get #1, , CB
     Text1.Text = Text1.Text + Chr$(CB)
     If InStr(Text1.Text, "Profile") Then
       Text1.Text = ""
       profilefound = 1
        I = 0
       cleartext2
       Get #1, , CB
       Get #1, , CB
     End If
     If LP > 2000 Then LP = 9999999
   End If
   If j > 206 Then LP = 9999999
   
  Next LP
  
  ypmax1 = ypmax1 + 5
  ypmax2 = ypmax2 + 5
  ypmax3 = ypmax3 + 5
  ypmax4 = ypmax4 + 5
  ypmax5 = ypmax5 + 5
  ypmax6 = ypmax6 + 5
  ypmax7 = ypmax7 + 5
  ypmax8 = ypmax8 + 5
  ypmax9 = ypmax9 + 5
  For I = 1 To j
    If Check1.Value = 1 Then Picture1.Line (xp(I - 1) * Picture1.Width / xpmax, yp1(I - 1) * Picture1.Height / ypmax1)-(xp(I) * Picture1.Width / xpmax, yp1(I) * Picture1.Height / ypmax1), &HFF
    If Check2.Value = 1 Then Picture1.Line (xp(I - 1) * Picture1.Width / xpmax, yp2(I - 1) * Picture1.Height / ypmax2)-(xp(I) * Picture1.Width / xpmax, yp2(I) * Picture1.Height / ypmax2), &HFF00
    If Check3.Value = 1 Then Picture1.Line (xp(I - 1) * Picture1.Width / xpmax, yp3(I - 1) * Picture1.Height / ypmax3)-(xp(I) * Picture1.Width / xpmax, yp3(I) * Picture1.Height / ypmax3), &HFF0000
    If Check4.Value = 1 Then Picture1.Line (xp(I - 1) * Picture1.Width / xpmax, yp4(I - 1) * Picture1.Height / ypmax4)-(xp(I) * Picture1.Width / xpmax, yp4(I) * Picture1.Height / ypmax4), &HFF000
  Next I
  fMain
  Close #1
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

Private Sub Form_Load()
Set DB = OpenDatabase(App.Path & "/rb.mdb")
  Dir1.Path = "C:\test"
Dim X As Single
Dim Y As Single

Screen.MousePointer = 11
Me.Left = (Screen.Width - Me.Width) / 2
Dchart1.PlotBackColor = &HE0E0E0
Dchart1.MaxDepth = 100
'Dchart1.SelectPoint(
tempserialno = "D0028"
txtdiveid = tempserialno
SQL = "SELECT * FROM main "
SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
Set RS = DB.OpenRecordset(SQL)
txttimedown = RS("startdate")
txttimeup = RS("finisheddate")
Txtduration = RS("duration")
TxtMaxdepth = RS("maxdepth")
txtinterval = RS("interval")
txtversion = RS("version")
SQL = "SELECT * FROM profile "
SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
Set RS = DB.OpenRecordset(SQL)
    MSFlexGrid1.FontSize = 7
    MSFlexGrid1.FontBold = False
        
            If RS.BOF And RS.EOF Then
                Screen.MousePointer = 0
                MsgBox "Empty record !", 48, Title
                Unload Me
                Exit Sub
            End If
            
                    
            For I = 0 To RS.Fields.Count - 1
                MSFlexGrid1.Col = 0
                MSFlexGrid1.Text = "Dive ID"
                MSFlexGrid1.Row = 0
                MSFlexGrid1.Cols = I + 1
                MSFlexGrid1.Col = I
                MSFlexGrid1.ColWidth(I) = 1000
                tempname = RS.Fields(I).Name
                 If IsNull(tempname) = True Or I = 0 Then
                 Else
                   If checkpfexist(tempname) = True Then
                      SQL = "SELECT * FROM pfindex "
                      SQL = SQL & "WHERE itemname =  '" & tempname & "'"
                      Set RS2 = DB.OpenRecordset(SQL)
                      MSFlexGrid1 = RS2("ITACNAME")
                   Else
                      MSFlexGrid1 = ""
                   End If
                   'MsgBox MSFlexGrid1.Text
                 End If
            Next I
            SQL = "SELECT * FROM profile "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            RS.MoveFirst
            numrow = 0
            While RS.EOF = False
               numrow = RS.RecordCount
               numrow = numrow + 1
               RS.MoveNext
            Wend
               RS.MoveFirst
            For I = 1 To numrow - 1
              If RS.EOF Then
                 Exit For
              End If
              'On Error Resume Next
              MSFlexGrid1.Row = I
                 For j = 0 To RS.Fields.Count - 1
                 'MsgBox j
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
            Next I
            MSFlexGrid1.Rows = numrow
    Totalcount = numrow - 1
    plotdchart
    Screen.MousePointer = 0
End Sub
Private Sub plotdchart()
  
  '  Dchart1.Xaxis.Interval = 4

End Sub
Private Sub Check_input()
  Do While 1
    If MSComm1.PortOpen = True Then
      S = MSComm1.Input
      If Len(S) = 0 Then
        Exit Do
      Else
        Text8.Text = Text8.Text + S
        TS = Hex(Asc(S))
        If InStr(S, "HW") Then
'          T2(1) = "Board 1"
'          W(0) = False
        End If
        T(0) = T(0) + TS
        If Len(T(0)) > 98 Then T(0) = TS
        Text1.Text = T(0)
        If Asc(S) > 9 Then
          T(1) = T(1) + S
          If (Len(T(1)) > 1398) Or S = "#" Then
            T2(0) = Time
            T(1) = "Board 1" + vbCrLf0 "Time=" + T2(0)
            T2(0) = Date
            T(1) = T(1) + vbCrLf + "Date=" + T2(0)
          End If
          Text7.Text = T(1)
          T2(1) = T(1)
        End If
      End If
    Else: Exit Do
    End If
  Loop
  Screen.MousePointer = 0
  'Print "DGDF"
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
Picture1.Cls
If Check2.Value = 1 Then
  RS.MoveFirst
  K = 0
  While RS.EOF = False
     K = K + 1
     test = Trim(RS("item6"))
     test = Right$(test, 4)
     yp1(K) = CInt(test)
     temperature = test
     If Val(temperature) > Val(maxtemp) Then
        maxtemp = Val(temperature)
     End If
     RS.MoveNext
  Wend
  For K = 1 To Totalcount
      Picture1.Line ((K - 1) * Picture1.Width / Totalcount, (yp1(K - 1)) * Picture1.Height / maxtemp)-(K * Picture1.Width / Totalcount, yp1(K) * Picture1.Height / maxtemp), &HFF
  Next K
End If

If Check1.Value = 1 Then
  RS.MoveFirst
  K = 0
  While RS.EOF = False
     K = K + 1
     yp1(K) = CInt(RS("item2"))
     tempdepth = RS("item2")
     If Val(tempdepth) > Val(MaxDepth) Then
        MaxDepth = Val(tempdepth)
     End If
     RS.MoveNext
  Wend
  For K = 1 To Totalcount
     'If Check1.Value = 1 Then Picture1.Line ((K - 1) * Picture1.Width / Totalcount, (yp1(K - 1)) * Picture1.Height / MaxDepth)-(K * Picture1.Width / Totalcount, yp1(K) * Picture1.Height / MaxDepth), &HFF
     Picture1.Line (K * Picture1.Width / Totalcount, yp1(K) * Picture1.Height / MaxDepth)-((K - 1) * Picture1.Width / Totalcount, (yp1(K - 1)) * Picture1.Height / MaxDepth), &HFF
  Next K
End If


If chkpo2.Value = 1 Then
RS.MoveFirst
  K = 0
  While RS.EOF = False
     K = K + 1
     yp1(K) = CInt(RS("item4"))
     tempdepth = RS("item4")
     If Val(tempdepth) > Val(MaxDepth) Then
        MaxDepth = Val(tempdepth)
     End If
     RS.MoveNext
  Wend
  For K = 1 To Totalcount
     Picture1.Line ((K - 1) * Picture1.Width / Totalcount, (yp1(K - 1)) * Picture1.Height / MaxDepth)-(K * Picture1.Width / Totalcount, yp1(K) * Picture1.Height / MaxDepth), &HFF
  Next K
End If
End Function
Private Sub Getduration()
  
End Sub
