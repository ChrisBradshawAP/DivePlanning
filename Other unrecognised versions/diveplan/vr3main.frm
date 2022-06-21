VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form vr3main 
   BackColor       =   &H80000013&
   Caption         =   "VR3 Main"
   ClientHeight    =   8475
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "vr3main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdelete 
      Caption         =   "Delete Dive"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "Go"
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
      TabIndex        =   7
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Main Menu"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   7680
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12938
      _Version        =   393216
      Rows            =   16384
      Cols            =   21
      FixedCols       =   0
      RowHeightMin    =   30
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save as csv"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   7680
      Width           =   1335
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
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   9000
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File2 
      Height          =   675
      Left            =   6720
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtsearch 
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnupopprofile 
         Caption         =   "&Profile"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu MNUSAVEDB 
         Caption         =   "&Save in DB"
      End
      Begin VB.Menu mnusavecsv 
         Caption         =   "Save in &CSV"
      End
   End
End
Attribute VB_Name = "vr3main"
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
Dim H As Integer
Dim W(4) As Boolean
Dim zxt As Double
Dim CT As Integer
Dim Auto_run As Integer



Private Sub cmdclose_Click()
Unload Me
main.Show

End Sub

Private Sub cmddelete_Click()
If MSFlexGrid1.CellBackColor = vbBlue Then
   MSFlexGrid1.Col = 0
   tempserialno = MSFlexGrid1.Text
   MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
   MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor
   ans = MsgBox("Are you sure you want to deleted the selected record(s)?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
      Select Case ans
         Case vbYes
            SQL = "DELETE FROM main "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            'On Error GoTo errorhandle
              'ComBurn.'CommandText = SQL
              '              Set RS = ComBurn.Execute
                    '    Next
                                           
                        Me.MousePointer = vbNormal
         Case vbNo
            Me.MousePointer = vbNormal
            Exit Sub
      End Select
    'On Error GoTo errorhandle:
    MsgBox "error"

End If
End Sub

Private Sub cmdelete_Click()
If MSFlexGrid1.CellBackColor = vbBlue Then
   MSFlexGrid1.Col = 0
   tempserialno = MSFlexGrid1.Text
  On Error GoTo errorhandle
   ans = MsgBox("Are you sure you want to deleted the selected record(s)?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
      Select Case ans
         Case vbYes
            SQL = "select * FROM main "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            RS.Delete
            RS.Close
            SQL = "select * FROM PROFILE "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            RS.Delete
            RS.Close
            SQL = "select * FROM gas "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            RS.Delete
            RS.Close
            SQL = "select * FROM tissue "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            RS.Delete
            RS.Close
            
            Me.MousePointer = vbNormal
            Unload Me
            rbmain.Show
         Case vbNo
            Me.MousePointer = vbNormal
            Exit Sub
      End Select
errorhandle:
 '  If Err.Number <> 0 Then
 '   MsgBox Error$
 '  End If
 Unload Me
 rbmain.Show
End If
End Sub

Private Sub cmdgo_Click()
If MSFlexGrid1.CellBackColor = vbBlue Then
   MSFlexGrid1.Col = 0
   tempserialno = MSFlexGrid1.Text
   MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
   MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor '
   Unload Me
   rbinterface.Show
 Else
   MSFlexGrid1.Col = 0
   MSFlexGrid1.Row = Totalcount
   tempserialno = MSFlexGrid1.Text
   MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
   MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor '
   Unload Me
   rbinterface.Show
 End If
End Sub

Private Sub cmdsave_Click()
 On Error GoTo ErrorHandler2
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
        'test = Len(rawtext)
        
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
             rowtext = Trim(rowtext)
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

Private Sub Command3_Click()
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

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Dim X As Single
Dim Y As Single
If filesource = "" Then
   Set DB = OpenDatabase(App.Path & "/rb.mdb")
   Dir1.Path = App.Path
Else
   Set DB = OpenDatabase(filesource)
End If
Screen.MousePointer = 11
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 30
SQL = "SELECT * FROM main "
SQL = SQL & "WHERE systemid = 'VR3' "
Set RS = DB.OpenRecordset(SQL)
MSFlexGrid1.FontSize = 7
MSFlexGrid1.FontBold = False
If RS.BOF And RS.EOF Then
  Screen.MousePointer = 0
  MsgBox "Empty record !", 48, Title
      
      Exit Sub
      Unload Me
 ' main.Show
End If
 For i = 0 To RS.Fields.Count - 1
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Cols = i + 1
    MSFlexGrid1.Col = i
    If i < 7 And i > 0 Then
       MSFlexGrid1.ColWidth(i) = 2000
    Else
       MSFlexGrid1.ColWidth(i) = 1000
    End If
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
            Next i
            MSFlexGrid1.Rows = numrow
    Totalcount = numrow - 1
     Screen.MousePointer = 0
End Sub
Private Sub plotdchart()
  
  '  Dchart1.Xaxis.Interval = 4

End Sub
Private Sub Check_input()
'  Do While 1
'    If MSComm1.PortOpen = True Then
'      S = MSComm1.Input
'      If Len(S) = 0 Then
'        Exit Do
'      Else
'        Text8.Text = Text8.Text + S
 '       TS = Hex(Asc(S))
 ''       If InStr(S, "HW") Then
'          T2(1) = "Board 1"
'          W(0) = False
'        End If
'        T(0) = T(0) + TS
'        If Len(T(0)) > 98 Then T(0) = TS
'        Text1.Text = T(0)
'        If Asc(S) > 9 Then
'          T(1) = T(1) + S
 ''         If (Len(T(1)) > 1398) Or S = "#" Then
 '           T2(0) = Time
 '           T(1) = "Board 1" + vbCrLf0 "Time=" + T2(0)
 '           T2(0) = Date
 '           T(1) = T(1) + vbCrLf + "Date=" + T2(0)
 '         End If
 '         Text7.Text = T(1)
 '         T2(1) = T(1)
 '       End If
 '     End If
 '   Else: Exit Do
 '   End If
 ' Loop
 ' Screen.MousePointer = 0
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

Private Sub mnuaddinfo_Click()
End Sub

Private Sub mnupoopgas_Click()
End Sub

Private Sub mnudelete_Click()
If MSFlexGrid1.CellBackColor = vbBlue Then
   MSFlexGrid1.Col = 0
   tempserialno = MSFlexGrid1.Text
  ' MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
  ' MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor
   On Error GoTo errorhandle
   ans = MsgBox("Are you sure you want to deleted the selected record(s)?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
      Select Case ans
         Case vbYes
            SQL = "select * FROM main "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            RS.Delete
            RS.Close
            SQL = "select * FROM PROFILE "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            RS.Delete
            RS.Close
            SQL = "select * FROM gas "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            RS.Delete
            RS.Close
            SQL = "select * FROM tissue "
            SQL = SQL & "WHERE DiveID = '" & tempserialno & "'"
            Set RS = DB.OpenRecordset(SQL)
            RS.Delete
            RS.Close
            
            Me.MousePointer = vbNormal
            Unload Me
            rbmain.Show
         Case vbNo
            Me.MousePointer = vbNormal
            Exit Sub
      End Select
errorhandle:
   If Err.Number <> 0 Then
    MsgBox Error$
   End If
End If
End Sub

Private Sub mnupopprofile_Click()
MSFlexGrid1.Col = 0
tempserialno = MSFlexGrid1.Text
 MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
 MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor '
Unload Me
rbinterface.Show
End Sub

Private Sub mnupoptissue_Click()
End Sub

Private Sub mnusavecsv_Click()
On Error GoTo ErrorHandler2
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
        'test = Len(rawtext)
        
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
             rowtext = Trim(rowtext)
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

Private Sub MSFlexGrid1_Click()
rowindentified = MSFlexGrid1.Row
For K = 0 To Totalcount
  For p = 0 To 0
    MSFlexGrid1.Row = K
    MSFlexGrid1.Col = p
    If MSFlexGrid1.CellBackColor = vbBlue Then
      For H = 0 To 16
        MSFlexGrid1.Row = K
        MSFlexGrid1.Col = H
        MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
        MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor
      Next H
    End If
  Next p
Next K
For p = 0 To 16
  MSFlexGrid1.Row = rowindentified
  MSFlexGrid1.Col = p
  MSFlexGrid1.CellForeColor = vbWhite
  MSFlexGrid1.CellBackColor = vbBlue
Next p
End Sub

Private Sub MSFlexGrid1_DblClick()
MSFlexGrid1.Col = 0
tempserialno = MSFlexGrid1.Text
MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColor
 MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColor '
Unload Me
rbinterface.Show
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  rbinterface.Show
End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnupopup
End If
End Sub
