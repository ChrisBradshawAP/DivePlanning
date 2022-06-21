VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main 
   Caption         =   "Dive Planning System"
   ClientHeight    =   7650
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7995
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7650
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   7695
      Left            =   -120
      ScaleHeight     =   7635
      ScaleWidth      =   8235
      TabIndex        =   5
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton cmdpdbsingle 
         Caption         =   "Sequential Dive"
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
         Left            =   480
         TabIndex        =   8
         Top             =   5280
         Width           =   1935
      End
      Begin VB.CommandButton cmdsingle 
         Caption         =   "Single Dive"
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
         Left            =   480
         TabIndex        =   7
         Top             =   6000
         Width           =   1935
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Exit"
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
         Left            =   480
         TabIndex        =   6
         Top             =   6720
         Width           =   1935
      End
      Begin VB.Label lblprogress 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   2520
         TabIndex        =   9
         Top             =   3360
         Width           =   4935
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   480
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   8760
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   8520
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Menu mnuDive 
      Caption         =   "&Dive"
      Begin VB.Menu mnuDsequential 
         Caption         =   "&Sequential"
      End
      Begin VB.Menu mnusingle 
         Caption         =   "S&ingle"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnugas 
      Caption         =   "&Gas"
      Begin VB.Menu mnusystem 
         Caption         =   "&System Default"
      End
      Begin VB.Menu mnufacdefault 
         Caption         =   "&Factory Default"
      End
   End
   Begin VB.Menu mnuutuility 
      Caption         =   "&Utility"
      Begin VB.Menu mnudisplay 
         Caption         =   "&Display format"
      End
      Begin VB.Menu mnulogo 
         Caption         =   "Add &Logo"
      End
      Begin VB.Menu mnutooltips 
         Caption         =   "Tool Tips"
         Begin VB.Menu mnutooltipson 
            Caption         =   "&On"
         End
         Begin VB.Menu mnutooltipsoff 
            Caption         =   "&Off"
         End
      End
      Begin VB.Menu mnumsgbox 
         Caption         =   "Message Box"
         Begin VB.Menu mnumsgboxon 
            Caption         =   "On"
         End
         Begin VB.Menu mnumsgboxoff 
            Caption         =   "Off"
         End
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Y(20) As Integer
Dim X As Integer
Dim fp As UserDocument
Dim CB As Byte
Dim C As String
Dim c1 As Byte
Dim i As Integer
Dim T As Integer
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
Dim j As Integer
Dim txt3(30) As String
Dim txt2(40) As String

Dim F1 As String


Private Sub CMDCLOSE_Click()
Unload Me
End
End Sub

Private Sub cmdpdbsingle_Click()
Unload Me
previousform = "Main"
Splanmain.Show
End Sub


Private Sub cmdsingle_Click()
planmain.Show
End Sub



Private Sub cmdviewvr3_Click()

End Sub

Private Sub Command2_Click()
  ans = Shell("pc_link.exe", vbNormalFocus)
End Sub

Private Sub Command3_Click()
 'frmdownload.Show
  Dim fileflags As FileOpenConstants
On Error GoTo ErrorHandler2
 Dim filefilter As String
 'Set the text in the dialog title bar
 CommonDialog1.DialogTitle = "Open"
 'Set the default file name and filter
 CommonDialog1.InitDir = "\"
 CommonDialog1.FileName = ""
 filefilter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
 CommonDialog1.Filter = filefilter
 CommonDialog1.FilterIndex = 0
 'Verify that the file exists
 fileflags = cdlOFNFileMustExist
 CommonDialog1.Flags = fileflags
 'Show the Open common dialog box
 CommonDialog1.ShowOpen
 'Return the path and file name selected or
 'Return an empty string if the user cancels the dialog
 test = CommonDialog1.FileName
 INITIALISE
Screen.MousePointer = 11
f = test
Open f For Binary As #1

profilefound = 0
maxdprofile = 0
a = 0
' ' is 44
     'If CB = 13 Then
'  Text1.Text = ""
  j = 0
  ypmax1 = 1
  ypmax2 = 1
  ypmax3 = 1
  lpp = 0
  For LP = 0 To 999 '99
   Get #1, , CB
      Text1.Text = Text1.Text + Chr$(CB)
      'Store Version
      If InStr(Text1.Text, "ver=") Then
         Text1.Text = ""
            For i = 1 To 13
               Get #1, , CB
               If CInt(CB) <> 13 Then
                 TEMPVERSION = TEMPVERSION + Chr$(CB)
               End If
               If CB = 13 Then
                 i = 12
                 Newrecord
                 updatelocation
               End If
            Next i
        
      End If
      
      'store Interval
      If InStr(Text1.Text, "Recint=") Then
        Text1.Text = ""
          For i = 1 To 5
             Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempinterval = tempinterval + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 4
                    tempinterval = Trim(tempinterval)
                    updateinterval
                 End If
               Next i
       End If
      
      'Store start date
      If InStr(Text1.Text, "Start") Then
         Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) <= 57 And CInt(CB) > 47 Then
                    tempstartdate = tempstartdate + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 12
                    stdategenerate
                  End If
               Next i
            End If
         Next K
      End If
      
      'Read Finished Time Info
      If InStr(Text1.Text, "Finish") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) <= 57 And CInt(CB) > 47 Then
                    tempfinishdate = tempfinishdate + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    fdategenerate
                    checkduration
                 End If
               Next i
            End If
         Next K
      End If
      If InStr(Text1.Text, "Status") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempstatus = tempstatus + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    updatestatus
                  End If
               Next i
            End If
         Next K
      End If
                 
      If InStr(Text1.Text, "Descend") Then
        Text1.Text = ""
         For K = 1 To 11
            Get #1, , CB
            
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempdescend = tempdescend + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    updatedescend
                  End If
               Next i
            End If
         Next K
      End If
                 
      If InStr(Text1.Text, "MaxD") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempmaxd = tempmaxd + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    updatemaxdepth
                  End If
               Next i
            End If
         Next K
      End If
      
      If InStr(Text1.Text, "OTU") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempotu = tempotu + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 12
                    updateotu
                 End If
               Next i
            End If
         Next K
      End If
      
      'Read gas
      If InStr(Text1.Text, "Gas") Then
        Text1 = ""
        i = 0
        gasprofile = False
        While gasprofile = False
          Get #1, , CB
           If CB = 13 Then
              gasprofile = True
              updategasprofile
            End If
            If CB = 44 Then
               test = txt2(i)
               i = i + 1
            Else
               txt2(i) = txt2(i) + Chr$(CB)
            End If
         Wend
       End If
       
       
       'Tissue status
       If InStr(Text1.Text, "Tissue") Then
        Text1 = ""
        'Get #1, , CB
        i = 0
        tissueprofile = False
        While tissueprofile = False
          Get #1, , CB
            If CB = 13 Then
              tissueprofile = True
              updatetissueprofile
            End If
            If CB = 44 Then
               i = i + 1
            Else
               txt2(i) = txt2(i) + Chr$(CB)
            End If
         Wend
       End If
       
             
        'Profile
       If InStr(Text1.Text, "Profile") Then
        K = 0
        Text1 = ""
        Get #1, , CB
        i = 0
        Diveprofile = False
        While Diveprofile = False
          Get #1, , CB
            If CB = 13 Then
            '  diveprofile = True
              'txt2(I) = txt2(I) + Chr$(CB)
              For K = 0 To i
                 
                 Select Case K
                 
                 Case 0
                      SQL = "SELECT * FROM profile"
                      Set RS = DB.OpenRecordset(SQL)
                      RS.AddNew
                      RS!DiveID = tempserialno
                      recordid = txt2(K)
                      For i = 1 To Len(recordid)
                        If Mid$(recordid, i, 1) = Chr(13) Or Mid$(recordid, i, 1) = Chr(32) Or Mid$(recordid, i, 1) = Chr(9) Or Mid$(recordid, i, 1) = Chr(10) Then
                        Else
                        Buff = Buff & Mid$(recordid, i, 1)
                        End If
                      Next
                      recordid = Buff
                      RS!ITEM1 = recordid
                      test = recordid
                      RS.Update
                      Buff = ""
                   
                 Case 1
                   If IsNumeric(CInt(txt2(K))) = True And Len(txt2(K)) = 4 Then
                      tempcode = CInt(txt2(K)) * 0.09375
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Depth : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM2 = tempcode
                   RS.Update
                   tempcode = tempcode * 3.28084
                   tempcode = Format(tempcode, "###.00")
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM3 = tempcode
                   RS.Update
                 Case 2
                      tempcode = txt2(K)
                      For i = 1 To Len(tempcode)
                        If Mid$(tempcode, i, 1) = Chr(13) Or Mid$(tempcode, i, 1) = Chr(32) Or Mid$(tempcode, i, 1) = Chr(9) Or Mid$(tempcode, i, 1) = Chr(10) Then
                        Else
                        Buff = Buff & Mid$(tempcode, i, 1)
                        End If
                      Next
                      tempcode = Buff
                      Buff = ""
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM4 = tempcode
                   RS.Update
                 Case 3
                   If IsNumeric(CInt(txt2(K))) = True And Len(txt2(K)) = 4 Then
                      tempcode = CInt(txt2(K)) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PO2 : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM5 = tempcode
                   RS.Update
                   
                 Case 4
                    
                   If Left(txt2(K), 1) = "M" And Len(txt2(K)) = 5 Then
                      tempcode = Right(txt2(K), 1)
                   Else
                      errmsg = errmsg & Chr(13) & "Mark : " & recordid
                   End If
                   If tempcode <> "1" Then
                      tempcode = "0"
                    End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM6 = tempcode
                   RS.Update
                 Case 5
                   tempcode = txt2(K)
                   tempcodeleft = Left(tempcode, 1)
                   tempcoderight = Right(tempcode, 4)
                   If IsNumeric(tempcoderight) = True And Len(txt2(K)) = 5 And tempcodeleft = "T" Then
                      tempcoderight = CInt(tempcoderight)
                      tempcoderightf = ((tempcoderight * 9) / 5) + 32
                      tempcoderightf = Format(tempcoderightf, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Temperature : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM7 = tempcoderight
                   RS!ITEM8 = tempcoderightf
                   RS.Update
                   If ITEM1 = test Then
                      SQL = "SELECT * FROM main "
                      SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                      Set RS = DB.OpenRecordset(SQL)
                      RS.Edit
                      RS!temperature = tempcoderight
                      RS.Update
                   End If
                 Case 6
                   If Left(txt2(K), 1) = "A" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 10
                      tempcode2 = tempcode * 3.28084
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "External Depth : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM9 = tempcode
                   RS!ITEM10 = tempcode2
                   RS.Update
                   tempsysid = "Rebrether"
                   updatesystemid
                 Case 7   ' vbattery
                   If Left(txt2(K), 1) = "B" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 156
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Val Battery : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM11 = tempcode
                   RS.Update
                 Case 8   ' ebattery
                   If Left(txt2(K), 1) = "C" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 156
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Electronic Battery : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM12 = tempcode
                   RS.Update
                 Case 9  'hp Diluent
                    If Left(txt2(K), 1) = "D" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode)
                      tempcode2 = tempcode * 14.7
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = "Error in hp Diluent"
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM13 = tempcode
                   RS!ITEM14 = tempcode2
                   RS.Update
                 Case 10 'HP O2
                   If Left(txt2(K), 1) = "E" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode)
                      tempcode2 = tempcode * 14.7
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "HP02 : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM15 = tempcode
                   RS!ITEM16 = tempcode2
                   RS.Update
                 Case 11 'PPO2 A
                   If Left(txt2(K), 1) = "F" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 A : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM17 = tempcode
                   RS.Update
                 Case 12 'PPO2 b
                   If Left(txt2(K), 1) = "G" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 B : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM18 = tempcode
                   RS.Update
                 Case 13 'PPO2 C
                   If Left(txt2(K), 1) = "H" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 C : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM19 = tempcode
                   RS.Update
                 Case 14 'To 25 'nick change this and case to 30
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item20 = txt2(K)
                   RS.Update
                 Case 15
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item21 = txt2(K)
                   RS.Update
                 Case 16
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item22 = txt2(K)
                   RS.Update
                 Case 17
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item23 = txt2(K)
                   RS.Update
                 Case 18
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item24 = txt2(K)
                   RS.Update
                 Case 19
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item25 = txt2(K)
                   RS.Update
                 Case 20
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item26 = txt2(K)
                   RS.Update
                 Case 21
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item27 = txt2(K)
                   RS.Update
                 Case 22
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item28 = txt2(K)
                   RS.Update
                  Case 23
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item29 = txt2(K)
                   RS.Update
                 Case 24
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item30 = txt2(K)
                   RS.Update
                 Case 25
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item31 = txt2(K)
                   RS.Update
                 Case 26
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item32 = txt2(K)
                   RS.Update
                 Case 27
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item33 = txt2(K)
                   RS.Update
                 Case 28 To 44
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item34 = txt2(K)
                   RS.Update
               End Select
             Next K
                i = 0
                 cleartext2
                 'txt2.Text = ""
            End If
          
            If CB = 44 Then
               i = i + 1
            Else
               txt2(i) = txt2(i) + Chr$(CB)
               test = txt2(i)
               
            End If
             Text1.Text = Text1.Text + Chr$(CB)
            If InStr(Text1.Text, "End") Then
              Screen.MousePointer = 0
              Diveprofile = True
            End If
         Wend
       End If
       
Next LP

Unload Me
If Trim(errmsg) <> "" Then
  MsgBox "System found the following Error during downloading : " & errmsg
End If
Screen.MousePointer = 0
Close #1
rbinterface.Show
ErrorHandler2:
   Screen.MousePointer = 0
End Sub

Private Sub Command4_Click()
Dim destinationsource As String

destinationsource = App.Path & "\old\"
File1.Path = App.Path & "\new\"
File1.Pattern = "*.txt"
If File1.ListCount > 0 Then
  For i = 1 To File1.ListCount
    File1.ListIndex = i - 1
    tempfileselected = File1.FileName
    List1.AddItem File1.FileName
    Source = File1.Path & "\" & tempfileselected
    destinationsource = App.Path & "\old\" & tempfileselected
    FileCopy Source, destinationsource
  Next i
totalfile = File1.ListCount

ans = MsgBox(totalfile & " Dives to download... " & Chr(13) & "Do you want to download all the Dives ?", vbYesNo + vbCritical + vbDefaultButton2, "Comfirmation")
Select Case ans
Case vbYes
Me.MousePointer = 11
For T = 1 To File1.ListCount
   File1.ListIndex = T - 1
   tempfileselected = File1.FileName
   List1.AddItem File1.FileName
   Source = File1.Path & "\" & tempfileselected
   f = Source
INITIALISE
lblprogress = "System downloading : " & tempfileselected & " " & File1.ListIndex + 1 & " of " & File1.ListCount & " files"
Cls
'lblprogress = "System downloading : " & File1.ListIndex & " of " & File1.ListCount - 1 & tempfilename
Open f For Binary As #1
profilefound = 0
maxdprofile = 0

a = 0
' ' is 44
     'If CB = 13 Then
'  Text1.Text = ""
  j = 0
  ypmax1 = 1
  ypmax2 = 1
  ypmax3 = 1
  lpp = 0
  For LP = 0 To 999 '99
   Get #1, , CB
      Text1.Text = Text1.Text + Chr$(CB)
      'Store Version
      If InStr(Text1.Text, "ver=") Then
         Text1.Text = ""
            For i = 1 To 13
               Get #1, , CB
               TEMPVERSION = TEMPVERSION + Chr$(CB)
               If CB = 13 Then
                 i = 12
                 Newrecord
                 updatelocation
               End If
            Next i
        
      End If
      
      'store Interval
      If InStr(Text1.Text, "Recint=") Then
        Text1.Text = ""
          For i = 1 To 5
             Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempinterval = tempinterval + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 4
                    tempinterval = Trim(tempinterval)
                    updateinterval
                 End If
               Next i
       End If
      
      'Store start date
      If InStr(Text1.Text, "Start") Then
         Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) <= 57 And CInt(CB) > 47 Then
                    tempstartdate = tempstartdate + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 12
                    stdategenerate
                  End If
               Next i
            End If
         Next K
      End If
      
      'Read Finished Time Info
      If InStr(Text1.Text, "Finish") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) <= 57 And CInt(CB) > 47 Then
                    tempfinishdate = tempfinishdate + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    fdategenerate
                    checkduration
                 End If
               Next i
            End If
         Next K
      End If
                 
      If InStr(Text1.Text, "Status") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempstatus = tempstatus + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    updatestatus
                  End If
               Next i
            End If
         Next K
      End If
                 
      If InStr(Text1.Text, "Descend") Then
        Text1.Text = ""
         For K = 1 To 11
            Get #1, , CB
            
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempdescend = tempdescend + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    updatedescend
                  End If
               Next i
            End If
         Next K
      End If
                 
      If InStr(Text1.Text, "MaxD") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempmaxd = tempmaxd + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    updatemaxdepth
                  End If
               Next i
            End If
         Next K
      End If
      
      If InStr(Text1.Text, "OTU") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempotu = tempotu + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 12
                    updateotu
                 End If
               Next i
            End If
         Next K
      End If
      
      'Read gas
      If InStr(Text1.Text, "Gas") Then
        Text1 = ""
        i = 0
        gasprofile = False
        While gasprofile = False
          Get #1, , CB
           If CB = 13 Then
              gasprofile = True
              updategasprofile
            End If
            If CB = 44 Then
               test = txt2(i)
               i = i + 1
            Else
               txt2(i) = txt2(i) + Chr$(CB)
            End If
         Wend
       End If
       
       
       'Tissue status
       If InStr(Text1.Text, "Tissue") Then
        Text1 = ""
        'Get #1, , CB
        i = 0
        tissueprofile = False
        While tissueprofile = False
          Get #1, , CB
            If CB = 13 Then
              tissueprofile = True
              updatetissueprofile
            End If
            If CB = 44 Then
               i = i + 1
            Else
               txt2(i) = txt2(i) + Chr$(CB)
            End If
         Wend
       End If
       
             
        'Profile
       If InStr(Text1.Text, "Profile") Then
        K = 0
        Text1 = ""
        Get #1, , CB
        i = 0
        Diveprofile = False
        While Diveprofile = False
          Get #1, , CB
            If CB = 13 Then
            '  diveprofile = True
              'txt2(I) = txt2(I) + Chr$(CB)
              For K = 0 To i
                 
                 Select Case K
                 
                 Case 0
                      SQL = "SELECT * FROM profile"
                      Set RS = DB.OpenRecordset(SQL)
                      RS.AddNew
                      RS!DiveID = tempserialno
                      recordid = txt2(K)
                      For i = 1 To Len(recordid)
                        If Mid$(recordid, i, 1) = Chr(13) Or Mid$(recordid, i, 1) = Chr(32) Or Mid$(recordid, i, 1) = Chr(9) Or Mid$(recordid, i, 1) = Chr(10) Then
                        Else
                        Buff = Buff & Mid$(recordid, i, 1)
                        End If
                      Next
                      recordid = Buff
                      RS!ITEM1 = recordid
                      RS.Update
                      test = recordid
                      Buff = ""
                   
                 Case 1
                   If IsNumeric(CInt(txt2(K))) = True And Len(txt2(K)) = 4 Then
                      tempcode = CInt(txt2(K)) * 0.09375
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Depth : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM2 = tempcode
                   RS.Update
                   tempcode = tempcode * 3.28084
                   tempcode = Format(tempcode, "###.00")
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM3 = tempcode
                   RS.Update
                 Case 2
                      tempcode = txt2(K)
                      For i = 1 To Len(tempcode)
                        If Mid$(tempcode, i, 1) = Chr(13) Or Mid$(tempcode, i, 1) = Chr(32) Or Mid$(tempcode, i, 1) = Chr(9) Or Mid$(tempcode, i, 1) = Chr(10) Then
                        Else
                        Buff = Buff & Mid$(tempcode, i, 1)
                        End If
                      Next
                      tempcode = Buff
                      Buff = ""
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM4 = tempcode
                   RS.Update
                 Case 3
                   If IsNumeric(CInt(txt2(K))) = True And Len(txt2(K)) = 4 Then
                      tempcode = CInt(txt2(K)) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PO2 : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM5 = tempcode
                   RS.Update
                   
                 Case 4
                    
                   If Left(txt2(K), 1) = "M" And Len(txt2(K)) = 5 Then
                      tempcode = Right(txt2(K), 1)
                   Else
                      errmsg = errmsg & Chr(13) & "Mark : " & recordid
                   End If
                   If tempcode <> "1" Then
                      tempcode = "0"
                    End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM6 = tempcode
                   RS.Update
                 Case 5
                   tempcode = txt2(K)
                   tempcodeleft = Left(tempcode, 1)
                   tempcoderight = Right(tempcode, 4)
                   If IsNumeric(tempcoderight) = True And Len(txt2(K)) = 5 And tempcodeleft = "T" Then
                      tempcoderight = CInt(tempcoderight)
                      tempcoderightf = ((tempcoderight * 9) / 5) + 32
                      tempcoderightf = Format(tempcoderightf, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Temperature : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM7 = tempcoderight
                   RS!ITEM8 = tempcoderightf
                   RS.Update
                   If ITEM1 = test Then
                      SQL = "SELECT * FROM main "
                      SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                      Set RS = DB.OpenRecordset(SQL)
                      RS.Edit
                      RS!temperature = tempcoderight
                      RS.Update
                   End If
                 Case 6
                   If Left(txt2(K), 1) = "A" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 10
                      tempcode2 = tempcode * 3.28084
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "External Depth : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM9 = tempcode
                   RS!ITEM10 = tempcode2
                   RS.Update
                   tempsysid = "Rebrether"
                   updatesystemid
                 Case 7   ' vbattery
                   If Left(txt2(K), 1) = "B" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 156
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Val Battery : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM11 = tempcode
                   RS.Update
                 Case 8   ' ebattery
                   If Left(txt2(K), 1) = "C" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 156
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Electronic Battery : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM12 = tempcode
                   RS.Update
                 Case 9  'hp Diluent
                    If Left(txt2(K), 1) = "D" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode)
                      tempcode2 = tempcode * 14.7
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = "Error in hp Diluent"
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM13 = tempcode
                   RS!ITEM14 = tempcode2
                   RS.Update
                 Case 10 'HP O2
                   If Left(txt2(K), 1) = "E" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode)
                      tempcode2 = tempcode * 14.7
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "HP02 : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM15 = tempcode
                   RS!ITEM16 = tempcode2
                   RS.Update
                 Case 11 'PPO2 A
                   If Left(txt2(K), 1) = "F" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 A : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM17 = tempcode
                   RS.Update
                 Case 12 'PPO2 b
                   If Left(txt2(K), 1) = "G" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 B : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM18 = tempcode
                   RS.Update
                 Case 13 'PPO2 C
                   If Left(txt2(K), 1) = "H" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 C : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM19 = tempcode
                   RS.Update
                 Case 14 'To 25 'nick change this and case to 30
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item20 = txt2(K)
                   RS.Update
                 Case 15
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item21 = txt2(K)
                   RS.Update
                 Case 16
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item22 = txt2(K)
                   RS.Update
                 Case 17
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item23 = txt2(K)
                   RS.Update
                 Case 18
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item24 = txt2(K)
                   RS.Update
                 Case 19
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item25 = txt2(K)
                   RS.Update
                 Case 20
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item26 = txt2(K)
                   RS.Update
                 Case 21
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item27 = txt2(K)
                   RS.Update
                 Case 22
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item28 = txt2(K)
                   RS.Update
                  Case 23
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item29 = txt2(K)
                   RS.Update
                 Case 24
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item30 = txt2(K)
                   RS.Update
                 Case 25
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item31 = txt2(K)
                   RS.Update
                 Case 26
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item32 = txt2(K)
                   RS.Update
                 Case 27
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item33 = txt2(K)
                   RS.Update
                 Case 28 To 44
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item34 = txt2(K)
                   RS.Update
               End Select
             Next K
                i = 0
                 cleartext2
                 'txt2.Text = ""
            End If
          
            If CB = 44 Then
               i = i + 1
               txt2(i) = "" 'nick
            Else
               txt2(i) = txt2(i) + Chr$(CB)
               test = txt2(i)
               
            End If
             Text1.Text = Text1.Text + Chr$(CB)
            If InStr(Text1.Text, "End") Then
              Screen.MousePointer = 0
              Diveprofile = True
            End If
         Wend
       End If
       
Next LP

'Unload Me
'If Trim(errmsg) <> "" Then
'  MsgBox "System found the following Error during downloading : " & errmsg
'End If
'Screen.MousePointer = 0
Close #1
Next T
For i = 1 To File1.ListCount
   File1.ListIndex = i - 1
   tempfileselected = File1.FileName
   List1.AddItem File1.FileName
   Source = File1.Path & "\" & tempfileselected
   Kill Source
Next i
   lblprogress = ""
   MsgBox "Download completed."
Case vbNo
  Me.MousePointer = 0
End Select
Else 'no files
  ans = MsgBox("No new dives", vbOKOnly, "Download")
End If
Me.MousePointer = 0
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Set DB = OpenDatabase(App.Path & "/planmain.mdb")
Dir1.Path = App.Path 'Drive1.Drive   ' Set directory path.
Me.Top = 30
SQL = "select * FROM dpserialno "
Set RS3 = DB.OpenRecordset(SQL)
temptooltips = RS3("dptooltips")
mnutooltipsClear
If temptooltips = "On" Then
   mnutooltipson.Checked = True
Else
   mnutooltipsoff.Checked = True
End If
tempmsgbox = RS3("dpmsgbox")
mnumsgboxClear
If tempmsgbox = "On" Then
   mnumsgboxon.Checked = True
Else
   mnumsgboxoff.Checked = True
End If
End Sub

Private Sub mnucurrentmain_Click()
Unload Me
rbmain.Show
End Sub

Private Sub mnucurrentprofile_Click()

End Sub

Private Sub mnudisplay_Click()
Unload Me
frmdisplay.Show
End Sub

Private Sub mnuimpfile_Click()
  'frmdownload.Show
  Dim fileflags As FileOpenConstants
 Dim filefilter As String
 'Set the text in the dialog title bar
On Error GoTo ErrorHandler2
 CommonDialog1.DialogTitle = "Open"
 'Set the default file name and filter
 CommonDialog1.InitDir = "\"
 CommonDialog1.FileName = ""
 filefilter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
 CommonDialog1.Filter = filefilter
 CommonDialog1.FilterIndex = 0
 'Verify that the file exists
 fileflags = cdlOFNFileMustExist
 CommonDialog1.Flags = fileflags
 'Show the Open common dialog box
 CommonDialog1.ShowOpen
 'Return the path and file name selected or
 'Return an empty string if the user cancels the dialog
 test = CommonDialog1.FileName
 INITIALISE
Screen.MousePointer = 11
f = test
Open f For Binary As #1

profilefound = 0
maxdprofile = 0
a = 0
' ' is 44
     'If CB = 13 Then
'  Text1.Text = ""
  j = 0
  ypmax1 = 1
  ypmax2 = 1
  ypmax3 = 1
  lpp = 0
  For LP = 0 To 999 '99
   Get #1, , CB
      Text1.Text = Text1.Text + Chr$(CB)
      'Store Version
      If InStr(Text1.Text, "ver=") Then
         Text1.Text = ""
            For i = 1 To 13
               Get #1, , CB
               TEMPVERSION = TEMPVERSION + Chr$(CB)
               If CB = 13 Then
                 i = 12
                 Newrecord
                 updatelocation
               End If
            Next i
        
      End If
      
      'store Interval
      If InStr(Text1.Text, "Recint=") Then
        Text1.Text = ""
          For i = 1 To 5
             Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempinterval = tempinterval + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 4
                    tempinterval = Trim(tempinterval)
                    updateinterval
                 End If
               Next i
       End If
      
      'Store start date
      If InStr(Text1.Text, "Start") Then
         Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) <= 57 And CInt(CB) > 47 Then
                    tempstartdate = tempstartdate + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 12
                    stdategenerate
                  End If
               Next i
            End If
         Next K
      End If
      
      'Read Finished Time Info
      If InStr(Text1.Text, "Finish") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) <= 57 And CInt(CB) > 47 Then
                    tempfinishdate = tempfinishdate + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    fdategenerate
                    checkduration
                 End If
               Next i
            End If
         Next K
      End If
      If InStr(Text1.Text, "Status") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempstatus = tempstatus + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    updatestatus
                  End If
               Next i
            End If
         Next K
      End If
                 
      If InStr(Text1.Text, "Descend") Then
        Text1.Text = ""
         For K = 1 To 11
            Get #1, , CB
            
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempdescend = tempdescend + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    updatedescend
                  End If
               Next i
            End If
         Next K
      End If
                 
      If InStr(Text1.Text, "MaxD") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempmaxd = tempmaxd + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 13
                    updatemaxdepth
                  End If
               Next i
            End If
         Next K
      End If
      
      If InStr(Text1.Text, "OTU") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For i = 1 To 13
                 Get #1, , CB
                 If CInt(CB) = 46 Or (CInt(CB) <= 57 And CInt(CB) > 47) Then
                    tempotu = tempotu + Chr$(CB)
                 End If
                 If CB = 13 Then
                    i = 12
                    updateotu
                 End If
               Next i
            End If
         Next K
      End If
      
      'Read gas
      If InStr(Text1.Text, "Gas") Then
        Text1 = ""
        i = 0
        gasprofile = False
        While gasprofile = False
          Get #1, , CB
           If CB = 13 Then
              gasprofile = True
              updategasprofile
            End If
            If CB = 44 Then
               test = txt2(i)
               i = i + 1
            Else
               txt2(i) = txt2(i) + Chr$(CB)
            End If
         Wend
       End If
       
       
       'Tissue status
       If InStr(Text1.Text, "Tissue") Then
        Text1 = ""
        'Get #1, , CB
        i = 0
        tissueprofile = False
        While tissueprofile = False
          Get #1, , CB
            If CB = 13 Then
              tissueprofile = True
              updatetissueprofile
            End If
            If CB = 44 Then
               i = i + 1
            Else
               txt2(i) = txt2(i) + Chr$(CB)
            End If
         Wend
       End If
       
             
        'Profile
       If InStr(Text1.Text, "Profile") Then
        K = 0
        Text1 = ""
        Get #1, , CB
        i = 0
        Diveprofile = False
        While Diveprofile = False
          Get #1, , CB
            If CB = 13 Then
            '  diveprofile = True
              'txt2(I) = txt2(I) + Chr$(CB)
              For K = 0 To i
                 
                 Select Case K
                 
                 Case 0
                      SQL = "SELECT * FROM profile"
                      Set RS = DB.OpenRecordset(SQL)
                      RS.AddNew
                      RS!DiveID = tempserialno
                      recordid = txt2(K)
                      For i = 1 To Len(recordid)
                        If Mid$(recordid, i, 1) = Chr(13) Or Mid$(recordid, i, 1) = Chr(32) Or Mid$(recordid, i, 1) = Chr(9) Or Mid$(recordid, i, 1) = Chr(10) Then
                        Else
                        Buff = Buff & Mid$(recordid, i, 1)
                        End If
                      Next
                      recordid = Buff
                      RS!ITEM1 = recordid
                      test = recordid
                      RS.Update
                      Buff = ""
                   
                 Case 1
                   If IsNumeric(CInt(txt2(K))) = True And Len(txt2(K)) = 4 Then
                      tempcode = CInt(txt2(K)) * 0.09375
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Depth : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM2 = tempcode
                   RS.Update
                   tempcode = tempcode * 3.28084
                   tempcode = Format(tempcode, "###.00")
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM3 = tempcode
                   RS.Update
                 Case 2
                      tempcode = txt2(K)
                      For i = 1 To Len(tempcode)
                        If Mid$(tempcode, i, 1) = Chr(13) Or Mid$(tempcode, i, 1) = Chr(32) Or Mid$(tempcode, i, 1) = Chr(9) Or Mid$(tempcode, i, 1) = Chr(10) Then
                        Else
                        Buff = Buff & Mid$(tempcode, i, 1)
                        End If
                      Next
                      tempcode = Buff
                      Buff = ""
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM4 = tempcode
                   RS.Update
                 Case 3
                   If IsNumeric(CInt(txt2(K))) = True And Len(txt2(K)) = 4 Then
                      tempcode = CInt(txt2(K)) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PO2 : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM5 = tempcode
                   RS.Update
                   
                 Case 4
                    
                   If Left(txt2(K), 1) = "M" And Len(txt2(K)) = 5 Then
                      tempcode = Right(txt2(K), 1)
                   Else
                      errmsg = errmsg & Chr(13) & "Mark : " & recordid
                   End If
                   If tempcode <> "1" Then
                      tempcode = "0"
                    End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM6 = tempcode
                   RS.Update
                 Case 5
                   tempcode = txt2(K)
                   tempcodeleft = Left(tempcode, 1)
                   tempcoderight = Right(tempcode, 4)
                   If IsNumeric(tempcoderight) = True And Len(txt2(K)) = 5 And tempcodeleft = "T" Then
                      tempcoderight = CInt(tempcoderight)
                      tempcoderightf = ((tempcoderight * 9) / 5) + 32
                      tempcoderightf = Format(tempcoderightf, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Temperature : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM7 = tempcoderight
                   RS!ITEM8 = tempcoderightf
                   RS.Update
                   If ITEM1 = test Then
                      SQL = "SELECT * FROM main "
                      SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                      Set RS = DB.OpenRecordset(SQL)
                      RS.Edit
                      RS!temperature = tempcoderight
                      RS.Update
                   End If
                 Case 6
                   If Left(txt2(K), 1) = "A" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 10
                      tempcode2 = tempcode * 3.28084
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "External Depth : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM9 = tempcode
                   RS!ITEM10 = tempcode2
                   RS.Update
                   tempsysid = "Rebrether"
                   updatesystemid
                 Case 7   ' vbattery
                   If Left(txt2(K), 1) = "B" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 156
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Val Battery : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM11 = tempcode
                   RS.Update
                 Case 8   ' ebattery
                   If Left(txt2(K), 1) = "C" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 156
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Electronic Battery : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM12 = tempcode
                   RS.Update
                 Case 9  'hp Diluent
                    If Left(txt2(K), 1) = "D" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode)
                      tempcode2 = tempcode * 14.7
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = "Error in hp Diluent"
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM13 = tempcode
                   RS!ITEM14 = tempcode2
                   RS.Update
                 Case 10 'HP O2
                   If Left(txt2(K), 1) = "E" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode)
                      tempcode2 = tempcode * 14.7
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "HP02 : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM15 = tempcode
                   RS!ITEM16 = tempcode2
                   RS.Update
                 Case 11 'PPO2 A
                   If Left(txt2(K), 1) = "F" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 A : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM17 = tempcode
                   RS.Update
                 Case 12 'PPO2 b
                   If Left(txt2(K), 1) = "G" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 B : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM18 = tempcode
                   RS.Update
                 Case 13 'PPO2 C
                   If Left(txt2(K), 1) = "H" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 C : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM19 = tempcode
                   RS.Update
                 Case 14 'To 25 'nick change this and case to 30
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item20 = txt2(K)
                   RS.Update
                 Case 15
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item21 = txt2(K)
                   RS.Update
                 Case 16
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item22 = txt2(K)
                   RS.Update
                 Case 17
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item23 = txt2(K)
                   RS.Update
                 Case 18
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item24 = txt2(K)
                   RS.Update
                 Case 19
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item25 = txt2(K)
                   RS.Update
                 Case 20
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item26 = txt2(K)
                   RS.Update
                 Case 21
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item27 = txt2(K)
                   RS.Update
                 Case 22
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item28 = txt2(K)
                   RS.Update
                  Case 23
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item29 = txt2(K)
                   RS.Update
                 Case 24
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item30 = txt2(K)
                   RS.Update
                 Case 25
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item31 = txt2(K)
                   RS.Update
                 Case 26
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item32 = txt2(K)
                   RS.Update
                 Case 27
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item33 = txt2(K)
                   RS.Update
                 Case 28 To 44
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item34 = txt2(K)
                   RS.Update
               End Select
             Next K
                i = 0
                 cleartext2
                 'txt2.Text = ""
            End If
          
            If CB = 44 Then
               i = i + 1
            Else
               txt2(i) = txt2(i) + Chr$(CB)
               test = txt2(i)
               
            End If
             Text1.Text = Text1.Text + Chr$(CB)
            If InStr(Text1.Text, "End") Then
              Screen.MousePointer = 0
              Diveprofile = True
            End If
         Wend
       End If
       
Next LP

Unload Me
If Trim(errmsg) <> "" Then
  MsgBox "System found the following Error during downloading : " & errmsg
End If
Screen.MousePointer = 0
Close #1
rbinterface.Show
ErrorHandler2:
   Screen.MousePointer = 0
End Sub



Private Sub mnuopen_Click()
On Error GoTo ErrorHandler2
Dim fileflags As FileOpenConstants
 Dim filefilter As String
 'Set the text in the dialog title bar
 CommonDialog1.DialogTitle = "Open"
 'Set the default file name and filter
 CommonDialog1.InitDir = "\"
 CommonDialog1.FileName = ""
 filefilter = "mdb (*.mdb)|*.mdb|All Files (*.*)|*.*"
 CommonDialog1.Filter = filefilter
 CommonDialog1.FilterIndex = 0
 'Verify that the file exists
 fileflags = cdlOFNFileMustExist
 CommonDialog1.Flags = fileflags
 'Show the Open common dialog box
 CommonDialog1.ShowOpen
 'Return the path and file name selected or
 'Return an empty string if the user cancels the dialog
 filesource = CommonDialog1.FileName
Screen.MousePointer = 11
rbmain.Show
ErrorHandler2:
If CInt(Err.Number) = 75 Then
   Screen.MousePointer = 0
   Exit Sub
Else
   Screen.MousePointer = 0
End If
End Sub
Private Function Newrecord()
Dim Buff As String
TEMPVERSION = Trim(TEMPVERSION)
SQL = "SELECT * FROM SERIALNO"
Set RS = DB.OpenRecordset(SQL)
tempserialno2 = RS("serial_no")
tempserialno = Right(tempserialno2, 7)
tempserialno = CInt(tempserialno) + 1
lengthsn = Len(tempserialno)

Select Case lengthsn
Case 1
   tempserialno = "D000000" & tempserialno
Case 2
   tempserialno = "D00000" & tempserialno
Case 3
   tempserialno = "D0000" & tempserialno
Case 4
   tempserialno = "D000" & tempserialno
Case 5
   tempserialno = "D00" & tempserialno
Case 6
   tempserialno = "D0" & tempserialno
Case 7
   tempserialno = "D" & tempserialno
End Select
SQL = "SELECT * FROM serialno "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!serial_no = tempserialno
RS.Update
SQL = "SELECT * FROM MAIN"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!DiveID = tempserialno
'RS!recordid = "R0001"
RS!Version = TEMPVERSION

RS.Update

Text1.Text = ""
End Function
Private Function Newrecord2()


Dim Buff As String
TEMPVERSION = Trim(TEMPVERSION)
SQL = "SELECT * FROM SERIALNO"
Set RS = DB.OpenRecordset(SQL)
tempserialno2 = RS("serial_no")
tempserialno = Right(tempserialno2, 7)
tempserialno = CInt(tempserialno) + 1
lengthsn = Len(tempserialno)

Select Case lengthsn
Case 1
   tempserialno = "D000000" & tempserialno
Case 2
   tempserialno = "D00000" & tempserialno
Case 3
   tempserialno = "D0000" & tempserialno
Case 4
   tempserialno = "D000" & tempserialno
Case 5
   tempserialno = "D00" & tempserialno
Case 6
   tempserialno = "D0" & tempserialno
Case 7
   tempserialno = "D" & tempserialno
End Select
SQL = "SELECT * FROM serialno "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!serial_no = tempserialno
RS.Update
SQL = "SELECT * FROM MAIN"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!DiveID = tempserialno
'RS!recordid = "R0001"

RS.Update

Text1.Text = ""
End Function
Private Function updateinterval()
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!Interval = tempinterval
RS.Update
Text1.Text = ""
End Function
Private Function updateversion()
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!Version = tempstring
RS.Update
'Text1.Text = ""
End Function

Private Function stdategenerate()
tempstartdate = Trim(tempstartdate)
tempstartdate2 = tempstartdate
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
dateconvert
RS!StartDate = tempstartdate
RS.Update
Text1.Text = ""
End Function
Private Function updatemaxdepth()
tempmaxd = Trim(tempmaxd)
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!MaxDepth = tempmaxd
RS.Update
Text1.Text = ""
End Function
Private Function updateotu()
tempotu = Trim(tempotu)
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!otu = tempotu
RS.Update
End Function
Private Function updatesystemid()
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!systemid = tempsysid
RS.Update
End Function
Private Function updatestatus()
tempstatus = Trim(tempstatus)
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!Status = tempstatus
RS.Update
End Function
Private Function updatedescend()
tempdescend = Trim(tempdescend)
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!descend = tempdescend
RS.Update
End Function
Private Function updategasprofile()
txt2(i) = txt2(i) + Chr$(CB)
SQL = "SELECT * FROM GAS"
Set RS = DB.OpenRecordset(SQL)
tempgsstype = txt2(0)
tempGasN2 = txt2(1)
temphe = txt2(2)
tempMod = txt2(3)
tempstat = txt2(4)
RS.AddNew
RS!DiveID = tempserialno
RS!Gastype = tempgsstype
RS!GasN2 = tempGasN2
RS!GasHe = temphe
RS!Gasmod = tempMod
RS!Gasstat = tempstat
RS.Update
cleartext2
Text1.Text = ""
End Function
Private Function updatetissueprofile()
txt2(i) = txt2(i) + Chr$(CB)
SQL = "SELECT * FROM tissue"
temptsstype = txt2(0)
tempGasN2 = txt2(1)
temphe = txt2(2)
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!DiveID = tempserialno
RS!Tissuetype = temptsstype
RS!TissueN2 = tempGasN2
RS!TissueHe = temphe
RS.Update
cleartext2
Text1.Text = ""
End Function
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

Set rsc = Nothing
End Function

Private Function dateconvert()
tempstartday = Int(CInt(tempstartdate) / 86400)
remainder = tempstartdate - (tempstartday * 86400)
temphour = Int(remainder / 3600)
minutesremainder = remainder - (temphour * 3600)
tempminutes = Int(minutesremainder / 60)
secondremainder = minutesremainder - tempminutes * 60
tempstartdate = CInt(tempstartday)
oldstartdate = Format("01/01/1992", "mm/dd/yyyy")
testdate = DateAdd("d", tempstartdate, oldstartdate)
testdate = testdate & " " & temphour & ":" & tempminutes & ":" & secondremainder
tempstartdate = Format(testdate, "mm/dd/yyyy hh:mm:ss")

End Function
Private Function finisheddateconvert()
tempfinishday = Int(CInt(tempfinishdate) / 86400)
fremainder = tempfinishdate - (tempfinishday * 86400)
tempfhour = Int(fremainder / 3600)
minutesfremainder = fremainder - (tempfhour * 3600)
tempfminutes = Int(minutesfremainder / 60)
secondfremainder = minutesfremainder - tempfminutes * 60
tempfinishdate = CInt(tempfinishday)
oldfinishdate = Format("01/01/1992", "mm/dd/yyyy")
testdate = DateAdd("d", tempfinishdate, oldfinishdate)
testdate = testdate & " " & tempfhour & ":" & tempfminutes & ":" & secondfremainder
tempfinishdate = Format(testdate, "mm/dd/yyyy hh:mm:ss")

End Function
Private Function fdategenerate()
tempfinishdate = Trim(tempfinishdate)
tempfinishdate2 = tempfinishdate
SQL = "SELECT * FROM MAIN "
SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
finisheddateconvert
RS!Finisheddate = tempfinishdate
RS.Update
Text1.Text = ""
End Function
Private Sub cleartext2()
Dim ind As Integer
 For ind = 0 To 18
        txt2(ind) = ""
 Next ind
End Sub
Private Function checkduration()
tempduration = tempfinishdate2 - tempstartdate2
tempdminutes = Int(tempduration / 60)
tempduremainder = tempduration - tempdminutes * 60
tempduration = tempdminutes & ":" & tempduremainder
SQL = "SELECT * FROM MAIN "
SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!Duration = tempduration
RS.Update
End Function
Private Function INITIALISE()
tempstartdate = ""
tempfinishdate = ""
TEMPVERSION = ""
tempinterval = ""
tempmaxd = ""
tempotu = ""

End Function
Private Function updatelocation()
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!Location = tempfilename
RS.Update
'Text1.Text = ""
End Function
Private Function updatelocation2()
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = '" & Trim(tempserialno) & "' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!Location = tempfilename2
RS.Update
'Text1.Text = ""
End Function
Private Function getitemrecord()
SQL = "SELECT * FROM profile "
SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
Set RS = DB.OpenRecordset(SQL)
End Function

Private Sub mnupdb_Click()
  'frmdownload.Show
  Dim fileflags As FileOpenConstants
 
On Error GoTo ErrorHandler2
 Dim filefilter As String
 'Set the text in the dialog title bar
 CommonDialog1.DialogTitle = "Open"
 'Set the default file name and filter
 CommonDialog1.InitDir = "\"
 CommonDialog1.FileName = ""
 filefilter = "Pdb Files (*.pdb)|*.pdb|All Files (*.*)|*.*"
 CommonDialog1.Filter = filefilter
 CommonDialog1.FilterIndex = 0
 'Verify that the file exists
 fileflags = cdlOFNFileMustExist
 CommonDialog1.Flags = fileflags
 'Show the Open common dialog box
 CommonDialog1.ShowOpen
 'Return the path and file name selected or
 'Return an empty string if the user cancels the dialog
 tempfileselected = CommonDialog1.FileName
 templocationfile = tempfileselected
 tempdestfile = ""
 tempfilename2 = ""
 lengthfilename = Len(tempfileselected)
       For i = 1 To lengthfilename
          test = Mid(tempfileselected, i, 1)
          test3 = Asc(test)
          If test3 = 46 Then
             i = lengthfilename
          Else
             tempfilename2 = tempfilename2 & test
          End If
       Next i
       For i = 1 To lengthfilename
          test = Mid(tempfileselected, i, 1)
          test3 = Asc(test)
          If test3 = 92 Then
             tempdestfile = ""
          Else
             If test3 = 46 Then
                i = lengthfilename
             Else
                tempdestfile = tempdestfile & test
             End If
          End If
       Next i
       tempdestfile = tempdestfile & ".pdb"
       tempfileselected = tempfilename2 '
      Source = templocationfile
      destinationsource = App.Path & "\old\" & tempdestfile
      FileCopy Source, destinationsource
 INITIALISE
Screen.MousePointer = 11
dbdivelist = PDBOpen(Byfilename, tempfileselected, 0, 0, 0, 0, afModeReadWrite)
'If OpendivelistDatabase = True Then
   'INITIALISE
   PDBMoveFirst dbdivelist
   PDBGetField dbdivelist, 0, txtdiveid
   tempserialno3 = txtdiveid
   PDBGetField dbdivelist, 1, tempstring
   tempstring = Trim(tempstring)
   Newrecord
   tempfilename2 = templocationfile
   updatelocation2
   tempsysid = "VR3"
   updatesystemid
      While PDBEOF(dbdivelist) = False 'Repeat until EOF = True.
         v = v + 1
         PDBGetField dbdivelist, 0, txtdiveid
         PDBGetField dbdivelist, 1, tempstring
            If Trim(txtdiveid) <> Trim(tempserialno3) Then
               Newrecord
               tempfilename2 = templocationfile
               updatelocation2
               tempsysid = "VR3"
               updatesystemid
               tempserialno3 = txtdiveid
            Else
              ' PDBMoveNext (dbdivelist)
               
               tempstring = Trim(tempstring)
               lentempstring = Len(tempstring)
               If InStr(tempstring, "ver=") Then
                  TEMPVERSION = ""
                  For K = 1 To lentempstring
                     If K > 4 Then
                        test = Mid(tempstring, K, 1)
                        TEMPVERSION = TEMPVERSION + test
                     End If
                  Next K
                  tempstring = TEMPVERSION
                  updateversion
               End If
               If InStr(tempstring, "Recint=") Then
                  tempinterval = ""
                  For K = 1 To lentempstring
                     If K > 7 Then
                        test = Mid(tempstring, K, 1)
                        tempinterval = tempinterval + test
                     End If
                  Next K
                  updateinterval
               End If
               If InStr(tempstring, "Start") Then
               tempstartdate = ""
                  For K = 1 To lentempstring
                     If K > 12 Then
                        test = Mid(tempstring, K, 1)
                        If IsNumeric(test) = True Then
                           tempstartdate = tempstartdate + test
                        End If
                     End If
                  Next K
                  
                  stdategenerate
               End If
               If InStr(tempstring, "Finish") Then
               tempfinishdate = ""
                  For K = 1 To lentempstring
                     If K > 12 Then
                        test = Mid(tempstring, K, 1)
                        If IsNumeric(test) = True Then
                           tempfinishdate = tempfinishdate + test
                        End If
                     End If
                  Next K
                  fdategenerate
                  checkduration
               End If
               If InStr(tempstring, "MaxD") Then
               tempmaxd = ""
                  For K = 1 To lentempstring
                     If K > 11 Then
                        test = Mid(tempstring, K, 1)
                        tempmaxd = tempmaxd + test
                     End If
                  Next K
                  
                  updatemaxdepth
               End If
               If InStr(tempstring, "OTU") Then
                  tempotu = ""
                  For K = 1 To lentempstring
                     If K > 10 Then
                        test = Mid(tempstring, K, 1)
                        tempotu = tempotu + test
                     End If
                  Next K
                  updateotu
               End If
               If InStr(tempstring, "Descend") Then
                  tempdescend = ""
                  For K = 1 To lentempstring
                     If K > 19 Then
                        test = Mid(tempstring, K, 1)
                        tempdescend = tempdescend + test
                     End If
                  Next K
                  updatedescend
               End If
               If InStr(tempstring, "Status") Then
                  tempstatus = ""
                  For K = 1 To lentempstring
                     If K > 13 Then
                        test = Mid(tempstring, K, 1)
                        tempstatus = tempstatus + test
                     End If
                  Next K
                  updatestatus
               End If
                
                'Read gas
               If InStr(tempstring, "Gas") Then
                  i = 0
                  For K = 5 To lentempstring
                     test = Mid(tempstring, K, 1)
                     If test = "," Then
                  '      test = txt2(i)
                    
                        i = i + 1
                     Else
                        test = Mid(tempstring, K, 1)
                        txt2(i) = txt2(i) + test
                        
                     End If
                  Next K
                  updategasprofile
               End If
               If InStr(tempstring, "Tissue") Then
                  i = 0
                  For K = 7 To lentempstring
                     test = Mid(tempstring, K, 1)
                     If test = "," Then
                        i = i + 1
                     Else
                        test = Mid(tempstring, K, 1)
                        txt2(i) = txt2(i) + test
                        
                     End If
                  Next K
                  updatetissueprofile
               End If
               If InStr(tempstring, "Profile") Then
                  Diveprofile = True
                  v = 0
               End If
               If InStr(tempstring, "End") Then
                  Diveprofile = False
               End If
               If Diveprofile = True And v > 0 Then
                  i = 0
                  For K = 1 To lentempstring
                     test = Mid(tempstring, K, 1)
                     If test = "," Then
                        i = i + 1
                     Else
                        test = Mid(tempstring, K, 1)
                        txt2(i) = txt2(i) + test
                        
                     End If
                  Next K
                  
                For K = 0 To i
                
                  Select Case K
                  Case 0
                      SQL = "SELECT * FROM profile"
                      Set RS = DB.OpenRecordset(SQL)
                      RS.AddNew
                      RS!DiveID = tempserialno
                      recordid = txt2(K)
                      RS!ITEM1 = recordid
                      test = recordid
                      RS.Update
                      Buff = ""
                   
                 Case 1
                   If IsNumeric(CInt(txt2(K))) = True And Len(txt2(K)) = 4 Then
                      tempcode = CInt(txt2(K)) * 0.09375
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Depth : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM2 = tempcode
                   RS.Update
                   tempcode = tempcode * 3.28084
                   tempcode = Format(tempcode, "###.00")
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM3 = tempcode
                   RS.Update
                 Case 2
                      tempcode = txt2(K)
                      For i = 1 To Len(tempcode)
                        If Mid$(tempcode, i, 1) = Chr(13) Or Mid$(tempcode, i, 1) = Chr(32) Or Mid$(tempcode, i, 1) = Chr(9) Or Mid$(tempcode, i, 1) = Chr(10) Then
                        Else
                        Buff = Buff & Mid$(tempcode, i, 1)
                        End If
                      Next
                      tempcode = Buff
                      Buff = ""
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM4 = tempcode
                   RS.Update
                 Case 3
                   If IsNumeric(CInt(txt2(K))) = True And Len(txt2(K)) = 4 Then
                      tempcode = CInt(txt2(K)) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PO2 : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM5 = tempcode
                   RS.Update
                   
                 Case 4
                    
                   If Left(txt2(K), 1) = "M" And Len(txt2(K)) = 5 Then
                      tempcode = Right(txt2(K), 1)
                   Else
                      errmsg = errmsg & Chr(13) & "Mark : " & recordid
                   End If
                   If tempcode <> "1" Then
                      tempcode = "0"
                    End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM6 = tempcode
                   RS.Update
                 Case 5
                   tempcode = txt2(K)
                   tempcodeleft = Left(tempcode, 1)
                   tempcoderight = Right(tempcode, 4)
                   If IsNumeric(tempcoderight) = True And Len(txt2(K)) = 5 And tempcodeleft = "T" Then
                      tempcoderight = CInt(tempcoderight)
                      tempcoderightf = ((tempcoderight * 9) / 5) + 32
                      tempcoderightf = Format(tempcoderightf, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Temperature : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM7 = tempcoderight
                   RS!ITEM8 = tempcoderightf
                   RS.Update
                   If ITEM1 = test Then
                      SQL = "SELECT * FROM main "
                      SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                      Set RS = DB.OpenRecordset(SQL)
                      RS.Edit
                      RS!temperature = tempcoderight
                      RS.Update
                   End If
                 Case 6
                   If Left(txt2(K), 1) = "A" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 10
                      tempcode2 = tempcode * 3.28084
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "External Depth : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM9 = tempcode
                   RS!ITEM10 = tempcode2
                   RS.Update
                   tempsysid = "Rebrether"
                   updatesystemid
                 Case 7   ' vbattery
                   If Left(txt2(K), 1) = "B" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 156
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Val Battery : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM11 = tempcode
                   RS.Update
                 Case 8   ' ebattery
                   If Left(txt2(K), 1) = "C" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 156
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "Electronic Battery : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM12 = tempcode
                   RS.Update
                 Case 9  'hp Diluent
                    If Left(txt2(K), 1) = "D" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode)
                      tempcode2 = tempcode * 14.7
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = "Error in hp Diluent"
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM13 = tempcode
                   RS!ITEM14 = tempcode2
                   RS.Update
                 Case 10 'HP O2
                   If Left(txt2(K), 1) = "E" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode)
                      tempcode2 = tempcode * 14.7
                      tempcode = Format(tempcode, "###.00")
                      tempcode2 = Format(tempcode2, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "HP02 : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM15 = tempcode
                   RS!ITEM16 = tempcode2
                   RS.Update
                 Case 11 'PPO2 A
                   If Left(txt2(K), 1) = "F" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 A : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM17 = tempcode
                   RS.Update
                 Case 12 'PPO2 b
                   If Left(txt2(K), 1) = "G" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 B : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM18 = tempcode
                   RS.Update
                 Case 13 'PPO2 C
                   If Left(txt2(K), 1) = "H" And Len(txt2(K)) = 6 Then
                      tempcode = Right(txt2(K), 5)
                      tempcode = CInt(tempcode) / 100
                      tempcode = Format(tempcode, "###.00")
                   Else
                      errmsg = errmsg & Chr(13) & "PPO2 C : " & recordid
                   End If
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM19 = tempcode
                   RS.Update
                 Case 14 'To 25 'nick change this and case to 30
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item20 = txt2(K)
                   RS.Update
                 Case 15
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item21 = txt2(K)
                   RS.Update
                 Case 16
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item22 = txt2(K)
                   RS.Update
                 Case 17
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item23 = txt2(K)
                   RS.Update
                 Case 18
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item24 = txt2(K)
                   RS.Update
                 Case 19
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item25 = txt2(K)
                   RS.Update
                 Case 20
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item26 = txt2(K)
                   RS.Update
                 Case 21
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item27 = txt2(K)
                   RS.Update
                 Case 22
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item28 = txt2(K)
                   RS.Update
                  Case 23
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item29 = txt2(K)
                   RS.Update
                 Case 24
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item30 = txt2(K)
                   RS.Update
                 Case 25
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item31 = txt2(K)
                   RS.Update
                 Case 26
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item32 = txt2(K)
                   RS.Update
                 Case 27
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item33 = txt2(K)
                   RS.Update
                 Case 28 To 44
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID ='" & Trim(tempserialno) & "'  AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!item34 = txt2(K)
                   RS.Update
               End Select
            Next K
            i = 0
            cleartext2
            End If
               
            End If
            PDBMoveNext (dbdivelist)
     Wend
        
ErrorHandler2:
   Screen.MousePointer = 0
End Sub
Private Sub mnutooltipsClear()
 mnutooltipson.Checked = False
 mnutooltipsoff.Checked = False
 SQL = "select * FROM dpserialno "
 Set RS = DB.OpenRecordset(SQL)
 RS.Edit
 RS!dpmsgbox = "Off"
 RS.Update
 tempmsgbox = "Off"
End Sub
Private Sub mnumsgboxClear()
 mnumsgboxon.Checked = False
 mnumsgboxoff.Checked = False

End Sub


Private Sub mnuDsequential_Click()
Splanmain.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnumsgboxoff_Click()
 mnumsgboxClear
 mnumsgboxoff.Checked = True
 SQL = "select * FROM dpserialno "
 Set RS = DB.OpenRecordset(SQL)
 RS.Edit
 RS!dpmsgbox = "Off"
 RS.Update
 tempmsgbox = "Off"
End Sub

Private Sub mnumsgboxon_Click()
 mnumsgboxClear
 mnumsgboxon.Checked = True
 SQL = "select * FROM dpserialno "
 Set RS = DB.OpenRecordset(SQL)
 RS.Edit
 RS!dpmsgbox = "On"
 RS.Update
 tempmsgbox = "On"
End Sub

Private Sub mnusingle_Click()
planmain.Show
End Sub

Private Sub mnutooltipsoff_Click()
   mnutooltipsClear
   mnutooltipsoff.Checked = True
   SQL = "select * FROM dpserialno "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   RS!dptooltips = "Off"
   RS.Update
   temptooltips = "Off"
End Sub

Private Sub mnutooltipson_Click()
   mnutooltipsClear
   mnutooltipson.Checked = True
   SQL = "select * FROM dpserialno "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   RS!dptooltips = "On"
   RS.Update
   temptooltips = "On"
End Sub

