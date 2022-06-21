VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form2"
   ScaleHeight     =   3300
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2040
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
   Begin VB.CommandButton Command2 
      Caption         =   "Read from Rebreather"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read from text file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim Y(20) As Integer
Dim X As Integer
Dim fp As UserDocument
'Dim fpp As CFileBinaryReadable
'Dim FileMgr As New CFileManager
'Dim TxtFile As CFileTextReadable

Dim CB As Byte
Dim C As String
Dim c1 As Byte
Dim I As Integer
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
Dim txt2(20) As String
'Dim #1 As Integer
Dim hOutFile As Integer
'
Dim F1 As String
Dim T(4) As String
Dim T2(4) As String
Dim S As String
Dim TS As String
Dim G As String
Dim H As String
Dim W(4) As Boolean
Dim zxt As Double
Dim CT As Integer
Dim Auto_run As Integer
Private Function Newrecord()
TEMPVERSION = Trim(TEMPVERSION)
SQL = "SELECT * FROM SERIALNO"
Set RS = DB.OpenRecordset(SQL)
tempserialno = RS("serial_no")
tempserialno = Val(tempserialno) + 1
SQL = "SELECT * FROM serialno "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!serial_no = tempserialno
RS.Update
SQL = "SELECT * FROM MAIN"
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!DiveID = TEMPSERIANLNO
RS!RECORDID = "R0003"
RS!Version = TEMPVERSION
MsgBox TEMPVERSION
RS.Update

Text1.Text = ""
End Function
Private Function updateinterval()
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = 'D0003' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!Interval = tempinterval
RS.Update
Text1.Text = ""
End Function
Private Function stdategenerate()
tempstartdate = Trim(tempstartdate)
tempstartdate2 = tempstartdate
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = 'D0003' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
dateconvert
RS!startdate = tempstartdate
RS.Update
Text1.Text = ""
End Function
Private Function updatemaxdepth()
tempmaxd = Trim(tempmaxd)
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = 'D0003' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!maxdepth = tempmaxd
RS.Update
Text1.Text = ""
End Function
Private Function updateotu()
tempotu = Trim(tempotu)
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = 'D0003' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!otu = tempotu
RS.Update
Text1.Text = ""
End Function
Private Function updategasprofile()
txt2(I) = txt2(I) + Chr$(CB)
SQL = "SELECT * FROM GAS"
tempgsstype = txt2(0)
tempGasN2 = txt2(1)
tempHe = txt2(2)
tempMod = txt2(3)
tempstat = txt2(4)
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!DiveID = "D0003"
RS!Gastype = tempgsstype
RS!GasN2 = tempGasN2
RS!GasHe = tempHe
RS!Gasmod = tempMod
RS!Gasstat = tempstat
RS.Update
cleartext2
Text1.Text = ""
End Function
Private Function updatetissueprofile()
txt2(I) = txt2(I) + Chr$(CB)
SQL = "SELECT * FROM tissue"
temptsstype = txt2(0)
tempGasN2 = txt2(1)
tempHe = txt2(2)
Set RS = DB.OpenRecordset(SQL)
RS.AddNew
RS!DiveID = "D0003"
RS!Tissuetype = temptsstype
RS!TissueN2 = tempGasN2
RS!TissueHe = tempHe
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
Private Sub Command1_Click()
Dim fileflags As FileOpenConstants
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
Screen.MousePointer = 11
F = test
Open F For Binary As #1

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
            For I = 1 To 13
               Get #1, , CB
               TEMPVERSION = TEMPVERSION + Chr$(CB)
               If CB = 13 Then
                 I = 12
                 Newrecord
               End If
            Next I
        
      End If
      
      'store Interval
      If InStr(Text1.Text, "Recint=") Then
        Text1.Text = ""
          For I = 1 To 5
             Get #1, , CB
                 If Val(CB) = 46 Or (Val(CB) < 57 And Val(CB) > 47) Then
                    tempinterval = tempinterval + Chr$(CB)
                 End If
                 If CB = 13 Then
                    I = 4
                    tempinterval = Trim(tempinterval)
                    updateinterval
                 End If
               Next I
       End If
      
      'Store start date
      If InStr(Text1.Text, "Start") Then
         Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For I = 1 To 13
                 Get #1, , CB
                 If Val(CB) < 57 And Val(CB) > 47 Then
                    tempstartdate = tempstartdate + Chr$(CB)
                 End If
                 If CB = 13 Then
                    I = 12
                    stdategenerate
                  End If
               Next I
            End If
         Next K
      End If
      
      'Read Finished Time Info
      If InStr(Text1.Text, "Finish") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For I = 1 To 13
                 Get #1, , CB
                 If Val(CB) < 57 And Val(CB) > 47 Then
                    tempfinishdate = tempfinishdate + Chr$(CB)
                 End If
                 If CB = 13 Then
                    I = 13
                    fdategenerate
                    checkduration
                 End If
               Next I
            End If
         Next K
      End If
                 
                 
      If InStr(Text1.Text, "MaxD") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            
            If CB = 44 Then
               For I = 1 To 13
                 Get #1, , CB
                 If Val(CB) = 46 Or (Val(CB) < 57 And Val(CB) > 47) Then
                    tempmaxd = tempmaxd + Chr$(CB)
                 End If
                 If CB = 13 Then
                    I = 13
                    updatemaxdepth
                  End If
               Next I
            End If
         Next K
      End If
      
      If InStr(Text1.Text, "OTU") Then
        Text1.Text = ""
         For K = 1 To 6
            Get #1, , CB
            If CB = 44 Then
               For I = 1 To 13
                 Get #1, , CB
                 If Val(CB) = 46 Or (Val(CB) < 57 And Val(CB) > 47) Then
                    tempotu = tempotu + Chr$(CB)
                 End If
                 If CB = 13 Then
                    I = 12
                    updateotu
                 End If
               Next I
            End If
         Next K
      End If
      
      'Read gas
      If InStr(Text1.Text, "Gas") Then
        Text1 = ""
        I = 0
        gasprofile = False
        While gasprofile = False
          Get #1, , CB
           If CB = 13 Then
              gasprofile = True
              updategasprofile
            End If
            If CB = 44 Then
               test = txt2(I)
               I = I + 1
            Else
               txt2(I) = txt2(I) + Chr$(CB)
            End If
         Wend
       End If
       
       
       'Tissue status
       If InStr(Text1.Text, "Tissue") Then
        Text1 = ""
        'Get #1, , CB
        I = 0
        tissueprofile = False
        While tissueprofile = False
          Get #1, , CB
            If CB = 13 Then
              tissueprofile = True
              updatetissueprofile
            End If
            If CB = 44 Then
               I = I + 1
            Else
               txt2(I) = txt2(I) + Chr$(CB)
            End If
         Wend
       End If
       
             
        'Profile
       If InStr(Text1.Text, "Profile") Then
        K = 0
        Text1 = ""
        Get #1, , CB
        I = 0
        diveprofile = False
        While diveprofile = False
          Get #1, , CB
            If CB = 13 Then
            '  diveprofile = True
              txt2(I) = txt2(I) + Chr$(CB)
              For K = 0 To I
                 
                 Select Case K
                 
                 Case 0
                   SQL = "SELECT * FROM profile"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.AddNew
                   RS!DiveID = "D0003"
                   RS!ITEM1 = txt2(K)
                   test = txt2(K)
                   RS.Update
                 Case 1
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM2 = txt2(K)
                   RS.Update
                 Case 2
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM3 = txt2(K)
                   RS.Update
                 Case 3
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM4 = txt2(K)
                   RS.Update
                 Case 4
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM5 = txt2(K)
                   RS.Update
                 Case 5
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM6 = txt2(K)
                   RS.Update
                 Case 6
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM7 = txt2(K)
                   RS.Update
                 Case 7
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM8 = txt2(K)
                   RS.Update
                 Case 8
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM9 = txt2(K)
                   RS.Update
                 Case 9
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM10 = txt2(K)
                   RS.Update
                 Case 10
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM11 = txt2(K)
                   RS.Update
                 Case 11
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM12 = txt2(K)
                   RS.Update
                 Case 12
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM13 = txt2(K)
                   RS.Update
                 Case 13
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM14 = txt2(K)
                   RS.Update
                 Case 14
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM15 = txt2(K)
                   RS.Update
                 Case 15
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM16 = txt2(K)
                   RS.Update
                 Case 16
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM17 = txt2(K)
                   RS.Update
                 Case 17
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM18 = txt2(K)
                   RS.Update
                 Case 18
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM19 = txt2(K)
                   RS.Update
                 Case 19
                   SQL = "SELECT * FROM profile "
                   SQL = SQL & "WHERE DiveID = 'D0003' AND ITEM1 = '" & test & "'"
                   Set RS = DB.OpenRecordset(SQL)
                   RS.Edit
                   RS!ITEM20 = txt2(K)
                   RS.Update
               End Select
             Next K
                I = 0
                 cleartext2
                 'txt2.Text = ""
            End If
          
            If CB = 44 Then
               I = I + 1
            Else
               txt2(I) = txt2(I) + Chr$(CB)
               test = txt2(I)
               'MsgBox TEST
            End If
             Text1.Text = Text1.Text + Chr$(CB)
            If InStr(Text1.Text, "End") Then
              Screen.MousePointer = 0
              
              MsgBox "ok"
              diveprofile = True
            End If
         Wend
       End If
       
            'For I = 1 To 13
            '     Get #1, , CB
            '     If Val(CB) < 57 And Val(CB) > 47 Then
            '        tempfinishdate = tempfinishdate + Chr$(CB)
            '     End If
            '     If CB = 13 Then
            '        I = 13
            '        tempfinishdate = Trim(tempfinishdate)
            '        SQL = "SELECT * FROM MAIN "
             '       SQL = SQL & " where DiveID = 'D0003' "
             '       Set RS = DB.OpenRecordset(SQL)
             '       RS.Edit
             ''       finisheddateconvert
             '       RS!Finisheddate = tempfinishdate
             '       RS.Update
             '       Text1.Text = ""
             '
             '    End If
             '  Next I
            'E 'nd If
         'Next K
     ' End If
      
      
      
      
        ' iNSERT RECORD HERE, A AS ROW, TXT2(i) AS READING
       '  I = 0
       '  txt2(ind) = ""
       '  Text1.Text = ""
       '  cleartext3
       '  A = A + 1
       '
       '  Get #1, , CB
     ' End If
        '  cleartext3
'          Text1.Text = Text1.Text + Chr$(CB)
'          Get #1, , CB
'
'
'         'LP = 9999998
'         ' MsgBox "cOMPLETED"
 '     End If
 ''  'If maxdprofile = 1 Then
 '    ' Text1.Text = Text1.Text + Chr$(CB)
 '     'If CB = 13 Then
 '     '  ' MsgBox Text1
 '      '  Text1.Text = ""
       '  cleartext3
'       '    End If
'   'End If
 ''  If profilefound = 1 Then
 ''     'j = j + 1
 '        'Text1.Text = TxtFile.Peek
 '     Get #1, , CB
 '     Text1.Text = Text1.Text + Chr$(CB)
 '     If InStr(Text1.Text, "End") Then
 '         LP = 9999998
 '         MsgBox "cOMPLETED"
 '     End If
 '     If CB = 13 Then
 '        For I = 0 To 12
 '          test = txt2(I)
 '          MsgBox test
 '        Next I
  '
  ''      ' iNSERT RECORD HERE, A AS ROW, TXT2(i) AS READING
  '       I = 0
  '       txt2(ind) = ""
  '       Text1.Text = ""
  '       cleartext3
   '      A = A + 1
   '
   ''      Get #1, , CB
   '   End If
    '  If CB = 44 Then
      ' Get #1, , CB
     'txt2(I) = txt2(I) + Chr$(CB)
   '    I = I + 1
   '   Else
  '       txt2(I) = txt2(I) + Chr$(CB)
         ' iNSERT RECORD HERE, A AS ROW, TXT2(i) AS READING
         
  '    End If
  '  Else
  '   Get #1, , CB
  '   Text1.Text = Text1.Text + Chr$(CB)
  '   If InStr(Text1.Text, "MaxD") Then
  '      maxdprofile = 1
  '     cleartext2
  '     Get #1, , CB
  '     Get #1, , CB
  '
   '  End If
''
'     If InStr(Text1.Text, "Profile") Then
 '      Text1.Text = ""
 ''      profilefound = 1
 '       I = 0
 '      cleartext2
  '     Get #1, , CB
  '     Get #1, , CB
  '
  '   End If
  '  If LP > 2000 Then LP = 9999998
  ' End If
 ''  If LP > 206 Then
 '    MsgBox LP
 '    MsgBox profilefound
 '
 '    LP = 9999998
 '  End If
Next LP

rbinterface.Show


End Sub
Private Function dateconvert()
tempstartday = Int(Val(tempstartdate) / 86400)
remainder = tempstartdate - (tempstartday * 86400)
temphour = Int(remainder / 3600)
minutesremainder = remainder - (temphour * 3600)
tempminutes = Int(minutesremainder / 60)
secondremainder = minutesremainder - tempminutes * 60
tempstartdate = Val(tempstartday)
oldstartdate = Format("01/01/1990", "mm/dd/yyyy")
testdate = DateAdd("d", tempstartdate, oldstartdate)
testdate = testdate & " " & temphour & ":" & tempminutes & ":" & secondremainder
tempstartdate = Format(testdate, "mm/dd/yyyy hh:mm:ss")

End Function
Private Function finisheddateconvert()
tempfinishday = Int(Val(tempfinishdate) / 86400)
fremainder = tempfinishdate - (tempfinishday * 86400)
tempfhour = Int(fremainder / 3600)
minutesfremainder = fremainder - (tempfhour * 3600)
tempfminutes = Int(minutesfremainder / 60)
secondfremainder = minutesfremainder - tempfminutes * 60
tempfinishdate = Val(tempfinishday)
oldfinishdate = Format("01/01/1990", "mm/dd/yyyy")
testdate = DateAdd("d", tempfinishdate, oldfinishdate)
testdate = testdate & " " & tempfhour & ":" & tempfminutes & ":" & secondfremainder
tempfinishdate = Format(testdate, "mm/dd/yyyy hh:mm:ss")

End Function
Private Function checkduration()
tempduration = tempfinishdate2 - tempstartdate2
tempdminutes = Int(tempduration / 60)
tempduremainder = tempduration - tempdminutes * 60
tempduration = tempdminutes & ":" & tempduremainder
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = 'D0003' "
Set RS = DB.OpenRecordset(SQL)
RS.Edit
RS!Duration = tempduration
RS.Update
End Function
Private Function fdategenerate()
tempfinishdate = Trim(tempfinishdate)
tempfinishdate2 = tempfinishdate
SQL = "SELECT * FROM MAIN "
SQL = SQL & " where DiveID = 'D0003' "
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

Private Sub Form_Load()
Set DB = OpenDatabase(App.Path & "/rb.mdb")
  Dir1.Path = "C:\test" 'Drive1.Drive   ' Set directory path.
 ' File2.Path = Dir1.Path   ' Set file path.End Sub
  'F = "c:\test\rbtest.txt"
  Me.WindowState = 2
End Sub
