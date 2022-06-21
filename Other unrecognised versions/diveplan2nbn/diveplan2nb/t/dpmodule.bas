Attribute VB_Name = "dpmodule"
Const STARTPATH = "\"
Global DB As Database
Global RS As Recordset
Global RS2 As Recordset
Global RS3 As Recordset
Global RS4 As Recordset
Global RS5 As Recordset
Global RS6 As Recordset
Global SQL As String
Global yPos As Integer
Global xPos As Integer
Global systemstarted As Boolean
Global tempseqduplicate As Boolean
Global previousform As String
Global temptooltips As String
Global tempmsgbox As String
Global fgactivate As String
Global colselected As String
Global cols0activated As String
Global p As Integer
Global T As Integer
Global comfirmDisplay  As String
Global tempchoice As String
Global rowindentified As String
Global rowidentified2 As String
Global tempstartdate As Variant
Global tempfinishdate As Variant
Global profilefound As Integer
Global maxdprofile As Integer
Global gasvalidation As String
Global tempmaxd As String
Global tempdiveserialno As String
Global tempseqdiveno As String
Global oldtempseqdiveno As String
Global tempotu As String
Global sortorder As String
Global tempfinishdate2 As String
Global tempstartdate2 As String
Global tempserialno, oldserialno, newserialno, newseqdiveno As String
Global tempniused, tempheused, tempmaxdused, tempppo2used As String
Global Totalcount As Variant
Global itemheader As String
Global displaydefaulted As String
Global itemselected As String
Global displaydefault As String
Global filesource As String
Global tempfilename As String
Global txtdiveid As String
Global tempstring As String
Global tempfilename2 As String
Global tempsysid As String
Global tempstatus As String
Global tempdescend As String
Global templocationfile As String
Global tempsysviewid As String
Global newseqserialno, DataChanged As String
Global checkgasselected, checkgasusedselected, formstarted, datachangedstatus, profilerecordexist As Boolean
Global feetormeter_factor As Double
Global psiorbar_factor As Double
Global feetormeter_string As String
Global feetormeter_feeton As Integer
Global feetormeter_shortstring As String
Global feetormeter_decostep As Double
Global do_not_load As Integer
Global cns_current As Double
Global otu_current As Double
Global deco_grid_display As Integer
Global deco_grid_display_last As Integer
Global deco_grid_display_rowlast As Integer
Global deco_grid_display_celllast As Long
Global buhl_mode As Integer
Global inc_depth As Double
Global inc_time As Double



Public Sub cleanserialno()
' test
End Sub

Public Sub cleandatabase()
cleanserialno
cleandiveplan
cleandivelist
End Sub

Public Sub cleandiveplan()
SQL = "select * FROM seqdpmain "
SQL = SQL & "order by DIVEPLANID "
Set RS3 = DB.OpenRecordset(SQL)
'tempdpid = RS3("diveplanid")
If RS3.BOF = True And RS3.EOF = True Then
   'do nothing
Else
   While RS3.EOF = False
      tempdpid = RS3("diveplanid")
      If tempdpid Like "T*" Then
         RS3.Delete
      End If
      RS3.MoveNext
   Wend
   RS3.Close
End If
SQL = "select * FROM dpmaingaslist "
SQL = SQL & "order by dpmainid "
Set RS3 = DB.OpenRecordset(SQL)
If RS3.BOF = True And RS3.EOF = True Then
   'do nothing
Else
   While RS3.EOF = False
      tempdpid = RS3("dpmainid")
      If tempdpid Like "T*" Then
         RS3.Delete
      End If
       RS3.MoveNext
   Wend
   RS3.Close
End If
SQL = "select * FROM seqdpprofile "
SQL = SQL & "order by dpprofileid "
Set RS3 = DB.OpenRecordset(SQL)
If RS3.BOF = True And RS3.EOF = True Then
   'do nothing
Else
   While RS3.EOF = False
      tempdpid = RS3("dpprofileid")
      If tempdpid Like "T*" Then
         RS3.Delete
      End If
        RS3.MoveNext
   Wend
   RS3.Close
End If
End Sub
Public Sub cleandivelist()
 SQL = "select * FROM seqdplist "
 Set RS3 = DB.OpenRecordset(SQL)
While RS3.EOF = False
   tempdpid = RS3("seqdiveidmain")
   If tempdpid Like "T*" Then
      tempdpid2 = tempdpid
      RS3.Delete
   End If
   RS3.MoveNext
Wend
 SQL = "select * FROM seqdplist "
 SQL = SQL & " order by seqdiveidmain "
 Set RS3 = DB.OpenRecordset(SQL)
 If RS3.EOF = True And RS3.BOF = True Then
    tempdpid = "SM00000000"
 Else
   While RS3.EOF = False
      RS3.MoveLast
      tempdpid = RS3("seqdiveidmain")
      RS3.MoveNext
   Wend
End If
   
   SQL = "select * FROM dpserialno "
   Set RS = DB.OpenRecordset(SQL)
   If RS.EOF = True And RS.BOF = True Then
      RS.AddNew
      RS!seqdiveserialno = tempdpid
      RS.Update
      RS.Close
   Else
      RS.Edit
      RS!seqdiveserialno = tempdpid
      RS.Update
      RS.Close
   End If
End Sub
Public Function validateserialno()
SQL = "SELECT * FROM seqdpmain "
SQL = SQL & "order by DiveplanID "
Set RS = DB.OpenRecordset(SQL)
If RS.EOF = False Then
   RS.MoveLast
   tempdiveid = RS("diveplanid")
   SQL = "SELECT * FROM dpserialno "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   RS!lastseqdserialno = tempdiveid
   RS.Update
Else
   tempdiveid = "SP00000000"
   SQL = "SELECT * FROM dpserialno "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   RS!lastseqdserialno = tempdiveid
   RS.Update
End If
SQL = "SELECT * FROM seqdplist "
SQL = SQL & "order by seqdiveidmain "
Set RS = DB.OpenRecordset(SQL)
If RS.EOF = False Then
   RS.MoveLast
   tempdiveid = RS("seqdiveidmain")
   SQL = "SELECT * FROM dpserialno "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   RS!seqdiveserialno = tempdiveid
   RS.Update
Else
   tempdiveid = "SM00000000"
   SQL = "SELECT * FROM dpserialno "
   Set RS = DB.OpenRecordset(SQL)
   RS.Edit
   RS!seqdiveserialno = tempdiveid
   RS.Update
End If
End Function
