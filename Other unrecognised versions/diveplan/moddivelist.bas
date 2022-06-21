Attribute VB_Name = "moddivelist"
'---------------------------------------------------------------
'     AppForge PDB Converter auto-generated code module
'
'     Source Database: C:\ppc\POCKETPC.MDB
'     Source Table   : divelist
'
'     Num Records    : 0
'
'     PDB Table Name : divelist
'          CreatorID : AFVM
'          TypeID    : DATA
'
'     Converted Time : 2/27/2004 10:53:57 AM
'
'---------------------------------------------------------------

Option Explicit
Global strSelection As String
Global lastserialno As String
Global testname As String
Global tempfileselected As String
' Use these constants for the CreatorID and TypeID
Public Const divelist_CreatorID As Long = &H4146564D
Public Const divelist_TypeID As Long = &H44415441

' Use this global to store the database handle
Public dbdivelist As Long
Public loadotherdb As String
Public pathname As String
Public tempfilename As String
' Use this enumeration to get access to the converted database Fields
Public Enum tdivelistDatabaseFields
        diveid_Field = 0
        divestring_Field = 1
End Enum

Public Type tdivelistRecord
        diveid As String
        divestring As String
End Type


Public Function OpendivelistDatabase() As Boolean
 ' PDBClose dbdivelist
  dbdivelist = 0
  
  dbdivelist = PDBOpen(Byfilename, tempfileselected, 0, 0, 0, 0, afModeReadWrite)
 ' dbdivelist = PDBOpen(Byfilename, pathname & tempfilename, 0, 0, 0, 0, afModeReadWrite)
  
     If dbdivelist <> 0 Then
                'We successfully opened the database
                OpendivelistDatabase = True

        Else
                'We failed to open the database
                OpendivelistDatabase = False
                #If APPFORGE Then
                MsgBox "Could not open database - " & tempfileselected & ".pdb", vbExclamation
                #Else
                MsgBox "Could not open database - " & tempfileselected & ".pdb" + vbCrLf + vbCrLf + "Potential causes are:" + vbCrLf + "1. Database file does not exist" + vbCrLf + "2. The database path in the PDBOpen call is incorrect", vbExclamation
                #End If
        End If
PDBMoveFirst dbdivelist
   PDBGetField dbdivelist, 0, txtdiveid
   'tempserialno = txtdiveid
  
End Function


Public Sub ClosedivelistDatabase()
        ' Close the database
        PDBClose dbdivelist
        dbdivelist = 0

End Sub


Public Function ReaddivelistRecord(MyRecord As tdivelistRecord) As Boolean
        ReaddivelistRecord = PDBReadRecord(dbdivelist, VarPtr(MyRecord))
End Function

Public Function WritedivelistRecord(MyRecord As tdivelistRecord) As Boolean
        WritedivelistRecord = PDBWriteRecord(dbdivelist, VarPtr(MyRecord))
End Function
Public Function CreatedivelistRecord(Newrecord2 As tdivelistRecord) As Boolean
   PDBCreateRecordBySchema dbdivelist
   CreatedivelistRecord = PDBWriteRecord(dbdivelist, VarPtr(Newrecord2))  'Write the data to the new record.
   PDBUpdateRecord dbdivelist             'Update the new record.
End Function
Public Sub DeletedivelistRecord()
  
   PDBDeleteRecordEx dbdivelist, afDeleteModeRemove
   'I used PDBDeleteRecordEx with the delete mode afDeleteModeRemove.
   'This method completely removes all traces of the record.
   
   'If you plan to use the AppForge Universal Conduit, use PDBDeleteRecord instead.
   'This will flag the data as being deleted, but the record will still be counted
   'in PDBNumRecords
      
End Sub
Public Sub CreateDivelistDatabase()
    
End Sub
