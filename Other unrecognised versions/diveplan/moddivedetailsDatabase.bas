Attribute VB_Name = "moddivedetailsDatabase"
'---------------------------------------------------------------
'     AppForge PDB Converter auto-generated code module
'
'     Source Database: C:\ppc\POCKETPC.MDB
'     Source Table   : divelist
'
'     Num Records    : 0
'
'     PDB Table Name : divedetails
'          CreatorID : AFVM
'          TypeID    : DATA
'
'     Converted Time : 4/14/2004 2:03:30 PM
'
'---------------------------------------------------------------

Option Explicit

' Use these constants for the CreatorID and TypeID
Public Const divedetails_CreatorID As Long = &H4146564D
Public Const divedetails_TypeID As Long = &H44415441

' Use this global to store the database handle
Public dbdivedetails As Long

' Use this enumeration to get access to the converted database Fields
Public Enum tdivedetailsDatabaseFields
        diveid_Field = 0
        divestring_Field = 1
End Enum

Public Type tdivedetailsRecord
        diveid As String
        divestring As String
End Type


Public Function OpendivedetailsDatabase() As Boolean

        ' Open the database
        #If APPFORGE Then
        dbdivedetails = PDBOpen(Byfilename, "divedetails", 0, 0, 0, 0, afModeReadWrite)
        #Else
        dbdivedetails = PDBOpen(Byfilename, App.Path & "\divedetails", 0, 0, 0, 0, afModeReadWrite)
        #End If

        If dbdivedetails <> 0 Then
                'We successfully opened the database
                OpendivedetailsDatabase = True

        Else
                'We failed to open the database
                OpendivedetailsDatabase = False
                #If APPFORGE Then
                MsgBox "Could not open database - divedetails", vbExclamation
                #Else
                MsgBox "Could not open database - " + App.Path + "\divedetails.pdb" + vbCrLf + vbCrLf + "Potential causes are:" + vbCrLf + "1. Database file does not exist" + vbCrLf + "2. The database path in the PDBOpen call is incorrect", vbExclamation
                #End If
        End If

End Function
Public Function CreatedivedetailsRecord(NewRecord3 As tdivedetailsRecord) As Boolean
   PDBCreateRecordBySchema dbdivedetails
   CreatedivedetailsRecord = PDBWriteRecord(dbdivedetails, VarPtr(NewRecord3))  'Write the data to the new record.
   PDBUpdateRecord dbdivedetails             'Update the new record.
End Function

Public Sub ClosedivedetailsDatabase()

        ' Close the database
        PDBClose dbdivedetails
        dbdivedetails = 0

End Sub


Public Function ReaddivedetailsRecord(MyRecord As tdivedetailsRecord) As Boolean

        ReaddivedetailsRecord = PDBReadRecord(dbdivedetails, VarPtr(MyRecord))

End Function


Public Function WritedivedetailsRecord(MyRecord As tdivedetailsRecord) As Boolean

        WritedivedetailsRecord = PDBWriteRecord(dbdivedetails, VarPtr(MyRecord))

End Function
