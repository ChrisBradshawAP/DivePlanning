Attribute VB_Name = "modDivemainDatabase"
'---------------------------------------------------------------
'     AppForge PDB Converter auto-generated code module
'
'     Source Database: C:\ppc\POCKETPC.MDB
'     Source Table   : Divemain
'
'     Num Records    : 0
'
'     PDB Table Name : Divemain
'          CreatorID : AFVM
'          TypeID    : DATA
'
'     Converted Time : 3/25/2004 12:45:20 AM
'
'---------------------------------------------------------------

Option Explicit

' Use these constants for the CreatorID and TypeID
Public Const Divemain_CreatorID As Long = &H4146564D
Public Const Divemain_TypeID As Long = &H44415441
Public tempserialno As String
Public tempdiveid As String
' Use this global to store the database handle
Public dbDivemain As Long

' Use this enumeration to get access to the converted database Fields
Public Enum tDivemainDatabaseFields
        Divemid_Field = 0
        Divemdate_Field = 1
        Divemlocation_Field = 2
        Divemmaxdepth_Field = 3
        Divemduration_Field = 4
        Divedatestring_Field = 5
End Enum

Public Type tDivemainRecord
        Divemid As String
        Divemdate As String
        Divemlocation As String
        Divemmaxdepth As String
        Divemduration As String
        Divedatestring As String
End Type


Public Function OpenDivemainDatabase() As Boolean

        ' Open the database
        #If APPFORGE Then
        dbDivemain = PDBOpen(Byfilename, "Divemain", 0, 0, 0, 0, afModeReadWrite)
        #Else
        dbDivemain = PDBOpen(Byfilename, App.Path & "\Divemain", 0, 0, 0, 0, afModeReadWrite)
        #End If
        If dbDivemain <> 0 Then
                'We successfully opened the database
                OpenDivemainDatabase = True

        Else
                'We failed to open the database
                OpenDivemainDatabase = False
                #If APPFORGE Then
                MsgBox "Could not open database - Divemain", vbExclamation
                #Else
                MsgBox "Could not open database - " + App.Path + "\Divemain.pdb" + vbCrLf + vbCrLf + "Potential causes are:" + vbCrLf + "1. Database file does not exist" + vbCrLf + "2. The database path in the PDBOpen call is incorrect", vbExclamation
                #End If
        End If

End Function


Public Sub CloseDivemainDatabase()

        ' Close the database
        PDBClose dbDivemain
        dbDivemain = 0

End Sub


Public Function ReadDivemainRecord(MyRecord As tDivemainRecord) As Boolean
  ReadDivemainRecord = PDBReadRecord(dbDivemain, VarPtr(MyRecord))
End Function


Public Function WriteDivemainRecord(MyRecord As tDivemainRecord) As Boolean

        WriteDivemainRecord = PDBWriteRecord(dbDivemain, VarPtr(MyRecord))

End Function
Public Function CreateDivemainRecord(NewRecord As tDivemainRecord) As Boolean
   PDBCreateRecordBySchema dbDivemain
   CreateDivemainRecord = PDBWriteRecord(dbDivemain, VarPtr(NewRecord))  'Write the data to the new record.
   PDBUpdateRecord dbDivemain             'Update the new record.
End Function
