Attribute VB_Name = "modpalmtbDatabase"
'---------------------------------------------------------------
'     AppForge PDB Converter auto-generated code module
'
'     Source Database: C:\prolink\PALM.MDB
'     Source Table   : palmtb
'
'     Num Records    : 11
'
'     PDB Table Name : palmtb
'          CreatorID : AFVM
'          TypeID    : DATA
'
'     Converted Time : 2/18/2004 11:33:56 AM
'
'---------------------------------------------------------------

Option Explicit

' Use these constants for the CreatorID and TypeID
Public Const palmtb_CreatorID As Long = &H4146564D
Public Const palmtb_TypeID As Long = &H44415441

' Use this global to store the database handle
Public dbpalmtb As Long

' Use this enumeration to get access to the converted database Fields
Public Enum tpalmtbDatabaseFields
	diveid_Field = 0
	divestring_Field = 1
End Enum

Public Type tpalmtbRecord
	diveid As String
	divestring As String
End Type


Public Function OpenpalmtbDatabase() As Boolean

	' Open the database
	#If APPFORGE Then
	dbpalmtb = PDBOpen(Byfilename, "palmtb", 0, 0, 0, 0, afModeReadWrite)
	#Else
	dbpalmtb = PDBOpen(Byfilename, App.Path & "\palmtb", 0, 0, 0, 0, afModeReadWrite)
	#End If

	If dbpalmtb <> 0 Then
		'We successfully opened the database
		OpenpalmtbDatabase = True

	Else
		'We failed to open the database
		OpenpalmtbDatabase = False
		#If APPFORGE Then
		MsgBox "Could not open database - palmtb", vbExclamation
		#Else
		MsgBox "Could not open database - " + App.Path + "\palmtb.pdb" + vbCrLf + vbCrLf + "Potential causes are:" + vbCrLf + "1. Database file does not exist" + vbCrLf + "2. The database path in the PDBOpen call is incorrect", vbExclamation
		#End If
	End If

End Function


Public Sub ClosepalmtbDatabase()

	' Close the database
	PDBClose dbpalmtb
	dbpalmtb = 0

End Sub


Public Function ReadpalmtbRecord(MyRecord As tpalmtbRecord) As Boolean

	ReadpalmtbRecord = PDBReadRecord(dbpalmtb, VarPtr(MyRecord))

End Function


Public Function WritepalmtbRecord(MyRecord As tpalmtbRecord) As Boolean

	WritepalmtbRecord = PDBWriteRecord(dbpalmtb, VarPtr(MyRecord))

End Function
