VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private myScheduleCol As New Collection

Public Sub Add(strHeader As String, intRecordType As Integer, strExtension As String, intRecordCount As Integer, strGenInfo As String)
        On Error GoTo OOPS
        Dim myClass As New ClsSchedule
        
        myClass.GetScheduleRecord strHeader, intRecordType, strExtension, intRecordCount, strGenInfo
        myScheduleCol.Add myClass
exiterror:
        Exit Sub
OOPS:
        LogMessage ErrorMessage, "ColSchedule_Add: " & Err.Description
        Resume exiterror
End Sub

Public Function Item(index As Variant)
      Set Item = myScheduleCol.Item(index)
End Function
Public Function Count() As Integer
    Count = myScheduleCol.Count
End Function
Public Sub Remove(index As Variant)
    myScheduleCol.Remove (index)
End Sub
Public Sub DeleteRecord(index As Variant)
    On Error GoTo OOPS
    Dim intRecCount As Integer
    Dim strExtension As String * 10
    Dim intRecordType As Integer
    intRecCount = myScheduleCol.Item(index).RecordCount
    RSet strExtension = Trim(myScheduleCol.Item(index).Extension)
    intRecordType = myScheduleCol.Item(index).RecordType
    myScheduleCol.Remove index
    DeleteScheduleRecord intRecordType, intRecCount, strExtension
    
exiterror:
    Exit Sub
    
OOPS:
    LogMessage ErrorMessage, "DeleteRecord :" & Err.Description
    Resume exiterror
End Sub
