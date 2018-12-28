VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Header1 As String
Private RecordType1 As Integer
Private Extension1 As String
Private RecordCount1 As Integer
Private GeneralInformation1 As String
Property Get Header() As String
             Header = Header1
End Property
Property Get RecordType() As Integer
             RecordType = RecordType1
End Property
Property Get Extension() As String
             Extension = Extension1
End Property
Property Get RecordCount() As Integer
             RecordCount = RecordCount1
End Property
Property Get Information() As String
             Information = GeneralInformation1
End Property

Public Sub GetScheduleRecord(strHeader As String, intRecordType As Integer, strExtension As String, intRecordCount As Integer, strGenInfo As String)
            On Error GoTo OOPS
            Header1 = strHeader
            RecordType1 = intRecordType
            Extension1 = strExtension
            RecordCount1 = intRecordCount
            GeneralInformation1 = strGenInfo
            
exiterror:
            Exit Sub
OOPS:
            LogMessage ErrorMessage, "GetScheduleRecord :" & Err.Description
            Resume exiterror
End Sub

