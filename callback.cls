VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public JobType As String     ' E=Email Server, F=FaxServer, P=PrintServer, G=paged by E-mail
Public ErrResult As Long     ' Err number from a server
Public ErrMessage As String  ' Err Message from a server
Public recNum As Long        ' This is InCounter field of the 'tblInfoRecord' table (Using this key I save the result back to this table)
Public txtFileName As String ' full path+filename+extension
Public MessageToDeliver As String
Public ProfileID As String

Public Event Sent(sProfile As String, lErr As Long, sErrMsg As String)





Public Function Notify(ByVal parErrNumber As Long, ByVal parErrMsg As String, _
    ByVal parJobType As String, ByVal parRecordNumber As Long, ByVal parTextFile As String)

    RaiseEvent Sent(Me.ProfileID, parErrNumber, parErrMsg)
    recNum = 123
    
End Function


