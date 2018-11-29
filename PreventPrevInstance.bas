Attribute VB_Name = "PreventPrevInstance"
Option Explicit

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private AppMutexHandle As Long
Private AppMutexName As String
Private Const WAIT_OBJECT_0 As Long = 0


Public Sub CleanUpMutex()

If AppMutexHandle <> 0 Then
    If ReleaseMutex(AppMutexHandle) Then
        ' mutex was released
    Else
        ' mutex was not released
    End If
End If

End Sub

Private Function IsIDE() As Boolean

On Error Resume Next
    
    Debug.Print 1 \ 0 'error, but only in IDE
    IsIDE = (Err.Number <> 0)


End Function

Public Function IsPrevInstanceRunning(ByVal parAppName As String) As Boolean
If IsIDE Then Exit Function  ' do nothing - this is VB environment !!!

AppMutexName = "Global\" & parAppName

' Create handle to the mutex:
AppMutexHandle = CreateMutex(0, False, AppMutexName & Chr(0))
If AppMutexHandle = Null Then
    ' Not OK
    IsPrevInstanceRunning = True
    Exit Function
Else
    ' OK, then
    ' attempt to secure ownership of the mutex:
    If (WAIT_OBJECT_0 <> WaitForSingleObject(AppMutexHandle, 1000)) Then
        ' instance of this app is already running
        IsPrevInstanceRunning = True
    Else
        IsPrevInstanceRunning = False
    End If
    
End If

End Function




