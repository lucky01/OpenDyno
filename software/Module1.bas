Attribute VB_Name = "Module1"
Declare Function OpenProcess Lib "kernel32.dll" _
                 ( _
                 ByVal dwDesiredAccess As Long, _
                 ByVal bInheritHandle As Long, _
                 ByVal dwProcessId As Long) As Long

Declare Function WaitForSingleObject Lib "kernel32" ( _
                 ByVal hHandle As Long, _
                 ByVal dwMilliseconds As Long) As Long
                 

Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long


 
