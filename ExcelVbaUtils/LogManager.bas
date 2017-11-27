Attribute VB_Name = "LogManager"
Option Explicit

Private logger As ILog

Public Function GetLogger() As ILog
    
    If logger Is Nothing Then
        Set logger = New LoggerCore
        Debug.Print "new LoggerCore created."
    End If
    
    Set GetLogger = logger

End Function


Public Function test()
    GetLogger.LogDebug ("test")
End Function
