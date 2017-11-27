VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILog

Private Property Get ILog_IsDebugEnabled() As Boolean

End Property

Private Property Get ILog_IsErrorEnabled() As Boolean

End Property

Private Property Get ILog_IsFatalEnabled() As Boolean

End Property

Private Property Get ILog_IsInfoEnabled() As Boolean

End Property

Private Property Get ILog_IsWarnEnabled() As Boolean

End Property

Private Sub ILog_LogDebug(message As String)
    Debug.Print "Hello" & message
End Sub

Private Sub ILog_LogError(message As String)

End Sub

Private Sub ILog_LogFatal(message As String)

End Sub

Private Sub ILog_LogInfo(message As String)

End Sub

Private Sub ILog_LogWarn(message As String)

End Sub
