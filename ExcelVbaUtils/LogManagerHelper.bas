Attribute VB_Name = "LogManagerHelper"
Option Explicit
' -----------------------------------------------------------------------------
'
' ブックにて LogManager を利用する際の初期化サンプル
'
' -----------------------------------------------------------------------------

#Const ENABLE_TEST_METHODS = 1

' ログファイル用のサブフォルダを表します。
Private Const DEFAULT_LOGPATH = "log"

' 操作ログを表すキーを表します。
Public Const KIND_OPERATION = "Operation"

' このブックにおけるログ出力の初期化が完了しているかどうかを保持します。
Private LogManagerHelper_Initialized As Boolean

' -----------------------------------------------------------------------------
' このブックにおけるログ出力の初期化をします。
' -----------------------------------------------------------------------------
Public Sub LogManagerHelper_LocalInit()

    If LogManagerHelper_Initialized = True Then
        Exit Sub
    End If

    ' ----- デバッグ用ログの初期化 -----
    If GetRangeValue("Config_DbgLogEnable", True) = True Then
        ' デフォルトの Logger を設定
        Call SetDefaultLogger(New LoggerCore)
        
        ' 各レベルの受信設定を反映
        GetDefaultLogger.IsDebugEnabled = GetRangeValue("Config_DbgLogDebug", False)
        GetDefaultLogger.IsInfoEnabled = GetRangeValue("Config_DbgLogInfo", True)
        GetDefaultLogger.IsWarnEnabled = GetRangeValue("Config_DbgLogWarn", True)
        GetDefaultLogger.IsErrorEnabled = GetRangeValue("Config_DbgLogError", True)
        GetDefaultLogger.IsFatalEnabled = GetRangeValue("Config_DbgLogFatal", True)
        
        ' Debug.Print 対応の Appender を追加
        If GetRangeValue("Config_DbgLogIdeOut", True) = True Then
            Call GetDefaultLogger.RegistAppender(New DebugPrintAppender)
        End If
        ' OutputDebugString 対応の Appender を追加
        If GetRangeValue("Config_DbgLogDebuggerOut", True) = True Then
            Call GetDefaultLogger.RegistAppender(New OutputDebugStringAppender)
        End If
        
        ' ログファイル用の Appender を追加
        If GetRangeValue("Config_DbgLogFileOut", True) = True Then
            Dim apLogFile As TextFileAppender
            Set apLogFile = New TextFileAppender
            ' ログフォルダを作成
            Call TryMakeDir(GetAbsolutePathNameFromThisWorkbookPath(GetRangeValue("Config_DbgLogFilePath", DEFAULT_LOGPATH)))
            ' ログファイル名の組み立て
            apLogFile.filePath = GetAbsolutePathNameFromThisWorkbookPath(GetRangeValue("Config_DbgLogFilePath", DEFAULT_LOGPATH)) & "\" & RemoveExtension(ThisWorkbook.Name) & "_dbg.log"
            Call GetDefaultLogger.RegistAppender(apLogFile)
        End If
    Else
        ' フォールバック用のの Logger を設定
        Call SetDefaultLogger(New FallbackLogger)
    End If
    
    ' ----- 操作ログの初期化 -----
    If GetRangeValue("Config_OpeLogEnable", True) = True Then
        ' 個別 Logger をクリア
        Call ClearLoggers
        ' 操作ログ用 Logger を追加
        Call RegistLogger(KIND_OPERATION, New LoggerCore)
          
        ' 各レベルの受信設定を反映
        GetLogger(KIND_OPERATION).IsDebugEnabled = GetRangeValue("Config_OpeLogDebug", False)
        GetLogger(KIND_OPERATION).IsInfoEnabled = GetRangeValue("Config_OpeLogInfo", True)
        GetLogger(KIND_OPERATION).IsWarnEnabled = GetRangeValue("Config_OpeLogWarn", True)
        GetLogger(KIND_OPERATION).IsErrorEnabled = GetRangeValue("Config_OpeLogError", True)
        GetLogger(KIND_OPERATION).IsFatalEnabled = GetRangeValue("Config_OpeLogFatal", True)
      
        ' 操作ログ用の Appender を追加
        If GetRangeValue("Config_OpeLogFileOut", True) = True Then
            Dim apOperationLogFile As TextFileAppender
            Set apOperationLogFile = New TextFileAppender
            ' 操作ログフォルダを作成
            Call TryMakeDir(GetAbsolutePathNameFromThisWorkbookPath(GetRangeValue("Config_OpeLogFilePath", DEFAULT_LOGPATH)))
            ' 操作ログ用ファイル名の組み立て
            apOperationLogFile.filePath = GetAbsolutePathNameFromThisWorkbookPath(GetRangeValue("Config_OpeLogFilePath", DEFAULT_LOGPATH)) & "\" & RemoveExtension(ThisWorkbook.Name) & "_ope.log"
            Call GetLogger(KIND_OPERATION).RegistAppender(apOperationLogFile)
        End If
        
        ' 操作ログ用の Logger に要求した際に、デフォルトの Logger に対しても
        ' 出力するように、操作ログの Logger の子としてデフォルトの Logger を設定
        If GetRangeValue("Config_OpeLogDbgRelay", True) = True Then
            Call GetLogger(KIND_OPERATION).RegistChild(GetDefaultLogger)
        End If
    End If
    
    GetDefaultLogger.LogInfo "[LogManagerHelper_LocalInit] ログ管理機能を初期化しました"
    
    LogManagerHelper_Initialized = True

End Sub

' -----------------------------------------------------------------------------
' このブックにおけるログ出力の初期化を強制します。
' -----------------------------------------------------------------------------
Public Sub LogManagerHelper_Reset()
    If LogManagerHelper_Initialized = True Then
        LogManagerHelper_Initialized = False
        GetDefaultLogger.LogInfo "[LogManagerHelper_LocalInit] ログ管理機能の初期化を要求しました"
    End If
    Call LogManagerHelper_LocalInit
End Sub

#If ENABLE_TEST_METHODS = 1 Then

Public Sub LogManagerHelper_LocalInitTest()

    Call LogManagerHelper_LocalInit
    
    GetDefaultLogger.LogDebug "LogManagerHelper_LocalInitTestDebug"
    GetDefaultLogger.LogFatal "LogManagerHelper_LocalInitTestFatal"
    GetLogger(KIND_OPERATION).LogInfo "OperationLogTest"
    
End Sub

#End If
