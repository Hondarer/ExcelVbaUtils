Attribute VB_Name = "LogManagerHelper"
Option Explicit
' -----------------------------------------------------------------------------
'
' ブックにて LogManager を利用する際の初期化サンプル
'
' -----------------------------------------------------------------------------

#Const ENABLE_TEST_METHODS = 1

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
    ' デフォルトの Logger に紐づく Appender をクリア
    Call GetDefaultLogger.ClearAppenders
    ' Debug.Print 対応の Appender を追加
    Call GetDefaultLogger.RegistAppender(New DebugPrintAppender)
    ' OutputDebugString 対応の Appender を追加
    Call GetDefaultLogger.RegistAppender(New OutputDebugStringAppender)
    
    ' ログファイル用の Appender を追加
    Dim apLogFile As textFileAppender
    Set apLogFile = New textFileAppender
    ' ログファイル名の組み立て
    apLogFile.filePath = ThisWorkbook.Path & "\" & RemoveExtension(ThisWorkbook.Name) & "_debug.log"
    Call GetDefaultLogger.RegistAppender(apLogFile)
    
    ' ----- 操作ログの初期化 -----
    ' 個別 Logger をクリア
    Call ClearLoggers
    ' 操作ログ用 Logger を追加
    Call RegistLogger(KIND_OPERATION, New LoggerCore)
    
    ' 操作ログ用の Appender を追加
    Dim apOperationLogFile As textFileAppender
    Set apOperationLogFile = New textFileAppender
    ' 操作ログ用ファイル名の組み立て
    apOperationLogFile.filePath = ThisWorkbook.Path & "\" & RemoveExtension(ThisWorkbook.Name) & "_operation.log"
    Call GetLogger(KIND_OPERATION).RegistAppender(apOperationLogFile)
    
    ' 操作ログ用の Logger に要求した際に、デフォルトの Logger に対しても
    ' 出力するように、操作ログの Logger の子としてデフォルトの Logger を設定
    Call GetLogger(KIND_OPERATION).RegistChild(GetDefaultLogger)
    
    LogManagerHelper_Initialized = True

End Sub

#If ENABLE_TEST_METHODS = 1 Then

Public Sub LogManagerHelper_LocalInitTest()

    Call LogManagerHelper_LocalInit
    
    GetDefaultLogger.LogDebug "LogManagerHelper_LocalInitTestDebug"
    GetDefaultLogger.LogFatal "LogManagerHelper_LocalInitTestFatal"
    GetLogger(KIND_OPERATION).LogInfo "OperationLogTest"
    
End Sub

#End If
