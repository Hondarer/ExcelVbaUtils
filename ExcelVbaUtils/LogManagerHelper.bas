Attribute VB_Name = "LogManagerHelper"
Option Explicit
' -----------------------------------------------------------------------------
'
' �u�b�N�ɂ� LogManager �𗘗p����ۂ̏������T���v��
'
' -----------------------------------------------------------------------------

#Const ENABLE_TEST_METHODS = 1

' ���O�t�@�C���p�̃T�u�t�H���_��\���܂��B
Private Const DEFAULT_LOGPATH = "log"

' ���샍�O��\���L�[��\���܂��B
Public Const KIND_OPERATION = "Operation"

' ���̃u�b�N�ɂ����郍�O�o�͂̏��������������Ă��邩�ǂ�����ێ����܂��B
Private LogManagerHelper_Initialized As Boolean

' -----------------------------------------------------------------------------
' ���̃u�b�N�ɂ����郍�O�o�͂̏����������܂��B
' -----------------------------------------------------------------------------
Public Sub LogManagerHelper_LocalInit()

    If LogManagerHelper_Initialized = True Then
        Exit Sub
    End If

    ' ----- �f�o�b�O�p���O�̏����� -----
    If GetRangeValue("Config_DbgLogEnable", True) = True Then
        ' �f�t�H���g�� Logger ��ݒ�
        Call SetDefaultLogger(New LoggerCore)
        
        ' �e���x���̎�M�ݒ�𔽉f
        GetDefaultLogger.IsDebugEnabled = GetRangeValue("Config_DbgLogDebug", False)
        GetDefaultLogger.IsInfoEnabled = GetRangeValue("Config_DbgLogInfo", True)
        GetDefaultLogger.IsWarnEnabled = GetRangeValue("Config_DbgLogWarn", True)
        GetDefaultLogger.IsErrorEnabled = GetRangeValue("Config_DbgLogError", True)
        GetDefaultLogger.IsFatalEnabled = GetRangeValue("Config_DbgLogFatal", True)
        
        ' Debug.Print �Ή��� Appender ��ǉ�
        If GetRangeValue("Config_DbgLogIdeOut", True) = True Then
            Call GetDefaultLogger.RegistAppender(New DebugPrintAppender)
        End If
        ' OutputDebugString �Ή��� Appender ��ǉ�
        If GetRangeValue("Config_DbgLogDebuggerOut", True) = True Then
            Call GetDefaultLogger.RegistAppender(New OutputDebugStringAppender)
        End If
        
        ' ���O�t�@�C���p�� Appender ��ǉ�
        If GetRangeValue("Config_DbgLogFileOut", True) = True Then
            Dim apLogFile As TextFileAppender
            Set apLogFile = New TextFileAppender
            ' ���O�t�H���_���쐬
            Call TryMakeDir(GetAbsolutePathNameFromThisWorkbookPath(GetRangeValue("Config_DbgLogFilePath", DEFAULT_LOGPATH)))
            ' ���O�t�@�C�����̑g�ݗ���
            apLogFile.filePath = GetAbsolutePathNameFromThisWorkbookPath(GetRangeValue("Config_DbgLogFilePath", DEFAULT_LOGPATH)) & "\" & RemoveExtension(ThisWorkbook.Name) & "_dbg.log"
            Call GetDefaultLogger.RegistAppender(apLogFile)
        End If
    Else
        ' �t�H�[���o�b�N�p�̂� Logger ��ݒ�
        Call SetDefaultLogger(New FallbackLogger)
    End If
    
    ' ----- ���샍�O�̏����� -----
    If GetRangeValue("Config_OpeLogEnable", True) = True Then
        ' �� Logger ���N���A
        Call ClearLoggers
        ' ���샍�O�p Logger ��ǉ�
        Call RegistLogger(KIND_OPERATION, New LoggerCore)
          
        ' �e���x���̎�M�ݒ�𔽉f
        GetLogger(KIND_OPERATION).IsDebugEnabled = GetRangeValue("Config_OpeLogDebug", False)
        GetLogger(KIND_OPERATION).IsInfoEnabled = GetRangeValue("Config_OpeLogInfo", True)
        GetLogger(KIND_OPERATION).IsWarnEnabled = GetRangeValue("Config_OpeLogWarn", True)
        GetLogger(KIND_OPERATION).IsErrorEnabled = GetRangeValue("Config_OpeLogError", True)
        GetLogger(KIND_OPERATION).IsFatalEnabled = GetRangeValue("Config_OpeLogFatal", True)
      
        ' ���샍�O�p�� Appender ��ǉ�
        If GetRangeValue("Config_OpeLogFileOut", True) = True Then
            Dim apOperationLogFile As TextFileAppender
            Set apOperationLogFile = New TextFileAppender
            ' ���샍�O�t�H���_���쐬
            Call TryMakeDir(GetAbsolutePathNameFromThisWorkbookPath(GetRangeValue("Config_OpeLogFilePath", DEFAULT_LOGPATH)))
            ' ���샍�O�p�t�@�C�����̑g�ݗ���
            apOperationLogFile.filePath = GetAbsolutePathNameFromThisWorkbookPath(GetRangeValue("Config_OpeLogFilePath", DEFAULT_LOGPATH)) & "\" & RemoveExtension(ThisWorkbook.Name) & "_ope.log"
            Call GetLogger(KIND_OPERATION).RegistAppender(apOperationLogFile)
        End If
        
        ' ���샍�O�p�� Logger �ɗv�������ۂɁA�f�t�H���g�� Logger �ɑ΂��Ă�
        ' �o�͂���悤�ɁA���샍�O�� Logger �̎q�Ƃ��ăf�t�H���g�� Logger ��ݒ�
        If GetRangeValue("Config_OpeLogDbgRelay", True) = True Then
            Call GetLogger(KIND_OPERATION).RegistChild(GetDefaultLogger)
        End If
    End If
    
    GetDefaultLogger.LogInfo "[LogManagerHelper_LocalInit] ���O�Ǘ��@�\�����������܂���"
    
    LogManagerHelper_Initialized = True

End Sub

' -----------------------------------------------------------------------------
' ���̃u�b�N�ɂ����郍�O�o�͂̏��������������܂��B
' -----------------------------------------------------------------------------
Public Sub LogManagerHelper_Reset()
    If LogManagerHelper_Initialized = True Then
        LogManagerHelper_Initialized = False
        GetDefaultLogger.LogInfo "[LogManagerHelper_LocalInit] ���O�Ǘ��@�\�̏�������v�����܂���"
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
