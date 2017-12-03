Attribute VB_Name = "LogManagerHelper"
Option Explicit
' -----------------------------------------------------------------------------
'
' �u�b�N�ɂ� LogManager �𗘗p����ۂ̏������T���v��
'
' -----------------------------------------------------------------------------

#Const ENABLE_TEST_METHODS = 1

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
    ' �f�t�H���g�� Logger �ɕR�Â� Appender ���N���A
    Call GetDefaultLogger.ClearAppenders
    ' Debug.Print �Ή��� Appender ��ǉ�
    Call GetDefaultLogger.RegistAppender(New DebugPrintAppender)
    ' OutputDebugString �Ή��� Appender ��ǉ�
    Call GetDefaultLogger.RegistAppender(New OutputDebugStringAppender)
    
    ' ���O�t�@�C���p�� Appender ��ǉ�
    Dim apLogFile As textFileAppender
    Set apLogFile = New textFileAppender
    ' ���O�t�@�C�����̑g�ݗ���
    apLogFile.filePath = ThisWorkbook.Path & "\" & RemoveExtension(ThisWorkbook.Name) & "_debug.log"
    Call GetDefaultLogger.RegistAppender(apLogFile)
    
    ' ----- ���샍�O�̏����� -----
    ' �� Logger ���N���A
    Call ClearLoggers
    ' ���샍�O�p Logger ��ǉ�
    Call RegistLogger(KIND_OPERATION, New LoggerCore)
    
    ' ���샍�O�p�� Appender ��ǉ�
    Dim apOperationLogFile As textFileAppender
    Set apOperationLogFile = New textFileAppender
    ' ���샍�O�p�t�@�C�����̑g�ݗ���
    apOperationLogFile.filePath = ThisWorkbook.Path & "\" & RemoveExtension(ThisWorkbook.Name) & "_operation.log"
    Call GetLogger(KIND_OPERATION).RegistAppender(apOperationLogFile)
    
    ' ���샍�O�p�� Logger �ɗv�������ۂɁA�f�t�H���g�� Logger �ɑ΂��Ă�
    ' �o�͂���悤�ɁA���샍�O�� Logger �̎q�Ƃ��ăf�t�H���g�� Logger ��ݒ�
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
