Attribute VB_Name = "PInvoke"
Option Explicit
' -----------------------------------------------------------------------------
' ExcelChildProcessController
' https://github.com/Hondarer/ExcelChildProcessController
' -----------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2017 Tetsuo Honda
' t-honda@hondarer-soft.com
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
' 以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル
' （以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを
' 無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、
' 変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、
' およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。
'
' 上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な
' 部分に記載するものとします。
'
' ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、
' 何らの保証もなく提供されます。
' ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての
' 保証も含みますが、それに限定されるものではありません。
' 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、
' ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の
' 扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わない
' ものとします。
'
' -----------------------------------------------------------------------------

Public Const STARTF_USESHOWWINDOW = &H1
Public Const STARTF_USESTDHANDLES = &H100

Public Const SW_HIDE = 0

Public Const INFINITE As Long = &HFFFF

Public Const NORMAL_PRIORITY_CLASS As Long = &H20

Public Const HANDLE_FLAG_INHERIT = &H1

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const STD_ERROR_HANDLE = -12&

Public Const STILL_ACTIVE = &H103

Public Const EXIT_FAILURE = 1

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" ( _
    ByRef dest As Any, _
    ByVal numBytes As Long)

Public Declare Function CreatePipe Lib "kernel32" ( _
    ByRef phReadPipe As Long, _
    ByRef phWritePipe As Long, _
    ByRef lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long

Public Declare Function SetHandleInformation Lib "kernel32" ( _
    ByVal hObject As Long, _
    ByVal dwMask As Long, _
    ByVal dwFlags As Long) As Long

Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As Any, _
    lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    lpEnvironment As Any, _
    ByVal lpCurrentDriectory As String, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long
    
Public Declare Function GetStdHandle Lib "kernel32" ( _
    ByVal nStdHandle As Long) As Long

Public Declare Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As Long, _
    ByRef lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    ByRef lpBytesRead As Long, _
    ByRef lpTotalBytesAvail As Long, _
    ByRef lpBytesLeftThisMessage As Long) As Long

Public Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, _
    ByRef lpOverlapped As OVERLAPPED) As Long

Public Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    ByRef lpOverlapped As OVERLAPPED) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
Public Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Public Declare Function GetExitCodeProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    lpExitCode As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long

Public Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Byte, _
    ByRef Source As Byte, _
    ByVal Length As Long)

Public Declare Sub GetSystemTime Lib "kernel32" ( _
    ByRef lpSystemTime As SYSTEMTIME)

Public Declare Sub GetLocalTime Lib "kernel32" ( _
    ByRef lpSystemTime As SYSTEMTIME)

Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" ( _
    ByVal lpOutputString As String)



