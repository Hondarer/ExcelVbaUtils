Attribute VB_Name = "FileUtility"
Option Explicit
' -----------------------------------------------------------------------------
' ExcelVbaUtils
' https://github.com/Hondarer/ExcelVbaUtils
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

' Dependency: None

#Const ENABLE_TEST_METHODS = 1

' FileSystemObject を保持します。
Dim fso As Object

' -----------------------------------------------------------------------------
' 指定されたフォルダが存在するか返します。
' <IN> folderName As String チェックするフォルダ名。
' <OUT> Boolean フォルダが存在する場合は True、存在しない場合は False。
' -----------------------------------------------------------------------------
Public Function FolderExists(folderName As String) As Boolean
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    FolderExists = fso.FolderExists(folderName)
    
    Set fso = Nothing

End Function

' -----------------------------------------------------------------------------
' 指定されたファイルが存在するか返します。
' <IN> folderName As String チェックするフォルダ名。
' <OUT> Boolean ファイルが存在する場合は True、存在しない場合は False。
' -----------------------------------------------------------------------------
Public Function FileExists(fileName As String) As Boolean
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    FileExists = fso.FileExists(fileName)
    
    Set fso = Nothing

End Function

' -----------------------------------------------------------------------------
' ファイル名の拡張子部分を返します。
' <IN> filePath As String 対象のファイルパス。
' <OUT> String ピリオドを含む拡張子部分。ピリオドが見つからない場合は、空文字。
' -----------------------------------------------------------------------------
Public Function GetExtension(filePath As String) As String

    If InStr(filePath, ".") = 0 Then
        Exit Function
    End If
    
    If InStrRev(filePath, ".") < InStrRev(filePath, "\") Then
        ' フルパスで、上位のフォルダ名にピリオドが含まれている場合
        Exit Function
    End If

    GetExtension = Mid(filePath, InStrRev(filePath, "."))

End Function

' -----------------------------------------------------------------------------
' ファイル名の拡張子を除いた部分を返します｡
' <IN> filePath As String 対象のファイルパス。
' <OUT> String 拡張子を取り除いたファイルパス。ピリオドが見つからない場合は、入力をそのまま返します。
' -----------------------------------------------------------------------------
Public Function RemoveExtension(filePath As String) As String

    RemoveExtension = Left(filePath, Len(filePath) - Len(GetExtension(filePath)))

End Function

' -----------------------------------------------------------------------------
' ワークブックからのパスを、絶対パスに変換します。
' <IN> workbookPath As String ワークブックからの相対パスか、絶対パス。
' <OUT> String 解決された絶対パス。
' -----------------------------------------------------------------------------
Public Function GetAbsolutePathNameFromThisWorkbookPath(workbookPath As String) As String

    ' カレントディレクトリをブックのパスに設定する
    Call SetCurrentDirectory(ThisWorkbook.Path)
    ' パス名を解決する
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    GetAbsolutePathNameFromThisWorkbookPath = fso.GetAbsolutePathName(workbookPath)

End Function

' -----------------------------------------------------------------------------
' 指定されたサブフォルダーが存在するかチェックし、
' 存在しない場合は作成します。
' <IN> dirPath As String チェックするフォルダのパス。論理パスの場合はカレントディレクトリのカレントフォルダを基準にします。
' <OUT> Boolean ディレクトリが存在するか、作成に成功した場合は True。作成に失敗した場合は False。True の場合でも、出力に失敗する可能性があるため、出力時のエラーチェックは必ず実施してください。
' -----------------------------------------------------------------------------
Public Function TryMakeDir(dirPath As String) As Boolean
    
    Dim rtc As Long
    
    ' すでに目的のフォルダがあるか
    If PathIsDirectory(dirPath) = True Then
        TryMakeDir = True
        Exit Function
    End If
    
    rtc = SHCreateDirectoryEx(0&, dirPath, 0&)
    
    ' 正常に作成できた場合 NO_ERROR(0)
    ' 途中がファイルで再帰作成に失敗した場合 ERROR_PATH_NOT_FOUND(3)
    ' 既にディレクトリがある場合 ERROR_ALREADY_EXISTS(183) (最終階層がファイルの場合も ERROR_ALREADY_EXISTS のため、ただちに成功とはいえない)
    ' ただし、当該フォルダにファイルの生成権があるかどうかは、ここではチェックしていない。
    
    If rtc <> NO_ERROR Then
        ' log
    End If
    
    ' 最終階層がファイルの場合などを想定して、API で最終チェック
    TryMakeDir = PathIsDirectory(dirPath)

End Function

#If ENABLE_TEST_METHODS = 1 Then

Public Sub RemoveExtensionTest()
    Debug.Print GetExtension("aaa.txt")
    Debug.Print RemoveExtension("aaa.txt")
    Debug.Print GetExtension("bbb")
    Debug.Print RemoveExtension("bbb")
    Debug.Print GetExtension(".\ccc.txt")
    Debug.Print RemoveExtension(".\ccc.txt")
    Debug.Print GetExtension(".\ddd")
    Debug.Print RemoveExtension(".\ddd")
End Sub

Public Sub mkdirtest()
    
    Debug.Print TryMakeDir(GetAbsolutePathNameFromThisWorkbookPath("log\sub1\sub2"))
    
End Sub

#End If


