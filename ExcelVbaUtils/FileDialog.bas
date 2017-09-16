Attribute VB_Name = "FileDialog"
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

' Dependency: FileUtility

#Const ENABLE_TEST_METHODS = 1

' reference:
' Excel ファイルパス、フォルダパスを選択しセルに格納する参照ボタン
' http://qiita.com/boss_ape/items/1733fe6317e4566fdebb

' -----------------------------------------------------------------------------
' フォルダの選択ダイアログを表示します。
' 初期フォルダが存在しない場合は、指定されたスペシャルフォルダを初期フォルダとして表示します。
' <IN> defaultPath As String 初期フォルダ。指定しない場合は空文字を指定します。
' <IN> fallbackSpecialFolder As String SPECIALFOLDERS_ で始まるスペシャルフォルダの識別子。defaultPath が空白または無効な場合に採用されます。
' <OUT> String 選択されたフォルダ。キャンセルされた場合は空文字を返します。
' -----------------------------------------------------------------------------
Public Function SelectFolderWithDialog(defaultPath As String, fallbackSpecialFolder As String) As String
    
    Dim ofdFolderDlg As Office.FileDialog
    Dim openPath As String

    ' 初期フォルダの設定
    If Len(defaultPath) > 0 Then
        ' 末尾の "\" 削除
        If Right(defaultPath, 1) = "\" Then
            openPath = Left(defaultPath, Len(defaultPath) - 1)
        Else
            openPath = defaultPath
        End If

        ' フォルダ存在チェック
        If Not FolderExists(openPath) Then
            openPath = GetSpecialFolder(fallbackSpecialFolder)
        End If
    Else
        openPath = GetSpecialFolder(fallbackSpecialFolder)
    End If

    ' フォルダ選択ダイアログ設定
    Set ofdFolderDlg = Application.FileDialog(msoFileDialogFolderPicker)
    With ofdFolderDlg
        ' 表示するアイコンの大きさを指定
        .InitialView = msoFileDialogViewDetails
        ' フォルダ初期位置
        .InitialFileName = openPath & "\"
        ' 複数選択不可
        .AllowMultiSelect = False
    End With

    ' フォルダ選択ダイアログ表示
    If ofdFolderDlg.Show() = -1 Then
        ' フォルダパス設定
        SelectFolderWithDialog = ofdFolderDlg.SelectedItems(1)
    End If

    Set ofdFolderDlg = Nothing
    
End Function

#If ENABLE_TEST_METHODS = 1 Then

' -----------------------------------------------------------------------------
' SelectFolderWithDialog メソッドのテストを行います。
' -----------------------------------------------------------------------------
Public Sub SelectFolderWithDialogTest()

    Debug.Print "SelectFolderWithDialog=" & SelectFolderWithDialog("", SPECIALFOLDERS_MYDOCUMENTS)
    
End Sub

#End If
