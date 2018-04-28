' ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
' [ファイル名] dirsize.vbs
' [  概  要  ] 指定されたフォルダ内の、サブフォルダ・ファイルの各容量を取得します。
' [ 使用方法 ] cscript dirsize.vbs <フォルダパス>
'              もしくは、
'              dirsize.vbs のアイコンへ、調べたいフォルダを、ドラッグ＆ドロップ
'
' -更新日- ---更新者--- ------------------------内容---------------------------
' 12/08/09 seki-moto    新規作成
' ───────────────────────────────────────
Option Explicit

' サイズの表示桁数
Public Const SPACE_SIZE = 19

' 主処理実行
WScript.Quit( Main() )


'*****************************************************************************
'[関数名] Main
'[ 概要 ] 主処理
'
'-更新日- ---更新者--- ------------------------内容---------------------------
'12/08/09 seki-moto    新規作成
'=============================================================================
Function Main()

    ' 実行環境チェック
    Call RunExeCheck()

    Dim Path
    Dim FSO
    Dim Folder
    Dim FolderSub
    Dim Files
    Dim FormatDir

    If WScript.Arguments.Count <> 1 Then
        ' %ERRORLEVEL% == 1
        Main = 1
        Exit Function
    End If

    ' パラメータの取得
    Path = WScript.Arguments(0)

    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' パラメータの正当性チェック
    If FSO.FolderExists(Path) = True Then

        Set Folder = FSO.GetFolder(Path)

        WScript.echo Folder.Path & " のディレクトリ"
        WScript.echo ""

        ' フォルダのサイズ取得・表示
        For Each FolderSub In Folder.SubFolders
            On Error Resume Next
                WScript.echo EchoFormatDir( FolderSub.Size, FolderSub.Path )
            On Error Goto 0
        Next
        Set FolderSub = Nothing

        ' ファイルのサイズ取得・表示
        For Each Files In Folder.Files
            WScript.echo EchoFormatFile( Files.Size, Files.Path )
        Next
        Set Files = Nothing

        ' %ERRORLEVEL% == 0
        Main = 0

    Else

        WScript.echo Path & " は、フォルダではないか、存在しません。"
        ' %ERRORLEVEL% == 2
        Main = 2

    End If

    Set Folder = Nothing
    Set FSO = Nothing

End Function
'*****************************************************************************

'*****************************************************************************
'[関数名] EchoFormatDir
'[ 概要 ] フォルダの表示書式
'
'[ 引数 ] size：フォルダ容量
'         name：フォルダのパス
'
'[返り値] 改行なしの表示書式
'
'-更新日- ---更新者--- ------------------------内容---------------------------
'12/08/09 seki-moto    新規作成
'=============================================================================
Function EchoFormatDir(ByVal size, ByVal name)

    Dim strRet
    Dim strSize
    strSize = FormatNumber(size, 0, -2, -2, -1)

    strRet = Space( SPACE_SIZE - Len(strSize) ) & strSize
    strRet = strRet & " <DIR> "
    strRet = strRet & name & "\"

    EchoFormatDir = strRet

End Function
'*****************************************************************************

'*****************************************************************************
'[関数名] EchoFormatFile
'[ 概要 ] ファイルの表示書式
'
'[ 引数 ] size：フォルダ容量
'         name：フォルダのパス
'
'[返り値] 改行なしの表示書式
'
'-更新日- ---更新者--- ------------------------内容---------------------------
'12/08/09 seki-moto    新規作成
'=============================================================================
Function EchoFormatFile(ByVal size, ByVal name)

    Dim strRet
    Dim strSize
    strSize = FormatNumber(size, 0, -2, -2, -1)

    strRet = Space( SPACE_SIZE - Len(strSize) ) & strSize
    strRet = strRet & "       "
    strRet = strRet & name

    EchoFormatFile = strRet

End Function
'*****************************************************************************


'*****************************************************************************
'[関数名] RunExeCheck
'[ 概要 ] 実行環境チェック
'
'-更新日- ---更新者--- ------------------------内容---------------------------
'12/08/09 seki-moto    新規作成
'=============================================================================
Sub RunExeCheck()
    If LCase(Right(WScript.FullName, 11)) <> "cscript.exe" Then

        Dim FSO
        Dim wkDir
        Dim Path
        Dim objWShell
        Dim strCmd
        Dim retQuit

        Set FSO = CreateObject("Scripting.FileSystemObject")
        wkDir = FSO.GetFile( WScript.ScriptFullName ).ParentFolder.Path
        Set FSO = Nothing

        If WScript.Arguments.Count = 1 Then
            ' パラメータの取得
            Path = """" & WScript.Arguments(0) & """"
        Else
            Path = ""
        End If

        Set objWShell = Createobject("WScript.Shell")

        If Path = "" Then
            strCmd = "cmd /K ""cd """ & wkDir & """ && echo 使い方：cscript dirsize.vbs フォルダパス """
        Else
            strCmd = "cmd /K ""cd """ & wkDir & """ && cscript /nologo dirsize.vbs " & Path & " """
        End If

        ' cscript の別プロセスを立ち上げ
        retQuit = objWShell.Run( strCmd, 1, True )

        Set objWShell = Nothing

        WScript.Quit(retQuit)
    End If
End Sub
'*****************************************************************************
' ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
