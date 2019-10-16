Sub cmd_xcopy()

    'コマンドプロンプトのオブジェクト
    Dim wsh As IWshRuntimeLibrary.WshShell
    Set wsh = New IWshRuntimeLibrary.WshShell
    
    Dim OriginalFolderPath As String, NewFolderPath As String
    OriginalFolderPath = Range("B1").Value
    NewFolderPath = Range("B2").Value
    
    '実行コマンドを作る
    Dim command As String
    command = "xcopy /t /e " & OriginalFolderPath & " " & NewFolderPath
    
    'コマンドを実行
    wsh.Run "%ComSpec% /c " & command

End Sub
