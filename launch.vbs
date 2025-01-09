Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' VBSファイルのパスを取得
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)

' バッチファイルの相対パス（VBSの場所を基準に指定）
batchFile = currentDir & "\launch.bat"

' 非表示で実行
WshShell.Run batchFile, 0, False
