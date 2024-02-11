Option Explicit
Dim ArgsLocalFile
Dim ArgsVersionNo
Dim CurrentDir
Dim objWshShell
Dim objFS
Dim CopySrcFullName
  '5秒(5000ミリ秒)処理を止めます。
  WScript.Sleep 5000
  '引数で渡されてきた値を変数に格納します。
  ArgsLocalFile = Wscript.Arguments(0)
  ArgsVersionNo = Wscript.Arguments(1)
  CurrentDir = "\\ファイルサーバ名\共有フォルダ名\○○システムAccess自動バージョンアップ"
  'コピー元のフルパスをセットします。
  CopySrcFullName = CurrentDir & "\" & ArgsVersionNo & "\○○システム.mdb"
  Set objFS = CreateObject("Scripting.FileSystemObject")
  Call objFS.CopyFile(CopySrcFullName, ArgsLocalFile)
  'オブジェクトを解放します。
  Set objFS = Nothing
  Set objWshShell = Nothing
  msgbox "アップデートが正常に終了しました。"