Option Explicit
dim objShell, objWshShell, objFolder, ZipFile, i

if WScript.Arguments.Count < 1 or WScript.Arguments.Count > 2 then
WScript.Echo "Usage: CScript.exe UnZip.VBS ZIPFile [objFolder]"
WScript.Quit
end if

Set objShell = CreateObject("shell.application")
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set ZipFile = objShell.NameSpace (WScript.Arguments(0)).items
if WScript.Arguments.Count = 2 then
Set objFolder = objShell.NameSpace (WScript.Arguments(1)) '指定された解凍先フォルダ
else
Set objFolder = objShell.NameSpace (objWshShell.CurrentDirectory) '省略時はカレントディレクトリへ
end if

objFolder.CopyHere ZipFile, &H14 '進行状況ダイアログボックス非表示 + ダイアログボックスは[すべてはい]
