Set shell = CreateObject("WScript.Shell")
shell.CurrentDirectory = "current working directory"
shell.Run "python AI_News.py",0
Set WshShell = Nothing 