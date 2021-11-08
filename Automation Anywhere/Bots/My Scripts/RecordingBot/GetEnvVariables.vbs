Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("System")

WScript.StdOut.WriteLine objEnv("rpa_path")