On Error Resume Next
Dim Args, Filename, Path, Fullpath, Pathdir, Shellobj, Sourceobj, Targetobj, intOptions
Set Args = WScript.Arguments
WScript.Echo "Command Line Unzip Utilty by Govind Parmar"
If(WScript.Arguments.Count > 0) Then
	Filename = Args(0)
Else
	WScript.Echo "No filename specified"
	WScript.Quit
End If
Path = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
Fullpath = Path & "\" & Filename
Pathdir = Path & "\"
Set Shellobj = CreateObject("Shell.Application")
Set Sourceobj = Shellobj.NameSpace(Fullpath).Items()
If Err.Number <> 0 Then
	WScript.Echo "Couldn't open '" + Filename + "'"
	WScript.Quit
End If
Set Targetobj = Shellobj.NameSpace(Pathdir)
WScript.Echo("Extracting file: " & Filename)
intOptions = 256
Targetobj.CopyHere Sourceobj, intOptions
WScript.Echo("Done.")
WScript.Quit

