Option Explicit

Dim Shell, root
Set Shell = CreateObject("WScript.Shell")

Private Sub Init()
	Dim FSO, Dirs(4), Y
	root = Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\Root")
	Set Shell = CreateObject("WScript.Shell")
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Dirs(0) = root
	Dirs(1) = FSO.BuildPath(root, "spool")
	Dirs(2) = FSO.BuildPath(root, "spool\sent")
	Dirs(3) = FSO.BuildPath(root, "spool\outgoing")
	Dirs(4) = FSO.BuildPath(root, "logs")
	For Y = 0 To 4
		If FSO.FolderExists(Dirs(Y)) = False Then
			FSO.CreateFolder Dirs(Y)
		End If
	Next
End Sub

Private Sub Main()
	Init()
	ProcessMail()
End Sub

Private Sub ProcessMail()
	Dim FSO, Folder, FileInfo, Mailer, File, FileContent, strLine, numLine, SentDir
	
	Set Mailer = CreateObject("CDO.Message") 
	root = Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\Root")
	Mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") 		= Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\SmtpServer")
	Mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") 	= Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\SmtpServerPort")
	Mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") 		= Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\SendUsing")
	Mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\SmtpAuthenticate")
	Mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") 		= Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\SmtpUseSSL")
	Mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") 	= Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\SendUsername")
	Mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") 	= Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\SendPassword")
	Mailer.Configuration.Fields.Update

	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set Folder = FSO.GetFolder(FSO.BuildPath(root, "spool\outgoing"))
	For Each FileInfo In Folder.Files
		Log("Process " & FileInfo.Path)
		Mailer.From = "risk@byterox.com" 
		Mailer.To = "achernov@byterox.com" 
		'Mailer.Cc = "ester@esterholdings.com"  
		Set File = FSO.OpenTextFile(FileInfo.Path)
		numLine = 0
		Do Until File.AtEndOfStream
			If numLine = 0 Then
				Mailer.Subject = File.ReadLine()
			Else
				strLine = File.ReadLine()
				FileContent = FileContent & strLine & vbCrLf
			End If
			numLine = numLine + 1
		Loop
		File.Close()
		Mailer.TextBody = FileContent
		Mailer.Send()
		SentDir = FSO.BuildPath(root, "spool\sent")
		FSO.CopyFile FileInfo.Path, FSO.BuildPath(SentDir, FileInfo.Name), True
		FileInfo.Delete()
	Next
End Sub


Private Sub Log(strLogMessage)
	Dim FSO, LogFile, FileName, LogDir, FullPath
	FileName = "app_" & Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & ".log"
	root = Shell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\DevAndreas\VbMailer\Root")
	Set FSO = CreateObject("Scripting.FileSystemObject")
	LogDir = FSO.BuildPath(root, "logs")
	FullPath = FSO.BuildPath(LogDir, FileName)
	Set LogFile = FSO.OpenTextFile(FullPath, 8, True)	' 8 - ForAppending, 1 - ForReading
	LogFile.Write Right("0" & Hour(Now()), 2) & ":" & Right("0" & Minute(Now()), 2) & ":" & Right("0" & Second(Now()), 2) & vbTab & strLogMessage & vbCrLf
	LogFile.Close()
End Sub

Main()