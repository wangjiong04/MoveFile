'====================================================================================
'========	Created by 	: Rudy Pringadi <hansip101@yahoo.com>
'========	Date		: March 1, 2015
'========	Function	: Notify by Email if size of file is 1GB or More
'====================================================================================

SrcFolder = "D:\EDI2\Problem\julia\Output\"  

SrcFolder = "\\Covhouapp03\C$\Program Files\LexiCom\outbox\HP\FromHere\"
NotifyFileSizeIsZero SrcFolder

SrcFolder = "\\covhouapp03\C$\Program Files\LexiCom\outbox\Covisint\Tecumseh-Corrugated-Box\"
NotifyFileSizeIsZero SrcFolder

Sub NotifyFileSizeIsZero(SrcFldr)
	Set mfsox = CreateObject("Scripting.FileSystemobject")
	Set folder = mfsox.GetFolder(SrcFldr)
	brs = -1
	For Each file In folder.Files
		If file.Size = 0 Then
			brs = brs + 1
			ReDim Preserve cell(1, brs)
			cell(0, brs) = file.Name
			cell(1, brs) = file.Size
		End If
	Next
	If brs > -1 Then
		
		Maxi = Maxi & " " & UOM
		SubjectEmail = "DongNote/AmyNote: Please check it some file have size 0B and delete."
		TextB = "Please check folder " & SrcFolder & "test\ some file have size 0B and delete " & vbCrLf
		For x = 0 To brs
			mfile = Srcfolder & cell(0, x)
			DstFldr = SrcFolder & "Test\"
			mfsox.CopyFile mfile, DstFldr
			mfsox.DeleteFile mfile, True
			TextB =  TextB & cell(0, x) & " have size " & cell(1, x) & vbCrLf
		Next
		
		SendEmailCDO SubjectEmail, TextB, ""
	End If
	
End Sub



Function SendEmailCDO(EmailSubject,TextBody, AttchFile)
 
    Set objMessage = CreateObject("CDO.Message")
    objMessage.Subject = EmailSubJect
    objMessage.From = "customercare@covalentworks.com"
    objMessage.To = "customercare@covalentworks.com"
    objMessage.CC = "b2bsupport@covalentworks.com"
    objMessage.TextBody = TextBody
   
   
    If attchFile <> "" then
                    objMessage.AddAttachment attchFile
    End If
   
   '==This section provides the configuration information for the remote SMTP server.
    '==Normally you will only change the server name or IP.

    objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    'Name or IP of Remote SMTP Server
    objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.covalentworks.com"

    'Server port (typically 25)
    objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objMessage.Configuration.Fields.Update
    '==End remote SMTP server configuration section==
    objMessage.Send
 
End Function


