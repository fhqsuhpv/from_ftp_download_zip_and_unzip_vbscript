strFileURL = "http://<yourpath>/[youzipname].zip"
strWinRARFileURL = "http://http://<yourpath>/[youzipname].zip"

Function DownloadTo(url, dest)
	' Fetch the file
	Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

	objXMLHTTP.open "GET", url, false
	objXMLHTTP.send()

	' Save the file
	If objXMLHTTP.Status = 200 Then
		Set objADOStream = CreateObject("ADODB.Stream")
		objADOStream.Open
		objADOStream.Type = 1

		objADOStream.Write objXMLHTTP.ResponseBody
		objADOStream.Position = 0

		objADOStream.SaveToFile dest
		objADOStream.Close
		Set objADOStream = Nothing
	End If

	Set objXMLHTTP = Nothing

	' Download complete
End Function

Function ExtractTo(file, dest)
	Set objShell = CreateObject("Shell.Application")
	Set objSource = objShell.NameSpace(file).Items()

	Set objTarget = objShell.NameSpace(dest)
	objTarget.CopyHere objSource, 256
End Function

' Set your settings
strLocalFile = "nwjs.zip"


Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(strLocalFile) Then
	'WScript.Echo "nwjs.zip exists, quiting"
	WScript.Quit
End If

'WScript.Echo "Downloading " & strFileURL
DownloadTo strFileURL, strLocalFile

sCurPath = fso.GetAbsolutePathName(".")
strZipFile = sCurPath & "\" & strLocalFile
outFolder = sCurPath & "\nwjs"

If fso.FolderExists(outFolder) Then
	fso.DeleteFolder(outFolder)
End If

fso.CreateFolder(outFolder)

'WScript.Echo "Extracting items to " & outFolder

ExtractTo strZipFile, outFolder


' Move files 
Set folder = fso.GetFolder(outFolder)

If folder.Files.Count = 0 AND folder.SubFolders.Count = 1 Then
	For Each item In folder.SubFolders
		subFolder = Item
	Next
	fso.MoveFile subFolder & "/*.*", outFolder
	fso.MoveFolder subFolder & "/*.*", outFolder
	fso.DeleteFolder(subFolder)
end if


strWinRARLocalFile = "winRAR.zip"
strWinRARZipFile = sCurPath & "\" & strWinRARLocalFile

If fso.FileExists(strWinRARLocalFile) Then
	'WScript.Echo "winRAR.zip exists, quiting"
	WScript.Quit
End If

'WScript.Echo "Downloading " & strWinRARFileURL
DownloadTo strWinRARFileURL, strWinRARLocalFile

ExtractTo strWinRARZipFile, sCurPath

'WScript.Echo("download and unzip complete")