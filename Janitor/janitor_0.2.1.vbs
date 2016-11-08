''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DESCRIPTION:	This script will clean up any files that have a last modified 
'		date greater than or equal to the number specified. 
'CREATED BY:	Brian Plexico - microISV.com
'MODIFIED BY: Gina Trapani (ginatrapani@gmail.com)
'CREATE DATE:	07/19/2005
'UPDATED:	08/10/2007
'
'INSTRUCTIONS:	1.  Enter the path to the directory you'd like to clean on line 28.
'			Make sure to only change the text that currently reads:
'			"C:\DIRECTORY_TO_CLEAN".  The path must be surrounded by quotes.
'		2.  Enter the number of days that must have passed for a file to be deleted
'			on line 38.  The default is 30 but can be changed to any number.
'		3.  If you want to retain some of the subfolders so that they aren't deleted,
'			you have two choices.  See A & B below.
'			A.  Add the folder names you want to skip to the foldersToSkip line.
'			    Separate foldernames with a semicolon.
'				Example:  foldersToSkip = "Folder1;Folder2"
'			B.  Create a file named 'janitor.skip' in the folder(s) you want to 
'				leave alone.
'
'
'LEGAL MUMBO JUMBO:  	-THIS SCRIPT HAS THE POTENTIAL TO BE VERY DAMAGING TO YOUR COMPUTER.  
'			-USE THIS SCRIPT AT YOUR OWN RISK.
'			-I'M NOT RESPONSIBLE FOR ANY BAD EFFECTS OF THIS SCRIPT SUCH AS DIZZINESS,
'			 NOSE BLEEDS, OR A COMPLETELY DESTROYED OPERATING SYSTEM.  
'			-IF YOU HAVE DOUBTS AS TO WHAT YOU'RE DOING, THEN DON'T DO IT.  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DO NOT EDIT THE SECTION BELOW '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
On Error Resume Next

Dim fso, PathToClean, numberOfDays, folder, rootFolder, objFolder, objSubfolders, objFiles, folderToClean, folderToCheck, fts, foldersToSkip, skippedfolders
Set fso = CreateObject("Scripting.FileSystemObject")
Set fts = CreateObject("Scripting.Dictionary")
'DO NOT EDIT THE SECTION ABOVE ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ENTER THE PATH THAT CONTAINS THE FILES YOU WANT TO CLEAN UP
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Path to the root directory that you're cleaning up
PathToClean = "\\SERVER\Videos\Television"
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ENTER THE NUMBER OF DAYS SINCE THE FILE WAS LAST MODIFIED
'
'ANY FILE WITH A DATE LAST MODIFIED THAT IS GREATER OR EQUAL TO 
'THIS NUMBER WILL BE DELETED.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Specify the how many days old a file must be in order to be deleted.
numberOfDays = 90
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ENTER THE NAMES OF THE FOLDERS YOU DO NOT WANT TO BE DELETED
'
'ALL FOLDERS WITH THE SPECIFIED NAME WILL NOT BE DELETED.
'ALL FILES IN THE SPECIFIED FOLDERS WILL NOT BE DELETED.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' INSTRUCTIONS TO ADD FOLDERS YOU WISH TO SKIP
'Add the names of the folders you wish to skip.
'Separate each name with a semicolon.
'If you don't wish to use this feature, the line below must look like this:  foldersToSkip = ""
'The line must look like the following to skip folders:  foldersToSkip = "Folder1;Folder2"

foldersToSkip = ""

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DON'T CHANGE ANYTHING BELOW THIS LINE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Check to make sure path is not a drive root
If Right(PathToClean, 2) = ":\" or Right(PathToClean, 1) = ":" Then
	msgbox "Whoa Nelly!  It's best not to run the Janitor on a drive root like " + PathToClean, vbOkOnly, "Don't Do That!"
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Start at the folder specified and walk down the directory tree
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set rootFolder = fso.GetFolder(PathToClean)
If Err.Number > 0 Then
	msgbox PathToClean + "is not a valid directory path.  Please correct the path and run the script again.", vbOkOnly, "Path Not Found"
	Wscript.Quit
End If

EnumerateFoldersToSkip(foldersToSkip)
GetSubfolders(rootFolder)
CleanupFiles(rootFolder)

'Let person know when the cleanup is complete
MsgBox "Files older than " + CStr(numberOfDays) + " days have been deleted from " + PathToClean, vbOkOnly, "Windows Janitor Cleanup Complete"

'Clean up
Set fso = Nothing

Wscript.Quit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub EnumerateFoldersToSkip(skippedfolders)
	If skippedfolders <> "" Then
		If InStr(1, skippedfolders, ";", 1) <> 0 Then
			Dim arrSkippedFolders, sf
			arrSkippedFolders = Split(skippedfolders, ";")
			For each sf In arrSkippedFolders
				fts.Add UCase(Trim(sf)), ""
			Next
		Else
			fts.Add UCase(skippedfolders), ""
		End If
	End If
End Sub

Sub GetSubfolders(folder)
	If CheckForSkip(folder) Then Exit Sub

	Dim oSubfolder
	Set objFolder = fso.GetFolder(folder)
	Set objSubfolders = objFolder.Subfolders

	For Each oSubfolder in objSubfolders
		If fts.Exists(UCase(oSubfolder.Name)) = False Then
			'Recursively go down the directory tree
			GetSubfolders(oSubfolder.Path)
	
			'Cleanup any files that meet the criteria
			CleanupFiles(oSubfolder.Path)
	
			'Delete the folder if its empty
			CleanupFolder(oSubfolder.Path)	
		End If
	Next
End Sub

Sub CleanupFiles(folderToClean)
	If CheckForSkip(folderToClean) Then Exit Sub

	dim objFile
	Set objFolder = fso.GetFolder(folderToClean)
	Set objFiles = objFolder.Files

	For Each objFile in objFiles
		If DateDiff("d", objFile.DateLastModified, Now) > numberOfDays Then
			objFile.Delete
		End If
	Next

	Set objFolder = Nothing
	Set objFiles = Nothing
End Sub

Sub CleanupFolder(folderToCheck)
	If CheckForSkip(folderToCheck) Then Exit Sub

	Set objFolder = fso.GetFolder(folderToCheck)
	Set objSubfolders = objFolder.Subfolders
	Set objFiles = objFolder.Files

	If objFiles.Count = 0 and objSubfolders.Count = 0 Then
		objFolder.Delete
	End If

	Set objFolder = Nothing
	Set objSubfolders = Nothing
	Set objFiles = Nothing
End Sub

Function CheckForSkip(folderToCheck)
	CheckForSkip = fso.FileExists(folderToCheck & "\JANITOR.SKIP")
End Function

