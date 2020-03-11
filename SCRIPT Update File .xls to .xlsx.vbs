' Constant variables used by Excel methods
CONST xlOpenXMLWorkbook = 51
CONST xlLocalSessionChanges = 2

' Subroutine that gets called after user cancels the script
SUB EndScript
	Message = "Script cancelled."
	Assistant = MsgBox(Message,vbOKOnly,"END")
	WScript.Quit
END SUB

' Message telling user what to check before proceeding
Message = "*** UPDATE FILE .XLS TO .XLSX ***" _
& vbCrLf & "" _
& vbCrLf & "This script will update every .xls file in the current folder to a .xlsx file." _
& vbCrLf & "" _
& vbCrLf & "Make sure the SCRIPT file is in the same folder as the files you want to update." _
& vbCrLf & "" _
& vbCrLf & "Click OK to continue, or Cancel to stop script."

Assistant = MsgBox(Message,vbOKCancel+vbInformation,"ATTENTION")

' Check if user cancels the script at opening message
IF Assistant = vbCancel THEN
	EndScript
END IF

Message = ""

' Create file system object needed to navigate files, set variables to orient script to its location
SET objFSO = CreateObject("Scripting.FileSystemObject")
scriptPath = WScript.ScriptFullName
scriptFolder = objFSO.GetParentFolderName(scriptPath)
scriptLog = scriptFolder & "\XLS to XLSX Log.txt"
SET objFolder = objFSO.GetFolder(scriptFolder)
scriptLegacyArchive = scriptFolder & "\LegacyArchive"

' Check for LegacyArchive folder, create if it doesn't exist
IF objFSO.FolderExists(scriptLegacyArchive) = FALSE THEN 
	objFSO.CreateFolder(scriptLegacyArchive)
END IF 

' Create Excel object, make Excel visible
SET objExcel = CreateObject("Excel.Application")
objExcel.Visible = TRUE
objExcel.DisplayAlerts = FALSE

' Cycle through all files in the current folder, save as .xlsx file, move .xls file to LegacyArchive folder, update message
FOR EACH Fil IN objFolder.Files 
	IF RIGHT(Fil.Name,3) = "xls" THEN 
		SET objBOOK = objExcel.Workbooks.Open(scriptFolder & "\" & Fil.Name)
		objBOOK.SaveAs scriptFolder & "\" & Fil.Name & "x",xlOpenXMLWorkbook,,,,,,xlLocalSessionChanges
		Message = Message & Fil.Name & " changed to " & objBOOK.Name & vbCrLf
		objBOOK.Close
		objFSO.MoveFile scriptFolder & "\" & Fil.Name, scriptFolder & "\LegacyArchive\" & YEAR(NOW()) & " " & RIGHT("00" & MONTH(NOW()),2) & "-" & RIGHT("00" & DAY(NOW()),2) & " " & Fil.Name
	END IF 
NEXT 

' Re-enable display alerts, quit excel
objExcel.DisplayAlerts = TRUE
objExcel.Quit

' Write message to log file
SET WriteLog = objFSO.OpenTextFile(scriptLog,8,TRUE)
WriteLog.WriteLine NOW()
WriteLog.Write Message
WriteLog.Close

' Display success message
Assistant = MsgBox("Script finished! Check " & scriptLog & " for changes.",vbOKOnly,"SUCCESS")
