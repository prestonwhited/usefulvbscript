' Collect printer list from computer
strComputer = "."
SET objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
SET colPrinters = objWMIService.ExecQuery("SELECT * FROM Win32_Printer")

' Load printer list into Message
Counter = 0
Message = ""
FOR EACH objPrinter IN colPrinters 
	Counter = Counter + 1
	Message = Message & Counter & ") " & objPrinter.Name & vbCrLf
NEXT 

' Show list of printers
Assistant = MsgBox(Message,64,"PRINTER LIST")
