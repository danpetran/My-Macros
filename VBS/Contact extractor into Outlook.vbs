Option Explicit

class Contact
	dim Name
	dim Phone
	dim Email
	dim City
	dim Notes
	dim Country
	dim BirthDay
end class


dim objFS, inFILE, fileNAME, msgLength, strLine, i, objOutlook, objContact
dim nCounter, strBirthDay, arrBirthday (10000), strTime

dim Contacts(10000)
Const olContactItem = 2
fileNAME = "Export.txt"

set objFS = CreateObject ("Scripting.FileSystemObject")
if not objFS.FileExists(fileNAME) then
	SendMessage "Fisierul """ & FileName & """nu a fost gasit in directorul curent!", 2
	Wscript.Quit
end if

set inFILE = objFS.OpenTextFile (fileNAME)

nCounter = 0
i = 0

while NOT inFILE.AtEndOfStream
	strLine = inFILE.ReadLine

	if inStr(strLine, "Sent:"&vbTAB) then
		i = i + 1
		arrBirthday(i) = DateValue(remove_date(Right(strLine, len(strLine)-len("Sent:"&vbTAB))))
	end if
	if inStr(strLine, "Enviada:") then
		i = i + 1
		arrBirthday(i) = ConvertDate(Right(strLine, len(strLine)-len("Enviada:")))
	end if

	if inStr(strLine, "Nume: ") then
		nCounter = nCounter + 1
		Set Contacts(nCounter) = new Contact
		Contacts(nCounter).Name =  Right(strLine, len(strLine)-len("Nume: "))
	end if
	if inStr(strLine, "Name: ") then
		nCounter = nCounter + 1
		Set Contacts(nCounter) = new Contact
		Contacts(nCounter).Name =  Right(strLine, len(strLine)-len("Name: "))
	end if
	if inStr(strLine, "Nombre: ") then
		nCounter = nCounter + 1
		Set Contacts(nCounter) = new Contact
		Contacts(nCounter).Name =  Right(strLine, len(strLine)-len("Nombre: "))
	end if
	if (inStr(strLine, "Nome: ")) then
		nCounter = nCounter + 1
		Set Contacts(nCounter) = new Contact
		Contacts(nCounter).Name =  Right(strLine, len(strLine)-len("Nome: "))
	end if

	if inStr(strLine, "E-mail: ") then
		Contacts(nCounter).Email = Right(strLine, len(strLine)-len("E-mail: "))
	end if
	if inStr(strLine, "Email: ") then
		Contacts(nCounter).Email = Right(strLine, len(strLine)-len("E-mail: "))
	end if

	if (inStr(strLine, "Country: ")) then
		Contacts(nCounter).Country = Right(strLine, len(strLine)-len("Country: "))
	end if
	if (inStr(strLine, "Tara: ")) then
		Contacts(nCounter).Country = Right(strLine, len(strLine)-len("Tara: "))
	end if
	if (inStr(strLine, "Pays: ")) then
		Contacts(nCounter).Country = Right(strLine, len(strLine)-len("Pays: "))
	end if

	if (inStr(strLine, "Telefon: ")) then
		Contacts(nCounter).Phone = Right(strLine, len(strLine)-len("Telefon: "))
	end if
	if (inStr(strLine, "Phone: ")) then
		Contacts(nCounter).Phone = Right(strLine, len(strLine)-len("Phone: "))
	end if
	if inStr(strLine, "Telefone: ") then
		Contacts(nCounter).Phone = Right(strLine, len(strLine)-len("Telefone: "))
	end if
	if inStr(strLine, "Telefono: ") then
		Contacts(nCounter).Phone = Right(strLine, len(strLine)-len("Telefono: "))
	end if

	if inStr(strLine, "Oras: ") then
		Contacts(nCounter).City = Right(strLine, len(strLine)-len("Oras: "))
	end if
	if inStr(strLine, "Cidade: ") then
		Contacts(nCounter).City = Right(strLine, len(strLine)-len("Cidade: "))
	end if
	if inStr(strLine, "City: ") then
		Contacts(nCounter).City = Right(strLine, len(strLine)-len("City: "))
	end if
	if inStr(strLine, "Ciudad: ") then
		Contacts(nCounter).City = Right(strLine, len(strLine)-len("Ciudad: "))
	end if

	if inStr(strLine, "Quiere empezar: ") then
		Contacts(nCounter).Notes = Contacts(nCounter).Notes & "Quire empezar: " & Right(strLine, len(strLine)-len("Quiere empezar: ")) & vbCRLF
	end if
	if inStr(strLine, "How To Start: ") then
		Contacts(nCounter).Notes = Contacts(nCounter).Notes & "How to start: " & Right(strLine, len(strLine)-len("How To Start: ")) & vbCRLF
	end if
	if inStr(strLine, "Best Call Time: ") then
		Contacts(nCounter).Notes = Contacts(nCounter).Notes & "Best call time: " & Right(strLine, len(strLine)-len("Best Call Time: ")) & vbCRLF
	end if
	if inStr(strLine, "Interest: ") then
		Contacts(nCounter).Notes = Contacts(nCounter).Notes & "Interest: " & Right(strLine, len(strLine)-len("Interest: ")) & vbCRLF
	end if
	if inStr(strLine, "Vreau sa Încep") then
		Contacts(nCounter).Notes = Contacts(nCounter).Notes & "Cum sa incep: " & Right(strLine, len(strLine)-len("Vreau sa Încep :")) & vbCRLF
	end if

	if nCounter>1 then
		Contacts(nCounter-1).BirthDay = strBirthDay
	end if
wend

If ProcessRunning("outlook") = False then
	SendMessage "Outlook nu ruleaza! Va rugam sa il porniti pentru a adauga contactele!", 1
	Wscript.Quit
else
	Set objOutlook = CreateObject("Outlook.Application")
End If

for i = 1 to nCounter
	Set objContact = objOutlook.CreateItem(olContactItem)

	objContact.FullName = Contacts(i).Name
	objContact.Email1Address = Contacts(i).Email
	objContact.HomeTelephoneNumber = Contacts(i).Phone
	objContact.HomeAddress = Contacts(i).City
	objContact.Body = Contacts(i).Notes
	Contacts(i).BirthDay = Trim(arrBirthday(i))
	objContact.BirthDay = Contacts(i).Birthday

	objContact.Save
next

SendMessage "Am adaugat " & nCounter & " contacte. Va rugam verificati...", 0

Function ProcessRunning(sProcess)	'Check for a process running
'--------------------------------------------------------------------------------------------------------------
	Dim Process, strComputer, objWMIService, colProcesses

	ProcessRunning = False
	strComputer = "."

	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcesses = objWMIService.ExecQuery ("SELECT * FROM Win32_Process")

	For each process in colprocesses
		if instr(lcase (process.name), sProcess) then
			ProcessRunning = True
		end if
	next
End Function

Sub SendMessage (msgText, intStatus)
	dim wshShell
	set wshShell = CreateObject("wscript.shell")

	msgText = String(len(msgText), "_") & vbCRLF & vbCRLF & msgText & vbCRLF & vbCRLF &  String(len(msgText), "¯")
	if intStatus = 0 then
		wshShell.Popup msgText, , "Informatie", vbInformation
	end if
	if intStatus = 1 then 
		wshShell.Popup msgText, ,"Avertisment!", vbExclamation
	end if
	if intStatus = 2 then
		wshShell.Popup msgText, , "Eroare!", vbCritical
	end if
End Sub

Function convertDate(strDate)
	dim wrongFormat, c, temp
	dim sDate, strMonth, strYear

	wrongFormat = True

	if inStr(LCase(strDate), "segunda-feira,") then
		strDate = trim(Right(strDate, len(strDate)-len("segunda-feira, ")))
		wrongFormat = False
	end if
	if inStr(LCase(strDate), "terça-feira,") then
		strDate = trim(Right(strDate, len(strDate)-len("terça-feira, ")))
		wrongFormat = False
	end if
	if inStr(LCase(strDate), "quarta-feira,") then
		strDate = trim(Right(strDate, len(strDate)-len("quarta-feira, ")))
		wrongFormat = False
	end if
	if inStr(LCase(strDate), "quinta-feira,") then
		strDate = trim(Right(strDate, len(strDate)-len("quinta-feira, ")))
		wrongFormat = False
	end if
	if inStr(LCase(strDate), "sexta-feira,") then
		strDate = trim(Right(strDate, len(strDate)-len("sexta-feira, ")))
		wrongFormat = False
	end if
	if inStr(LCase(strDate), "sábado,") then
		strDate = trim(Right(strDate, len(strDate)-len("sábado, ")))
		wrongFormat = False
	end if
	if inStr(LCase(strDate), "domingo,") then
		strDate = trim(Right(strDate, len(strDate)-len("domingo, ")))
		wrongFormat = False
	end if

	if wrongFormat = True then
		wscript.echo "Wrong format detected!"
	end if

	c = 0
	while isNumeric(c)
		c = left(strDate, 1)
		strDate = Right(strDate, len(strDate)-1)
		sDate = sDate & c
	wend
	
	sDate = cInt(trim(sDate))

	c = ""
	while NOT c = " "
		c = left(strDate, 1)
		strDate = Right(strDate, len(strDate) - 1)
		temp = temp & c
	wend

	c = ""
	strMonth = ""

	while NOT c = " "
		c = left(strDate, 1)
		strDate = Right(strDate, len(strDate) - 1)
		strMonth = strMonth & c
	wend

	strMonth = Trim(strMonth)

	if strMonth = "Janeiro" then
		strMonth = Month(Year(date()) & "-01-"& day(date()))
	end if
	if strMonth = "Fevereiro" then
		strMonth = Month(Year(date()) & "-02-"& day(date()))
	end if
	if strMonth = "Março" then
		strMonth = Month(Year(date()) & "-03-"& day(date()))
	end if
	if strMonth = "Abril" then
		strMonth = Month(Year(date()) & "-04-"& day(date()))
	end if
	if strMonth = "Maio" then
		strMonth = Month(Year(date()) & "-05-"& day(date()))
	end if
	if strMonth = "Junho" then
		strMonth = Month(Year(date()) & "-06-"& day(date()))
	end if
	if strMonth = "Julho" then
		strMonth = Month(Year(date()) & "-07-"& day(date()))
	end if
	if strMonth = "Agosto" then
		strMonth = Month(Year(date()) & "-08-"& day(date()))
	end if
	if strMonth = "Setembro" then
		strMonth = Month(Year(date()) & "-09-"& day(date()))
	end if
	if strMonth = "Outubro" then
		strMonth = Month(Year(date()) & "-10-"& day(date()))
	end if
	if strMonth = "Novembro" then
		strMonth = Month(Year(date()) & "-11-"& day(date()))
	end if
	if strMonth = "Dezembro" then
		strMonth = Month(Year(date()) & "-12-"& day(date()))
	end if

	c = ""
	while NOT c = " "
		c = left(strDate, 1)
		strDate = right(strDate, len(strDate) - 1)
	wend

	strDate = trim(strDate)

	strYear = cInt(left(strDate, 4))

	strDate = trim(right(strDate, len(strDate)-4))
	strTime = TimeValue(strDate)

	convertDate = DateSerial(strYear, strMonth, sDate)
End Function

function remove_date(strDate)
	dim wrongFormat

	wrongFormat = TRUE

	if inStr(lcase(strDate), "monday") > 0 then
		remove_date = trim(right(strDate, len(strDate) - len("monday, ")))
		wrongFormat = FALSE
	end if
	if inStr(lcase(strDate), "tuesday") > 0 then
		remove_date = trim(right(strDate, len(strDate) - len("tuesday, ")))
		wrongFormat = FALSE
	end if
	if inStr(lcase(strDate), "wednesday") > 0 then
		remove_date = trim(right(strDate, len(strDate) - len("wednesday, ")))
		wrongFormat = FALSE
	end if
	if inStr(lcase(strDate), "thursday") > 0 then
		remove_date = trim(right(strDate, len(strDate) - len("thursday, ")))
		wrongFormat = FALSE
	end if
	if inStr(lcase(strDate), "friday") > 0 then
		remove_date = trim(right(strDate, len(strDate) - len("friday, ")))
		wrongFormat = FALSE
	end if
	if inStr(lcase(strDate), "saturday") > 0 then
		remove_date = trim(right(strDate, len(strDate) - len("saturday, ")))
		wrongFormat = FALSE
	end if
	if inStr(lcase(strDate), "sunday") > 0 then
		remove_date = trim(right(strDate, len(strDate) - len("sunday, ")))
		wrongFormat = FALSE
	end if

	if wrongFormat = TRUE then
		wscript.echo "Wrong format detected in English Date"
	end if
end function