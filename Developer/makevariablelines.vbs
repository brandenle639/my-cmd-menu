On error resume next
'Creates the shell
set oShell = Wscript.CreateObject("Wscript.Shell")

'Creates the file system
set filesys = CreateObject("Scripting.FileSystemObject")

'Gets the current directory
thisdir = oShell.CurrentDirectory

'Declaring the variables
dim rline
dim oline
counter = 1

'Writes the lines to the Vars text file
set owrite = filesys.OpenTextFile("Vars.txt",2,True)

'writes a blank line for the different sections of the dosfullscreen vbs filee for the installation file
owrite.writeline("Variable line for customcmd.bat in the installation script")

'writes a blank line
owrite.writeline("")

'Closes the vars the text file
owrite.Close

'Opens the customcmd bat file for reading
set fread = filesys.OpenTextFile(thisdir & "\customcmd.bat",1)

'Loops through the file rad line by line
do while not fread.AtEndOfStream
	'Reads the line
	rline = fread.ReadLine()
	
	'Replaces the quotes, and new lines with key characters
	rline = Replace(rline,"""","^")
	rline = Replace(rline,"'","$")
	oline=oline + rline + "#n#"
	
	'Looks for ":: VBS Break" in the customcmd bat file
	if InStr(rline,":: VBS Break") then
	
		'Writes the lines to the Vars text file
		set owrite = filesys.OpenTextFile("Vars.txt",8,True)
		
		'writes the lines for the different sections of the customcmd batch file for the installation file
		owrite.writeline("sect" + cstr(counter) + "=" + """" + oline + """")
		
		'Closes the vars the text file
		owrite.Close
		
		'Clears the oline variable
		oline = ""
		
		'Increases the counter for each section
		counter = counter + 1
	end if
loop
'Closes the customcmd file
fread.close

'Opens the dosfullscreen vbs file
set fread = filesys.OpenTextFile(thisdir & "\customcmdaddons\dosfullscreen.vbs",1)

'Reads the file
do while not fread.AtEndOfStream
	'Reads the line
	rline = fread.ReadLine()
	
	'Replaces the quotes, and new lines with key characters
	rline = Replace(rline,"""","^")
	rline = Replace(rline,"'","$")
	oline=oline + rline + "#n#"
loop
'Writes the lines to the Vars text file
set owrite = filesys.OpenTextFile("Vars.txt",8,True)

'writes a blank line
owrite.writeline("")

'writes a blank line for the different sections of the dosfullscreen vbs filee for the installation file
owrite.writeline("Variable line for dosfullscreen.vbs in the installation script")

'writes a blank line
owrite.writeline("")

'writes the lines for the different sections of the dosfullscreen vbs filee for the installation file
owrite.writeline("ln=" + """" + oline + """")

'Closes the vars the text file
owrite.Close

'Closes the customcmd file
fread.Close
