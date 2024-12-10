cls
@echo off & SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
rem Declares echo output and expansions

rem sets the title of the batch
TITLE Custom CMD

rem Sets the color scheme of the terminal
color 71

rem The start function is used when the script loads
:start
	rem Gets the default web browser from the registry
	for /f "tokens=2*" %%a in ('reg QUERY HKEY_CLASSES_ROOT\http\shell\open\command /ve') do ( set ibrow=%%b )
	
	rem Sets the path for full screen script from the customcmdaddons folder
	set fullscreen=%appdata%customcmdaddons\dosfullscreen.vbs
	
	rem Checks to see if the full screen script exists and if it does then it run it
	if EXIST "%fullscreen%" (
		start "" "%fullscreen%"
	)
	
	rem Changes to the users directory
	cd C:\Users\%USERNAME%\
	
	rem Goes to the listcommands to print out the menu of commands
	goto :listincomms
	
rem The beginning of the script for user input
:begin
	
	rem Prints a blank line
	echo.
	
	rem Sets the variables to appropriate pre values
	set incomm=
	set runthis=
	set skip=
	set count=0
	
	rem Asks for user input
	set /P incomm=Enter incomm:
	
	rem Gets the first character from the input
	set qchar=%incomm:~0,1%
	
	rem Checks to see if the qchar variable is a quote and is so it goes to rcomm function to run command
	if "%qchar%%qchar%" == """" (
		set runthis=%incomm%
		goto :rcomm
	)
	
	rem Looks for quotes in string and if they exist it goes to hasarguments function
	echo %incomm% | findstr /C:"""" 1>nul
	if errorlevel 1 (
		echo.
	) else (
		set runthis=%incomm%
		goto :hasarguments
	)
	
	rem If input is empty it loops back to the beginning
	if "%incomm%"=="" goto :begin
	
	rem Calls the lower case function to set the input to lowercase
	call :tolower incomm
	
	rem general commands built into the script
	if "%incomm%"=="reload" set runthis=exit & start "" "%~f0"
	if "%incomm%"=="home" set runthis=cd C:\Users\%USERNAME%\
	if "%incomm%"=="gdate" set runthis=echo %date% %time%
	if "%incomm%"=="listcommands" goto :listincomms
	if "%incomm%"=="web" set runthis=%ibrow%
	if "%incomm%"=="exit" exit
	
	rem custom commands
	:: VBS Break

	rem programs
	if "%incomm%"=="notes" set runthis="C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.EXE"
	if "%incomm%"=="snippet" set runthis="C:\WINDOWS\system32\SnippingTool.exe"
	if "%incomm%"=="word" set runthis="C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"
	if "%incomm%"=="excel" set runthis="C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"
	:: VBS Break

	rem websites
	:: VBS Break
	
	rem Gets the first character from the runthis variable
	set qchar=%runthis:~0,1%
	
	rem Checks to see if the qchar variable is a quote and is so it goes to rcomm function to run command
	if "%qchar%%qchar%" == """" (
		set runthis=%runthis%
		goto :rcomm
	)
	
	rem Looks for quotes in string and if they exist it goes to hasarguments function
	echo %runthis% | findstr /C:"""" 1>nul
	if errorlevel 1 (
		echo.
	) else (
		set runthis=%runthis%
		goto :hasarguments
	)
	
	rem If runthis value is empty it just gets the command from the user input
	if "%runthis%" == "" set runthis=%incomm%
	goto :hasarguments
	
	rem If input is empty it loops back to the beginning
	if "%incomm%"=="" goto :begin
	
	rem Goes to hasarguments function
	goto :hasarguments
	
rem Arguments functions
:hasarguments
	
	rem Gets the first part of the input
	for /F "tokens=1" %%a in ("%runthis%") do set b=%%a
	
	rem Gets the remainder part of the input
	call set args=%%runthis:%b%=%%
	
	rem Checks the first part of the input and calls the proper function
	if "%b%" == "word" (
		set runthis="C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE" %args%
	)
	if "%b%" == "excel" (
		set runthis="C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE" %args%
	)
	if "%b%" == "notes" (
		set runthis="C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.EXE" %args%
	)
	if "%b%" == "web" (
		set runthis=%ibrow% %args%
	)
	
	rem Goes to run the command function
	goto :rcomm
	
rem Run commands function
:rcomm
	
	rem Sets the counter to negative 1
	set count=-1
	
	rem Counts the number of periods
	for %%a in (%runthis:.= %) do set /a count+=1
	
	rem Looks for a period in string
	echo %runthis% | findstr /C:"." 1>nul
	
	rem If error level is 1 then it runs the command internally
	rem If error level is not 1 then it looks to see if count is 4 then it runs the command internally
	rem If count is not 4 then it runs the command externally
	if errorlevel 1 (
	 	%runthis%
	) else (
		if %count% == 4 (
			%runthis%
		) else (
			start "" %runthis%
		)
	)
	
	rem Loops back to the beginning
	goto :begin

rem Function to convert input to lowercase	
:tolower
for %%L in (a b c d e f g h i j k l m n o p q r s t u v w x y z) DO SET %1=!%1:%%L=%%L!
goto :EOF

rem The menu function
:listincomms
	echo ---------------
	echo General Commands
	echo ----------------
	echo 	exit - Exits this script
	echo 	listincomms - Shows this list of Commands
	echo 	reload - Reloads the script
	echo 	web - Loads a web page
	echo ----------------
	echo Custom Commands
	echo ---------------
:: VBS Break
	echo ---------------
	echo Websites
	echo ---------------
:: VBS Break
	echo ---------------
	echo Programs
	echo ---------------
	echo 	excel - Opens Microsoft Excel
	echo 	notes - Opens Microsoft Access
	echo 	snippet - Opens the Snipping Tool
	echo 	word - Opens Microsoft Word
:: VBS Break
	echo ---------------
	goto :begin
:: VBS Break
