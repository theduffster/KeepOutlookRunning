@echo off
echo ___________________________________________________________
echo Welcome to the KeepOutlookRunning installer batch
echo ___________________________________________________________
echo This batch will install KeepOutlookRunning for the active user only
echo No system files will be touched
echo ___________________________________________________________
echo IMPORTANT: Do not run this install batch from the zip file
echo            directly. It won't work. Please extract the zip 
echo            file contents to a folder and run it from there.
echo            Must CD to the unzipped folder for this to work.
echo   
echo ___________________________________________________________
:Ask
echo Would you like install KeepOutlookRunning?(Y/N)
set INPUT=
set /P INPUT=Type input: %=%
If /I "%INPUT%"=="y" goto yes 
If /I "%INPUT%"=="n" goto no 
goto Ask
:yes
set "sourcepath=%cd%"
set installfolder="KeepOutlookRunning"
set installpath="%appdata%\KeepOutlookRunning"

cd %appdata%
if not exist "%installfolder%" mkdir %installfolder%

cd %sourcepath%

robocopy "KeepOutlookRunning" %installpath% /E

powershell "(Get-Content KeepOutlookRunning.reg) | foreach-object {($_ -replace \"FOLDER_APP_DATA\", (\"%appdata%\").replace(\"\\\",\"\\\\\"))} | foreach-object {($_ -replace \"FILEPATH_APP_DATA\", (\"%appdata%\").replace(\"\\\",\"/\"))} | Set-Content KeepOutlookRunningLocalized.reg"


reg import KeepOutlookRunningLocalized.reg


echo KeepOutlookRunning was installed
pause
exit
:no
echo KeepOutlookRunning was not installed
pause