@echo off
setlocal enabledelayedexpansion

:: Ask for the file to hide
set /p filePath="Enter the full path of the file to hide: "

:: Ask for the .docx file
set /p docxPath="Enter the full path of the .docx file: "

:: Check if the file exists
if not exist "!filePath!" (
    echo The file "!filePath!" does not exist.
    pause
    exit /b
)

:: Check if the .docx file exists
if not exist "!docxPath!" (
    echo The .docx file "!docxPath!" does not exist.
    pause
    exit /b
)

:: Rename .docx to .zip
set "zipPath=!docxPath:.docx=.zip!"
rename "!docxPath!" "*.zip"

:: Extract the contents of the .zip file
mkdir tempDocx
cd tempDocx
powershell -command "Expand-Archive -Path ..\!zipPath! -DestinationPath ."

:: Copy the file to the extracted directory
copy /y "!filePath!" .\

:: Repack the files into a new .zip file
powershell -command "Compress-Archive -Path .\* -DestinationPath ..\new.zip"

:: Rename the new .zip back to .docx
cd ..
del /q "!zipPath!"
rename new.zip "!docxPath!"

:: Clean up temporary files
rd /s /q tempDocx

echo File successfully hidden inside the .docx.
pause
