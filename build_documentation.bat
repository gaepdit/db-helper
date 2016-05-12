@echo off
echo -
echo    IMPORTANT: This only works if you have downloaded and built 
echo    "Docu" from https://github.com/jagregory/docu !
echo -
echo    Also, you must build a release version of this project first.
echo _

rmdir /S /Q documentation

..\docu\src\Docu.Console\bin\Release\docu EpdItDbHelper\bin\Release\EpdItDbHelper.dll --output=documentation

echo -
pause