@echo off
echo -
echo    IMPORTANT: This only works if you have downloaded and built 
echo    "Docu" from https://github.com/jagregory/docu !
echo -
echo    Also, you must build a release version of this project first.
echo _
echo Deleting old documentation...

rmdir /S /Q documentation

echo _
echo Done
echo _

..\docu\src\Docu.Console\bin\Release\docu EpdItDbHelper\bin\Release\EpdIt.DbHelper.dll --output=documentation

echo _
echo Done
echo -
pause