@echo off

set JAVA_HOME=jre\bin
set PATH=%JAVA_HOME%\bin;%PATH%

echo.
echo   Working...
echo.

jre\bin\java -cp ".;lib/*" prog.WordBatchProcess

echo.
echo   Done. Check the generated files in folder "result". Have a nice day! :)
echo.

pause
