@echo on

set JAVA_HOME=.\jdk\bin
set PATH=%JAVA_HOME%\bin;%PATH%

echo.
echo   Working...
echo.

rmdir /s /q result

java -cp ".;lib/*" prog.WordBatchProcess

echo.
echo   Done. Check the generated files in folder "result". Have a nice day! :)
echo.

pause
