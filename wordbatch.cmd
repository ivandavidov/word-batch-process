@echo off

set JAVA_HOME=jre\bin
set PATH=%JAVA_HOME%\bin;%PATH%

jre\bin\java -cp ".;lib/*" WordBatchProcess

pause
