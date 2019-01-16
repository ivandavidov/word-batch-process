@echo off

set JAVA_HOME=C:\Program Files\Java\jdk1.8.0_112
set PATH=%JAVA_HOME%\bin;%PATH%

javac -cp ".;lib/*" WordBatchProcess.java

pause
