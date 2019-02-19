@echo on

set JAVA_HOME=.\jdk
set PATH=%JAVA_HOME%\bin;%PATH%

javac -cp "lib/*" prog\WordBatchProcess.java

pause
