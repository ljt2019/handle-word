@echo off

@echo start build handle-word

call mvn clean install -Dskiptest=true

pause