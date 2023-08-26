@echo off

set /p "newComputerName=Enter the new computer name: "

wmic computersystem where name="%computername%" rename "%newComputerName%"