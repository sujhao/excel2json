@echo off
title [convert excel to json]
echo press any button to start.
@pause > nul
echo start converting ....
cd %~dp0
node main.js --export
echo convert over!
@pause