#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Persistent	; Script must continue to run to work


^h::
run "c:\Program Files\mosquitto\mosquitto_pub.exe" -h 192.168.1.15 -t MMLNR/RTUpdate -m 1
return
