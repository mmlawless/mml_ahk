; ------------------------------------------------------------
; Teams Presence Monitor via ImageSearch (highlighted + normal)
; AutoHotkey v1.x
; ------------------------------------------------------------

#NoEnv
#Persistent
SetBatchLines, -1
SetTitleMatchMode, 2
SendMode, Input

; ========== CONFIG: ICON PATHS ==========

; Folder containing your BMP icons
; e.g. C:\AHK\teams-icons\
iconDir := "C:\AHK\teams-icons\"   ; <-- CHANGE THIS

; For each status, we store [normal, highlighted] images.
; Make sure these files exist in iconDir.
icons := {}
icons["Available"] := [iconDir . "available.bmp",     iconDir . "available_sel.bmp"]
icons["Away"]      := [iconDir . "away.bmp",          iconDir . "away_sel.bmp"]
icons["Busy"]      := [iconDir . "busy.bmp",          iconDir . "busy_sel.bmp"]
icons["DND"]       := [iconDir . "dnd.bmp",           iconDir . "dnd_sel.bmp"]
; Optional: if you actually capture an explicit "unknown" dot
; icons["Unknown"]   := [iconDir . "unknown.bmp",       iconDir . "unknown_sel.bmp"]

; ========== CONFIG: CONTACT SEARCH REGIONS ==========

; These coordinates are OFFSETS from the top-left of the Teams window.
; Use Window Spy to find the dot position for each person, then create
; a small box around it (e.g. ±10 px).
;
; Example only – YOU MUST CHANGE these numbers for your setup.

contacts := []

contacts.push({name:"Alexandra", x1:20, y1:100, x2:40, y2:120})
contacts.push({name:"Clark",     x1:20, y1:130, x2:40, y2:150})
contacts.push({name:"Ellis",     x1:20, y1:160, x2:40, y2:180})
contacts.push({name:"Lorees",    x1:20, y1:190, x2:40, y2:210})
contacts.push({name:"Louaye",    x1:20, y1:220, x2:40, y2:240})
contacts.push({name:"Tom",       x1:20, y1:250, x2:40, y2:270})
contacts.push({name:"Paul",      x1:20, y1:280, x2:40, y2:300})
contacts.push({name:"Peter",     x1:20, y1:310, x2:40, y2:330})

; ========== OPTIONAL: MQTT (comment out if not needed) ==========

; mqttExe  := "mosquitto_pub"         ; must be in PATH or use full path
; mqttHost := "192.168.1.15"      ; e.g. 192.168.1.10
; mqttBase := "mymttteam"                 ; topic prefix: teams/<Name>/presence

; ========== GENERAL SETTINGS ==========

CheckIntervalMs := 5000   ; 5000 ms = 5 seconds
TeamsExeName    := "Teams.exe"   ; for classic Teams
; For new Teams you might need "ms-teams.exe" instead

; Hotkey to manually toggle always-on-top for active window (optional)
^!t::
    WinGet, active_id, ID, A
    WinSet, AlwaysOnTop, Toggle, ahk_id %active_id%
return

; Start periodic presence checks
SetTimer, CheckPresence, %CheckIntervalMs%
return


; ========== MAIN TIMER ROUTINE ==========

CheckPresence:
    ; Look for the Teams window
    if !WinExist("ahk_exe " . TeamsExeName)
    {
        ToolTip, Teams window not found (ahk_exe %TeamsExeName%).
        return
    }

    ; Keep Teams always on top so the list is visible
    WinSet, AlwaysOnTop, On, ahk_exe %TeamsExeName%

    ; Get Teams window position
    WinGetPos, winX, winY, winW, winH, ahk_exe %TeamsExeName%

    statusSummary := ""
    for i, contact in contacts
    {
        status := DetectStatus(contact, winX, winY)
        statusSummary .= contact.name ": " status "`n"

        ; OPTIONAL MQTT publish:
        ; PublishMQTT(contact.name, status)
    }

    ToolTip, %statusSummary%
return


; ========== STATUS DETECTION FUNCTION ==========

DetectStatus(contact, winX, winY)
{
    global icons

    ; Convert contact-relative offsets to screen coordinates
    x1 := winX + contact.x1
    y1 := winY + contact.y1
    x2 := winX + contact.x2
    y2 := winY + contact.y2

    ; Loop through each status and its icon variants
    for status, iconArray in icons
    {
        for i, iconPath in iconArray
        {
            if !FileExist(iconPath)
                continue

            ; *20 = tolerance for minor variations, adjust if needed
            ImageSearch, fx, fy, x1, y1, x2, y2, *20 %iconPath%
            if (ErrorLevel = 0)
                return status
        }
    }

    return "Unknown"
}


; ========== MQTT PUBLISH FUNCTION (OPTIONAL) ==========

PublishMQTT(name, status)
{
    global mqttExe, mqttHost, mqttBase

    if (!mqttExe || !mqttHost || !mqttBase)
        return  ; MQTT not configured

    if (status = "")
        status := "Unknown"

    topic := mqttBase "/" name "/presence"
    cmd   := mqttExe . " -h " . mqttHost . " -t """ . topic . """ -m """ . status . """ -r"

    RunWait, %ComSpec% /c "%cmd%",, Hide
}
