#NoEnv
#SingleInstance Force
SendMode Input
SetWorkingDir %A_ScriptDir%
#Persistent
FileEncoding, UTF-8-RAW

; -------- Config --------
global host    := "192.168.1.15"
global port    := 1883
global topic   := "MMLNR/MailStatusIN2"
global exe     := "C:\Program Files\mosquitto\mosquitto_pub.exe"  ; path that works for you
global logFile := A_ScriptDir "\mqtt_ahk.log"

; Dedupe window (milliseconds). If the same EntryID is seen again within this window, skip it.
global SEEN_TTL_MS := 15000

; ---- Logging ----
Log(msg) {
    global logFile
    FileAppend, %A_Now%`t%msg%`n, %logFile%
}

; ---- MQTT publish helpers (DIRECT RunWait, no cmd.exe) ----
; Simple text payloads (safe for strings without quotes, e.g. "Mail", "Clear")
PubSimple(payload) {
    global exe, host, port, topic
    q := Chr(34)
    cmd := q exe q " -d -h " host " -p " port " -t " topic " -m " q payload q
    RunWait, %cmd%,, UseErrorLevel
    Log("PUB(M) RC=" ErrorLevel " CMD=" cmd)
    return ErrorLevel
}

; JSON payloads (via -f file to avoid quoting issues)
PubJSON(json) {
    global exe, host, port, topic
    q := Chr(34)
    tmp := A_Temp "\mqtt_payload_" A_TickCount ".json"
    FileDelete, %tmp%
    FileAppend, %json%, %tmp%
    cmd := q exe q " -d -h " host " -p " port " -t " topic " -f " q tmp q
    RunWait, %cmd%,, UseErrorLevel
    Log("PUB(F) RC=" ErrorLevel " CMD=" cmd " FILE=" tmp)
    FileDelete, %tmp%
    return ErrorLevel
}

; ---- Outlook wiring ----
global olApp
try {
    olApp := ComObjActive("Outlook.Application")
} catch e {
    olApp := ComObjCreate("Outlook.Application")
}
if !IsObject(olApp) {
    MsgBox, 16, Outlook, Could not open Outlook Application COM object.
    ExitApp
}

; Keep references + counters
global cnt_NewMailEx := 0
global cnt_ItemAdd   := 0

ComObjConnect(olApp, "EventApp_")                  ; NewMailEx
global olNs := olApp.GetNamespace("MAPI")
global olInbox := olNs.GetDefaultFolder(6)         ; 6 = olFolderInbox
global inboxItems := olInbox.Items                 ; keep reference alive
ComObjConnect(inboxItems, "EventItems_")           ; ItemAdd fallback

SetTimer, __Tick, 2000

; ---- JSON helpers ----
JSON_Escape(s) {
    q := Chr(34)
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, q, "\" . q)
    s := StrReplace(s, "`r`n", "\n")
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`r", "\n")
    return s
}

; ---- Dedupe store ----
global seen := {}   ; EntryID -> last tick
PurgeSeen() {
    global seen, SEEN_TTL_MS
    now := A_TickCount
    for k, ts in seen
        if (now - ts > SEEN_TTL_MS)
            seen.Delete(k)
}
SeenRecently(id) {
    global seen, SEEN_TTL_MS
    if (id = "")
        return false
    now := A_TickCount
    if (seen.HasKey(id)) {
        if (now - seen[id] <= SEEN_TTL_MS)
            return true
        seen.Delete(id)  ; stale
    }
    return false
}
MarkSeen(id) {
    global seen
    if (id != "")
        seen[id] := A_TickCount
}

PublishMail(mail) {
    if !IsObject(mail)
        return

    ; Only MailItem (olMail = 43)
    try if (mail.Class != 43)
        return

    ; Outlook unique ID for this item
    id := ""
    try id := mail.EntryID

    ; If we've just published this ID, skip duplicates (NewMailEx + ItemAdd)
    PurgeSeen()
    if (SeenRecently(id)) {
        Log("Skip duplicate EntryID=" id)
        return
    }

    ; Optional ignore
    try if (mail.Subject = "XXXX")
        return

    ; Resolve SMTP sender
    from := ""
    try {
        if (mail.SenderEmailType = "EX") {
            exu := mail.Sender.GetExchangeUser()
            if (exu)
                from := exu.PrimarySmtpAddress
        }
    }
    if (!from)
        try from := mail.SenderEmailAddress

    subj := ""
    try subj := mail.Subject

    FormatTime, ts, %A_Now%, yyyy-MM-dd HH:mm:ss
    q := Chr(34)
    json := "{"
        . q "from" q ":" q JSON_Escape(from) q ","
        . q "subj" q ":" q JSON_Escape(subj) q ","
        . q "ts"   q ":" q ts q
        . "}"

    ; --- 1) Publish full JSON details ---
    rcJson := PubJSON(json)
    Log("PublishMail JSON rc=" rcJson " id=" id " from=" from " subj=" subj)

    ; --- 2) Also publish simple 'Mail' string ---
    rcSimple := PubSimple("Mail")
    Log("PublishMail Simple rc=" rcSimple " id=" id)

    ; Dedupe on JSON success (so the same EntryID isn't reprocessed)
    if (rcJson = 0)
        MarkSeen(id)
}



; ---- Robust "get current mail" for Ctrl+U ----
GetCurrentMailItem() {
    global olApp, olInbox
    ; 1) If an Inspector window is open (email popped out), use that
    try {
        insp := olApp.ActiveInspector
        if (IsObject(insp)) {
            item := insp.CurrentItem
            if (IsObject(item)) {
                Log("Ctrl+U: using ActiveInspector.CurrentItem")
                return item
            }
        }
    }
    ; 2) Otherwise use the selection in the Explorer (main Outlook window)
    try {
        expl := olApp.ActiveExplorer
        if (IsObject(expl)) {
            sel := expl.Selection
            if (sel.Count >= 1) {
                Log("Ctrl+U: using ActiveExplorer.Selection(1), count=" sel.Count)
                return sel.Item(1)
            } else {
                Log("Ctrl+U: selection count=0")
            }
        } else {
            Log("Ctrl+U: no ActiveExplorer")
        }
    }
    ; 3) Fall back to most recent Inbox item (so the hotkey always does something)
    try {
        items := olInbox.Items
        items.Sort("[ReceivedTime]", false) ; descending
        item := items.GetFirst()
        if (IsObject(item)) {
            Log("Ctrl+U: fallback to newest Inbox item")
            return item
        }
    }
    Log("Ctrl+U: no mail item found")
    return ""
}

; ---- Outlook events ----
EventApp_NewMailEx(IDs) {
    global cnt_NewMailEx, olApp
    cnt_NewMailEx++
    for _, ID in StrSplit(IDs, ",") {
        mail := ""
        try mail := olApp.Session.GetItemFromID(ID)
        PublishMail(mail)
    }
}
EventItems_ItemAdd(item) {
    global cnt_ItemAdd
    cnt_ItemAdd++
    PublishMail(item)
}

__Tick:
    ; REMOVED: TrayTip status display - script runs silently in background
    ; All status info is still logged to mqtt_ahk.log file
return

; ---- Hotkeys ----
^j::  ; Clear
    PubSimple("Clear")
return

^i::  ; Mail (legacy control)
    PubSimple("Mail")
return

^u::  ; try inspector -> selection -> newest inbox
    item := GetCurrentMailItem()
    if IsObject(item) {
        PublishMail(item)
    } else {
        ; REMOVED: TrayTip popup - now only logs to file
        Log("Ctrl+U: no mail item found")
    }
return

^k::  ; newest inbox -> JSON
    try {
        items := olInbox.Items
        items.Sort("[ReceivedTime]", false)
        mail := items.GetFirst()
        if !IsObject(mail) {
            Log("Ctrl+K: no items in Inbox")
        } else {
            Log("Ctrl+K: attempting PublishMail on newest Inbox item")
            PublishMail(mail)
        }
    }
return
