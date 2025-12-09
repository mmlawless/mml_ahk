CoordMode, Pixel, Screen
CheckStatusInterval := 1000
lastStatus := ""

availableIcon := "C:\\Users\\mlawless\\Dropbox\\001 Home\\800 Useful\\801 PC Setup\\003 AutoHotKey\\002 AHK Output Teams Status MQTT\\!MML AHK_Teams_Available.png"
busyIcon := "C:\\Users\\mlawless\\Dropbox\\001 Home\\800 Useful\\801 PC Setup\\003 AutoHotKey\\002 AHK Output Teams Status MQTT\\!MML AHK_Teams_Busy.png"
dndIcon := "C:\\Users\\mlawless\\Dropbox\\001 Home\\800 Useful\\801 PC Setup\\003 AutoHotKey\\002 AHK Output Teams Status MQTT\\!MML AHK_Teams_DoND.png"
appofflIcon := "C:\\Users\\mlawless\\Dropbox\\001 Home\\800 Useful\\801 PC Setup\\003 AutoHotKey\\002 AHK Output Teams Status MQTT\\!MML AHK_Teams_AppearOffline.png"
brbIcon := "C:\\Users\\mlawless\\Dropbox\\001 Home\\800 Useful\\801 PC Setup\\003 AutoHotKey\\002 AHK Output Teams Status MQTT\\!MML AHK_Teams_BeRightBack.png"

mqttPath := "C:\\Program Files\\mosquitto\\mosquitto_pub.exe"
mqttServer := "192.168.1.15"
mqttTopic := "MMLNR/TeamsStatusIN"

Loop {
    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *20 %availableIcon%
    if (ErrorLevel = 0) {
        if (lastStatus != "MML_T_Available") {
            lastStatus := "MML_T_Available"
            runMqttCommand("Available")
        }
    } else {
        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *20 %busyIcon%
        if (ErrorLevel = 0) {
            if (lastStatus != "MML_T_Busy") {
                lastStatus := "MML_T_Busy"
                runMqttCommand("Busy")
            }
        } else {
            ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *20 %dndIcon%
            if (ErrorLevel = 0) {
                if (lastStatus != "MML_T_DoNotDisturb") {
                    lastStatus := "MML_T_DoNotDisturb"
                    runMqttCommand("DoNotDisturb")
                }
            } else {
                ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *20 %appofflIcon%
                if (ErrorLevel = 0) {
                    if (lastStatus != "MML_T_AppearOffline") {
                        lastStatus := "MML_T_AppearOffline"
                        runMqttCommand("AppearOffline")
                    }
                } else {
                    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *20 %brbIcon%
                    if (ErrorLevel = 0) {
                        if (lastStatus != "MML_T_BeRightBack") {
                            lastStatus := "MML_T_BeRightBack"
                            runMqttCommand("BeRightBack")
                        }
                    } else {
                        if (lastStatus != "MML_T_Unknown") {
                            lastStatus := "MML_T_Unknown"
                            runMqttCommand("Unknown")
                        }
                    }
                }
            }
        }
    }
    Sleep, %CheckStatusInterval%
}

runMqttCommand(status) {
    global mqttPath, mqttServer, mqttTopic
    Run, %mqttPath% -h %mqttServer% -t %mqttTopic% -m %status%,, Hide
}
