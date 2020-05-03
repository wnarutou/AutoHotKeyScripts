
;Use Alt+Ctrl+n replace PageDown in explorer window.
^!n::
     WinGet, prcess_name, ProcessName, A
     if (prcess_name = "explorer.exe")
            Send, {PgDn}
Return

;Use Alt+Ctrl+p replace PageUp in explorer window.
^!p::
     WinGet, prcess_name, ProcessName, A
     if (prcess_name = "explorer.exe")
            Send, {PgUp}
Return

;Use Ctrl+n replace DownArrow in explorer window.
^n::
     WinGet, prcess_name, ProcessName, A
     if (prcess_name = "explorer.exe")
            Send, {Down}
Return

;Use Ctrl+p replace UpArrow in explorer window.
^p::
     WinGet, prcess_name, ProcessName, A
     if (prcess_name = "explorer.exe")
            Send, {Up}
Return

;Use Ctrl+Alt+2 run everything with path parameter current path 
#Include, explorer.ahk
~^!2::
     curpath := Explorer_GetWndCurPath()
     if %curpath%
     {
          if (curpath = A_Desktop)
               return
          Run, C:\Program Files\Everything\Everything.exe -path "%curpath%"
     }
return 

;Use Ctrl+h to locate file or directory in current directory
#Include zh2py.ahk
~^h::
    WinGet, process, processName, % "ahk_id" hwnd := WinExist("A")
    if (process!="explorer.exe")
        return
    WinGetClass class, ahk_id %hwnd%
    filenamekeys := "abcdefghijklmnopqrstuvwxyz0123456789!@#$%^&()_[]{}.'``~"
    if (class ~= "(Cabinet|Explore)WClass")
    {
        nameList := Explorer_GetWndItemsNameList()
        pyNameList := Array()
        for index, name in nameList
        {
            lowername := zh2py(name)
            StringLower, lowername, lowername
            pyNameList.push(lowername)
        }

        matchedList := Array()
        matchStr :=
        next := 1

        LOOP
        {
            Input, keystroke, L1 M, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{Capslock}{Numlock}{PrintScreen}{Pause}
            If InStr(ErrorLevel, "EndKey:")
                break

            curHwnd := WinExist("A")
            if (hwnd != curHwnd)
                break
            StringLower, keystroke, keystroke

            if (keystroke = "`t")
            {
                if (StrLen(matchStr) = 0)
                    Continue
                next += 1
            }
            else if (instr(filenamekeys, keystroke))
            {
                matchStr .= keystroke

                matchedList := []
                for index, name in pyNameList
                    IfInString, name, %matchStr%
                    {
                        matchedList.push(index)
                    }
            }
            else
            {
                break
            }
            Explorer_SetWndItemSelection("", matchedList[next] - 1)
        }
    }
return