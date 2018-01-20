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