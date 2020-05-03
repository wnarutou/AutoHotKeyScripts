/*
  Library for getting info from a specific explorer window (if window handle not specified, the currently active
  window will be used).  Requires AHK_L or similar.  Works with the desktop.  Does not currently work with save
  dialogs and such.
 
 
  Explorer_GetSelected(hwnd="")   - paths of target window's selected items
  Explorer_GetAll(hwnd="")        - paths of all items in the target window's folder
  Explorer_GetPath(hwnd="")       - path of target window's folder
 
  example:
   F1::
    path := Explorer_GetPath()
    all := Explorer_GetAll()
    sel := Explorer_GetSelected()
    MsgBox % path
    MsgBox % all
    MsgBox % sel
   return
  
  original version by Joshua A. Kinnison
    2011-04-27, 16:12


  1. Do not return Error message
  2. Change function names:
       Explorer_GetSelected -> Explorer_GetSelectedPath
       Explorer_GetAll -> Explorer_GetAllPath
       Explorer_GetPath -> Explorer_GetWndCurPath
       Explorer_Get -> Explorer_GetWndItemsPath
  3. Add functions:
       Explorer_GetWndItems
  edit by wnarutou
    2020-05-01, 23:06
*/

Explorer_GetWndCurPath(hwnd="")
{
  if !(window := Explorer_GetWindow(hwnd))
    return
  if (window="desktop")
    return A_Desktop
  path := window.LocationURL
  path := RegExReplace(path, "ftp://.*@","ftp://")
  StringReplace, path, path, file:///
  StringReplace, path, path, /, \, All 
  
  ; 存在URL编码的情况，需要将URL编码进行转码
  ; thanks to polyethene
  Loop
    If RegExMatch(path, "i)(?<=%)[\da-f]{1,2}", hex)
      StringReplace, path, path, `%%hex%, % Chr("0x" . hex), All
    Else Break
  return path
}

Explorer_GetAllPath(hwnd="")
{
  return Explorer_GetWndItemsPath(hwnd)
}

Explorer_GetSelectedPath(hwnd="")
{
  return Explorer_GetWndItemsPath(hwnd,true)
}

; 获取explorer窗口对象
Explorer_GetWindow(hwnd="")
{
  ; thanks to jethrow for some pointers here
  ; 没有参数则取当前活动窗口的进程名
  WinGet, process, processName, % "ahk_id" hwnd := hwnd? hwnd:WinExist("A")

  ; 获取类名
  WinGetClass class, ahk_id %hwnd%
 
  ; 仅支持explorer
  if (process!="explorer.exe")
    return

  ; explorer存在2种类
  if (class ~= "(Cabinet|Explore)WClass")
  {
    ; 从所有explorer中寻找
    for window in ComObjCreate("Shell.Application").Windows
      if (window.hwnd==hwnd)
        return window
  }
  ; 桌面类
  else if (class ~= "Progman|WorkerW") 
    return "desktop" ; desktop found
}

Explorer_GetWndItemsPath(hwnd="",selection=false)
{
  if !(window := Explorer_GetWindow(hwnd))
    return
  if (window="desktop")
  {
    ; 获取桌面对象的ListView
    ControlGet, hwWindow, HWND,, SysListView321, ahk_class Progman
    if !hwWindow ; #D mode
      ControlGet, hwWindow, HWND,, SysListView321, A
    ; 获取ListView中的Item内容
    ControlGet, files, List, % ( selection ? "Selected":"") "Col1",,ahk_id %hwWindow%
    base := SubStr(A_Desktop,0,1)=="\" ? SubStr(A_Desktop,1,-1) : A_Desktop
    Loop, Parse, files, `n, `r
    {
      path := base "\" A_LoopField
      IfExist %path% ; ignore special icons like Computer (at least for now)
        ret .= path "`n"
   }
  }
  else
  {
    if selection
      collection := window.document.SelectedItems
    else
      collection := window.document.Folder.Items
    for item in collection
      ret .= item.path "`n"
  }
  return Trim(ret,"`n")
}

Explorer_GetWndItemsNameList(hwnd="",selection=false)
{
  if !(window := Explorer_GetWindow(hwnd))
    return

  nameList := Array()
  if (window="desktop")
  {
    ; 获取桌面对象的ListView
    ControlGet, hwWindow, HWND,, SysListView321, ahk_class Progman
    if !hwWindow ; #D mode
      ControlGet, hwWindow, HWND,, SysListView321, A
    ; 获取ListView中的Item内容
    ControlGet, files, List, % ( selection ? "Selected":"") "Col1",,ahk_id %hwWindow%
    base := SubStr(A_Desktop,0,1)=="\" ? SubStr(A_Desktop,1,-1) : A_Desktop
    Loop, Parse, files, `n, `r
    {
      path := base "\" A_LoopField
      IfExist %path% ; ignore special icons like Computer (at least for now)
        nameList.push(A_LoopField)
   }
  }
  else
  {
    if selection
      collection := window.document.SelectedItems
    else
      collection := window.document.Folder.Items
    for item in collection
      nameList.push(item.Name)
  }
  return nameList
}

Explorer_SetWndItemSelection(hwnd="", itemindex=0)
{
  if !(window := Explorer_GetWindow(hwnd))
    return

  if (window="desktop")
  {
    ; This does not work
    ;LVM_FIRST := 0x1000
    ;LVM_SETITEMSTATE := LVM_FIRST+43
    ;LVIS_SELECTED := 0x2
    ;LVIS_FOCUSED := 0x1
    ;WinId := WinExist("A")
    ;VarSetCapacity(LVITEM, 20, 0) ;to receive LVITEM
    ;NumPut(LVIS_FOCUSED | LVIS_SELECTED, LVITEM, 12)  ; state
    ;NumPut(LVIS_FOCUSED | LVIS_SELECTED, LVITEM, 16)  ; stateMask
    ;RemoteBuf_Open(hLVITEM, WinId, 20)  ; MASTER_ID = the ahk_id of the process owning the SysListView32 control
    ;RemoteBuf_Write(hLVITEM, LVITEM,20)
    ;SendMessage,LVM_SETITEMSTATE,3,RemoteBuf_Get(hLVITEM),SysListView321,ahk_id %WinId%  ; supposing the cntrlNN is SysListView321
    ;RemoteBuf_Close(hLVITEM)
  }
  else
  {
    collection := window.document.Folder.Items
    ; 4 means Deselect all item but the specified item.
    window.document.SelectItem(collection.Item[itemindex], 4)
    ; 16 means Give the item the focus.
    window.document.SelectItem(collection.Item[itemindex], 16)
    ; 1 means Select the item.
    window.document.SelectItem(collection.Item[itemindex], 1)
  }
}
