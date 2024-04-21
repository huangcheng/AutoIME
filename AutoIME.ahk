#Requires AutoHotkey v2.0 64-bit

#SingleInstance Force

;************************** Config & Settings *******************************

InitGlobals()
{
    global RuleSetsDir := A_WorkingDir . "\rules"

    global ConfigFile := A_WorkingDir . "\config.ini"

    global Running := true
}

InitConfigs()
{
    global CurrentRule := IniRead(ConfigFile, "Config", "CurrentRule", "")

    global RuleSetsDir

    global RuleSets := []

    global StartStopHotkey := IniRead(ConfigFile, "Hotkeys", "StartStop", "")
    global RuleSetsHotkey := IniRead(ConfigFile, "Hotkeys", "RuleSets", "")
}

ScanRuleSets()
{
    global RuleSetsDir

    global RuleSets := []

    Loop Files, RuleSetsDir . "\*.ini", "F"
    {
        FileName := A_LoopFileName
        FileName := SubStr(FileName, 1, StrLen(FileName) - StrLen(A_LoopFileExt) - 1)
        RuleSets.Push(FileName)
    }
}

;************************** GUI Creating *******************************

CreateGUI()
{
    GuiWidth := "w248"
    InputWdith := "W130"
    ControlHeight := "H20"

    global Win := Gui()
    Win.Opt("+LastFound -AlwaysOnTop -Caption +ToolWindow")

    ;************************** Hotkeys GroupBox *******************************
    global StartStopHotkey
    global RuleSetsHotkey

    Win.Add("GroupBox", "Wrap x10 r3 h0 " . GuiWidth, "快捷键")
    Win.Add("Text", "X20 YP+20 W60 " . ControlHeight, "开关快捷键")
    Win.Add("Hotkey", "XP+66 YP-2 vChosenStartSopHotkey " . ControlHeight, StartStopHotkey).OnEvent("Change", HotkeyHandler)

    Win.Add("Text", "X20 Y54 W60 " . ControlHeight, "预设快捷键")
    Win.Add("Hotkey", "XP+66 YP-2 vChosenRuleSetsHotkey " . ControlHeight, RuleSetsHotkey).OnEvent("Change", HotkeyHandler)

    global LV := Win.Add("ListView", "+Checked +Redraw +Report R14 X10 " . GuiWidth, ["进程名", "输入状态"])
    LV.OnEvent("ContextMenu", ShowContextMenu)
    LV.Focus()

    ;************************** RuleSets GroupBox *******************************
    Win.Add("GroupBox", "Wrap x10 r3 h1 " . GuiWidth, "预设")
    Win.Add("Text", "X20 YP+20 W50 " . ControlHeight, "预设名称")
    global RuleNameEdit := Win.Add("Edit", "XP+50 YP-3 " . InputWdith . " " . ControlHeight)
    Win.Add("Button", "XP+135 W20 " . ControlHeight, "＋").OnEvent("Click", AddRule)
    Win.Add("Button", "XP+25 W20 " . ControlHeight, "－").OnEvent("Click", RemoveRule)

    Win.Add("Text", "X20 YP+30 W50 " . ControlHeight, "预设选择")
    global RuleSetsList := Win.Add("DropDownList", "XP+50 YP-3 " . InputWdith, RuleSets)
    RuleSetsList.OnEvent("Change", RuleSetsChangeHandler)

    loop RuleSets.Length
    {
        if (RuleSets[A_Index] = CurrentRule)
        {
            RuleSetsList.Choose(A_Index)
        }
    }

    Win.Show
}

CreateContextMenu()
{
    global ContextMenu := Menu()

    count := 256

    buffer_size := A_PtrSize * count

    buf := Buffer(buffer_size, 0)

    DllCall("lib\ime\GetIMEs", "Ptr", buf, "UInt", count)

    IMEs := StrGet(buf)
    IMEs := StrSplit(IMEs, "|")

    for i, IME in IMEs
    {
        ContextMenu.Add(IME, AssociateInputMethod)
    }

}

CreateTrayMenu()
{
    A_TrayMenu.Delete()

    A_TrayMenu.Add("重启", TrayMenuHandler)
    A_TrayMenu.Add("退出", TrayMenuHandler)

    OnMessage(0x404, NotifyIcon)
}

AddProcessToListView()
{
    global LV

    store := Map()

    HWNDs := WinGetList(, , "Task Manager")

    ImageListID := IL_Create(HWNDs.Length)

    LV.SetImageList(ImageListID)

    sfi_size := A_PtrSize + 688
    sfi := Buffer(sfi_size)

    for HWND in HWNDs
    {
        hProcess := DllCall("Oleacc\GetProcessHandleFromHwnd", "Ptr", HWND)

        pid := DllCall("GetProcessId", "Ptr", hProcess, "UInt")

        Name := ProcessGetName(pid)
        Path := ProcessGetPath(pid)

        if store.Has(Name)
            continue

        store[Name] := 1

        if not DllCall(
            "Shell32\SHGetFileInfoW",
            "Str",
            Path,
            "Uint",
            0,
            "Ptr",
            sfi,
            "UInt",
            sfi_size,
            "UInt",
            0x101
        )
        {
            IconNumber := 9999999
        }
        else {
            hIcon := NumGet(sfi, 0, "Ptr")

            IconNumber := DllCall("ImageList_ReplaceIcon", "Ptr", ImageListID, "Int", -1, "Ptr", hIcon) + 1

            DllCall("DestroyIcon", "Ptr", hIcon)
        }

        LV.Add("Vis Icon" . IconNumber, Name)

        DllCall("CloseHandle", "Ptr", hProcess)
    }

    LV.ModifyCol()
}

;************************** Callbacks *******************************

AssociateInputMethod(ItemName, *)
{
    global LV

    FocusedRowNumber := LV.GetNext(0, "F")

    if ( not FocusedRowNumber)
    {
        return
    }


    LV.Modify(FocusedRowNumber, , , ItemName)

    LV.ModifyCol()
}

ShowContextMenu(LV, Item, IsRightClick, X, Y)
{
    global ContextMenu

    ContextMenu.Show(X, Y)
}

HideWindow(HotkeyName)
{
    global Win

    Win.Hide
}

StartStopAction(HotkeyName)
{
    global Running

    Running := !Running

    TrayTip(Running ? "输入法自动切换已启动" : "输入法自动切换已停止")
}

RuleSetsAction(HotkeyName)
{
    global Win

    MsgBox("RuleSetsAction")
}

TrayMenuHandler(ItemName, ItemPos, MyMenu)
{
    if (ItemName = "退出")
    {
        ExitApp()
    } else if (ItemName = "重启")
    {
        Reload()
    }
}

NotifyIcon(wParam, lParam, msg, hwnd)
{
    if (lParam = 0x202)
    {
        Win.Show
    }
}

HotkeyHandler(HotkeyControl, Info)
{
    global Win

    if (HotkeyControl.name = "ChosenStartSopHotkey")
    {
        IniWrite(HotkeyControl.value, ConfigFile, "Hotkeys", "StartStop")
    } else if (HotkeyControl.name = "ChosenRuleSetsHotkey")
    {
        IniWrite(HotkeyControl.value, ConfigFile, "Hotkeys", "RuleSets")
    }
}

RuleSetsChangeHandler(Control, Info)
{
    global CurrentRule

    RuleNameEdit.Value := Control.Text

    CurrentRule := Control.Value

    IniWrite(CurrentRule, ConfigFile, "Config", "CurrentRule")
}

AddRule(Control, Info)
{
    global RuleSetsDir
    global RuleSets

    FileName := RuleNameEdit.value

    if ( not FileName)
    {
        return
    }

    FileFullName := RuleSetsDir . "\" . FileName . ".ini"

    if (FileExist(FileFullName))
    {
        return
    }

    if ( not DirExist(RuleSetsDir))
    {
        DirCreate(RuleSetsDir)
    }

    FileAppend("", FileFullName, "UTF-16")

    ScanRuleSets()

    RuleSetsList.Delete()
    RuleSetsList.Add(RuleSets)
}

RemoveRule(Control, Info)
{
    global RuleSetsDir
    global RuleSets

    FileName := RuleNameEdit.value

    if ( not FileName)
    {
        return
    }

    FileFullName := RuleSetsDir . "\" . FileName . ".ini"

    if ( not FileExist(FileFullName))
    {
        return
    }

    FileDelete(FileFullName)

    ScanRuleSets()

    RuleSetsList.Delete()
    RuleSetsList.Add(RuleSets)
}

;************************** Hotkeys *******************************
RegisterHotkeys()
{
    Hotkey("Esc", HideWindow)

    global StartStopHotkey
    global RuleSetsHotkey

    if (StartStopHotkey)
    {
        Hotkey(StartStopHotkey, StartStopAction)
    }

    if (RuleSetsHotkey)
    {
        Hotkey(RuleSetsHotkey, RuleSetsAction)
    }
}

;************************** Entry *******************************


Main()
{
    InitGlobals()

    InitConfigs()

    ScanRuleSets()

    CreateGUI()

    CreateContextMenu()

    CreateTrayMenu()

    AddProcessToListView()

    RegisterHotkeys()
}

Main()