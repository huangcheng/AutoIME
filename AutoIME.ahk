#Requires AutoHotkey v2.0 64-bit

#SingleInstance Force

;************************** Config & Settings *******************************

InitGlobals()
{
    global RuleSetsDir := A_WorkingDir . "\rules"

    global ConfigFile := A_WorkingDir . "\config.ini"
}

InitConfigs()
{
    global CurrentRule := IniRead(ConfigFile, "Config", "CurrentRule", "")

    global RuleSetsDir

    global RuleSets := []

    Loop Files, RuleSetsDir . "\*.ini", "F"
    {
        RuleSets.Push(A_LoopFileFullPath)
    }
}

;************************** GUI Creating *******************************

CreateGUI()
{
    GuiWidth := "w248"
    InputWdith := "W130"
    ControlHeight := "H20"

    global Win := Gui()
    Win.Opt("+LastFound +AlwaysOnTop -Caption +ToolWindow")

    ;************************** Hotkeys GroupBox *******************************
    Win.Add("GroupBox", "Wrap x10 r3 h0 " . GuiWidth, "快捷键")
    Win.Add("Text", "X20 YP+20 W60 " . ControlHeight, "开关快捷键")
    Win.Add("Hotkey", "XP+66 YP-2 vChosenShowHideHotkey " . ControlHeight)

    Win.Add("Text", "X20 Y54 W60 " . ControlHeight, "预设快捷键")
    Win.Add("Hotkey", "XP+66 YP-2 vChosenStartStpHotkey " . ControlHeight)

    global LV := Win.Add("ListView", "+Checked +Redraw +Report R14 X10 " . GuiWidth, ["进程名", "输入状态"])
    LV.OnEvent("ContextMenu", ShowContextMenu)

    ;************************** RuleSets GroupBox *******************************
    Win.Add("GroupBox", "Wrap x10 r3 h1 " . GuiWidth, "预设")
    Win.Add("Text", "X20 YP+20 W50 " . ControlHeight, "预设名称")
    global RuleEdit := Win.Add("Edit", "XP+50 YP-3 " . InputWdith . " " . ControlHeight)
    Win.Add("Button", "XP+135 W20 " . ControlHeight, "＋")
    Win.Add("Button", "XP+25 W20 " . ControlHeight, "－")

    Win.Add("Text", "X20 YP+30 W50 " . ControlHeight, "预设选择")
    Win.Add("ComboBox", "XP+50 YP-3 vColorChoice " . InputWdith, RuleSets)

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

    global RuleEdit

    OutputDebug(RuleEdit.Value)
}

ShowContextMenu(LV, Item, IsRightClick, X, Y)  ; In response to right-click or Apps key.
{
    global ContextMenu

    ContextMenu.Show(X, Y)
}

ShowHideHotkeyChange()
{
    global Win

    OutputDebug(Win.ChosenShowHideHotkey.Value)
}

HideWindow(HotkeyName)
{
    global Win

    Win.Hide
}

TrayMenuHandler(ItemName, ItemPos, MyMenu)
{
    if (ItemName = "退出")
    {
        ExitApp()
    }
}

NotifyIcon(wParam, lParam, msg, hwnd)
{
    if (lParam = 0x202)
    {
        Win.Show
    }
}

;************************** Hotkeys *******************************
RegisterHotkeys()
{
    Hotkey("Esc", HideWindow)
}

;************************** Entry *******************************


Main()
{
    InitGlobals()

    InitConfigs()

    CreateGUI()

    CreateContextMenu()

    CreateTrayMenu()

    AddProcessToListView()

    RegisterHotkeys()
}

Main()