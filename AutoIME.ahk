#Requires AutoHotkey v2.0 64-bit

#SingleInstance Force

;************************** Config & Settings *******************************

InitGlobals()
{
    global ProfileSetsDir := A_WorkingDir . "\Profiles"

    global ConfigFile := A_WorkingDir . "\config.ini"

    global Running := true

    global Visible := true

    global IMEs := []
}

InitConfigs()
{
    global CurrentProfile := IniRead(ConfigFile, "Config", "CurrentProfile", "default")

    if ( not CurrentProfile)
    {
        CurrentProfile := "default"
    }

    global ProfileSetsDir

    global ProfileSets := []

    global StartStopHotkey := IniRead(ConfigFile, "Hotkeys", "StartStop", "")
    global ProfileSetsHotkey := IniRead(ConfigFile, "Hotkeys", "ProfileSets", "")
    global ShowHideHotkey := IniRead(ConfigFile, "Hotkeys", "ShowHide", "")
    global EnableGlobalDefaultIME := IniRead(ConfigFile, "Config", "EnableGlobalDefaultIME", "0")

    global CurrentGlobalIME := IniRead(ConfigFile, "Config", "CurrentGlobalIME", "")

    global StartMinimized := IniRead(ConfigFile, "Config", "StartMinimized", "0")
    global ShowTrayIcon := IniRead(ConfigFile, "Config", "ShowTrayIcon", "1")

    DefaultProfile := ProfileSetsDir . "\default.ini"

    if (ProfileSets.Length = 0)
    {
        if ( not DirExist(ProfileSetsDir))
        {
            DirCreate(ProfileSetsDir)
        }

        FileAppend("", DefaultProfile, "UTF-16")
    }

    if (ShowTrayIcon = "0")
    {
        A_IconHidden := true
    }
}

ScanProfileSets()
{
    global ProfileSetsDir

    global ProfileSets := []

    Loop Files, ProfileSetsDir . "\*.ini", "F"
    {
        FileName := A_LoopFileName
        FileName := SubStr(FileName, 1, StrLen(FileName) - StrLen(A_LoopFileExt) - 1)
        ProfileSets.Push(FileName)
    }
}

ScanIMEs()
{
    global IMEs

    count := 256

    buffer_size := A_PtrSize * count

    buf := Buffer(buffer_size, 0)

    DllCall("lib\ime\GetIMEs", "Ptr", buf, "UInt", count)

    IMEs := StrGet(buf)
    IMEs := StrSplit(IMEs, "|")
}

;************************** GUI Creating *******************************

CreateGUI()
{
    global StartStopHotkey
    global ProfileSetsHotkey
    global ShowHideHotkey
    global CurrentProfile
    global IMEs
    global StartMinimized
    global Visible
    global ShowTrayIcon
    global StartMinimized
    global EnableGlobalDefaultIME
    global CurrentGlobalIME

    GuiWidth := "w248"
    InputWdith := "W130"
    ControlHeight := "H20"

    global Win := Gui()
    Win.Opt("+LastFound -AlwaysOnTop -Caption +ToolWindow")
    Win.SetFont("s8", "Microsoft YaHei")

    ;************************** Hotkeys GroupBox *******************************
    Win.Add("GroupBox", "Wrap x10 r6 h0 " . GuiWidth, "快捷键")

    Win.Add("Text", "X20 YP+20 W60 " . ControlHeight, "开关快捷键")
    Win.Add("Hotkey", "XP+66 YP-2 vChosenStartSopHotkey " . ControlHeight, StartStopHotkey).OnEvent("Change", HotkeyHandler)

    Win.Add("Text", "X20 YP+30 W60 " . ControlHeight, "预设快捷键")
    Win.Add("Hotkey", "XP+66 YP-2 vChosenProfileSetsHotkey " . ControlHeight, ProfileSetsHotkey).OnEvent("Change", HotkeyHandler)

    Win.Add("Text", "X20 YP+30 W60 " . ControlHeight, "显隐快捷键")
    Win.Add("Hotkey", "XP+66 YP-2 vChosenShowHideHotkey " . ControlHeight, ShowHideHotkey).OnEvent("Change", HotkeyHandler)

    Win.Add("Checkbox", "X20 YP+30 vShowTrayIcon W90 " . ControlHeight . (ShowTrayIcon = "0" ? " Checked" : " "), "隐藏托盘图标").OnEvent("Click", CheckHandler)

    Win.Add("Checkbox", "XP110 YP vStartMinimized W90 " . ControlHeight . (StartMinimized = "1" ? " Checked" : " "), "最小化启动").OnEvent("Click", CheckHandler)

    ;************************** Process List View *******************************
    global LV := Win.Add("ListView", "+Checked +Redraw +Report R14 X10 " . GuiWidth, ["进程名", "输入状态"])
    LV.OnEvent("ContextMenu", ShowContextMenu)
    LV.OnEvent("ItemCheck", LVItemCheckHandler)
    LV.Focus()

    ;************************** ProfileSets GroupBox *******************************
    Win.Add("GroupBox", "Wrap x10 r4.5 h1 " . GuiWidth, "预设")
    Win.Add("Text", "X20 YP+20 W50 " . ControlHeight, "预设名称")
    global ProfileNameEdit := Win.Add("Edit", "XP+50 YP-3 " . InputWdith . " " . ControlHeight, CurrentProfile)
    Win.Add("Button", "XP+135 W20 " . ControlHeight, "＋").OnEvent("Click", AddProfile)
    Win.Add("Button", "XP+25 W20 " . ControlHeight, "－").OnEvent("Click", RemoveProfile)

    Win.Add("Text", "X20 YP+30 W50 " . ControlHeight, "预设选择")
    global ProfileSetsList := Win.Add("DropDownList", "XP+50 YP-3 " . InputWdith, ProfileSets)
    ProfileSetsList.OnEvent("Change", ProfileSetsChangeHandler)

    Win.Add("Checkbox", "X20 YP+30 vEnableGlobalDefaultIME W100 " . ControlHeight . (EnableGlobalDefaultIME = "1" ? " Checked" : " "), "设置全局默认").OnEvent("Click", CheckHandler)
    global IMEsList := Win.Add("DropDownList", "XP+100 " . InputWdith, IMEs)
    IMEsList.OnEvent("Change", IMEsListChangeHandler)


    loop ProfileSets.Length
    {
        if (ProfileSets[A_Index] = CurrentProfile)
        {
            ProfileSetsList.Choose(A_Index)
        }
    }

    loop IMEs.Length
    {
        if (IMEs[A_Index] = CurrentGlobalIME)
        {
            IMEsList.Choose(A_Index)
        }
    }

    if (StartMinimized = "1")
    {
        Win.Hide()

        Visible := false
    }
    else
    {
        Win.Show()

        Visible := true
    }
}

CreateContextMenu()
{
    global IMEs
    global IMEsList
    global ContextMenu := Menu()

    ScanIMEs()

    for i, IME in IMEs
    {
        ContextMenu.Add(IME, ContextMenuHandler)
    }

    ContextMenu.Add()
    ContextMenu.Add("清除", ContextMenuHandler)

    ContextMenu.Add()
    ContextMenu.Add("刷新", ContextMenuHandler)

    IMEsList.Delete()
    IMEsList.Add(IMEs)
}

CreateTrayMenu()
{
    global Running
    global CurrentProfile

    A_TrayMenu.Delete()

    A_TrayMenu.Add(Running ? "开" : "关", TrayMenuHandler)

    SubMenu := Menu()

    for i, Profile in ProfileSets
    {
        SubMenu.Add(Profile, SwitchProfileHandler)

        if (Profile = CurrentProfile)
        {
            SubMenu.Check(Profile)
        }
    }

    A_TrayMenu.Add("预设", SubMenu)

    A_TrayMenu.Add("重启", TrayMenuHandler)
    A_TrayMenu.Add("退出", TrayMenuHandler)

    OnMessage(0x404, NotifyIcon)
}

RefreshTrayMenu := CreateTrayMenu

AddProcessToListView()
{
    global LV
    global CurrentProfile

    store := Map()

    HWNDs := WinGetList(, , "Task Manager")

    ImageListID := IL_Create(HWNDs.Length)

    LV.SetImageList(ImageListID)

    sfi_size := A_PtrSize + 688
    sfi := Buffer(sfi_size)

    for HWND in HWNDs
    {
        try
        {
            hProcess := DllCall("Oleacc\GetProcessHandleFromHwnd", "Ptr", HWND)

            pid := DllCall("GetProcessId", "Ptr", hProcess, "UInt")

            Name := ProcessGetName(pid)
            Path := ProcessGetPath(pid)

            if store.Has(Name)
            {
                continue
            }

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
            else
            {
                hIcon := NumGet(sfi, 0, "Ptr")

                IconNumber := DllCall("ImageList_ReplaceIcon", "Ptr", ImageListID, "Int", -1, "Ptr", hIcon) + 1

                DllCall("DestroyIcon", "Ptr", hIcon)
            }

            if ( not CurrentProfile) {
                LV.Add("Vis Icon" . IconNumber, Name)
            }
            else
            {
                Config := ProfileSetsDir . "\" . CurrentProfile . ".ini"

                IME := IniRead(Config, Name, "IME", "")
                Enable := IniRead(Config, Name, "Enable", "0")

                Check := Enable = "1" ? "Check" : ""

                if (IME) {
                    LV.Add("Vis Icon" . IconNumber . " " . Check, Name, IME)
                } else {
                    LV.Add("Vis Icon" . IconNumber, Name)
                }
            }

            DllCall("CloseHandle", "Ptr", hProcess)
        }
        catch
        {
        }
    }

    LV.ModifyCol()
}

;************************** Callbacks *******************************

ContextMenuHandler(ItemName, *)
{
    global LV

    if (ItemName = "刷新")
    {
        Refresh()

        return
    }

    RowNumber := 0

    Loop
    {
        RowNumber := LV.GetNext(RowNumber)

        if ( not RowNumber)
        {
            break
        }

        ItemState := SendMessage(0x102C, RowNumber - 1, 0xF000, LV)
        IsChecked := (ItemState >> 12) - 1

        SetImeForProcess(LV.GetText(RowNumber, 1), ItemName = "清除" ? "" : ItemName, ItemName = "清除" ? false : IsChecked)

        if (ItemName = "清除")
        {
            LV.Modify(RowNumber, , , "")
        }
        else
        {
            LV.Modify(RowNumber, , , ItemName)
        }
    }

    LV.ModifyCol()
}

ShowContextMenu(LV, Item, IsRightClick, X, Y)
{
    global ContextMenu

    CreateContextMenu()

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

    RefreshTrayMenu()

    TrayTip(Running ? "输入法自动切换已启动" : "输入法自动切换已停止")
}

ProfileSetsAction(HotkeyName)
{
    global Win
    global ProfileSets
    global CurrentProfile
    global ProfileSetsList
    global ProfileNameEdit

    if ( not ProfileSets.Length)
    {
        return
    }

    Index := 0

    loop ProfileSets.Length
    {
        if (ProfileSets[A_Index] = CurrentProfile)
        {
            Index := A_Index
        }
    }

    if Index = ProfileSets.Length
    {
        Index := 1
    }
    else
    {
        Index++
    }

    CurrentProfile := ProfileSets[Index]

    ProfileSetsList.Choose(Index)
    ProfileNameEdit.Value := CurrentProfile

    IniWrite(CurrentProfile, ConfigFile, "Config", "CurrentProfile")

    Refresh()

    TrayTip("已切换至预设：" . CurrentProfile)
}

ShowHideWindowAction(HotkeyName)
{
    global Win
    global Visible

    if (Visible)
    {
        Win.Hide()

        Visible := false
    }
    else
    {
        Win.Show()

        Visible := true
    }
}

TrayMenuHandler(ItemName, ItemPos, MyMenu)
{
    if (ItemName = "退出")
    {
        ExitApp()

        return
    }
    else if (ItemName = "重启")
    {
        Reload()

        return
    }
}

SwitchProfileHandler(ItemName, ItemPos, MyMenu)
{
    global CurrentProfile
    global ProfileSetsList
    global ProfileNameEdit

    CurrentProfile := ItemName

    for i, Profile in ProfileSets {
        if (Profile = CurrentProfile) {
            ProfileSetsList.Choose(i)
        }
    }

    ProfileNameEdit.Value := CurrentProfile

    IniWrite(CurrentProfile, ConfigFile, "Config", "CurrentProfile")

    Refresh()

    RefreshTrayMenu()

    TrayTip("已切换至预设：" . CurrentProfile)
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
    }
    else if (HotkeyControl.name = "ChosenProfileSetsHotkey")
    {
        IniWrite(HotkeyControl.value, ConfigFile, "Hotkeys", "ProfileSets")
    }
    else if (HotkeyControl.name = "ChosenShowHideHotkey")
    {
        IniWrite(HotkeyControl.value, ConfigFile, "Hotkeys", "ShowHide")
    }
}

ProfileSetsChangeHandler(Control, Info)
{
    global CurrentProfile

    ProfileNameEdit.Value := Control.Text

    CurrentProfile := Control.Text

    IniWrite(CurrentProfile, ConfigFile, "Config", "CurrentProfile")

    Refresh()
}

IMEsListChangeHandler(Control, Info)
{
    global CurrentGlobalIME

    CurrentGlobalIME := Control.Text

    IniWrite(CurrentGlobalIME, ConfigFile, "Config", "CurrentGlobalIME")
}

AddProfile(Control, Info)
{
    global ProfileSetsDir
    global ProfileSets
    global CurrentProfile
    global ProfileSetsList
    global ProfileNameEdit

    FileName := ProfileNameEdit.value

    if ( not FileName)
    {
        return
    }

    FileFullName := ProfileSetsDir . "\" . FileName . ".ini"

    if (FileExist(FileFullName))
    {
        return
    }

    if ( not DirExist(ProfileSetsDir))
    {
        DirCreate(ProfileSetsDir)
    }

    FileAppend("", FileFullName, "UTF-16")

    SetCurrentProfile(FileName)
}

RemoveProfile(Control, Info)
{
    global ProfileSetsDir
    global ProfileSets
    global ProfileNameEdit
    global ProfileSetsList

    FileName := ProfileNameEdit.value

    if ( not FileName)
    {
        return
    }

    FileFullName := ProfileSetsDir . "\" . FileName . ".ini"

    if ( not FileExist(FileFullName))
    {
        return
    }

    FileDelete(FileFullName)

    SetCurrentProfile("")
}

LVItemCheckHandler(Control, Item, Checked)
{
    global CurrentProfile
    global ProfileSetsDir

    Config := ProfileSetsDir . "\" . CurrentProfile . ".ini"

    if ( not CurrentProfile)
    {
        MsgBox("请先选择预设")
    }

    Process := Control.GetText(Item, 1)
    IME := Control.GetText(Item, 2)

    SetImeForProcess(Process, IME, Checked)
}

CheckHandler(Control, Info)
{
    if (Control.Name = "ShowTrayIcon")
    {
        IniWrite(Control.Value = 1 ? "0" : "1", ConfigFile, "Config", "ShowTrayIcon")

        if (Control.Value = 1)
        {
            A_IconHidden := true
        }
        else
        {
            A_IconHidden := false
        }
    }
    else if (Control.Name = "StartMinimized")
    {
        IniWrite(Control.Value, ConfigFile, "Config", "StartMinimized")
    }
    else if (Control.Name = "EnableGlobalDefaultIME")
    {
        IniWrite(Control.Value, ConfigFile, "Config", "EnableGlobalDefaultIME")
    }
}
;************************** Utils ********************************
SetImeForProcess(Process, IME, Checked)
{
    global CurrentProfile
    global ProfileSetsDir

    if ( not CurrentProfile)
    {
        return
    }

    Config := ProfileSetsDir . "\" . CurrentProfile . ".ini"

    IniWrite(IME, Config, Process, "IME")

    if (Checked)
    {
        IniWrite("1", Config, Process, "Enable")
    }
    else
    {
        IniWrite("0", Config, Process, "Enable")
    }
}

SetCurrentProfile(ProfileName)
{
    global CurrentProfile
    global ProfileSetsList
    global ProfileNameEdit
    global ProfileSets
    global ConfigFile

    ScanProfileSets()

    ProfileSetsList.Delete()
    ProfileSetsList.Add(ProfileSets)

    for i, Profile in ProfileSets
    {
        if (Profile = ProfileName)
        {
            ProfileSetsList.Choose(i)
        }
    }

    CurrentProfile := ProfileName

    ProfileNameEdit.Value := ProfileName

    IniWrite(CurrentProfile, ConfigFile, "Config", "CurrentProfile")

    Refresh()

    RefreshTrayMenu()
}
;************************** Hotkeys *******************************
RegisterHotkeys()
{
    Hotkey("Esc", HideWindow)

    global StartStopHotkey
    global ProfileSetsHotkey
    global ShowHideHotkey

    if (StartStopHotkey)
    {
        Hotkey(StartStopHotkey, StartStopAction)
    }

    if (ProfileSetsHotkey)
    {
        Hotkey(ProfileSetsHotkey, ProfileSetsAction)
    }

    if (ShowHideHotkey)
    {
        Hotkey(ShowHideHotkey, ShowHideWindowAction)
    }
}
;************************** Entry *******************************
Refresh()
{
    global LV

    LV.Delete()
    AddProcessToListView()
}

;************************** Entry *******************************
Start()
{
    SetTimer(Entry, 2000)

    Entry() {
        global Running
        global CurrentProfile
        global EnableGlobalDefaultIME
        global CurrentGlobalIME

        if (Running && CurrentProfile)
        {
            ConfigFile := ProfileSetsDir . "\" . CurrentProfile . ".ini"

            try
            {
                HWND := DllCall("GetForegroundWindow")

                if ( not HWND)
                {
                    return
                }

                Process := DllCall("Oleacc\GetProcessHandleFromHwnd", "Ptr", HWND)

                pid := DllCall("GetProcessId", "Ptr", Process, "UInt")

                Name := ProcessGetName(pid)

                IME := IniRead(ConfigFile, Name, "IME", "")

                Enable := IniRead(ConfigFile, Name, "Enable", "0")

                if (IME && Enable = "1")
                {
                    DllCall("lib\ime\SetIME", "Str", IME)
                }
                else if (EnableGlobalDefaultIME = "1" && CurrentGlobalIME)
                {
                    DllCall("lib\ime\SetIME", "Str", CurrentGlobalIME)
                }

                DllCall("CloseHandle", "Ptr", Process)
            }
            catch
            {
            }
        }
    }
}


Main()
{
    InitGlobals()

    InitConfigs()

    ScanProfileSets()

    ScanIMEs()

    CreateGUI()

    CreateTrayMenu()

    AddProcessToListView()

    RegisterHotkeys()

    Start()
}

Main()