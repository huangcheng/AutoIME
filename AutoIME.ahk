#Requires AutoHotkey v2.0 64-bit

#SingleInstance Force

; Initialize GUI
CreateGUI()
{
    global Win := Gui()
    Win.Opt("+LastFound +AlwaysOnTop -Caption +ToolWindow")

    global LV := Win.Add("ListView", "+Checked +Redraw +Report r14 x10 w208", ["进程名", "输入状态"])

    LV.OnEvent("ContextMenu", ShowContextMenu)

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

AssociateInputMethod(ItemName, *)
{
    global LV

    FocusedRowNumber := LV.GetNext(0, "F")
    if not FocusedRowNumber
        return

    LV.Modify(FocusedRowNumber, , , ItemName)

    LV.ModifyCol()
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

ShowContextMenu(LV, Item, IsRightClick, X, Y)  ; In response to right-click or Apps key.
{
    global ContextMenu

    ContextMenu.Show(X, Y)
}


Main()
{
    CreateGUI()
    CreateContextMenu()
    AddProcessToListView()
}

Main()