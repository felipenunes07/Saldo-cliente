toggle := false

F8::
{
    global toggle
    toggle := true

    ToolTip("Automação ON")

    while (toggle)
    {
        Send("{Right}")
        Sleep(17000)
        Send("{Right}")
        Sleep(17000)
    }
}

F9::
{
    global toggle
    toggle := false

    ToolTip("Automação OFF")
    Sleep(1000)
    ToolTip()
}