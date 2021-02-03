$MyCount = 0
$PinCounter = 1
"Looking for your pin every 3 seconds"

Do{

    $wshell = New-Object -ComObject wscript.shell;
    if($wshell.AppActivate('title of the application window'))
    {
        "# times pin found:" + $PinCounter
        "@ $(Get-Date)"
        $PinCounter++
        Sleep 1
        $wshell.SendKeys('EnterPinHere~')
    }
    else
    {
    }
    Sleep 3
    $MyCount++
}While(1)
