option explicit

private hHook as long

' Win API {

type MouseInput        ' derived from WinAPI structure INPUT, which is a union {
   type_       as long ' 0 = mouse. (1 = keyboard, 2 = hardware, but requires different definition)

   dx          as long ' Relative or absolute
   dy          as long ' position, dependent on flags

   mouseData   as long ' Dependent on dwFlags
   dwFlags     as long ' we use this (see MOUSEEVENT flags below)
   time_       as long
   dwExtraInfo as long
end type ' }

private type RECT ' {
   Left   as long
   Top    as long
   Right  as long
   Bottom as long
end type ' }

' SetWindowsHookEx / UnhookWindowsHookEx {

private declare function SetWindowsHookEx _
    lib   "user32"                        _
    alias "SetWindowsHookExA"             _
    (byVal idHook     as long,            _
     byVal lpfn       as long,            _
     byVal hmod       as long,            _
     byVal dwThreadId as long) as long

private declare function UnhookWindowsHookEx lib "user32" _
    (byVal hHook as long) as long

' }

' GetWindowText / GetWindowTextLength {
private declare function GetWindowText       _
   lib   "user32"                            _
   alias "GetWindowTextA" (                  _
   byVal hwnd     as long  ,                 _
   byVal lpString as string,                 _
   byVal cch      as long                    _
) as long

private declare function GetWindowTextLength _
   lib   "user32"                            _
   alias "GetWindowTextLengthA"  (           _
   byVal hWnd    as long                     _
) as long

' }

' { GetWindowRect
private declare function GetWindowRect       _
   lib "user32.dll"                        ( _
   byVal hwnd   as long                    , _
   byRef lpRect as RECT                      _
) as long

' }

' { SendInput
private declare function SendInput _
   lib "User32.dll"                _
  (byVal nCommands as long       , _
         iCommand  as MouseInput , _
   byVal cSize     as long) as long

' }

' { SetCursorPos
private declare function SetCursorPos lib "user32" ( _
   byVal x as long,                                  _
   byVal y as long                                   _
) as long

' }

private declare function GetCurrentThreadId lib "kernel32" () as long

private declare sub      Sleep              lib "kernel32" (byVal dwMilliseconds as long)

private const WH_CBT        = 5
private const HCBT_ACTIVATE = 5

private const MOUSEEVENTF_LEFTDOWN   = &H2
private const MOUSEEVENTF_LEFTUP     = &H4
private const MOUSEEVENTF_MOVE       = &H1
private const MOUSEEVENTF_RIGHTDOWN  = &H8
private const MOUSEEVENTF_RIGHTUP    = &H10
private const MOUSEEVENTF_MIDDLEDOWN = &H20
private const MOUSEEVENTF_MIDDLEUP   = &H40
private const MOUSEEVENTF_ABSOLUTE   = &H8000
private const MOUSEEVENTF_WHEEL      = &H800

' }

' getWindowText_ {
function getWindowText_(hWnd as long) as string ' {
 '
 '  getWindowText_() is a convenience function to wrap
 '  the call to the WinAPI function GetWindowText()
 '
    dim lenCopied   as long
    dim len_        as long

    len_           = GetWindowTextLength(hWnd)
    getWindowText_ = space(len_)
    lenCopied      = GetWindowText(hWnd, getWindowText_, len_+1)

end function ' }

' }

private sub mouseClick(x as long, y as long) ' {
 '
 '   Simulate a mouse click on coordinates x, y
 '

Sleep 500
    if SetCursorPos(x, y) = 0 then ' { Internal Use Only
        writeLog  "    SetCursorPos failed"
    end if ' }

    dim cmd as MouseInput

    cmd.type_       = 0 ' 0 = INPUT_MOUSE
    cmd.dx          = 0
    cmd.dy          = 0
  ' cmd.mouseData   =   ' not used
    cmd.time_       = 0 ' Let System provide timestamp
  ' cmd.dwExtraInfo =   ' not used

    cmd.dwFlags = MOUSEEVENTF_LEFTDOWN
Sleep 500
    SendInput 1&, cmd, len(cmd)

Sleep 500
    cmd.dwFlags = MOUSEEVENTF_LEFTUP
    SendInput 1&, cmd, len(cmd)

end sub ' }

public sub startHook() ' {
    writeLog "started"

    hHook = SetWindowsHookEx(WH_CBT                  , _
                             addressOf cbtHook       , _
                             0                       , _
                             GetCurrentThreadId)

    writeLog "hHook = " & hHook

end sub ' }

public sub stopHook() ' {

    if hHook = 0 then
       exit sub
    end if

    UnhookWindowsHookEx hHook
    hHook = 0
    debug.print "unhooked"

end sub ' }

' cbtHook {
private function cbtHook(byVal lMsg   as long, _
                         byVal wParam as long, _
                         byVal lParam as long) as long


    writeLog "cbtHook       "

    if lMsg = HCBT_ACTIVATE then
       writeLog "  HCBT_ACTIVATE"

       dim winTxt as string

       winTxt = getWindowText_(wParam)
       writeLog "  winTxt  " & winTxt

       if winTxt = "Microsoft Azure Information Protection" then ' {

          writeLog "    MAIP"

          dim r as rect
          GetWindowRect wParam, r

          writeLog "    r.left = " & r.left

          mouseClick r.left+200, r.top+ 75
          mouseClick r.left+420, r.top+120

          writeLog "    Cursor set"

       end if ' }

    end if

    writeLog "exit cbtHook       "

end function ' }

sub writeLog(txt as string) ' {

   dim f as integer : f = freeFile()
   open environ("temp") & "\cbtHook.log" for append as #f
   print# f, txt
   close  f

end sub ' }
