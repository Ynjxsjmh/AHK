
;; Keyboard map
;; LControl LAlt Space

; https://superuser.com/questions/726988/how-to-remap-a-program-to-lock-windows-winl
; WARNING: Programs that use User32\LockWorkStation (i.e. programmatically locking the operating system) may not work correctly! 
; This includes Windows itself (i.e. using start menu or task manager to lock will also not work).
; Script changes Win-L to show a msgbox and Ctrl-Alt-L to lock windows

; The following 3 code lines are auto-executed upon script run, the return line marks an end to the auto-executed code section.
; Register user defined subroutine 'OnExitSub' to be executed when this script is terminating
OnExit, OnExitSub

; Disable LockWorkStation, so Windows doesn't intercept Win+L and this script can act on that key combination 
SetDisableLockWorkstationRegKeyValue( 1 )
return

LControl::LWin
LWin::LControl

OnExitSub:
    ; Enable LockWorkStation, because this script is ending (so other applications aren't further disturbed)
    SetDisableLockWorkstationRegKeyValue( 0 )
    ExitApp
    return

SetDisableLockWorkstationRegKeyValue( value )
{
    ;; https://www.autohotkey.com/docs/commands/RegWrite.htm
    ;; It has two syntaxes, the new syntax won't create the nonexisting item, so you should create it yourself and reboot.
    ;; https://superuser.com/questions/180541/is-it-possible-to-apply-a-registry-change-without-rebooting
    ;; Changing registry values does NOT require a reboot, they are "applied" immediately.
    RegWrite, REG_DWORD, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System, DisableLockWorkstation, %value%
} 
