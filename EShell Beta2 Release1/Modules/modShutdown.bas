Attribute VB_Name = "modShutdown"
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public Enum ShutdownModes
    SM_Shutdown
    SM_Reboot
    SM_Logoff
End Enum

Public Sub Shutdown(mode As ShutdownModes)
    Select Case mode
        Case SM_Shutdown
            ExitWindowsEx EWX_SHUTDOWN, 0
        Case SM_Reboot
            ExitWindowsEx EWX_REBOOT, 0
        Case SM_Logoff
            ExitWindowsEx EWX_LOGOFF, 0
    End Select
End Sub
