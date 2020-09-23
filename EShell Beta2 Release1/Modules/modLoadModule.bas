Attribute VB_Name = "modLoadModule"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim mExe(255) As String
Dim mScript(255) As String
Public mName(255) As String
Public mCount As Integer
Public failLoad(255) As Boolean

Public Function LoadModuleConfig()

Dim ff As Integer
Dim l As String
Dim d(2) As String
Dim p As Integer

ff = FreeFile

Open App.path & "\module.cfg" For Input As #ff

Do Until EOF(ff)

    Line Input #ff, l

    If Left(l, 1) <> "#" And Trim(l) <> "" Then

        p = InStr(1, l, ",")
        d(0) = Left(l, p - 1)
        l = Right(l, Len(l) - p)
        
        p = InStr(1, l, ",")
        d(1) = Left(l, p - 1)
        l = Right(l, Len(l) - p)
        
        d(2) = l
        
        mExe(mCount) = App.path & "\eshellmodules\" & d(0)
        mScript(mCount) = App.path & "\eshellmodules\" & d(1)
        mName(mCount) = d(2)
        
        mCount = mCount + 1
        
    End If

Loop

Close #ff

End Function

Public Function LoadModule(num As Integer) As Boolean

On Error GoTo fail

ShellFile mExe(num)
'ExecScript mScript(num)

LoadModule = True

Exit Function

fail:
LoadModule = False

End Function

Public Function FindPort(name As String) As Long

Dim i As Long

For i = 0 To modLoadModule.mCount - 1

    If LCase(modLoadModule.mName(i)) = LCase(name) Then GoTo foundport

Next i

FindPort = 0
Exit Function

foundport:

FindPort = 1200 + i

End Function
