Option Explicit On
Option Compare Text

Imports System
Imports System.Windows.Forms

Public Module VbeOnKey

    '***************************************************************************
    '*
    '* PROJECT NAME:    VBEONKEY
    '* AUTHOR & DATE:   STEPHEN BULLEN, Office Automation Ltd.
    '*                  10 May 2000
    '*
    '*                  COPYRIGHT © 2000 BY Office Automation LTD
    '*
    '* CONTACT:         stephen@oaltd.co.uk
    '* WEB SITE:        http://www.oaltd.co.uk
    '*
    '* DESCRIPTION:     Provides functionality similar to Application.OnKey, for the VBE
    '*
    '* USAGE:           To use in other projects, copy the following components into your
    '*                  project, in their entirety:
    '*                     - modVBEOnKey
    '*                     - frmVBEOnKey
    '*
    '*                  You can then use the following lines to turn key trapping on and off:
    '*                     VBEOnKey "%X", "RunProcedureX"
    '*                     VBEOnKey "%X"
    '*
    '* THIS MODULE:     API functions to provide the shortcut key functionality.
    '*
    '* PROCEDURES:
    '*
    '*  VBEOnKey        Main entry point to register a key/procedure combination, same as Application.OnKey
    '* (HookKey)        Register a key combination with Windows
    '* (UnHookKey)      Unregister a key combination
    '*  UnHookAll       Called if the hook form goes out of scope, to remove all our hooks
    '* (HookWindow)     Create a Windows message hook on the frmVBEOnKey userform
    '* (UnhookWindow)   Release the windows message hook
    '* (WindowProc)     The window message callback function, processes Windows messages
    '* (TimerCallback)  The windows timer callback function, to check between switching apps
    '* (LoWord)         Get the integer portion of a Word
    '* (GetHookInfo)    Convert the stored hook info string into a UDT
    '* (GetKeyCode)     Convert an 'OnKey' keycode to a vbKey number
    '*
    '***************************************************************************
    '*
    '* CHANGE HISTORY
    '*
    '*  DATE        NAME                DESCRIPTION
    '*  10/05/2000  Stephen Bullen      Initial version
    '*  11/26/2015  Mathieu Guindon     Slight modifications to build in VB.NET
    '*
    '***************************************************************************

    'Stuff for registering hot-keys
    Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
    Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
    Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Long
    Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Long) As Long

    'Stuff for handling windows
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long

    'Stuff for the tmer to check for switching between the VBE and the host
    Dim miTimerID As Long
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

    'Used by the Windows callback
    Private Const GWL_WNDPROC As Long = (-4)
    Private Const WM_HOTKEY As Long = &H312
    Private Const WM_ACTIVATEAPP As Long = &H1C
    Private Const WA_INACTIVE As Long = 0
    Private Const WA_ACTIVE As Long = 1

    'Used by the RegisterHotKey API
    Private Const MOD_ALT = &H1
    Private Const MOD_CONTROL = &H2
    Private Const MOD_SHIFT = &H4

    Dim miOldWndProc As Long
    Dim mhWndForm As Long
    Dim mhWndVBE As Long

    'Collection to store our key hook information
    Dim moKeys As New Collection
    Dim mbRegistered As Boolean

    'UDT to describe our key hooks
    Structure HookInfo
        public HookID As Long
        public KeyCode As Long
        public Shift As Long
        public Proc As String
    End Structure

    '
    '  Main entry point to register a key/procedure combination, same as Application.OnKey
    '  e.g.  VBEOnKey "%P", "MyProcedureForAltP"
    '
    '  See OnKey or SendKeys in online help for valid key codes
    '
    Public Sub VBEOnKey(ByVal Key As String, Optional Procedure As Object = Nothing)

        Dim lShift As Long, i As Integer, lKey As Long

        'Work out if Ctrl/Alt/Shift included
        For i = 1 To 3
            Select Case Left$(Key, 1)
            Case "+": lShift = lShift Or MOD_SHIFT
            Case "%": lShift = lShift Or MOD_ALT
            Case "^": lShift = lShift Or MOD_CONTROL
            Case Else: Exit For
            End Select
            Key = Mid$(Key, 2)
        Next

        lKey = GetKeyCode(Key)

        If lKey > 0 Then
            If Procedure Is Nothing Then
                'Unhook the key combination
                UnHookKey(lKey, lShift)
            Else
                'Hook the key combination
                HookKey(lKey, lShift, CStr(Procedure))
            End If
        End If

        'Store away the VBE's hWnd for future comparisons
        mhWndVBE = poVBE.MainWindow.hWnd

    End Sub

    '
    ' Register a key combination with Windows
    '
    Private Sub HookKey(lKeyCode As Long, lShift As Long, sProc As String)

        Dim hHookID As Long, sKey As String

        On Error Resume Next

        'Unhook it if already hooked
        UnHookKey(lKeyCode, lShift)

        'Hook the userform if we haven't already
        If moKeys.Count = 0 Then HookWindow()

        'Get a unique ID for this hook
        hHookID = GlobalAddAtom(CStr(Now()) & Rnd)

        'Register the hot-key with windows
        RegisterHotKey(mhWndForm, hHookID, lShift, lKeyCode)

        'Store all the info in a string and add it to our collection
        sKey = hHookID & "," & lKeyCode & "," & lShift & "," & sProc
        moKeys.Add(sKey, CStr(hHookID))

        mbRegistered = True

    End Sub

    '
    ' Unregister a key combination
    '
    Private Sub UnHookKey(lKeyCode As Long, lShift As Long)

        Dim iKey As Integer, sKeyItem As Object, uHook As HookInfo

        On Error Resume Next

        'Loop through our key combinations
        iKey = 0
        For Each sKeyItem In moKeys
            iKey = iKey + 1

            'Is it the one we mean?
            uHook = GetHookInfo(sKeyItem)
            If uHook.KeyCode = lKeyCode And uHook.Shift = lShift Then

                'Yes, so unregister it and tidy up
                UnregisterHotKey mhWndForm, uHook.HookID
                GlobalDeleteAtom uHook.HookID
                moKeys.Remove iKey

                'If no hooks left, unhook the form
                If moKeys.Count = 0 Then UnhookWindow
                Exit For
            End If
        Next

    End Sub

    '
    ' Called if the hook form goes out of scope, to remove all our hooks
    '
    Public Sub UnHookAll()

        Dim sKeyItem As Variant, hHookID As Long

        On Error Resume Next

        'Loop through all our keys, unregistering them and tidying up
        For Each sKeyItem In moKeys
            hHookID = GetHookInfo(sKeyItem).HookID
            UnregisterHotKey mhWndForm, hHookID
            GlobalDeleteAtom hHookID
        Next

        'Clear out our collection and unhook the window
        Set moKeys = Nothing
        mbRegistered = False
        UnhookWindow
    End Sub

    '
    ' Create a Windows message hook on the frmVBEOnKey userform
    '
    Private Sub HookWindow()

        On Error Resume Next

        'Establish a hook
        mhWndForm = frmVBEOnKey.hWnd
        miOldWndProc = SetWindowLong(mhWndForm, GWL_WNDPROC, AddressOf WindowProc)

        'Establish a timer proc to check for switching between apps
        miTimerID = GlobalAddAtom(CStr(Now()) & Rnd)
        SetTimer(mhWndForm, miTimerID, 500, AddressOf TimerCallback)

    End Sub

    '
    ' Release the windows message hook
    '
    Private Sub UnhookWindow()

        On Error Resume Next

        'Reset the message handler
        SetWindowLong(mhWndForm, GWL_WNDPROC, miOldWndProc)
        mhWndForm = 0
        mbRegistered = False

        'Kill the timer
        KillTimer(mhWndForm, miTimerID)
        GlobalDeleteAtom(miTimerID)
        miTimerID = 0

        Unload(frmVBEOnKey)

    End Sub

    '
    ' The window message callback function, processes Windows messages
    '
    Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

        Dim sKeyItem As Object, uHook As HookInfo
        Dim bProcessed As Boolean

        On Error Resume Next

        bProcessed = False

        'Does it concern our form?
        If hWnd = mhWndForm Then

            'Yes, so which message is it?
            Select Case uMsg
            Case WM_HOTKEY

                'It's a hot-key!, so see if the VBE is the active window
                If GetWindowThread(GetForegroundWindow) = GetWindowThread(mhWndVBE) Then

                    'Run the required procedure
                    Select Case GetHookInfo(moKeys(CStr(wParam))).Proc
                    Case "Indent2KProc"
                        IndentProcedure()

                    Case "Indent2KMod"
                        IndentModule()
                    End Select

                    'We handled it!
                    bProcessed = True
                End If

            Case WM_ACTIVATEAPP

                Select Case LoWord(wParam)
                Case WA_INACTIVE
                    'App lost focus, so unregister keys
                    For Each sKeyItem In moKeys
                        UnregisterHotKey(mhWndForm, GetHookInfo(sKeyItem).HookID)
                    Next

                    mbRegistered = False

                Case WA_ACTIVE
                    'App has focus, so reregister keys
                    For Each sKeyItem In moKeys
                        uHook = GetHookInfo(sKeyItem)
                        RegisterHotKey(mhWndForm, uHook.HookID, uHook.Shift, uHook.KeyCode)
                    Next

                    mbRegistered = True
                End Select
            Case Else
                'Ignore it
            End Select
        End If

        'Pass message to the original window message handler if we haven't handled it already
        If Not bProcessed Then WindowProc = CallWindowProc(miOldWndProc, hWnd, uMsg, wParam, lParam)

    End Function

    Private Function GetWindowThread(ByVal hWnd As Long) As Long
        Dim hThread As Long

        GetWindowThreadProcessId(hWnd, hThread)
        GetWindowThread = hThread

    End Function

    '
    '   Called by the Windows Timer routine, checks if the user switched to/from the VBE
    '
    Private Sub TimerCallback(ByVal hWnd As Long, ByVal lngMsg As Long, ByVal lngID As Long, ByVal lngTime As Long)

        Dim sKeyItem As Object, uHook As HookInfo

        'Check if the VBE is still in the foreground
        If GetForegroundWindow = mhWndVBE Then
            If Not mbRegistered Then
                'App got focus, so reregister keys
                For Each sKeyItem In moKeys
                    uHook = GetHookInfo(sKeyItem)
                    RegisterHotKey(mhWndForm, uHook.HookID, uHook.Shift, uHook.KeyCode)
                Next

                mbRegistered = True
            End If
        Else
            If mbRegistered Then
                'App lost focus, so unregister keys
                For Each sKeyItem In moKeys
                    UnregisterHotKey(mhWndForm, GetHookInfo(sKeyItem).HookID)
                Next

                mbRegistered = False
            End If
        End If

    End Sub

    '
    ' Get the Integer portion of a Word
    '
    Private Function LoWord(dw As Long) As Integer

        On Error Resume Next

        If dw And &H8000& Then
            LoWord = &H8000 Or (dw And &H7FFF&)
        Else
            LoWord = dw And &HFFFF&
        End If

    End Function

    '
    ' Convert the stored hook info string into a UDT
    '
    Private Function GetHookInfo(ByVal sString As String) As HookInfo

        Dim i As Integer

        On Error Resume Next

        i = InStr(1, sString, ",")
        GetHookInfo.HookID = CLng(Left$(sString, i - 1))

        sString = Mid$(sString, i + 1)
        i = InStr(1, sString, ",")
        GetHookInfo.KeyCode = CLng(Left$(sString, i - 1))

        sString = Mid$(sString, i + 1)
        i = InStr(1, sString, ",")
        GetHookInfo.Shift = CLng(Left$(sString, i - 1))

        GetHookInfo.Proc = Mid$(sString, i + 1)

    End Function

    '
    ' Convert an 'OnKey' keycode to a vbKey number
    '
    Function GetKeyCode(sOnKey As String) As Long

        Select Case Left$(sOnKey, 1)
        Case "{"
            Select Case sOnKey
            Case "{BACKSPACE}", "{BS}", "{BKSP}": Return Keys.Back
            Case "{CAPSLOCK}": Return Keys.CapsLock
            Case "{DELETE}", "{DEL}": Return Keys.Delete
            Case "{DOWN}": Return Keys.Down
            Case "{END}": Return Keys.End
            Case "{ENTER}", "{RETURN}": Return Keys.Return
            Case "{ESC}": Return Keys.Escape
            Case "{HELP}": Return Keys.Help
            Case "{HOME}": Return Keys.Home
            Case "{INSERT}", "{INS}": Return Keys.Insert
            Case "{LEFT}": Return Keys.Left
            Case "{NUMLOCK}": Return Keys.NumLock
            Case "{PGDN}": Return Keys.PageDown
            Case "{PGUP}": Return Keys.PageUp
            Case "{PRTSC}": Return Keys.PrintScreen
            Case "{RIGHT}": Return Keys.Right
            Case "{TAB}": Return Keys.Tab
            Case "{UP}": Return Keys.Up
            Case "{F1}": Return Keys.F1
            Case "{F2}": Return Keys.F2
            Case "{F3}": Return Keys.F3
            Case "{F4}": Return Keys.F4
            Case "{F5}": Return Keys.F5
            Case "{F6}": Return Keys.F6
            Case "{F7}": Return Keys.F7
            Case "{F8}": Return Keys.F8
            Case "{F9}": Return Keys.F9
            Case "{F10}": Return Keys.F10
            Case "{F11}": Return Keys.F11
            Case "{F12}": Return Keys.F12
            Case "{F13}": Return Keys.F13
            Case "{F14}": Return Keys.F14
            Case "{F15}": Return Keys.F15
            Case "{F16}": Return Keys.F16
            End Select
        Case "~": Return Keys.Return
        Case Else
            If sOnKey <> String.Empty Then Return Asc(Left$(sOnKey, 1))
        End Select

        Return Keys.None

    End Function


End Module
