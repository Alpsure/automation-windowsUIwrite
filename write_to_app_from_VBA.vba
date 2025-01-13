' Tested on 13 Jan 2025 with Notepad
' Adjust windowName and text_to_write as needed

Option Explicit

Private Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const BM_CLICK = &HF5&
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Declare PtrSafe Function GetWindow Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal wCmd As Long) _
    As LongPtr

Private Declare PtrSafe Function SendMessageByString Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As LongPtr, _
    ByVal wMsg As Long, _
    ByVal wParam As LongPtr, _
    ByVal lParam As String) _
    As Long

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As LongPtr, _
    ByVal wMsg As Long, _
    ByVal wParam As LongPtr, _
    lParam As Any) _
    As Long

Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) _
    As Long

Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) _
    As LongPtr

Sub SaveAs_PopUp()
    Dim appHwnd As LongPtr
    Dim childHwnd As LongPtr
    Dim sChildText As String
    Dim sChildClass As String
    Dim text_to_write As String
    Dim bNext As Boolean

    ' Define the application window name
    Dim windowName As String
    windowName = "Untitled - Notepad"

    ' Find the application window
    appHwnd = FindWindow(vbNullString, windowName)
    If appHwnd = 0 Then
        MsgBox "Window not found: " & windowName, vbCritical
        Exit Sub
    End If

    ' Initialize variables
    childHwnd = GetWindow(appHwnd, GW_CHILD)
    text_to_write = "Mr Excel is working."
    bNext = False

    ' Loop through child windows
    Do While childHwnd <> 0
        sChildText = WindowText(childHwnd)
        sChildClass = Get_Class(childHwnd)
        Debug.Print sChildText
        Debug.Print sChildClass
        Debug.Print "============="

        ' Look for the Edit field and Save button
        If sChildClass = "Edit" Then
            bNext = False
            Call SendMessageByString(childHwnd, WM_SETTEXT, 0, text_to_write)
        End If

        ' Move to the next sibling window
        childHwnd = GetWindow(childHwnd, GW_HWNDNEXT)
    Loop

    MsgBox "Save As dialog not found!", vbExclamation
End Sub

Public Function WindowText(hwnd As LongPtr) As String
    ' Return the text associated with the window.
    Dim iTxtlen As Long
    Dim sTxt As String

    WindowText = ""
    If hwnd = 0 Then Exit Function

    iTxtlen = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
    If iTxtlen = 0 Then Exit Function

    iTxtlen = iTxtlen + 1
    sTxt = Space$(iTxtlen)
    iTxtlen = SendMessage(hwnd, WM_GETTEXT, iTxtlen, ByVal sTxt)
    WindowText = Left$(sTxt, iTxtlen)
End Function

Public Function Get_Class(hwnd As LongPtr) As String
    Dim sBuffer As String
    Dim lBuffer As Long

    lBuffer = 256
    sBuffer = Space$(lBuffer - 1)
    lBuffer = GetClassName(hwnd, sBuffer, lBuffer)
    Get_Class = Left$(sBuffer, lBuffer)
End Function
