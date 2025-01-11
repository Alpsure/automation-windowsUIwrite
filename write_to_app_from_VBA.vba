
Option Explicit

'General notes:
' - Looping through children works with the same principle as the "sVar = dir()" method in looping through files in a folder
' - API's like GetWindow and SendMessageByString act differently based on the (constant) variables used
'   * SendMessageByString with WM_GETTEXT will read the field
'   * SendMessageByString with WM_SETTEXT will write to the field
'   * SendMessageByString with BM_CLICK will click the field (preferably a button)

'Process:
' - Find "Edit" field to write in the directory
' - Find the correct Open Button and click

Private Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const BM_CLICK = &HF5&
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

#If Win64 Then 'depending on 64 or 32 bit windows
    Private Declare PtrSafe Function GetWindow Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal wCmd As Long) _
        As Long
        
    Private Declare PtrSafe Function SendMessageByString Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As String) _
        As Long
        
    Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) _
        As Long
#Else
    Private Declare Function GetWindow Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal wCmd As Long) _
        As Long
        
    Private Declare PtrSafe Function SendMessageByString Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As String) _
        As Long
        
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) _
        As Long
        
    Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
        ByVal hwnd As Long, _
        ByVal lpClassName As String, _
        ByVal nMaxCount As Long) _
        As Long
#End If

Declare Function GetDesktopWindow Lib "user32" () As Long

Sub SaveAs_PopUp()

Dim lMain As Long
Dim lChild1 As Long
Dim lChild2 As Long
Dim sChild1 As String
Dim sChild2 As String
Dim sClass As String
Dim bNext As Boolean
Dim sFileName As String

lMain = GetDesktopWindow
lChild1 = GetWindow(lMain, GW_CHILD)
bNext = False
sFileName = "Myfile.pdf"

Do While lChild1 <> 0
    sChild1 = WindowText(lChild1)
    
    If sChild1 = "Save As" Then 'if we find the Save As window...
        lChild2 = GetWindow(lChild1, GW_CHILD)
        
        Do While lChild2 <> 0 'Loop through the children of the Save As window
            sChild2 = WindowText(lChild2)
            sClass = Get_Class(lChild2)
            
            If bNext = True And sClass = "Edit" Then
                bNext = False
                Call SendMessageByString(lChild2, WM_SETTEXT, 0, sFileName)
            End If
            If sChild2 = "File &name:" Then
                bNext = True
            End If
                        
            If sChild2 = "&Save" Then
                Call SendMessageByString(lChild2, BM_CLICK, 0, ByVal 0&)
                End
            End If
            
            lChild2 = GetWindow(lChild2, GW_HWNDNEXT)
        Loop
    End If
    
    lChild1 = GetWindow(lChild1, GW_HWNDNEXT)
Loop

End Sub
Public Function WindowText(window_hwnd As Long) As String

'Return the text associated with the window.

Dim iTxtlen As Long
Dim sTxt As String

WindowText = ""
If window_hwnd = 0 Then Exit Function

iTxtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
If iTxtlen = 0 Then Exit Function

iTxtlen = iTxtlen + 1
sTxt = Space$(iTxtlen)
iTxtlen = SendMessage(window_hwnd, WM_GETTEXT, iTxtlen, ByVal sTxt)
WindowText = Left$(sTxt, iTxtlen)
    
End Function
Public Function Get_Class(hwnd As Long) As String

Dim sBuffer As String
Dim lBuffer As Long

lBuffer = 256
sBuffer = Space$(lBuffer - 1)
lBuffer = GetClassName(hwnd, sBuffer, lBuffer)
Get_Class = Left$(sBuffer, lBuffer)

End Function


