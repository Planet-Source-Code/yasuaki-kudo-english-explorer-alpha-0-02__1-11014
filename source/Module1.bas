Attribute VB_Name = "Module1"

Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long


Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long


Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long


Declare Function GetCurrentThreadId Lib "kernel32" () As Long


Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Public Const WM_RBUTTONUP = &H205
    Public Const WH_MOUSE = 7


Type POINTAPI
    x As Long
    y As Long
    End Type


Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
    End Type
    Public l_hMouseHook As Long
                            
 



                        
'**************************************
' Name: Disable Right Mouse click
' Description:Disable Right Mouse click
'     in the web browser control.

' By: Newsgroup Posting
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None
'
'Warranty:
'code provided by Planet Source Code(tm)
'     (http://www.Planet-Source-Code.com) 'as
'     is', without warranties as to performanc
'     e, fitness, merchantability,and any othe
'     r warranty (whether expressed or implied
'     ).
'Terms of Agreement:
'By using this source code, you agree to
'     the following terms...
' 1) You may use this source code in per
'     sonal projects and may compile it into a
'     n .exe/.dll/.ocx and distribute it in bi
'     nary format freely and with no charge.
' 2) You MAY NOT redistribute this sourc
'     e code (for example to a web site) witho
'     ut written permission from the original
'     author.Failure to do so is a violation o
'     f copyright laws.
' 3) You may link to this code from anot
'     her website, provided it is not wrapped
'     in a frame.
' 4) The author of this code may have re
'     tained certain additional copyright righ
'     ts.If so, this is indicated in the autho
'     r's description.
'**************************************



Public Function MouseHookProc(ByVal nCode As Long, ByVal wParam As Long, mhs As MOUSEHOOKSTRUCT) As Long
    'Prevent Right-Mouse Clicks in WebBrowse
    '     r Control:


    If (nCode >= 0 And wParam = WM_RBUTTONUP) Then
        Dim sClassName As String
        Dim sTestClass As String
        sTestClass = "HTML_Internet Explorer"
        sClassName = String$(256, 0)


        If GetClassName(mhs.hwnd, sClassName, Len(sClassName)) > 0 Then


            If Left$(sClassName, Len(sTestClass)) = sTestClass Then
                MouseHookProc = 1
                Exit Function
            End If
        End If
    End If
    MouseHookProc = CallNextHookEx(l_hMouseHook, nCode, wParam, mhs)
End Function


Public Sub BeginRightMouseTrap()
    'Start Trapping Right-Mouse clicks in We
    '     bBrowser Control:
    l_hMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseHookProc, App.hInstance, GetCurrentThreadId)
End Sub


Public Sub EndRightMouseTrap()
    'End Trapping Right-Mouse clicks in WebB
    '     rowser Control:
    UnhookWindowsHookEx l_hMouseHook
End Sub

 
 

